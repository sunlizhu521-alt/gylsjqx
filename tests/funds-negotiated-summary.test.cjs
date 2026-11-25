const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const HEADERS=['创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','是否谈判分批','谈判明细','是否结清','2026年11月30日','2027年1月4日','2026年12月付款','2027年1月付款'];
const row=()=>['李奇','汇总合成供应商','迈德斯特','不支付','十二月货款',100,'是','12月10日-40\n2027年1月20日-60','是',40,60,40,60];
function sample(rows=[row()],headers=HEADERS){
 const input=fixture(),name=input.negotiated.book.SheetNames[0];
 input.negotiated.book.Sheets[name]=X.utils.aoa_to_sheet([headers,...rows]);
 return input;
}
const sheet=i=>i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];
const summaryIssues=r=>r.issues.filter(i=>i.id.startsWith('summary-'));
const text=c=>c.t==='n'&&X.SSF.is_date(c.z||'')?X.SSF.format('yyyy-mm-dd',c.v):c.v;

test('新汇总固定九列，按谈判原行分类，保留不支付、已结清、重复、零和空白金额',()=>{
 const rows=[row(),row(),row(),row()];rows[1][6]='否';rows[2][5]=0;rows[2][7]='';rows[3][5]='';rows[3][7]='';
 const input=sample(rows),before=JSON.stringify(input),result=C.build(input);
 assert.deepEqual(result.negotiatedSummaryHeaders,['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','付款日期','申请纸质单日期']);
 assert.equal(result.negotiatedSummaryRows.length,4);
 assert.deepEqual(result.negotiatedSummaryRows[0].slice(0,7).map(c=>c.v),['谈判付款-谈判付款表','李奇','汇总合成供应商','迈德斯特','不支付','十二月货款',100]);
 assert.equal(result.negotiatedSummaryRows[1][0].v,'正常付款-谈判付款表');
 assert.equal(result.negotiatedSummaryRows[2][6].v,0);assert.equal(result.negotiatedSummaryRows[2][6].t,'n');
 assert.equal(result.negotiatedSummaryRows[3][6].v,'');assert.equal(result.negotiatedSummaryRows[0][6].z,'#,##0.00');
 assert.deepEqual(result.negotiatedSummaryRows[0].slice(7).map(text),['2026-12-10\n2027-01-20','2026-11-30\n2027-01-04']);
 assert.equal(JSON.stringify(input),before);
 assert(!result.negotiated.managedColumns.includes(6),'只读分类列不能成为原表受管理列');
});

test('跟单员别名独立识别，重复列留空并提示，缺失不从创建人或OA申请人推断',()=>{
 for(const name of ['创跟单员','采购跟单员','跟单员']){
  const headers=HEADERS.slice();headers[0]=name;
  assert.equal(C.build(sample([row()],headers)).negotiatedSummaryRows[0][1].v,'李奇');
 }
 const absent=sample();sheet(absent).A1.v='创建人';assert.equal(C.build(absent).negotiatedSummaryRows[0][1].v,'');
 const duplicate=sample();X.utils.sheet_add_aoa(sheet(duplicate),[['采购跟单员'],['徐英']],{origin:'N1'});
 const result=C.build(duplicate);assert.equal(result.negotiatedSummaryRows[0][1].v,'');assert(summaryIssues(result).some(i=>i.id.startsWith('summary-owner-header:')));
});

test('分类空白、异常、公式错误和缺列不当正常付款；确认保留空白仍保留所有行',()=>{
 for(const value of ['', '未知', {t:'e',v:15,w:'#VALUE!'}, {t:'n',f:'A99'}]){
  const input=sample();sheet(input).G2=typeof value==='object'?value:{t:'s',v:value};
  const result=C.build(input),issue=summaryIssues(result).find(i=>i.id.startsWith('summary-classification:'));
  assert(issue&&!issue.resolved);assert.equal(result.negotiatedSummaryRows[0][0].v,'');
  assert.equal(C.build(input,{[issue.id]:['keep']}).issues.find(i=>i.id===issue.id).resolved,true);
 }
 for(const duplicate of [false,true]){
  const input=sample();if(duplicate)X.utils.sheet_add_aoa(sheet(input),[['是否谈判分批'],['否']],{origin:'N1'});else sheet(input).G1.v='其他字段';
  const result=C.build(input);assert.equal(result.negotiatedSummaryRows[0][0].v,'');assert(summaryIssues(result).some(i=>i.id.startsWith('summary-classification-header:')));
 }
 assert.equal(summaryIssues(C.build(sample([]))).length,0);
});

test('单期使用真实日期，多期付款与申请日期分别按来源顺序保留，不将申请月当付款月',()=>{
 const one=row();one[7]='2026年12月10日-100';one[9]=100;one[10]=0;
 const result=C.build(sample([one])),dates=result.negotiatedSummaryRows[0].slice(7);
 assert.deepEqual(dates.map(text),['2026-12-10','2026-11-30']);
 for(const c of dates){assert.equal(c.t,'n');assert.equal(c.z,'yyyy-mm-dd');assert.equal(c.f,undefined);}
 const multiple=row();multiple[7]='2027年1月20日-60\n2026年12月10日-20\n2026年12月10日-20';
 assert.equal(C.build(sample([multiple])).negotiatedSummaryRows[0][7].v,'2027-01-20\n2026-12-10\n2026-12-10');
});

test('缺日、无效日期、缺少年份及跨年冲突提示精确来源，不默认补日或年份',()=>{
 for(const detail of ['2026年12月-100','2026年2月30日-100','3月10日-100']){
  const input=sample();sheet(input).H2.v=detail;const result=C.build(input);
  assert.equal(result.negotiatedSummaryRows[0][7].v,'');assert(summaryIssues(result).some(i=>i.id.startsWith('summary-dates:')&&i.message.includes('H2')));
 }
 const input=sample();X.utils.sheet_add_aoa(sheet(input),[['2027年12月付款'],[0]],{origin:'N1'});
 const result=C.build(input);assert.equal(result.negotiatedSummaryRows[0][7].t,'n');
 assert.equal(text(result.negotiatedSummaryRows[0][7]),'2027-01-20');assert(summaryIssues(result).some(i=>/缺少年份/.test(i.message)));
});

test('计划空白和零不取日期，异常金额及合并归属提示；使用原表头而非批次排序年份',()=>{
 const input=sample(),s=sheet(input);s.J2.v=0;s.K2.v='';
 assert.equal(C.build(input).negotiatedSummaryRows[0][8].v,'');
 s.J2={t:'e',v:15,w:'#VALUE!'};s.K2.v=60;s['!merges']=[{s:{r:1,c:10},e:{r:2,c:10}}];
 let result=C.build(input);assert.equal(result.negotiatedSummaryRows[0][8].v,'');assert(summaryIssues(result).some(i=>/J2/.test(i.message)&&/K2/.test(i.message)));
 const changed=sample(),name=changed.negotiated.book.SheetNames[0];sheet(changed).L1.v='2027年11月付款';
 result=C.build(changed,{['batch-year:'+name+'!J1']:['year:2027']});
 assert.equal(result.negotiatedSummaryRows[0][8].v,'2026-11-30\n2027-01-04');
 assert.equal(result.negotiated.batchYearAdjustments[0].year,2027);
});

test('公式只取缓存值，错误或缺缓存不当零；文本OA及前导零完整保留',()=>{
 const input=sample(),s=sheet(input);s.D2.v='000123';s.F2={t:'n',v:100,f:'SUM(J2:K2)'};s.H2.f='A1';
 let result=C.build(input),summary=result.negotiatedSummaryRows[0];assert.equal(summary[4].v,'000123');assert.equal(summary[4].t,'s');assert.equal(summary[6].v,100);
 for(const c of summary)assert.equal(c.f,undefined);
 s.F2={t:'n',f:'SUM(J2:K2)'};result=C.build(input);assert.equal(result.negotiatedSummaryRows[0][6].v,'公式结果不可用');assert(summaryIssues(result).some(i=>i.id.includes('summary-value:')));
});

test('重上传只保留一个新结果表，不将新表误作谈判主源',()=>{
 const names=['谈判付款','谈判付款付款表付款节奏','谈判付款（原表）','辅助说明'];
 assert.deepEqual(C.negotiatedSheets({SheetNames:names}),['谈判付款','谈判付款（原表）']);
 const plan=C.outputSheetPlan(names,'谈判付款');assert.equal(plan[1].name,'谈判付款付款表付款节奏');
 assert.equal(plan.filter(p=>p.name==='谈判付款付款表付款节奏').length,1);
});

test('1904来源日期表头转为正确日历日期，核心结果仍统一1900日期数值',()=>{
 const input=sample(),s=sheet(input),one=row();one[7]='2026年12月10日-100';one[9]=100;one[10]=0;
 X.utils.sheet_add_aoa(s,[one],{origin:'A2'});
 input.negotiated.book.Workbook={WBProps:{date1904:true}};
 for(const [address,stamp]of [['J1','2026-11-30'],['K1','2027-01-04']])s[address]={t:'n',v:C.parsePaymentDate(stamp).v-1462,z:'yyyy-mm-dd'};
 const result=C.build(input);assert.deepEqual(result.negotiatedSummaryRows[0].slice(7).map(text),['2026-12-10','2026-11-30']);
 assert.equal(result.negotiatedSummaryRows[0][8].v,C.parsePaymentDate('2026-11-30').v);
});

test('双层表头分类与跟单员读取原合并表头，保留数据来源及计划原日期',()=>{
 const input=sample(),original=sheet(input),next={'!merges':[]};
 for(const [address,c]of Object.entries(original))if(/^[A-Z]+\d+$/.test(address)){const p=X.utils.decode_cell(address);next[X.utils.encode_cell({r:p.r+1,c:p.c})]={...c};}
 for(let c=0;c<HEADERS.length;c++)if(![9,10].includes(c)){
  const top=X.utils.encode_cell({r:0,c}),lower=X.utils.encode_cell({r:1,c});next[top]=next[lower];delete next[lower];next['!merges'].push({s:{r:0,c},e:{r:1,c}});
 }
 next.J1={t:'s',v:'2026年12月付款'};next.K1={t:'s',v:'2027年1月付款'};next['!ref']='A1:M3';
 input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]]=next;
 const before=JSON.stringify(input),result=C.build(input);assert.equal(result.negotiated.headerTopRow,0);
 assert.deepEqual(result.negotiatedSummaryRows[0].slice(0,2).map(c=>c.v),['谈判付款-谈判付款表','李奇']);
 assert.deepEqual(result.negotiatedSummaryRows[0].slice(7).map(text),['2026-12-10\n2027-01-20','2026-11-30\n2027-01-04']);
 assert.equal(JSON.stringify(input),before);assert.equal(result.negotiatedSummaryRows.length,1);
});
