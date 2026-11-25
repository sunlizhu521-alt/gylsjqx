const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,set,ID,AH,book}=require('./funds-progress-fixtures.cjs');
const text=cell=>cell?.t==='n'&&X.SSF.is_date(cell.z||'')?X.SSF.format('yyyy-mm-dd',cell.v):cell?.v;
const source=input=>input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];
function sample({details=[['2026-11-12',100]],plans=[['2026-11-02',100]],applications=[{amount:100,paid:'是'}]}={}){
 const input=fixture(),total=details.reduce((sum,[,amount])=>sum+amount,0);
 const headers=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','已付金额','剩余付款金额','是否结清','应付款时间','备注','谈判明细','分批总金额','核对',...plans.map(([date])=>date),'是否谈判分批','跟单员'];
 const row=['合成供应商1','迈德斯特',ID(1),total,0,total,'否','2026-11-30','保留备注',details.map(([date,amount])=>date+' - '+amount).join('\n'),total,'一致',...plans.map(([,amount])=>amount),'是','李奇'];
 input.negotiated.book=book('十一月谈判付款主表',[headers,row]);
 input.applications.book=book('付款申请汇总',[[...AH,'是否已付款','核对说明','应付时间'],...applications.map((a,i)=>[a.date||'2026-11-'+String(i+1).padStart(2,'0'),'合成供应商1','迈德斯特',a.oa||ID(1),a.amount,'谈判付款',a.paid??'',a.note||'','2026-11-30'])]);
 for(const header of ['申请付款金额','剩余金额'])set(input,'ledger',header,2,total,'对账明细-李奇');
 for(const row of [2,6])set(input,'oa','付款金额',row,total);
 return input;
}
const rows=result=>result.negotiatedSummaryRows.map(row=>({amount:row[6]?.v,date:text(row[7]),paperDate:text(row[8]),paper:row[9]?.v,paid:row[10]?.v}));
const noDatesBlocked=result=>assert.equal(result.issues.some(issue=>issue.id.startsWith('summary-dates:')),false,'仅映射歧义不增加日期阻断提示');

test('日期按上周一计划唯一匹配，同额多期只使用本期当前申请，不广播状态',()=>{
 const input=sample({details:[['2026-11-12',50],['2026-11-19',50]],plans:[['2026-11-02',50],['2026-11-09',50]],applications:[{amount:50,paid:'是',note:'第1批'}]}),before=JSON.stringify(input),result=C.build(input);
 assert.equal(result.applicationBatchMatches['付款申请汇总!2'].unique,true);
 assert.deepEqual(rows(result),[
  {amount:50,date:'2026-11-12',paperDate:'2026-11-02',paper:'是',paid:'是'},
  {amount:50,date:'2026-11-19',paperDate:'2026-11-09',paper:'待核对',paid:'待核对'}
 ]);
 assert.equal(JSON.stringify(input),before);
});

test('同额申请没有唯一批次归属，即使原表被推测标记也不自动确认任何日期组状态',()=>{
 const input=sample({details:[['2026-11-12',50],['2026-11-19',50]],plans:[['2026-11-02',50],['2026-11-09',50]],applications:[{amount:50,paid:'是'}]}),result=C.build(input);
 assert.equal(result.applicationBatchMatches['付款申请汇总!2'].unique,false);
 assert.deepEqual(rows(result).map(row=>[row.paper,row.paid]),[['待核对','待核对'],['待核对','待核对']]);
 assert.deepEqual(rows(result).map(row=>row.paperDate),['2026-11-02','2026-11-09']);
});

test('相同付款日期合并金额，多个计划全部有申请且已付才显示是，部分或混合状态待核对',()=>{
 for(const [applications,paper,paid] of [
  [[{amount:40,paid:'是'},{amount:60,paid:'是'}],'是','是'],
  [[{amount:40,paid:'否'},{amount:60,paid:'否'}],'是','否'],
  [[{amount:40,paid:'是'},{amount:60,paid:'否'}],'是','待核对'],
  [[{amount:40,paid:'是'},{amount:60,paid:''}],'是','待核对'],
  [[{amount:40,paid:'是'}],'待核对','待核对']
 ]){
  const input=sample({details:[['2026-11-19',40],['2026-11-19',60]],plans:[['2026-11-02',40],['2026-11-09',60]],applications}),result=C.build(input);
  assert.deepEqual(rows(result),[{amount:100,date:'2026-11-19',paperDate:'2026-11-02\n2026-11-09',paper,paid}],JSON.stringify(applications));
 }
});

test('付款状态仅按唯一对应当前申请的是否及日期；空白错误冲突不借OA归档补判',()=>{
 for(const [paid,wanted] of [['是','是'],['2026-11-13','是'],[C.parsePaymentDate('2026-11-13').v,'是'],['否','否'],['','待核对'],['待确认','待核对'],[{t:'e',v:23,w:'#REF!'},'待核对'],[{t:'n',f:'A99'},'待核对']]){
  const result=C.build(sample({applications:[{amount:100,paid}]}));
  assert.deepEqual(rows(result).map(row=>[row.paper,row.paid]),[['是',wanted]],JSON.stringify(paid));
 }
 for(const [states,wanted] of [[['是','否'],'待核对'],[['是','待确认'],'待核对'],[['否','否'],'否'],[['是',''],'是'],[['',''],'待核对']]){
  const result=C.build(sample({applications:states.map(paid=>({amount:100,paid}))}));
  assert.deepEqual(rows(result).map(row=>[row.paper,row.paid]),[['是',wanted]],JSON.stringify(states));
 }
});

test('同日分组金额合计相等不把一笔整额申请拆作多个计划的申请或付款证据',()=>{
 const result=C.build(sample({details:[['2026-11-19',40],['2026-11-19',60]],plans:[['2026-11-02',40],['2026-11-09',60]],applications:[{amount:100,paid:'是'}]}));
 assert.deepEqual(result.applicationBatchMatches,{});
 assert.deepEqual(rows(result),[{amount:100,date:'2026-11-19',paperDate:'2026-11-02\n2026-11-09',paper:'待核对',paid:'待核对'}]);
});

test('重新申请仍证明纸质单已申请，其付款状态不参与确认，普通同批申请按本次状态判断',()=>{
 for(const [applications,wanted] of [
  [[{amount:100,paid:'是',note:'重新申请'}],'待核对'],
  [[{amount:100,paid:'是',note:'重新申请'},{amount:100,paid:'否'}],'否'],
  [[{amount:100,paid:'否',note:'重新申请'},{amount:100,paid:'是'}],'是']
 ]){
  const result=C.build(sample({applications}));
  assert.deepEqual(rows(result).map(row=>[row.paper,row.paid]),[['是',wanted]]);
 }
});

test('只有历史红色批注或OA归档，没有当前唯一申请时两状态待核对',()=>{
 const input=sample({applications:[]}),sheet=source(input);
 sheet.M2.s={patternType:'solid',fgColor:{rgb:'FF0000'}};
 sheet.M2.c=[{a:'人工',t:'已支付\n申请日期（付款节奏跟进）：2026年11月1日\n付款时间（付款节奏跟进）：2026年11月12日'}];
 const result=C.build(input);
 assert.deepEqual(rows(result),[{amount:100,date:'2026-11-12',paperDate:'2026-11-02',paper:'待核对',paid:'待核对'}]);
 assert.deepEqual(result.applicationBatchMatches,{});
});

test('跨年上周一唯一匹配及唯一金额回退保留原申请表头日期',()=>{
 for(const [paymentDate,planDate] of [['2027-01-07','2026-12-28'],['2026-11-12','2026-11-16']]){
  const result=C.build(sample({details:[[paymentDate,100]],plans:[[planDate,100]]}));
  assert.deepEqual(rows(result),[{amount:100,date:paymentDate,paperDate:planDate,paper:'是',paid:'是'}]);
 }
});

test('两个日期组共用唯一同额同周计划候选不能按处理先后分配，无映射日期留空且不新增阻断',()=>{
 const result=C.build(sample({details:[['2026-11-12',100],['2026-11-13',100]],plans:[['2026-11-02',100]],applications:[{amount:100,paid:'是'}]}));
 assert.equal(rows(result).length,2);
 assert.deepEqual(rows(result).map(row=>[row.paperDate,row.paper,row.paid]),[['','待核对','待核对'],['','待核对','待核对']]);
 noDatesBlocked(result);
});

test('同供应商相同付款日的不同原行独立输出，不合并金额也不共享申请归属',()=>{
 const input=sample(),sheet=source(input),original=X.utils.sheet_to_json(sheet,{header:1})[1];
 X.utils.sheet_add_aoa(sheet,[original],{origin:-1});
 const result=C.build(input);
 assert.deepEqual(rows(result).map(row=>[row.amount,row.date]),[[100,'2026-11-12'],[100,'2026-11-12']]);
 assert.deepEqual(rows(result).map(row=>[row.paper,row.paid]),[['待核对','待核对'],['待核对','待核对']]);
});
