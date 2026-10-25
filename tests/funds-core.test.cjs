const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const LH=['月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付款金额','是否结清','剩余付款金额','是否谈判付款','是否提交付款申请单'];
const NH=['供应商','付款主体','对公付款0A流程号','合同总金额','已付金额','摘要','申请付款金额','单据号【财务应付单】','创建人','是否结清','剩余付款金额','原备注','付款计划'];
const AH=['申请日期','供应商全称','付款主体','合同号','合同金额','已付金额','付款事由','备注','本次申请付款金额'];
const OH=['流程编号','当前节点','未操作者','流程状态','申请人','收款单位','付款金额'];
test('V19 batch columns ignore styled tails and preserve financial formulas, blank paid amounts and manual settlement',()=>{
 const input=fixture(),sheet=input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];
 sheet.N1={t:'s',v:'谈判明细'};sheet.O1={t:'s',v:'分批总金额'};sheet.P1={t:'s',v:'核对'};
 sheet.N2={t:'s',v:'1、10月30日－100.00\n2、11月30日－0.00'};sheet.O2={t:'n',v:100,f:'SUM(M2:M2)'};sheet.P2={t:'s',v:'一致',f:'IF(O2=G2,"一致","不一致")'};
 sheet.K2={t:'n',v:100,f:'G2-E2'};sheet.J2={t:'s',v:'是'};delete sheet.E2;
 sheet.XFD1={t:'z',v:undefined};sheet['!ref']='A1:XFD1048576';
 const original=JSON.stringify(sheet),result=C.build(input),n=result.negotiated;
 assert.equal(n.batchLayout,true);assert.equal(n.sourceRange.e.c,15);assert.equal(n.sourceRange.e.r,1);
 for(const h of ['提交付款申请单日期','提交时间汇总','付款自然周'])assert(!n.headers.includes(h));
 assert.equal(n.rows[0][10].f,'G2-E2');assert.equal(n.rows[0][9].v,'是');assert.equal(n.rows[0][4].v,'');
 assert.equal(n.rows[0][14].f,'SUM(M2:M2)');assert.equal(n.rows[0][15].f,'IF(O2=G2,"一致","不一致")');
 assert(!n.updates.some(u=>/^(E|G|J|K|M|N|O|P)2$/.test(u.address)));assert.equal(JSON.stringify(sheet),original);
 sheet.XFD1000={t:'n',v:0};assert.throws(()=>C.build(input),/数据范围过大/,'a real far-right value must not be silently dropped');
});
function book(name,headers,rows){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows.map(r=>headers.map(h=>r[h]??''))]),name);return b;}
function fixture(){
 const main={'月份':'2026-10-01','供应商名称':'测试供应商甲','付款主体':'迈德斯特','对公付款0A流程号':'MDST-DGFK20261001001','开票时间':'2026-10-01','开票金额':100,'摘要':'九月货款','申请付款金额':100,'应付款时间':'2026-10-31','单据号【财务应付单】':'AP000001','OA合同编号':'MDST-HT20260101','合同总金额':1000,'已付款金额':20,'是否结清':'否','剩余付款金额':80,'是否谈判付款':'是','是否提交付款申请单':''};
 const normal={...main,'供应商名称':'测试供应商乙','对公付款0A流程号':'MDST-DGFK20261001002','单据号【财务应付单】':'AP000002','OA合同编号':'MDST-HT20260102','是否谈判付款':'否','已付款金额':100,'是否结清':'是','剩余付款金额':0};
 const ledger=book('对账明细-李奇',LH,[main,normal]);for(const person of C.PEOPLE.slice(1))X.utils.book_append_sheet(ledger,X.utils.aoa_to_sheet([LH]),'对账明细-'+person);
 const n={'供应商':'测试供应商甲','付款主体':'迈德斯特','对公付款0A流程号':main['对公付款0A流程号'],'合同总金额':1000,'已付金额':0,'摘要':'九月货款','申请付款金额':100,'单据号【财务应付单】':'AP000001','创建人':'李奇','是否结清':'否','剩余付款金额':100,'原备注':'保留备注','付款计划':100};
 const negotiated=book('10月谈判付款主表',NH,[n]);negotiated.Sheets[negotiated.SheetNames[0]].M2={t:'n',v:100,f:'G2'};
 const appRows=[main,normal].map(r=>({'申请日期':'2026-10-10','供应商全称':r['供应商名称'],'付款主体':r['付款主体'],'合同号':r['对公付款0A流程号'],'合同金额':r['合同总金额'],'已付金额':0,'付款事由':r['摘要'],'备注':r['单据号【财务应付单】'],'本次申请付款金额':r['申请付款金额']}));
 const oaRows=[main,normal].map(r=>({'流程编号':r['对公付款0A流程号'],'当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇','收款单位':r['供应商名称'],'付款金额':r['申请付款金额']}));
 return {ledger:{book:ledger,name:'资金.xlsx'},negotiated:{book:negotiated,name:'谈判.xlsx'},applications:{book:book('付款申请汇总',AH,appRows),name:'申请.xlsx'},oa:{book:book('Sheet0',OH,oaRows),name:'OA.xlsx'}};
}
function set(input,role,header,row,value,sheet){const b=input[role].book,s=b.Sheets[sheet||b.SheetNames[0]],h=X.utils.sheet_to_json(s,{header:1})[0],c=h.indexOf(header);assert(c>=0);s[X.utils.encode_cell({r:row-1,c})]=typeof value==='object'&&value!==null?value:{t:typeof value==='number'?'n':'s',v:value};}
const update=(r,address)=>r.negotiated.updates.find(u=>u.address===address)?.cell;
function resolveKeep(input){let d={};for(let i=0;i<4;i++){const r=C.build(input,d);if(!r.stats.pending)return r;for(const issue of r.issues)if(!issue.resolved)d[issue.id]=['keep'];}return C.build(input,d);}
test('正常付款提交付款申请单周按应付款时间提前一周计算，跨月跨年且不使用申请日期',()=>{
 const serial=(Date.UTC(2026,9,28)-Date.UTC(1899,11,30))/86400000;
 for(const [raw,expected,date1904] of [
  ['2026-10-26','2026/10/19周'],['2026-10-27','2026/10/19周'],['2026年10月28日','2026/10/19周'],['2026/11/1 23:59:59','2026/10/19周'],['2026-11-02','2026/10/26周'],
  ['2026-01-01','2025/12/22周'],['2026-12-31','2026/12/21周'],['2024-02-29','2024/02/19周'],[serial+.75,'2026/10/19周'],[serial-1462,'2026/10/19周',true],
  ['', '日期待确认'],['日期待定','日期待确认'],['2026-02-30','日期待确认'],[0,'日期待确认'],[{t:'e',v:23,w:'#REF!'},'日期待确认'],[{t:'n',f:'A99'},'日期待确认'],[{t:'n',v:serial,f:'A99'},'2026/10/19周']
 ]){
  const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付时间','付款自然周'],[raw,'旧周标记']],{origin:'N1'});set(i,'ledger','应付款时间',3,raw);
  if(date1904){i.negotiated.book.Workbook={WBProps:{date1904:true}};i.ledger.book.Workbook={WBProps:{date1904:true}};}
  const before=JSON.stringify(i),r=C.build(i),column=r.negotiated.headers.indexOf('付款自然周');assert.equal(r.negotiated.rows[0][column].v,'日期待确认');assert.equal(r.normalRows[0][r.normalHeaders.indexOf('提交付款申请单周')].v,expected);assert.equal(r.normalHeaders.at(-1),'提交付款申请单周');assert.equal(r.negotiated.headers.filter(h=>h==='付款自然周').length,1);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(JSON.stringify(i),before);
 }
 const i=fixture(),r=C.build(i);assert.equal(r.negotiated.rows[0].at(-1).v,'日期待确认');assert.equal(r.normalRows[0][r.normalHeaders.indexOf('提交付款申请单周')].v,'2026/10/19周');assert.equal(r.reviewHeaders.includes('付款自然周'),false);
});

test('谈判付款只按有金额的表头日期归周，多周去重排序，原计划和公式完整保留',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets['10月谈判付款主表'];
 X.utils.sheet_add_aoa(s,[['应付款时间',46325,'2026-11-01','2026-11-03','2026-12-31','2027-01-02','2026-10-05','2026-10-12'],['2030-01-01',10,{t:'n',v:20,f:'G2'},30,40,50,0,'']],{origin:'N1'});
 const before=JSON.stringify(i),r=C.build(i);assert.equal(r.negotiated.rows[0].at(-1).v,'2026/10/26周；2026/11/02周；2026/12/28周');assert.equal(JSON.stringify(i),before);assert.equal(r.normalRows[0][r.normalHeaders.indexOf('提交付款申请单周')].v,'2026/10/19周');
 for(const [header,amount,expected] of [
  ['2026-10-28',0,'日期待确认'],['2026-10-28','', '日期待确认'],['2026-10-28','￥1,234.56','2026/10/26周'],['2026-10-28',{t:'n',v:12,f:'G2'},'2026/10/26周'],
  ['2026-10-28',{t:'n',f:'G2'},'付款计划待确认'],['2026-10-28',{t:'e',v:23},'付款计划待确认'],['2026-10-28',-10,'付款计划待确认'],['2026-10-28','待定','付款计划待确认'],
  ['10月20日沟通','待定','日期待确认'],['10月21日沟通',500,'日期待确认'],['3月25',100,'付款计划待确认'],['2026-02-30',100,'付款计划待确认'],['2026-10-28',null,'日期待确认']
 ]){const input=fixture(),sheet=input.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(sheet,[[header],[amount]],{origin:'N1'});assert.equal(C.build(input).negotiated.rows[0].at(-1).v,expected);}
 const input=fixture(),sheet=input.negotiated.book.Sheets['10月谈判付款主表'];input.negotiated.book.Workbook={WBProps:{date1904:true}};X.utils.sheet_add_aoa(sheet,[[46325-1462,'3月25'],[50,10]],{origin:'N1'});assert.equal(C.build(input).negotiated.rows[0].at(-1).v,'2026/10/26周；付款计划待确认');
});
test('正常付款申请状态旁增加时间汇总，精确名称金额匹配并保留每笔来源',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2];
 const earlier=row.slice();earlier[0]='2026-10-08';earlier[3]='其他合同编号';
 const partial=row.slice();partial[8]=50;
 const other=row.slice();other[1]='其他供应商';
 X.utils.sheet_add_aoa(s,[earlier,earlier,partial,other],{origin:-1});
 const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.normalHeaders.slice(17,20),['是否提交付款申请单','提交时间汇总','当前节点']);
 assert.equal(r.normalRows[0][17].v,'已提交');
 assert.equal(r.normalRows[0][18].v,'2026-10-08-100.00\n2026-10-08-100.00\n2026-10-10-100.00');
 assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('正常付款未申请、金额或名称无法匹配时汇总为空，不沿用旧摘要',()=>{
 for(const [header,v] of [['供应商全称','其他供应商'],['供应商全称',''],['本次申请付款金额',99],['本次申请付款金额',''],['本次申请付款金额','异常'],['本次申请付款金额',0.001]]){
  const i=fixture();set(i,'applications',header,3,v);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['提交时间汇总'],[],['旧摘要']],{origin:'R1'});
  const r=C.build(i);assert.equal(r.normalRows[0][17].v,'未申请');assert.equal(r.normalRows[0][18].v,'');assert.equal(r.reviewRows.length,0);
 }
});
test('正常付款汇总保留零元及日期空白或异常提示，1904日期正确转换',()=>{
 for(const [raw,expected] of [['','日期为空'],['日期待定','日期待定'],['2026-02-30','2026-02-30'],[44834,'2026-10-01']]){
  const i=fixture();set(i,'ledger','申请付款金额',3,0);set(i,'ledger','已付款金额',3,0);set(i,'applications','本次申请付款金额',3,0);set(i,'applications','申请日期',3,raw);
  if(typeof raw==='number')i.applications.book.Workbook={WBProps:{date1904:true}};
  const r=C.build(i);assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(r.normalRows[0][18].v,expected+'-0.00');
 }
});
test('五张台账、谈判付款以自身为准、正常付款字段顺序、OA状态与不修改源表',()=>{
 const input=fixture(),before=JSON.stringify(input),r=C.build(input);assert.equal(r.stats.pending,0);assert.equal(r.stats.normalRows,1);assert.equal(r.stats.negotiatedRows,1);assert.deepEqual(r.normalHeaders,C.NORMAL_HEADERS);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);assert(!r.negotiated.updates.some(u=>['E2','J2','K2','L2','M2'].includes(u.address)));assert.equal(r.normalRows[0][15].f,'I2-N2');assert.equal(r.normalRows[0][15].v,0);assert.equal(r.normalRows[0][14].v,'是');assert.equal(r.normalRows[0][22].v,'李奇');assert.equal(r.normalRows[0][0].t,'n');assert.equal(r.normalRows[0][5].z,C.DATE_FORMAT);assert.equal(JSON.stringify(input),before);assert.equal(r.negotiated.rows[0][12].f,'G2');
});
test('谈判工作表名称包含匹配：零、一、多张及选择合法性',()=>{
 const i=fixture();assert.deepEqual(C.negotiatedSheets(i.negotiated.book),['10月谈判付款主表']);X.utils.book_append_sheet(i.negotiated.book,X.utils.aoa_to_sheet([NH]),'谈判付款-反馈');assert.throws(()=>C.build(i),/请选择/);i.negotiated.sheetName='10月谈判付款主表';assert.doesNotThrow(()=>C.build(i));i.negotiated.sheetName='不存在';assert.throws(()=>C.build(i),/请选择/);i.negotiated.book=book('普通表',NH,[]);assert.throws(()=>C.build(i),/未找到名称/);
});
test('同一OA多行一致时折叠，不将归档当成付款或申请提交',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇','测试供应商甲','￥100.00']],{origin:-1});set(i,'oa','流程状态',3,'归档');set(i,'ledger','已付款金额',3,0);set(i,'ledger','是否结清',3,'');const r=C.build(i);assert.equal(r.stats.oaFlows,2);assert.equal(r.normalRows[0][13].v,0);assert.equal(r.normalRows[0][14].v,'否');assert.equal(r.normalRows[0][21].v,'归档');assert.equal(r.stats.pending,0);assert.equal(r.stats.oaConsistent,2);assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');
});
test('正常付款空白实付直接保留原情况，零实付正常计算，斜线不转零',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,'');let r=C.build(i);assert.equal(r.normalRows[0][13].v,'');assert.equal(r.normalRows[0][14].v,'是');assert.equal(r.normalRows[0][15].v,0);assert.equal(r.issues.some(x=>x.kind==='amount'),false);assert.equal(resolveKeep(i).stats.pending,0);
 set(i,'ledger','已付款金额',3,0);set(i,'ledger','是否结清',3,'');r=C.build(i);assert.equal(r.normalRows[0][13].v,0);assert.equal(r.normalRows[0][15].v,100);set(i,'ledger','已付款金额',3,'/');r=C.build(i);assert.equal(r.normalRows[0][13].v,'/');assert(r.issues.some(x=>x.kind==='amount'));
});
test('明确金额按分计算且不静默截断第三位或越界金额',()=>{
 assert.equal(C.money({v:'￥1,000.20'}).cents,100020n);assert.equal(C.money({v:0.1+0.2}).cents,30n);assert(C.money({v:0.001}).invalid);assert(C.money({v:'90071992547409.92'}).invalid);assert(C.money({v:'待核对'}).invalid);
 const i=fixture();set(i,'ledger','申请付款金额',3,0.3);set(i,'ledger','已付款金额',3,0.1);set(i,'ledger','是否结清',3,'');const r=C.build(i);assert.equal(r.normalRows[0][15].v,0.2);
});
test('正常付款结清与金额冲突需要明确选择，负余额不得自动更新',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,20);let r=C.build(i),q=r.issues.find(x=>x.kind==='settlement');assert(q);assert.equal(r.normalRows[0][15].v,0);r=C.build(i,{[q.id]:['recalculate']});assert.equal(r.normalRows[0][15].v,80);assert.equal(r.normalRows[0][14].v,'否');set(i,'ledger','已付款金额',3,101);r=C.build(i);assert.equal(r.normalRows[0][15].v,0);assert(r.issues.some(x=>x.kind==='amount'));
});
test('付款申请多次日期全部保留，不使用申请表已付覆盖资金台账',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;const row=X.utils.sheet_to_json(s,{header:1})[1];row[0]='2026-10-08';row[5]=90;X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const r=C.build(i);assert.equal(r.stats.pending,0);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);const dateCol=r.negotiated.headers.indexOf('提交付款申请单日期');assert.equal(r.negotiated.rows[0][dateCol].v,'2026-10-08\n2026-10-10');
});
test('无对应付款OA且单次金额不一致时不累加同供应商申请，旧人工选择不覆盖',()=>{
 const i=fixture();set(i,'applications','合同号',2,'MDST-DGFK20990101999');set(i,'applications','本次申请付款金额',2,50);const s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const r=C.build(i,{'application:10月谈判付款主表!2':['match:付款申请汇总!2']});assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未提交');assert.equal(r.reviewRows.length,0);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);
});
test('无付款申请记录时两类付款均自动未申请，忽略旧申请确认选择',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i,{'application:10月谈判付款主表!2':['keep','not-submitted']});assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未提交');assert.equal(r.normalRows[0][17].v,'未申请');assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('提交付款申请单日期')].v,'');
});
test('谈判重复记录逐条完整保留，不从资金台账重复分配付款金额',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const r=C.build(i,{'payment-match:10月谈判付款主表!2':['match:对账明细-李奇!2']});assert.equal(r.negotiated.rows.length,2);assert.deepEqual(r.negotiated.rows.map(row=>[row[4].v,row[9].v,row[10].v]),[[0,'否',100],[0,'否',100]]);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'E3'),undefined);
});
test('旧分类是或空但名称金额未命中时纳入正常付款，已结清也保留',()=>{
 for(const flag of ['是','', '待定']){const i=fixture();set(i,'ledger','是否谈判付款',3,flag);const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][14].v,'是');assert.equal(r.normalRows[0][16].v,'否');}
});
test('OA重复冲突与多个编号保留每个编号对应状态',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');assert(q);const candidate=q.choices.find(c=>c.label.includes('出纳'));r=C.build(i,{[q.id]:[candidate.value]});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('当前节点')].v,'出纳');
 const j=fixture();set(j,'ledger','对公付款0A流程号',3,'MDST-DGFK20261001001、MDST-DGFK20261001002');set(j,'oa','收款单位',2,'测试供应商乙');r=C.build(j);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001001：财务审核\nMDST-DGFK20261001002：财务审核');assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001001\nMDST-DGFK20261001002');
});
test('正常付款公式仅使用有效缓存，缺缓存保留错误并阻止未核对导出',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,{t:'n',f:'10+10',v:20});set(i,'ledger','是否结清',3,'');let r=C.build(i);assert.equal(r.normalRows[0][13].v,20);assert.equal(r.normalRows[0][13].f,undefined);set(i,'ledger','已付款金额',3,{t:'n',f:'10+10'});r=C.build(i);assert.equal(r.normalRows[0][13].v,'公式结果不可用');assert(r.stats.pending>0);set(i,'ledger','OA合同编号',3,{t:'e',v:23,w:'#REF!'});r=C.build(i);assert.equal(r.normalRows[0][11].v,'#REF!');assert(r.issues.some(x=>x.kind==='formula'));
});
test('已有追加字段不重复新增，源表空行和非授权字段保持完整',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号']],{origin:'N1'});s['!ref']='A1:T4';s.L4={t:'s',v:'尾注'};const r=C.build(i);assert.equal(r.negotiated.headers.length,23);assert.equal(r.negotiated.headers[20],'应付款时间');assert.equal(r.negotiated.headers[21],'提交时间汇总');assert.equal(r.negotiated.headers.at(-1),'付款自然周');assert.equal(r.negotiated.rows.length,3);assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);assert.equal(r.negotiated.rows[2][11].v,'尾注');assert(!r.negotiated.updates.some(u=>u.address==='L4'));
});
test('1904来源日期转换为1900体系，原表不变',()=>{
 const i=fixture();i.ledger.book.Workbook={WBProps:{date1904:true}};set(i,'ledger','开票时间',3,{t:'n',v:44834,z:'yyyy-mm-dd'});const r=C.build(i);assert.equal(r.normalRows[0][5].v,44834+1462);assert.equal(i.ledger.book.Sheets['对账明细-李奇'].E3.v,44834);
});
test('登记月份和应付款时间无法识别时完整保留原文，不产生日期核对阻断',()=>{
 for(const [month,due] of [['10月登记',' 月底前付款\n以通知为准 '],['2026-13','2026-02-30'],['/','待定'],['','']]){
  const i=fixture();set(i,'ledger','月份',3,month);set(i,'ledger','应付款时间',3,due);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.normalRows[0][0].v,month);assert.equal(r.normalRows[0][9].v,due);assert.equal(r.normalRows[0][0].t,'s');assert.equal(r.normalRows[0][9].t,'s');assert.equal(r.issues.some(x=>x.kind==='date'),false);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('登记月份和应付款时间有效时仍导出真实Excel日期及相应格式',()=>{
 const i=fixture();set(i,'ledger','月份',3,'2026-10');set(i,'ledger','应付款时间',3,'2026年10月31日');const r=C.build(i),month=r.normalRows[0][0],due=r.normalRows[0][9];assert.equal(month.t,'n');assert.equal(X.SSF.format('yyyy-mm-dd',month.v),'2026-10-01');assert.equal(month.z,'yyyy"年"m"月"');assert.equal(due.t,'n');assert.equal(X.SSF.format('yyyy-mm-dd',due.v),'2026-10-31');assert.equal(due.z,C.DATE_FORMAT);assert.equal(r.stats.pending,0);
});
test('开票时间无法识别仍保留原文并要求日期核对',()=>{
 const i=fixture();set(i,'ledger','月份',3,'10月登记');set(i,'ledger','应付款时间',3,'月末付款');set(i,'ledger','开票时间',3,'开票待确认');const r=C.build(i),issues=r.issues.filter(x=>x.kind==='date');assert.equal(r.normalRows[0][5].v,'开票待确认');assert.equal(issues.length,1);assert.equal(issues[0].id,'date:对账明细-李奇!3:开票时间');assert.equal(issues[0].resolved,false);assert.equal(r.stats.pending,1);
});
test('OA只读取指定字段，其他重复明细列不构成冲突',()=>{
 const i=fixture();i.oa.book=book('Sheet0',[...OH,'备注','备注'],[{'流程编号':'MDST-DGFK20261001001','当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇'}]);assert.doesNotThrow(()=>C.build(i));
});
test('名称金额未命中时仍保留谈判原付款，旧来源选择不能覆盖原值',()=>{
 const i=fixture();set(i,'ledger','申请付款金额',2,101);const r=C.build(i,{'payment-match:10月谈判付款主表!2':['match:对账明细-李奇!2']});assert.equal(r.normalRows.length,2);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);
});
test('谈判原付款状态与金额即使不一致也原样保留，不受资金台账覆盖',()=>{
 const i=fixture();set(i,'ledger','是否结清',2,'');set(i,'negotiated','是否结清',2,'是');set(i,'negotiated','剩余付款金额',2,{t:'n',v:99,f:'G2-E2-1'});const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='settlement'),false);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'J2'),undefined);assert.equal(update(r,'K2'),undefined);assert.equal(r.negotiated.rows[0][9].v,'是');assert.deepEqual(r.negotiated.rows[0][10],i.negotiated.book.Sheets['10月谈判付款主表'].K2);
});
test('资金台账与谈判表财务单号冲突不能补空，申请状态不依赖财务单号',()=>{
 const i=withoutMatchingOA(fixture());set(i,'negotiated','单据号【财务应付单】',2,'AP999999');set(i,'negotiated','对公付款0A流程号',2,'');const r=C.build(i);assert.equal(update(r,'C2'),undefined);assert.equal(update(r,'E2'),undefined);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);
 const j=withoutMatchingOA(fixture());set(j,'applications','备注',2,'AP999999');assert.equal(C.build(j).issues.some(x=>x.kind==='application'),false);assert.equal(C.build(j).negotiated.rows[0][C.build(j).negotiated.headers.indexOf('是否提交付款申请单')].v,'已经提交全部');
});
test('供应商或主体为空直接保留原值，不产生身份核对',()=>{
 for(const field of ['供应商名称','付款主体']){const i=fixture();set(i,'ledger',field,3,'');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='classification'||x.kind==='identity'),false);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][r.normalHeaders.indexOf(field)].v,'');}
 const i=fixture();set(i,'negotiated','供应商',2,'');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='identity'),false);assert.equal(r.negotiated.rows[0][0].v,'');assert.equal(r.negotiated.rows[0][1].v,'');
});
test('旧分类否或空但名称金额命中谈判表时只保留谈判原表，不混入正常付款',()=>{
 for(const flag of ['否','']){const i=fixture();set(i,'ledger','是否谈判付款',2,flag);const before=JSON.stringify(i);const r=C.build(i);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][2].v,'测试供应商乙');assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(update(r,'E2'),undefined);assert.equal(JSON.stringify(i),before);}
});
test('同名不同金额分别分类，同名同金额跨主体按出现性归类且谈判付款原样保留',()=>{
 const i=fixture();set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);let r=C.build(i);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][8].v,101);
 set(i,'ledger','申请付款金额',3,100);set(i,'ledger','付款主体',3,'国源');r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.negotiatedLedgerRows,2);assert.equal(r.stats.negotiatedRows,1);assert.equal(update(r,'E2'),undefined);
});
test('分类按分精确比较，规范空白、全角名称和金额表示，零值有效而空金额不算零',()=>{
 const i=fixture();set(i,'ledger','供应商名称',2,' 测试供应商Ａ ');set(i,'negotiated','供应商',2,'测试供应商A');set(i,'ledger','申请付款金额',2,'￥1,000.20');set(i,'negotiated','申请付款金额',2,1000.2);let r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.normalRows.length,1);
 set(i,'ledger','申请付款金额',2,0);set(i,'negotiated','申请付款金额',2,0);r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);set(i,'ledger','申请付款金额',2,'');r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,0);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.normalRows.length,2);assert.equal(r.normalRows[0][8].v,'');assert.equal(r.issues.some(x=>x.id==='classification:对账明细-李奇!2'),false);
});
test('空白及异常金额直接纳入正常付款，保留金额并忽略旧分类决策',()=>{
 for(const amount of ['', '/', -1, '0.001', {t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'ledger','申请付款金额',3,amount);const before=JSON.stringify(i),r=C.build(i,{'classification:对账明细-李奇!3':['keep']});assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(r.normalRows.length,1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.normalRows[0][8].v,typeof amount==='object'?amount.w||'公式结果不可用':amount);assert.equal(r.normalRows[0][16].v,'否');assert.equal(r.negotiated.rows.length,1);assert.equal(JSON.stringify(i),before);
 }
});
test('供应商为空白或错误时自动正常付款，不删除谈判原记录',()=>{
 for(const supplier of ['', '　\n', {t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'ledger','供应商名称',2,supplier);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.stats.negotiatedRows,1);assert.equal(r.stats.negotiatedLedgerRows,0);assert.equal(r.normalRows.length,2);assert.equal(r.normalRows[0][8].v,100);assert.equal(r.normalRows[0][16].v,'否');assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(JSON.stringify(i),before);
 }
});
test('正常付款仅保留原OA有值的记录，匹配流程及旧对账值不能替代原OA',()=>{
 const cases=[['',false],['　\n',false],[{t:'s',f:'A99',v:''},false],[{t:'n',f:'A99'},false],[{t:'e',v:23,w:'#REF!'},false],['MDST-DGFK20261001002',true],[{t:'s',f:'A99',v:'MDST-DGFK20261001002'},true],['MDST-DGFK20261001002、MDST-DGFK20261001003',true]];
 for(const [oa,included] of cases){
  const i=fixture();set(i,'ledger','对公付款0A流程号',3,oa);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['是否对账'],['已对账'],[included?'未对账':'已对账']],{origin:'R1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.normalHeaders[24],'是否对账');assert.equal(r.normalRows.length,included?1:0);assert.equal(r.stats.excludedNormalRows,included?0:1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,included&&typeof oa==='string'&&oa.includes('、')?1:0);if(included){assert.equal(r.normalRows[0][24].v,'已对账');assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');}assert.equal(r.normalHeaders.filter(h=>h==='是否对账').length,1);assert.equal(r.negotiated.headers.includes('是否对账'),false);assert.equal(JSON.stringify(i),before);
 }
 const i=fixture();delete i.ledger.book.Sheets['对账明细-李奇'].D3;assert.equal(C.build(i).normalRows.length,0);delete i.ledger.book.Sheets['对账明细-李奇'].D1;const r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.negotiatedRows,1);assert(r.issues.some(q=>q.kind==='missing-header'&&q.message.includes('对公付款0A流程号')));
});
test('无OA正常记录不进入金额日期及申请核对，保留后续顺序并重排公式行号',()=>{
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[2];X.utils.sheet_add_aoa(s,[row,row],{origin:-1});set(i,'ledger','供应商名称',4,'保留供应商丙');set(i,'ledger','供应商名称',5,'保留供应商丁');set(i,'ledger','对公付款0A流程号',3,'');set(i,'ledger','申请付款金额',3,'异常金额');set(i,'ledger','开票时间',3,'异常日期');i.applications.book=book('付款申请汇总',AH,[]);const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.normalRows.map(row=>row[2].v),['保留供应商丙','保留供应商丁']);assert.deepEqual(r.normalRows.map(row=>row[15].f),['I2-N2','I3-N3']);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.filter(row=>row[0].v==='正常付款').length,0);assert.equal(r.reviewRows.some(row=>row[1].v==='测试供应商乙'),false);assert.equal(JSON.stringify(i),before);
});
test('未带原OA的记录不触发分类待核对，谈判原表和谈判匹配记录仍完整保留',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];row[6]='';X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);for(const r of [2,3])set(i,'ledger','对公付款0A流程号',r,'');set(i,'negotiated','对公付款0A流程号',2,'');const before=JSON.stringify(i),r=C.build(i);
 assert.equal(r.normalRows.length,0);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.stats.negotiatedRows,2);assert.equal(r.negotiated.rows.length,2);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.issues.some(q=>q.kind==='classification'),false);assert.equal(JSON.stringify(i),before);
});
test('谈判表同供应商金额缺失时未命中的款项需核对，已命中的款项仍可确定分类',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];row[6]='';X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);let r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.normalRows.length,0);const q=r.issues.find(x=>x.id==='classification:对账明细-李奇!3');assert.match(q.message,/谈判来源/);assert.match(q.message,/第 3 行/);r=C.build(i,{[q.id]:['normal']});assert.equal(r.normalRows.length,1);
});
test('重复明细保持原顺序，未选中的谈判反馈表不参与分类，缺失旧分类列不提示',()=>{
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[2];X.utils.sheet_add_aoa(s,[row],{origin:-1});delete s.P1;
 const feedback=book('谈判付款-反馈',NH,[{'供应商':'测试供应商乙','申请付款金额':100}]);X.utils.book_append_sheet(i.negotiated.book,feedback.Sheets['谈判付款-反馈'],'谈判付款-反馈');i.negotiated.sheetName='10月谈判付款主表';const r=C.build(i);assert.equal(r.normalRows.length,2);assert.deepEqual(r.normalRows.map(row=>row[2].v),['测试供应商乙','测试供应商乙']);assert.equal(r.issues.some(x=>x.kind==='missing-header'&&x.message.includes('是否谈判付款')),false);assert.equal(r.stats.negotiatedRows,1);
});
test('同AP同金额但已有OA冲突时不补空，不产生付款来源核对',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261024999');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(r.negotiated.rows[0][2].v,'MDST-DGFK20261001001');assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);
});
test('财务单号部分相交不能补空，集合相同顺序不同可以补空',()=>{
 const i=fixture();set(i,'ledger','单据号【财务应付单】',2,'AP000001、AP000003');set(i,'negotiated','单据号【财务应付单】',2,'AP000001、AP000004');set(i,'negotiated','付款主体',2,'');let r=C.build(i);assert.equal(update(r,'B2'),undefined);
 set(i,'negotiated','单据号【财务应付单】',2,'AP000003\nAP000001');r=C.build(i);assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'E2'),undefined);
 const j=fixture();set(j,'applications','备注',2,'AP000001、AP000099');r=C.build(j);assert.equal(r.issues.some(x=>x.kind==='application'),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'已经提交全部');
});
test('付款OA组部分相交不能补空且不能改变已付金额',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261001001、MDST-DGFK20261024999');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'E2'),undefined);
});
test('谈判仅供应商金额一致不能绕过OA号，正常付款仍按供应商金额匹配',()=>{
 const i=fixture();X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['OA合同编号'],['MDST-HT20260101']],{origin:'N1'});set(i,'applications','合同号',2,'MDST-HT20990101');set(i,'applications','付款主体',2,'国源');set(i,'applications','付款事由',2,'其他事由');set(i,'applications','备注',2,'AP999999');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未提交');
 const j=fixture();set(j,'ledger','OA合同编号',3,'MDST-HT20990101');set(j,'ledger','付款主体',3,'国源');assert.equal(C.build(j).normalRows[0][17].v,'已提交');
});
function historyFixture(value){const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单'],[value,'是']],{origin:'N1'});return i;}
test('已有提交日期与本次日期合并去重排序，反复处理不丢历史',()=>{
 const i=historyFixture('2026-10-09\n2026-10-01、2026-10-10');let r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10');
 set(i,'negotiated','提交付款申请单日期',2,update(r,'N2'));set(i,'applications','申请日期',2,'2026-10-12');r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10\n2026-10-12');
});
test('历史提交日期不能解析时保留原日期并列核对，不阻止已提交状态',()=>{
 const i=historyFixture('10月初提交过'),r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].v,'10月初提交过');assert.equal(update(r,'O2').v,'已经提交全部');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.reviewRows[0][4].v,'已提交');assert.equal(r.stats.pending,0);
});
test('未匹配申请时历史日期保留，提交状态重新计算为未申请',()=>{
 const i=historyFixture('2026-09-10、2026-09-12');i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i);assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2').v,'未提交');assert.equal(r.negotiated.rows[0][13].v,'2026-09-10、2026-09-12');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);
});
test('历史提交日期为Excel日期数值时正确转换后合并',()=>{
 const i=historyFixture(46294);const r=C.build(i);assert.equal(update(r,'N2').v,[X.SSF.format('yyyy-mm-dd',46294),'2026-10-10'].sort().join('\n'));
});
test('历史提交日期公式缺缓存时原样保留并列核对，匹配仍为已提交',()=>{
 const i=historyFixture('2026-09-10');set(i,'negotiated','提交付款申请单日期',2,{t:'n',f:'DATE(2026,9,10)'});const r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].f,'DATE(2026,9,10)');assert.equal(update(r,'O2').v,'已经提交全部');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);
});
test('正常付款原OA有值但填写错误时按供应商金额自动匹配，纠正原OA并保留是否对账',()=>{
 for(const old of ['待核对', 'MDST-DGFK20990101001']){
  const i=fixture();set(i,'ledger','对公付款0A流程号',3,old);const before=JSON.stringify(i),r=C.build(i,{'oa:对账明细-李奇!3':['keep']});assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);
  assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][13].v,100);assert.equal(r.normalRows[0][19].v,'财务审核');assert.equal(r.normalRows[0][20].v,'测试人员');assert.equal(r.normalRows[0][21].v,'未归档');assert.equal(r.normalRows[0][22].v,'李奇');assert.equal(r.normalRows[0][24].v,old?'已对账':'未对账');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/OA号复核.*已更新/);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判原OA空白错误时按唯一供应商金额补齐OA及流程信息',()=>{
 for(const old of ['', '　\n', 'MDST-DGFK20990101001', {t:'e',v:23,w:'#REF!'}, {t:'s',f:'A99'}]){
  const i=fixture();set(i,'negotiated','对公付款0A流程号',2,old);set(i,'ledger','对公付款0A流程号',2,old);const headers=['流程编号','当前节点','未操作者','流程状态','申请人'],values=['原流程','原节点','原人员','原状态','原申请人'];X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇']);assert.deepEqual([r.negotiated.rows[0][4].v,r.negotiated.rows[0][9].v,r.negotiated.rows[0][10].v],[0,'否',100]);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/OA号复核.*已更新/);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判按付款OA号匹配，不以收款单位或付款金额作为条件',()=>{
 for(const [field,value] of [['收款单位','其他供应商'],['付款金额',20],['付款金额',''],['付款金额',{t:'e',v:23,w:'#REF!'}],['收款单位',{t:'s',f:'A99'}]]){
  const i=fixture();set(i,'oa',field,2,value);const r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'财务审核');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'已付金额'),0);
 }
});
test('无精确OA匹配时两类付款原流程编号及全部进度信息自动保留，旧决定不影响',()=>{
 const headers=['当前节点','未操作者','流程状态','申请人','流程编号'];
 for(const values of [[' 原节点\n复核中 ','原未操作者','原流程状态','原申请人','原流程编号'],['','','','','']]){
  const i=fixture();set(i,'oa','流程编号',2,'MDST-DGFK20990101999');set(i,'oa','付款金额',2,99);set(i,'oa','付款金额',3,999);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[headers,[],values],{origin:'R1'});X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});const before=JSON.stringify(i);
  for(const answer of [[],['keep'],['match:Sheet0!2','match:Sheet0!3']]){
   const r=C.build(i,{'oa:对账明细-李奇!3':answer,'oa:10月谈判付款主表!2':answer});assert.deepEqual(r.normalRows[0].slice(19,24).map(c=>c.v),values);assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),values);assert.equal(r.issues.some(q=>q.kind==='oa'),false);assert.equal(r.stats.pending,0);assert.equal(r.stats.oaMatched,0);assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(r.normalRows[0][13].v,100);assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(JSON.stringify(i),before);
  }
 }
});
test('名称金额不可用时正常付款保留原流程，谈判付款仍按原OA匹配',()=>{
 for(const [role,field,value] of [['ledger','供应商名称',''],['ledger','供应商名称',{t:'e',v:23,w:'#REF!'}],['ledger','申请付款金额',''],['ledger','申请付款金额','金额待确认'],['negotiated','供应商',''],['negotiated','申请付款金额','']]){
  const i=fixture(),headers=['当前节点','未操作者','流程状态','申请人','流程编号'],values=['原节点','原未操作者','原状态','原申请人','原流程'];X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[headers,[],values],{origin:'R1'});X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});set(i,role,field,role==='ledger'?3:2,value);const before=JSON.stringify(i),r=C.build(i),targetId=role==='ledger'?'对账明细-李奇!3':'10月谈判付款主表!2';assert.equal(r.issues.some(q=>q.id==='oa:'+targetId),false);if(role==='ledger')assert.deepEqual(r.normalRows.find(row=>row[4].v==='MDST-DGFK20261001002').slice(19,24).map(c=>c.v),values);else assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),['财务审核','测试人员','未归档','李奇','MDST-DGFK20261001001']);assert.equal(JSON.stringify(i),before);
 }
});
test('同供应商同金额不同流程自动全部写入，重复展开行不重复写',()=>{
 for(const old of ['MDST-DGFK20990101999', 'MDST-DGFK20261001002']){
  const i=fixture(),s=i.oa.book.Sheets.Sheet0,row=['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','测试供应商乙',100];set(i,'ledger','对公付款0A流程号',3,old);X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][4].v,old);assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001002：财务审核\nMDST-DGFK20261001003：出纳');assert.equal(r.normalRows[0][20].v,'MDST-DGFK20261001002：测试人员\nMDST-DGFK20261001003：人员乙');assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002：未归档\nMDST-DGFK20261001003：归档');assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002：李奇\nMDST-DGFK20261001003：徐英');assert.equal(r.normalRows[0][13].v,100);assert.equal(r.normalRows[0][24].v,old?'已对账':'未对账');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('正常付款同流程号的收款单位或付款金额冲突仍须核对',()=>{
 for(const [payee,amount] of [['其他供应商',100],['测试供应商乙',101]]){
  const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','财务审核','测试人员','未归档','李奇',payee,amount]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/存在冲突/);assert.equal(r.normalRows[0][23].v,'');r=C.build(i,{[q.id]:[q.choices.find(c=>c.value.startsWith('match:')).value]});assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');
 }
});
test('谈判同OA的明细收款单位金额不同但流程信息一致时只读一次',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇','其他单位',20]],{origin:-1});const r=C.build(i);assert.equal(r.stats.pending,0);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'财务审核');assert.equal(negotiatedValue(r,'已付金额'),0);
});
test('谈判多个原OA逐号匹配，忽略同名同金额的其他流程并保留全部申请关联',()=>{
 const i=installmentFixture([{'本次申请付款金额':20},{'合同号':'MDST-DGFK20261001004','本次申请付款金额':30}]),ids='MDST-DGFK20261001003、MDST-DGFK20261001001、MDST-DGFK20261001004';set(i,'negotiated','对公付款0A流程号',2,ids);X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','其他单位',20],['MDST-DGFK20261001999','不应读取','其他人','未归档','其他人','测试供应商甲',100]],{origin:-1});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001003\nMDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'MDST-DGFK20261001003：出纳\nMDST-DGFK20261001001：财务审核');assert.equal(negotiatedValue(r,'对公付款0A流程号'),ids);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-20.00\n2026-10-10-30.00');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('谈判匹配OA的流程字段异常保留原值，不能用同名同金额的其他编号替代',()=>{
 const i=fixture();set(i,'oa','当前节点',2,{t:'e',v:23,w:'#REF!'});X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001999','出纳','其他人','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:10月谈判付款主表!2');assert(q);assert.match(q.message,/当前节点/);assert.equal(q.choices.length,1);assert.equal(negotiatedValue(r,'流程编号'),'');r=C.build(i,{[q.id]:['keep']});assert.equal(r.stats.pending,0);assert.equal(negotiatedValue(r,'流程编号'),'');
});
test('同一流程进度冲突不能同时选两版本，选择保留时原流程编号及进度都不覆盖',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0,n=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(n,[['流程编号','当前节点'],['旧流程记录','原节点']],{origin:'N1'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');r=C.build(i,{[q.id]:q.choices.filter(c=>c.value.startsWith('match:')).map(c=>c.value)});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2'),undefined);r=C.build(i,{[q.id]:['keep']});assert.equal(r.negotiated.rows[0][13].v,'旧流程记录');assert.equal(r.negotiated.rows[0][14].v,'原节点');assert.equal(r.negotiated.rows[0][2].v,'MDST-DGFK20261001001');assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);
});
test('选择冲突流程版本后仍完整写入其他唯一匹配流程，不能选唯一流程绕过冲突',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','测试供应商乙',100],['MDST-DGFK20261001003','复核','人员丙','未归档','徐英','测试供应商乙',100]],{origin:-1});const before=JSON.stringify(i);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);
 r=C.build(i,{[q.id]:['match:Sheet0!3']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(r.normalRows[0][23].v,'');
 const selected=q.choices.find(c=>c.label.includes('MDST-DGFK20261001003')&&c.label.includes('出纳'));assert(selected);r=C.build(i,{[q.id]:[selected.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001002：财务审核\nMDST-DGFK20261001003：出纳');assert.equal(r.normalRows[0][20].v,'MDST-DGFK20261001002：测试人员\nMDST-DGFK20261001003：人员乙');assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002：未归档\nMDST-DGFK20261001003：归档');assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002：李奇\nMDST-DGFK20261001003：徐英');assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][13].v,100);assert.equal(JSON.stringify(i),before);
});
test('有多个冲突流程时每个流程均需选择一个版本，部分选择不能完成核对',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','出纳','人员乙','归档','李奇','测试供应商乙',100],['MDST-DGFK20261001003','复核','人员丙','未归档','徐英','测试供应商乙',100],['MDST-DGFK20261001003','审批','人员丁','未归档','徐英','测试供应商乙',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);const chosen2=q.choices.find(c=>c.label.includes('MDST-DGFK20261001002')&&c.label.includes('出纳')),chosen3=q.choices.find(c=>c.label.includes('MDST-DGFK20261001003')&&c.label.includes('复核'));assert(chosen2&&chosen3);
 r=C.build(i,{[q.id]:[chosen2.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(r.normalRows[0][23].v,'');r=C.build(i,{[q.id]:[chosen2.value,chosen3.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001002：出纳\nMDST-DGFK20261001003：复核');
});
test('金额空白、错误、负数或超过分精度不参与自动匹配，明确零可以匹配',()=>{
 for(const amount of ['', '/', 0.001, -100, {t:'n',f:'100+0'}, {t:'e',v:23}]){
  const i=fixture();set(i,'oa','付款金额',3,amount);const r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][23].v,'');
 }
 const i=fixture();set(i,'ledger','申请付款金额',3,0);set(i,'ledger','已付款金额',3,0);set(i,'oa','付款金额',3,0);let r=C.build(i);assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);set(i,'oa','付款金额',3,'');r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][23].v,'');
});
test('OA匹配不累加多个流程金额，缺少匹配字段明确提示且不回退原OA',()=>{
 const i=fixture();set(i,'oa','付款金额',2,50);set(i,'oa','付款金额',3,50);set(i,'oa','收款单位',2,'测试供应商乙');let r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][23].v,'');
 const j=fixture();j.oa.book=book('Sheet0',OH.slice(0,5),[{'流程编号':'MDST-DGFK20261001002','当前节点':'出纳'}]);r=C.build(j);assert.equal(r.normalRows[0][23].v,'');assert.equal(r.normalRows[0][19].v,'');assert(r.issues.some(x=>x.kind==='missing-header'&&x.message.includes('收款单位、付款金额')));
});
test('原OA中的多余编号不能混入名称金额精确匹配结果',()=>{
 const i=fixture(),old='MDST-DGFK20261001002、MDST-DGFK20990101999';set(i,'ledger','对公付款0A流程号',3,old);const r=C.build(i);assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][19].v,'财务审核');assert.equal(r.normalRows[0][21].v,'未归档');assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);
});
test('旧OA不在源表时仍按名称金额匹配并更新流程信息，纠正原OA',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',3,'MDST-DGFK20260723032');X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['当前节点','流程状态'],[],['原节点','原状态']],{origin:'R1'});const r=C.build(i);assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][23].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][19].v,'财务审核');assert.equal(r.normalRows[0][21].v,'未归档');
});
test('原OA为Excel错误或公式无缓存时排除正常付款，已匹配流程不能使其入选',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){const i=fixture();set(i,'ledger','对公付款0A流程号',3,bad);const r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.oaMatched,1);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);}
});
test('OA进度错误或无缓存公式不得转成23或空白覆盖，重复去重不能掩盖错误',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const i=fixture();set(i,'oa','当前节点',3,bad);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/当前节点/);assert.equal(q.choices.length,1);assert.equal(r.normalRows[0][19].v,'');assert.equal(r.normalRows[0][23].v,'');r=C.build(i,{[q.id]:['keep']});assert.equal(r.normalRows[0][19].v,'');
 }
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;set(i,'oa','当前节点',3,{t:'e',v:23,w:'#REF!'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','23','测试人员','未归档','李奇','测试供应商乙',100]],{origin:-1});const r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/#REF!/);assert.equal(q.choices.length,2);assert.equal(r.normalRows[0][19].v,'');
});

test('谈判唯一同供应商同金额来源只补空OA主体财务单号，不导入资金付款金额或其他字段',()=>{
 const i=fixture();for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,'');set(i,'ledger','摘要',2,'不应覆盖的台账事由');set(i,'ledger','合同总金额',2,99999);const before=JSON.stringify(i),r=C.build(i);
 assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'H2').v,'AP000001');assert.equal(r.negotiated.rows[0][3].v,1000);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][5].v,'九月货款');assert.equal(r.negotiated.rows[0][10].v,100);assert.equal(r.issues.some(x=>['payment-match','identity'].includes(x.kind)),false);assert.equal(r.issues.some(x=>x.kind==='application'&&x.source.sheet.includes('谈判')),false);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.equal(JSON.stringify(i),before);
});
test('谈判补空候选不唯一时保持空白，主体已有值可排除另一主体',()=>{
 const i=withoutMatchingOA(fixture()),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[1];for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,'');X.utils.sheet_add_aoa(s,[row],{origin:-1});let r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'C2'),undefined);assert.equal(update(r,'H2'),undefined);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);
 set(i,'ledger','付款主体',4,'国源');set(i,'negotiated','付款主体',2,'迈德斯特');r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'H2').v,'AP000001');
});
test('谈判已有身份字段含文字、空值公式或错误时不能被覆盖',()=>{
 for(const old of ['手工保留编号',{t:'s',v:'',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=withoutMatchingOA(fixture());for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,old);const before=JSON.stringify(i),r=C.build(i);for(const a of ['B2','C2','H2'])assert.equal(update(r,a),undefined);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判原已付、结清及剩余空白错误与公式全部保留，不产生付款情况核对',()=>{
 for(const original of ['', '/', {t:'n',v:5,f:'G2-95'}, {t:'n',f:'A99'}, {t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'negotiated','已付金额',2,original);set(i,'negotiated','剩余付款金额',2,original);set(i,'negotiated','是否结清',2,'人工确认');const source=i.negotiated.book.Sheets['10月谈判付款主表'],r=C.build(i);for(const [a,c] of [['E2',4],['J2',9],['K2',10]]){assert.equal(update(r,a),undefined);assert.deepEqual(r.negotiated.rows[0][c],source[a]);}assert.equal(r.issues.some(x=>['amount','settlement','payment-match'].includes(x.kind)&&x.source.sheet.includes('谈判')),false);
 }
});
test('谈判缺少补充字段列时在表尾添加，原列和原付款保持不变',()=>{
 const i=fixture(),removed=['付款主体','对公付款0A流程号','单据号【财务应付单】'],oldSheet=i.negotiated.book.Sheets['10月谈判付款主表'],values=X.utils.sheet_to_json(oldSheet,{header:1})[1],headers=NH.filter(h=>!removed.includes(h));i.negotiated.book=book('10月谈判付款主表',headers,[Object.fromEntries(NH.map((h,c)=>[h,values[c]]))]);const r=C.build(i);assert.deepEqual(r.negotiated.headers.slice(0,headers.length),headers);for(const [h,v] of [['付款主体','迈德斯特'],['对公付款0A流程号','MDST-DGFK20261001001'],['单据号【财务应付单】','AP000001']]){assert.equal(r.negotiated.headers.filter(x=>x===h).length,1);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf(h)].v,v);}assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('已付金额')].v,0);
});
test('谈判补空不得忽略明确合同号冲突，也不得使用异常来源单元格',()=>{
 const i=withoutMatchingOA(fixture());set(i,'negotiated','付款主体',2,'');X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['OA合同编号'],['MDST-HT20990101']],{origin:'N1'});assert.equal(update(C.build(i),'B2'),undefined);
 for(const bad of [{t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const j=withoutMatchingOA(fixture());set(j,'negotiated','对公付款0A流程号',2,'');set(j,'ledger','对公付款0A流程号',2,bad);assert.equal(update(C.build(j),'C2'),undefined);}
});

test('未匹配付款申请仅标记未申请，不列入核对sheet，不改变两类付款金额',()=>{
 const i=fixture();set(i,'applications','合同号',2,'MDST-DGFK20990101999');for(const row of [2,3])set(i,'applications','本次申请付款金额',row,50);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,0);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(r.normalRows[0][17].v,'未申请');assert.deepEqual([negotiatedValue(r,'已付金额'),negotiatedValue(r,'剩余付款金额')],[0,100]);assert.deepEqual([r.normalRows[0][13].v,r.normalRows[0][15].v],[100,0]);assert.equal(JSON.stringify(i),before);
 set(i,'applications','合同号',2,'MDST-DGFK20261001001');for(const row of [2,3])set(i,'applications','本次申请付款金额',row,100);const confirmed=C.build(i);assert.equal(confirmed.reviewRows.length,0);assert.equal(negotiatedValue(confirmed,'是否提交付款申请单'),'已经提交全部');assert.equal(confirmed.normalRows[0][17].v,'已提交');
});
test('无付款申请候选时两类付款重新标记未申请且不列核对，谈判原提交日期保留',()=>{
 const i=historyFixture('2026-09-10');i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i);assert.equal(r.reviewRows.length,0);assert.equal(r.normalRows[0][17].v,'未申请');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-09-10');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,0);
});
test('多个谈判明细可接收同一申请时不重复分配，保留明细并列出归属核对',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const a=i.applications.book.Sheets.付款申请汇总,arow=X.utils.sheet_to_json(a,{header:1})[1];X.utils.sheet_add_aoa(a,[arow],{origin:-1});const r=C.build(i);assert.equal(r.reviewRows.length,2);const col=r.negotiated.headers.indexOf('是否提交付款申请单');assert.deepEqual(r.negotiated.rows.map(row=>row[col].v),['','']);assert(r.reviewRows.every(row=>row[5].v.includes('同一流程对应多笔付款')));assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,2);assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('谈判原合并区域只补锚点，不在合并区域其他单元格写值',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'negotiated','付款主体',2,'');set(i,'negotiated','付款主体',3,'');s['!merges']=[{s:{r:1,c:1},e:{r:2,c:1}}];const before=JSON.stringify(i),r=C.build(i);assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'B3'),undefined);assert.equal(JSON.stringify(i),before);
});
test('无法识别的既有OA或财务编号不被当作空白，以有效公式缓存补空时不复制源公式',()=>{
 for(const field of ['对公付款0A流程号','单据号【财务应付单】']){const i=fixture();set(i,'negotiated','付款主体',2,'');set(i,'negotiated',field,2,'待核对');assert.equal(update(C.build(i),'B2'),undefined);}
 const i=fixture();set(i,'negotiated','对公付款0A流程号',2,'');set(i,'ledger','对公付款0A流程号',2,{t:'s',v:'MDST-DGFK20261001001',f:'A99'});const r=C.build(i);assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'C2').f,undefined);assert.equal(i.ledger.book.Sheets['对账明细-李奇'].D2.f,'A99');
});

test('正常付款申请仍按名称及金额精确匹配，零有效',()=>{
 const i=fixture();set(i,'applications','供应商全称',3,'其他供应商');assert.equal(C.build(i).normalRows[0][17].v,'未申请');set(i,'applications','供应商全称',3,'　测试供应商乙 ');set(i,'applications','本次申请付款金额',3,'￥100.00');assert.equal(C.build(i).normalRows[0][17].v,'已提交');set(i,'applications','本次申请付款金额',3,100.01);assert.equal(C.build(i).normalRows[0][17].v,'未申请');set(i,'ledger','申请付款金额',3,0);set(i,'applications','本次申请付款金额',3,0);assert.equal(C.build(i).normalRows[0][17].v,'已提交');
});
test('正常付款申请名称或金额为空错误缺缓存时不命中',()=>{
 for(const amount of ['', '/', -100, 0.001, {t:'n',f:'A99'}, {t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','本次申请付款金额',3,amount);assert.equal(C.build(i).normalRows[0][17].v,'未申请');}
 for(const name of ['', {t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','供应商全称',3,name);assert.equal(C.build(i).normalRows[0][17].v,'未申请');}
});
test('谈判按OA号归集，缺编号按名称金额补判，错误或冲突编号不自动覆盖',()=>{
 for(const name of ['其他供应商','',{t:'e',v:23,w:'#REF!'}]){const i=installmentFixture([{'供应商全称':name,'本次申请付款金额':40},{'供应商全称':'不同名称','本次申请付款金额':60}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-40.00\n2026-10-10-60.00');assert.equal(r.reviewRows.length,0);}
 for(const id of ['', 'MDST-DGFK20990101999',{t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','合同号',2,id);const r=C.build(i);const missing=id===''||id?.f;assert.equal(negotiatedValue(r,'是否提交付款申请单'),missing?'已经提交全部':'未提交');assert.equal(negotiatedValue(r,'提交时间汇总'),missing?'2026-10-10-100.00':'');assert.equal(r.reviewRows.length,0);}
});
test('提交日期与汇总相邻，已有字段只重排列顺序，不重复新增或改源表',()=>{
 for(const existing of [[],['提交时间汇总'],['提交付款申请单日期'],['提交付款申请单日期','人工备注','提交时间汇总']]){
  const i=fixture();if(existing.length)X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[existing,existing.map(h=>h==='人工备注'?'必须保留':'')],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i),h=r.negotiated.columnOrder.map(c=>r.negotiated.headers[c]);assert.equal(Math.abs(h.indexOf('提交付款申请单日期')-h.indexOf('提交时间汇总')),1);assert.equal(h.filter(x=>x==='提交时间汇总').length,1);assert.equal(h.filter(x=>x==='提交付款申请单日期').length,1);assert.equal(JSON.stringify(i),before);if(existing.includes('人工备注'))assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('人工备注')].v,'必须保留');
 }
});
test('仅有付款OA号的申请也保留，空金额仍确认已提交并核对金额',()=>{
 const i=installmentFixture([{'供应商全称':'','本次申请付款金额':''}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-金额为空');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/申请金额为空或异常/);
});
test('匹配申请日期空白错误或缺缓存时保留原日期，已提交状态有效并进入核对表',()=>{
 for(const d of ['', '近期', {t:'n',f:'DATE(2026,10,10)'},{t:'e',v:23,w:'#REF!'}]){
  const i=historyFixture('2026-09-10');set(i,'applications','申请日期',2,d);const r=C.build(i);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].v,'2026-09-10');assert.equal(update(r,'O2').v,'已经提交全部');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.stats.pending,0);
 }
});
test('申请日期列缺失仍按OA号标记谈判已提交，日期保留并列入核对表',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',['供应商全称','本次申请付款金额','合同号'],[{'供应商全称':'测试供应商甲','本次申请付款金额':100,'合同号':'MDST-DGFK20261001001'},{'供应商全称':'测试供应商乙','本次申请付款金额':100,'合同号':'MDST-DGFK20261001002'}]);const r=C.build(i);assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'已经提交全部');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.stats.pending,0);
});

const negotiatedValue=(r,key,row=0)=>r.negotiated.rows[row][r.negotiated.headers.indexOf(key)].v;
function installmentFixture(entries){
 const i=fixture(),existing=X.utils.sheet_to_json(i.applications.book.Sheets.付款申请汇总);i.applications.book=book('付款申请汇总',AH,[...entries.map(e=>({...existing[0],'合同号':'MDST-DGFK20261001001',...e})),existing[1]]);return i;
}
test('付款申请流程编码替代旧合同号，分批状态和时间汇总一致且原表不变',()=>{
 for(const amounts of [[30,20],[40,60],[60,60],[0]]){
  const i=installmentFixture(amounts.map((amount,n)=>({'申请日期':n?'2026-10-12':'2026-10-10','本次申请付款金额':amount}))),legacy=C.build(i);
  i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:' 流程编码 '};const before=JSON.stringify(i),r=C.build(i);
  assert.deepEqual(r.normalRows,legacy.normalRows);assert.deepEqual(r.negotiated,legacy.negotiated);assert.deepEqual(r.stats,legacy.stats);
  if(r.reviewRows.length)assert.equal(r.reviewRows[0][6].v,'谈判.xlsx');
  assert.equal(JSON.stringify(i),before);
 }
});
test('新旧字段共存仅用流程编码，空白错误不同编号不回退旧合同号',()=>{
 for(const code of ['MDST-DGFK20261001001','MDST-DGFK20990101999','',{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]),s=i.applications.book.Sheets.付款申请汇总;
  X.utils.sheet_add_aoa(s,[['流程编码'],[''],['MDST-DGFK20261001002']],{origin:'J1'});set(i,'applications','流程编码',2,code);
  const r=C.build(i),matched=code==='MDST-DGFK20261001001';
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),matched?'已经提交部分':'未提交');
  assert.equal(negotiatedValue(r,'提交时间汇总'),matched?'2026-10-10-30.00':'');
  assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(r.normalRows[0][18].v,'2026-10-10-100.00');
  assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,0);
 }
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;
 set(i,'applications','合同号',2,'不匹配的旧合同');X.utils.sheet_add_aoa(s,[['流程编码'],['MDST-DGFK20261001001']],{origin:'J1'});
 assert.equal(negotiatedValue(C.build(i),'是否提交付款申请单'),'已经提交全部');
});
test('仅有流程编码的申请不能跳过，金额为空保留并列出核对来源',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;s.D1={t:'s',v:'流程编码'};
 for(const field of ['供应商全称','本次申请付款金额'])set(i,'applications',field,2,'');
 const r=C.build(i);assert.equal(r.stats.applications,2);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'');assert.equal(r.reviewRows.length,1);assert.equal(r.reviewRows[0][6].v,'谈判.xlsx');assert.equal(negotiatedValue(r,'已付金额'),0);
});
test('核对表流程编码对应即为已提交，名称金额异常和超额不抹去提交事实',()=>{
 for(const amount of ['',0.001,-1,120,{t:'e',v:23,w:'#REF!'}]){
  const i=installmentFixture([{'供应商全称':'不同名称','本次申请付款金额':amount}]);i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
  const before=JSON.stringify(i),r=C.build(i);assert.equal(r.reviewRows.length,1);assert.equal(r.reviewRows[0][4].v,'已提交');assert(r.reviewRows[0][5].v);assert.equal(r.reviewRows[0][6].v,'谈判.xlsx');
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),amount===120?'已经提交全部':'');assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
 }
});
test('同流程归属多条谈判记录时核对表标记已提交，仍保留原记录和归属说明',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];
 X.utils.sheet_add_aoa(s,[row],{origin:-1});i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
 const r=C.build(i);assert.equal(r.negotiated.rows.length,2);assert.equal(r.reviewRows.length,2);assert(r.reviewRows.every(row=>row[4].v==='已提交'&&row[5].v.includes('同一流程对应多笔付款')));assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('旧状态只写已提交但归属不清时须确认进度，保留核对来源',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});X.utils.sheet_add_aoa(s,[['是否提交付款申请单'],['已提交付款申请单']],{origin:'N1'});
 const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'');assert.equal(r.reviewRows[0][5].v,'同一流程对应多笔付款，请核对归属。');assert.deepEqual(r.reviewRows[0].slice(6,9).map(c=>c.v),['谈判.xlsx','10月谈判付款主表',2]);assert.equal(JSON.stringify(i),before);
});
test('付款申请核对已提交不等于全部提交，日期异常仍保留说明及累计进度',()=>{
 const i=installmentFixture([{'本次申请付款金额':40,'申请日期':'日期待确认'}]);i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
 const r=C.build(i);assert.equal(r.reviewRows[0][4].v,'已提交');assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');assert.equal(negotiatedValue(r,'提交时间汇总'),'日期待确认-40.00');
 i.applications.book=book('付款申请汇总',AH,[]);const empty=C.build(i);assert.equal(empty.reviewRows.length,0);assert.equal(negotiatedValue(empty,'是否提交付款申请单'),'未提交');assert.equal(negotiatedValue(empty,'流程编号'),'MDST-DGFK20261001001','OA流程存在不代表提交过申请');
});
test('谈判分批按OA归集，重复金额只计一次，再判断部分或全部',()=>{
 for(const [amounts,status] of [[[30,20],'已经提交部分'],[[40,60],'已经提交全部'],[[60,60],'已经提交部分'],[[0],'已经提交部分']]){
  const i=installmentFixture(amounts.map((amount,n)=>({'申请日期':n?'2026-10-12':'2026-10-10','本次申请付款金额':amount}))),before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),status);assert.equal(negotiatedValue(r,'提交时间汇总'),amounts.map((v,n)=>(n?'2026-10-12':'2026-10-10')+'-'+v.toFixed(2)).join('\n'));assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(r.stats.pending,0);assert.deepEqual(r.negotiated.rows[0].slice(4,5).map(c=>c.v),[0]);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(negotiatedValue(r,'是否结清'),'否');assert.equal(r.reviewRows.length,[...new Set(amounts)].reduce((a,b)=>a+b,0)>100?1:0);if([...new Set(amounts)].reduce((a,b)=>a+b,0)>100)assert.match(r.reviewRows[0][5].v,/申请超额.*120.00.*100.00/);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判同时按单次名称金额和付款OA命中的申请只计一次，分精度累计和零额有效',()=>{
 let i=installmentFixture([{'本次申请付款金额':100}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');assert.equal(r.reviewRows.length,0);
 for(const [total,amounts] of [[0,[0]],[0.3,[0.1,0.2]]]){i=installmentFixture(amounts.map(amount=>({'本次申请付款金额':amount})));set(i,'negotiated','申请付款金额',2,total);set(i,'ledger','申请付款金额',2,total);set(i,'oa','付款金额',2,total);r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);}
});
test('提交时间汇总按日期排序且同日同额源行全部保留，已有汇总列不重复添加',()=>{
 const i=installmentFixture([{'申请日期':'2026-10-12','本次申请付款金额':30},{'申请日期':'2026-10-10','本次申请付款金额':20},{'申请日期':'2026-10-10','本次申请付款金额':20}]);X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['提交时间汇总'],['旧汇总']],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-10-10\n2026-10-12');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-20.00\n2026-10-10-20.00\n2026-10-12-30.00');assert.equal(r.negotiated.headers.filter(h=>h==='提交时间汇总').length,1);assert.equal(JSON.stringify(i),before);
});
test('同名同OA申请金额为空异常负数或精度错误时确认提交并核对金额，不当零不改原付款',()=>{
 for(const amount of ['', '/', -1, 0.001,{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]);set(i,'applications','本次申请付款金额',2,amount);const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/申请金额为空或异常/);assert.equal(r.stats.pending,1);assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
 }
});
test('同名OA明确关联优先于另一谈判记录的单次同额，不重复分配',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),n=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(n,{header:1})[1];row[2]='MDST-DGFK20261001003';row[6]=40;X.utils.sheet_add_aoa(n,[row],{origin:-1});X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','李奇','测试供应商甲',40]],{origin:-1});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单',0),'已经提交部分');assert.equal(negotiatedValue(r,'提交时间汇总',0),'2026-10-10-40.00');assert.equal(negotiatedValue(r,'是否提交付款申请单',1),'未提交');assert.equal(negotiatedValue(r,'提交时间汇总',1),'');assert.equal(r.reviewRows.length,0);assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('申请归属不清时保留原日期汇总，无法识别的进度须确认后下载',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),n=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(n,{header:1})[1];X.utils.sheet_add_aoa(n,[row],{origin:-1});X.utils.sheet_add_aoa(n,[['是否提交付款申请单','提交付款申请单日期','提交时间汇总'],['人工确认','2026-09-01','原汇总甲'],['','2026-09-02','原汇总乙']],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单',0),'');assert.equal(negotiatedValue(r,'是否提交付款申请单',1),'');assert.equal(negotiatedValue(r,'提交付款申请单日期',0),'2026-09-01');assert.equal(negotiatedValue(r,'提交付款申请单日期',1),'2026-09-02');assert.equal(negotiatedValue(r,'提交时间汇总',0),'原汇总甲');assert.equal(negotiatedValue(r,'提交时间汇总',1),'原汇总乙');assert.equal(r.reviewRows.length,2);assert(r.reviewRows.every(row=>/同一流程对应多笔付款/.test(row[5].v)));assert.equal(r.stats.pending,2);assert.equal(JSON.stringify(i),before);
});
test('谈判分批申请始终按原付款OA关联，OA表缺记录或总金额不同不影响关联',()=>{
 const i=withoutMatchingOA(installmentFixture([{'合同号':'MDST-DGFK20990101999','本次申请付款金额':30}]));set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20990101999');let r=C.build(i);assert.equal(negotiatedValue(r,'流程编号'),'');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');set(i,'applications','合同号',2,'MDST-DGFK20261001001');r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20990101999');
 const j=installmentFixture([{'本次申请付款金额':290000}]);set(j,'negotiated','申请付款金额',2,1202409);r=C.build(j);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-290000.00');assert.equal(r.issues.some(q=>q.id==='oa:10月谈判付款主表!2'),false);assert.equal(r.issues.some(q=>q.kind.startsWith('application')),false);
});
test('申请OA对应但供应商为空时仍更新提交汇总，保留历史提交日期',()=>{
 for(const supplier of ['',{t:'e',v:23,w:'#REF!'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]);set(i,'applications','供应商全称',2,supplier);X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['是否提交付款申请单','提交付款申请单日期','提交时间汇总'],['已经提交部分','2026-09-01','原汇总']],{origin:'N1'});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交部分');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-09-01\n2026-10-10');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-30.00');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);
 }
});
test('单次同额申请含未知多OA也不能直接归集，全部属于目标的多OA申请只计一次',()=>{
 for(const ids of ['MDST-DGFK20990101998、MDST-DGFK20990101999','MDST-DGFK20261001001、MDST-DGFK20990101999']){
  const i=installmentFixture([{'合同号':ids,'本次申请付款金额':100}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),ids.includes('MDST-DGFK20261001001')?'':'未提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'');assert.equal(r.reviewRows.length,ids.includes('MDST-DGFK20261001001')?1:0);if(r.reviewRows.length)assert.match(r.reviewRows[0][5].v,/多个流程|OA 号冲突/);assert.equal(r.stats.pending,ids.includes('MDST-DGFK20261001001')?1:0);
 }
 const i=installmentFixture([{'合同号':'MDST-DGFK20261001001、MDST-DGFK20261001003','本次申请付款金额':100}]);set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20261001001、MDST-DGFK20261001003');X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);
});

test('申请汇总流程编码标为谈判付款时优先分类，未在谈判源表中的记录仅排除不新增',()=>{
 for(const amount of [100,0,'','异常',-1]){
  const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;
  X.utils.sheet_add_aoa(s,[['流程编码','付款分类'],['MDST-DGFK20261001001','正常付款'],['MDST-DGFK20261001002','谈判付款']],{origin:'J1'});
  set(i,'applications','供应商全称',3,'不同供应商');set(i,'applications','本次申请付款金额',3,amount);
  const before=JSON.stringify(i),r=C.build(i);
  assert.equal(r.normalRows.length,0);assert.equal(r.stats.negotiatedLedgerRows,2);assert.equal(r.negotiated.rows.length,1);assert.equal(r.reviewRows.length,0);
  assert.equal(r.negotiated.rows[0][0].v,'测试供应商甲');assert.equal(r.negotiated.rows[0][12].f,'G2');assert.equal(JSON.stringify(i),before);
 }
});
test('付款分类必须为谈判付款且流程号完整匹配，新流程编码空白不回退旧合同号',()=>{
 for(const [id,type] of [['','谈判付款'],['MDST-DGFK202610010020','谈判付款'],['MDST-DGFK2026100100','谈判付款'],['MDST-DGFK20261001002','非谈判付款'],['MDST-DGFK20261001002',''],[{t:'e',v:23},'谈判付款'],['MDST-DGFK20261001002',{t:'e',v:23}]]){
  const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['流程编码','付款分类'],[],[id,type]],{origin:'J1'});
  const r=C.build(i);assert.equal(r.normalRows.length,1);assert.equal(r.stats.negotiatedLedgerRows,1);
 }
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;X.utils.sheet_add_aoa(s,[['付款分类'],[],['谈判付款']],{origin:'J1'});
 assert.equal(C.build(i).normalRows.length,0,'旧表合同号仍兼容');
});
test('已在谈判源表的分类命中记录不进入付款申请核对，其他异常继续列出',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;
 set(i,'applications','申请日期',2,'异常日期');set(i,'applications','本次申请付款金额',2,200);
 assert(C.build(i).reviewRows.length>0);
 X.utils.sheet_add_aoa(s,[['付款分类'],['谈判付款']],{origin:'J1'});
 const r=C.build(i);assert.equal(r.reviewRows.length,0);assert.equal(r.negotiated.rows.length,1);assert.equal(r.normalRows.length,1);
 assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);
 const duplicate=X.utils.sheet_to_json(s,{header:1})[1].slice();duplicate[9]='正常付款';X.utils.sheet_add_aoa(s,[duplicate],{origin:-1});
 assert.equal(C.build(i).reviewRows.length,0,'存在任一谈判分类即生效');
});

test('谈判应付款时间按供应商、申请金额和OA号更新，不依赖主体并保留财务字段',()=>{
 const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];
 X.utils.sheet_add_aoa(n,[['应付时间'],['旧日期']],{origin:'N1'});
 set(i,'ledger','付款主体',2,'其他主体');set(i,'ledger','应付款时间',2,'2026年11月12日');
 const before=JSON.stringify(i),r=C.build(i),d=r.negotiated.rows[0][13];
 assert.equal(X.SSF.format(C.DATE_FORMAT,d.v),'2026年11月12日');assert.equal(d.t,'n');assert.equal(d.z,C.DATE_FORMAT);
 assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);assert.equal(r.negotiated.rows[0][12].f,'G2');assert.equal(JSON.stringify(i),before);
});
test('谈判应付日期未匹配或源日期空白保留原值，不把空金额作为零匹配',()=>{
 for(const [field,v] of [['供应商名称','其他供应商'],['申请付款金额',101],['申请付款金额',''],['申请付款金额','异常'],['应付款时间','']]){
  const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['原日期']],{origin:'N1'});set(i,'ledger',field,2,v);
  const r=C.build(i);assert.equal(r.negotiated.rows[0][13].v,'原日期');assert(!r.negotiated.updates.some(u=>u.address==='N2'));
 }
 const i=fixture();set(i,'ledger','申请付款金额',2,0);set(i,'negotiated','申请付款金额',2,0);
 const r=C.build(i);assert.equal(X.SSF.format(C.DATE_FORMAT,r.negotiated.rows[0][r.negotiated.headers.indexOf('应付款时间')].v),'2026年10月31日');
});
test('同键多笔日期一致可更新，日期冲突或错误保留并允许核对选择',()=>{
 for(const [raw,conflict] of [['2026年10月31日',false],['2026-11-01',true],['',true],[{t:'e',v:23},true],[{t:'n',f:'A99'},true],['2026-02-30',true]]){
  const i=fixture(),l=i.ledger.book.Sheets['对账明细-李奇'],n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['原日期']],{origin:'N1'});
  const row=X.utils.sheet_to_json(l,{header:1})[1].slice();row[8]=raw;X.utils.sheet_add_aoa(l,[row],{origin:-1});
  const r=C.build(i),q=r.issues.find(q=>q.kind==='due-date');assert.equal(!!q,conflict);
  if(conflict){assert.equal(r.negotiated.rows[0][13].v,'原日期');assert(q.message.includes('对账明细-李奇'));assert.equal(C.build(i,{[q.id]:['keep']}).negotiated.rows[0][13].v,'原日期');const choice=q.choices.find(c=>c.value.startsWith('date:'));assert.equal(C.build(i,{[q.id]:[choice.value]}).negotiated.rows[0][13].v,Number(choice.value.slice(5)));}
  else assert.equal(X.SSF.format(C.DATE_FORMAT,r.negotiated.rows[0][13].v),'2026年10月31日');
 }
});
test('谈判应付日期正确处理不同日期体系，保留合并单元格',()=>{
 for(const from1904 of [false,true])for(const to1904 of [false,true]){
  const i=fixture();i.ledger.book.Workbook={WBProps:{date1904:from1904}};i.negotiated.book.Workbook={WBProps:{date1904:to1904}};set(i,'ledger','应付款时间',2,{t:'n',v:46325-(from1904?1462:0)});
  const r=C.build(i),d=r.negotiated.rows[0][r.negotiated.headers.indexOf('应付款时间')];assert.equal(d.v,46325-(to1904?1462:0));
 }
 const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['合并日期']],{origin:'N1'});n['!merges']=[{s:{r:1,c:13},e:{r:2,c:13}}];
 const r=C.build(i);assert.equal(r.negotiated.rows[0][13].v,'合并日期');assert(r.issues.some(q=>q.kind==='due-date'&&q.message.includes('合并')));
});

test('更新谈判应付日期必须同时匹配原付款OA号，缺失或不一致不退回供应商金额匹配',()=>{
 for(const role of ['ledger','negotiated'])for(const id of ['', '   ', 'MDST-DGFK202610010010', 'OTHER-123', {t:'e',v:23}, {t:'n',f:'A99'}]){
  const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['原日期']],{origin:'N1'});set(i,role,'对公付款0A流程号',2,id);
  const before=JSON.stringify(i),r=C.build(i);assert.equal(r.negotiated.rows[0][13].v,'原日期');assert(!r.negotiated.updates.some(u=>u.address==='N2'));assert.equal(JSON.stringify(i),before);
 }
});
test('同供应商同金额不同OA时仅取匹配流程日期，同OA多日期仍核对',()=>{
 const i=fixture(),l=i.ledger.book.Sheets['对账明细-李奇'],n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['原日期']],{origin:'N1'});
 const row=X.utils.sheet_to_json(l,{header:1})[1].slice();row[3]='OTHER-123';row[8]='2026-12-31';X.utils.sheet_add_aoa(l,[row],{origin:-1});
 let r=C.build(i);assert.equal(X.SSF.format(C.DATE_FORMAT,r.negotiated.rows[0][13].v),'2026年10月31日');assert(!r.issues.some(q=>q.kind==='due-date'));
 set(i,'ledger','对公付款0A流程号',4,'MDST-DGFK20261001001');r=C.build(i);assert.equal(r.negotiated.rows[0][13].v,'原日期');assert(r.issues.some(q=>q.kind==='due-date'));
});
test('应付日期匹配多个OA号必须完整相同，忽略顺序和重复但不接受部分重叠',()=>{
 for(const [id,match] of [['OTHER-123；MDST-DGFK20261001001',true],['MDST-DGFK20261001001；OTHER-123；OTHER-123',true],['MDST-DGFK20261001001',false],['OTHER-123；THIRD-456',false]]){
  const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(n,[['应付款时间'],['原日期']],{origin:'N1'});set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20261001001；OTHER-123');set(i,'ledger','对公付款0A流程号',2,id);
  const r=C.build(i);assert.equal(r.negotiated.rows[0][13].v==='原日期',!match);
 }
});

test('全额申请后同OA又有分批申请，超额仍确认已提交，原申请和财务情况不变',()=>{
 const i=installmentFixture([{'申请日期':'2026-09-08','本次申请付款金额':1268598.1},{'申请日期':'2026-10-21','本次申请付款金额':668598.1}]);
 set(i,'negotiated','申请付款金额',2,1268598.1);set(i,'ledger','申请付款金额',2,1268598.1);
 const before=JSON.stringify(i),r=C.build(i);
 assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');
 assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-09-08-1268598.10\n2026-10-21-668598.10');
 assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/1937196.20/);
 assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
 X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['付款分类'],['谈判付款'],['谈判付款']],{origin:'J1'});
 const classified=C.build(i);assert.equal(negotiatedValue(classified,'是否提交付款申请单'),'已经提交全部');assert.equal(classified.reviewRows.length,0);
});

test('同OA多笔谈判付款按唯一供应商金额归属，逐笔计算提交进度不重复分配',()=>{
 const i=installmentFixture([{'本次申请付款金额':270840},{'本次申请付款金额':361120}]),n=i.negotiated.book.Sheets['10月谈判付款主表'];
 set(i,'negotiated','申请付款金额',2,270840);const row=X.utils.sheet_to_json(n,{header:1})[1].slice();row[6]=361120;X.utils.sheet_add_aoa(n,[row],{origin:-1});
 const before=JSON.stringify(i),r=C.build(i);assert.deepEqual(r.negotiated.rows.map((_,j)=>negotiatedValue(r,'是否提交付款申请单',j)),['已经提交全部','已经提交全部']);
 assert.deepEqual(r.negotiated.rows.map((_,j)=>negotiatedValue(r,'提交时间汇总',j)),['2026-10-10-270840.00','2026-10-10-361120.00']);assert.equal(r.reviewRows.length,0);assert.equal(JSON.stringify(i),before);
});
test('缺OA号允许唯一供应商金额补判，空值不互相匹配且重复目标不分摊',()=>{
 for(const id of ['', '未提交OA']){
  const i=withoutMatchingOA(fixture());set(i,'negotiated','对公付款0A流程号',2,id);set(i,'ledger','对公付款0A流程号',2,'');const r=C.build(i);
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'对公付款0A流程号'),id);assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');
 }
 const i=withoutMatchingOA(fixture()),n=i.negotiated.book.Sheets['10月谈判付款主表'];set(i,'negotiated','对公付款0A流程号',2,'未提交OA');set(i,'ledger','对公付款0A流程号',2,'');const row=X.utils.sheet_to_json(n,{header:1})[1].slice();X.utils.sheet_add_aoa(n,[row],{origin:-1});const r=C.build(i);assert.equal(r.reviewRows.length,2);assert(r.negotiated.rows.every((_,j)=>negotiatedValue(r,'是否提交付款申请单',j)===''));
 for(const [field,v] of [['供应商',''],['申请付款金额',''],['申请付款金额','异常']]){const input=withoutMatchingOA(fixture());set(input,'negotiated','对公付款0A流程号',2,'未提交OA');set(input,'negotiated',field,2,v);assert.equal(negotiatedValue(C.build(input),'是否提交付款申请单'),'未提交');}
});
test('两边OA号不同则未提交，不按名称金额回退，不列付款申请核对',()=>{
 const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];set(i,'applications','合同号',2,'OTHER-999');X.utils.sheet_add_aoa(n,[['是否提交付款申请单','提交付款申请单日期','提交时间汇总'],['未申请','2026-09-01','历史汇总']],{origin:'N1'});
 const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-09-01');assert.equal(negotiatedValue(r,'提交时间汇总'),'');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});

test('付款分类隔离OA归集、名称金额回退及冲突候选，不跨类别计算提交状态',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;
 X.utils.sheet_add_aoa(s,[['付款分类'],['谈判付款'],['正常付款']],{origin:'J1'});
 const rows=X.utils.sheet_to_json(s,{header:1}),wrong=rows[1].slice();wrong[8]=900;wrong[9]='正常付款';X.utils.sheet_add_aoa(s,[wrong],{origin:-1});
 let r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');assert.equal(r.normalRows[0][17].v,'已提交');
 s.J2={t:'s',v:'正常付款'};r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');
 s.D2={t:'s',v:'MDST-DGFK20990101999'};r=C.build(i);assert.equal(r.reviewRows.length,0,'异类申请不能成为OA冲突候选');
 s.D2={t:'s',v:''};r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交','缺OA也不能跨类别回退');
 s.J3={t:'s',v:'谈判付款'};s.D3={t:'s',v:''};r=C.build(i);assert.equal(r.normalRows[0][17].v,'未申请','正常付款不能借用谈判申请的名称金额');
 s.J2={t:'s',v:' 谈判付款 '};r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');
});
test('付款分类为空或异常须确认，分类决定仅用于本次匹配且原文件不变',()=>{
 for(const type of ['', '其他', {t:'e',v:23,w:'#REF!'}]){
  const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;X.utils.sheet_add_aoa(s,[['付款分类'],[type],['正常付款']],{origin:'J1'});const before=JSON.stringify(i);
  const r=C.build(i),problem=r.issues.find(x=>x.kind==='application-classification');assert(problem);assert.equal(r.stats.pending,1);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');
  for(const choice of ['negotiated','normal','keep']){const confirmed=C.build(i,{[problem.id]:[choice]});assert.equal(confirmed.stats.pending,0);assert.equal(negotiatedValue(confirmed,'是否提交付款申请单'),choice==='negotiated'?'已经提交全部':'未提交');}
  assert.equal(JSON.stringify(i),before);
 }
});

test('谈判提交进度仅有三个最终结果，旧状态兼容且无法判断时确认后才能下载',()=>{
 const states=['已经提交全部','已经提交部分','未提交'];
 for(const [old,expected] of [['已全部提交',states[0]],['已部分提交',states[1]],['未申请',states[2]],...states.map(s=>[s,s])]){
  const i=fixture();set(i,'applications','本次申请付款金额',2,'异常');X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['是否提交付款申请单'],[old]],{origin:'N1'});
  const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),expected);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
 for(const old of ['', '已提交', '已提交付款申请单', '待核对', '人工确认']){
  const i=fixture();set(i,'applications','本次申请付款金额',2,'异常');X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['是否提交付款申请单'],[old]],{origin:'N1'});
  const r=C.build(i),q=r.issues.find(q=>q.kind==='submission-status');assert(q);assert.equal(r.stats.pending,1);assert.deepEqual(q.choices.map(c=>c.value),states);
  for(const state of states){const resolved=C.build(i,{[q.id]:[state]});assert.equal(resolved.stats.pending,0);assert.equal(negotiatedValue(resolved,'是否提交付款申请单'),state);assert.equal(negotiatedValue(resolved,'已付金额'),0);}
  assert.equal(C.build(i,{[q.id]:['已提交']}).stats.pending,1,'旧的第四种状态不可用作确认结果');
 }
});

test('不同付款OA各自判定，无申请为未提交，有足额申请为全部，不合并不同OA并提示同名同额归属',()=>{
 const i=fixture(),n=i.negotiated.book.Sheets['10月谈判付款主表'];set(i,'applications','合同号',2,'MDST-DGFK20990101003');
 const second=X.utils.sheet_to_json(n,{header:1})[1].slice();second[2]='MDST-DGFK20990101003';X.utils.sheet_add_aoa(n,[second],{origin:-1});
 const r=C.build(i);assert.deepEqual(r.negotiated.rows.map((_,j)=>negotiatedValue(r,'是否提交付款申请单',j)),['未提交','已经提交全部']);assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/同一供应商和金额对应多笔付款/);assert.equal(r.stats.pending,0);
});
test('相同流程编码供应商金额为重复提交，只计一次金额，原记录及全部日期保留',()=>{
 for(const [amount,status] of [[100,'已经提交全部'],[60,'已经提交部分'],[0,'已经提交部分']]){
  const i=installmentFixture([{'申请日期':'2026-09-25','本次申请付款金额':amount},{'申请日期':'2026-10-21','本次申请付款金额':amount}]);const before=JSON.stringify(i),r=C.build(i);
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),status);assert.equal(r.reviewRows.length,0);assert.equal(negotiatedValue(r,'提交时间汇总'),`2026-09-25-${amount.toFixed(2)}\n2026-10-21-${amount.toFixed(2)}`);assert.equal(JSON.stringify(i),before);
 }
 const distinct=installmentFixture([{'本次申请付款金额':40},{'本次申请付款金额':60}]);assert.equal(negotiatedValue(C.build(distinct),'是否提交付款申请单'),'已经提交全部');
 const suppliers=installmentFixture([{'供应商全称':'测试甲','本次申请付款金额':50},{'供应商全称':'测试乙','本次申请付款金额':50}]);assert.equal(negotiatedValue(C.build(suppliers),'是否提交付款申请单'),'已经提交全部','供应商不同不能视为重复');
 const missing=installmentFixture([{'合同号':'','本次申请付款金额':100},{'合同号':'','本次申请付款金额':100}]);assert.equal(C.build(missing).reviewRows.length,1,'流程编码缺失不能凭相同金额认定重复');
});

function classifiedFixture(){const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['付款分类'],['谈判付款'],['正常付款']],{origin:'J1'});return i;}
const duplicateReviews=r=>r.reviewRows.filter(row=>row[5].v.startsWith('重复提交：'));
test('两类付款同日及跨日三次重复各只列一组，计额不变并保留来源信息',()=>{
 const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,rows=X.utils.sheet_to_json(s,{header:1});
 for(const n of [1,2]){const same=rows[n].slice(),later=rows[n].slice();later[0]='2026-10-21';later[2]='其他主体';X.utils.sheet_add_aoa(s,[same,later],{origin:-1});}
 const before=JSON.stringify(i),r=C.build(i),reviews=duplicateReviews(r);assert.equal(reviews.length,2);assert.equal(r.reviewRows.length,2);assert.equal(r.stats.pending,0);assert.deepEqual(reviews.map(row=>row[0].v),['谈判付款','正常付款']);
 for(const row of reviews){assert.match(row[5].v,/共 3 次，重复 2 次.*单笔金额 100.00.*重复计入金额 200.00/);assert.equal(row.length,10);assert.equal(row[6].v,'申请.xlsx');assert.equal(row[7].v,'付款申请汇总');assert(Number.isInteger(row[8].v));assert.equal(row[3].t,'n');assert.equal(row[3].v,100);assert.equal(row[2].v,'迈德斯特、其他主体');}
 assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(negotiatedValue(r,'提交时间汇总').split('\n').length,3);assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
});
test('扫描整份申请，未关联款项的重复也列核对但不新增付款明细',()=>{
 const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2].slice();row[1]='未关联供应商';row[3]='MDST-DGFK20991231099';row[8]=12.34;X.utils.sheet_add_aoa(s,[row,row],{origin:-1});
 const r=C.build(i),reviews=duplicateReviews(r);assert.equal(reviews.length,1);assert.equal(reviews[0][1].v,'未关联供应商');assert.equal(reviews[0][8].v,4);assert.match(reviews[0][5].v,/单笔金额 12.34.*重复计入金额 12.34/);assert.equal(r.normalRows.length,1);assert.equal(r.negotiated.rows.length,1);assert.equal(r.stats.pending,0);
});
test('重复组复用标准化名称编码金额，按分精确计算且零金额有效',()=>{
 for(const amount of [0,0.01,1234.56]){const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总;set(i,'applications','本次申请付款金额',3,amount);const row=X.utils.sheet_to_json(s,{header:1})[2].slice();row[1]='　测试供应商乙 ';row[3]=' mdst-dgfk20261001002 ';row[8]='￥'+amount.toFixed(2);row[9]=' 正常付款 ';X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const reviews=duplicateReviews(C.build(i));assert.equal(reviews.length,1);assert.equal(reviews[0][3].v,amount);assert(reviews[0][5].v.includes('重复计入金额 '+(amount*2).toFixed(2)));}
});
test('不同类别流程供应商金额不混合，无可靠分类或完整键不猜测重复',()=>{
 for(const [col,value] of [[9,'正常付款'],[3,'MDST-DGFK20990101099'],[1,'其他供应商'],[8,100.01]]){const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1].slice();row[col]=value;X.utils.sheet_add_aoa(s,[row],{origin:-1});assert.equal(duplicateReviews(C.build(i)).length,0);}
 for(const [col,bad] of [[3,''],[3,'未提交OA'],[3,{t:'e',v:23}],[3,{t:'s',f:'A99'}],[1,''],[1,{t:'e',v:23}],[8,''],[8,'异常'],[8,-1],[8,0.001],[8,{t:'n',f:'A99'}],[9,''],[9,'未知分类'],[9,{t:'e',v:23}]]){
  const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1].slice();row[col]=bad;X.utils.sheet_add_aoa(s,[row,row],{origin:-1});assert.equal(duplicateReviews(C.build(i)).length,0);
 }
 const legacy=fixture(),s=legacy.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});assert.equal(duplicateReviews(C.build(legacy)).length,0,'无付款分类列时不推断重复分类');
});
test('同组关联多条输出明细仍只列一次，日期异常仅保留来源不阻止重复识别',()=>{
 const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2].slice();row[0]='日期待定';X.utils.sheet_add_aoa(s,[row],{origin:-1});const ledger=i.ledger.book.Sheets['对账明细-李奇'],normal=X.utils.sheet_to_json(ledger,{header:1})[2];X.utils.sheet_add_aoa(ledger,[normal],{origin:-1});
 const r=C.build(i);assert.equal(r.normalRows.length,2);assert.equal(duplicateReviews(r).length,1);assert.equal(duplicateReviews(r)[0][6].v,'申请.xlsx');assert.equal(r.stats.pending,0);
});
test('多流程编码按完整集合识别重复，不把部分相同流程的申请合并',()=>{
 const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2].slice();row[3]='MDST-DGFK20990101001、MDST-DGFK20990101002';const reversed=row.slice();reversed[3]='MDST-DGFK20990101002、MDST-DGFK20990101001';const partial=row.slice();partial[3]='MDST-DGFK20990101001';X.utils.sheet_add_aoa(s,[row,reversed,partial],{origin:-1});const reviews=duplicateReviews(C.build(i));assert.equal(reviews.length,1);assert.match(reviews[0][5].v,/共 2 次/);assert.equal(reviews[0].length,10);
});

test('谈判已匹配申请但OA栏仍未提交时列核对，保留原OA和付款公式',()=>{
 for(const placeholder of ['未提交OA','未交OA',' 未提交 oa ','未交ＯＡ']){
  const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,placeholder);
  const s=i.negotiated.book.Sheets['10月谈判付款主表'];s.K2={t:'n',v:100,f:'G2-E2'};
  const before=JSON.stringify(i),r=C.build(i);assert.equal(r.reviewRows.length,1);assert.equal(r.stats.pending,0);
  const row=r.reviewRows[0];assert.equal(row[0].v,'谈判付款');assert.equal(row[4].v,'已提交');assert.match(row[5].v,/已提交付款申请单.*仍为/);
  assert.equal(r.reviewHeaders.includes('候选付款申请'),false);
  assert.deepEqual(row.slice(6,9).map(c=>c.v),['谈判.xlsx','10月谈判付款主表',2]);
  assert.equal(negotiatedValue(r,'对公付款0A流程号'),placeholder);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');
  assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'是否结清'),'否');assert.equal(r.negotiated.rows[0][10].f,'G2-E2');assert.equal(JSON.stringify(i),before);
 }
});
test('未交OA提示每条付款只列一次，可与重复提交提示并存且不重复计金额',()=>{
 const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,'未交OA');
 const s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1].slice();row[0]='2026-10-11';X.utils.sheet_add_aoa(s,[row],{origin:-1});
 const r=C.build(i);assert.equal(r.reviewRows.length,2);assert.equal(duplicateReviews(r).length,1);
 const review=r.reviewRows.find(row=>row[5].v.includes('流程号仍为'));assert.equal(review.length,10);assert.equal(negotiatedValue(r,'提交时间汇总').split('\n').length,2);
 assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(r.stats.pending,0);
});
test('未交OA提示不凭历史提交状态或异类申请触发，缺值、有效OA和归属不清不误报',()=>{
 for(const [field,value] of [['付款分类','正常付款'],['供应商全称','另一供应商'],['本次申请付款金额',99],['本次申请付款金额',''],['本次申请付款金额','异常']]){
  const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,'未提交OA');set(i,'applications',field,2,value);
  X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['是否提交付款申请单'],['已经提交全部']],{origin:'N1'});
  const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(r.reviewRows.length,0);
 }
 for(const placeholder of ['', 'MDST-DGFK20261001001', 'MDST-DGFK20990101099', {t:'e',v:23}]){
  const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,placeholder);assert(!C.build(i).reviewRows.some(row=>row[5].v.includes('流程号仍为')));
 }
 const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,'未交OA');const s=i.negotiated.book.Sheets['10月谈判付款主表'];X.utils.sheet_add_aoa(s,[X.utils.sheet_to_json(s,{header:1})[1]],{origin:-1});
 assert(!C.build(i).reviewRows.some(row=>row[5].v.includes('流程号仍为')));
});

// Isolate legacy ledger/application rules when the OA export has no supplier-amount match.
function withoutMatchingOA(input){set(input,'oa','付款金额',2,99);return input;}

test('OA补齐同时覆盖两类已有输出记录，按分匹配、保留原财务公式并使用新OA重算提交状态',()=>{
 for(const old of ['未提交OA','未交OA','未提交','MDST-DGFK20990101999','', {t:'s',f:'A99'}, {t:'e',v:23,w:'#REF!'}]){
  const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,old);set(i,'ledger','对公付款0A流程号',3,'未提交');
  const source=i.negotiated.book.Sheets['10月谈判付款主表'];source.K2={t:'n',v:100,f:'G2-E2'};
  const before=JSON.stringify(i),r=C.build(i),reviews=r.reviewRows.filter(row=>row[5].v.startsWith('OA号复核：'));
  assert.equal(reviews.length,2);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');
  assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][23].v,r.normalRows[0][4].v);assert.equal(r.normalRows[0][17].v,'已提交');
  assert.match(reviews[0][5].v,/原对公付款0A流程号.*已更新为“MDST-DGFK20261001001”/);assert.match(reviews[0][5].v,/OA 来源：OA.xlsx\/Sheet0!2/);
  assert.equal(reviews[0][8].v,2);assert.equal(r.stats.pending,0);assert.deepEqual(r.negotiated.rows[0][10],source.K2);assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'是否结清'),'否');assert.equal(JSON.stringify(i),before);
 }
 for(const amount of [0,0.01,5736.1]){const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未交OA');set(i,'negotiated','申请付款金额',2,amount);set(i,'oa','付款金额',2,'￥'+amount.toFixed(2));const r=C.build(i);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');}
});
test('OA补齐不把流程表记录当成付款申请，且不纳入原OA为空的正常付款',()=>{
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未提交OA');set(i,'ledger','对公付款0A流程号',3,'');set(i,'applications','付款分类',2,'正常付款');
 const r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.excludedNormalRows,1);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(r.reviewRows.length,1);assert.equal(r.reviewRows[0][4].v,'未提交');
});
test('多个OA候选或同编号身份冲突只列复核，不覆盖OA，重复导出行可唯一匹配',()=>{
 for(const [id,supplier,amount] of [['MDST-DGFK20261001999','测试供应商甲',100],['MDST-DGFK20261001001','另一收款单位',100],['MDST-DGFK20261001001','测试供应商甲',99]]){
  const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未提交OA');X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[[id,'财务审核','测试人员','未归档','李奇',supplier,amount]],{origin:-1});const r=C.build(i);
  assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未提交OA');const reviews=r.reviewRows.filter(row=>row[5].v.startsWith('OA号复核'));assert.equal(reviews.length,1);assert.match(reviews[0][5].v,/保留.*核对/);
 }
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未交OA');const s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[X.utils.sheet_to_json(s,{header:1})[1]],{origin:-1});const r=C.build(i);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/Sheet0!2/);assert.match(r.reviewRows[0][5].v,/Sheet0!4/);
});
test('OA自动补齐拒绝无效匹配键或无效流程编号，原表合并不拆分',()=>{
 for(const [role,field,v] of [['negotiated','供应商',''],['negotiated','申请付款金额',''],['negotiated','申请付款金额',-1],['negotiated','申请付款金额',0.001],['negotiated','申请付款金额',{t:'n',f:'A99'}],['oa','付款金额',''],['oa','收款单位',{t:'e',v:23}],['oa','流程编号','未提交OA'],['oa','流程编号',''],['oa','流程编号',{t:'e',v:23}],['oa','流程编号','MDST-001、MDST-002']]){
  const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未交OA');set(i,role,field,2,v);const r=C.build(i);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未交OA');
 }
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未交OA');const s=i.negotiated.book.Sheets['10月谈判付款主表'];s['!merges']=[X.utils.decode_range('C2:D2')];const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未交OA');assert(r.reviewRows.some(row=>/OA号位于合并单元格/.test(row[5].v)));assert.equal(JSON.stringify(i),before);
});

test('原OA已准确对应候选时不因同名同额其他流程误改或重复提示',()=>{
 const i=classifiedFixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001999','财务审核','另一人员','未归档','徐英','测试供应商甲',100]],{origin:-1});
 const r=C.build(i);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(r.reviewRows.length,0);
 set(i,'negotiated','对公付款0A流程号',2,' mdst-dgfk20261001001 ');assert.equal(C.build(i).reviewRows.length,0);
});

function plannedOAFixture(){
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未申请OA');set(i,'negotiated','申请付款金额',2,10);
 X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['2026-11-25'],[10]],{origin:'R1'});return i;
}
const plannedReviews=r=>r.reviewRows.filter(row=>row[5].v.startsWith('分批付款OA待核对'));
test('分批金额按台账日期列及OA总额交叉验证，仅复核不补齐，不以付款计划认定提交',()=>{
 const i=plannedOAFixture(),before=JSON.stringify(i),r=C.build(i),reviews=plannedReviews(r);assert.equal(reviews.length,1);
 assert.match(reviews[0][5].v,/谈判金额 10.00.*仅核对，不自动补齐/);assert.match(reviews[0][5].v,/MDST-DGFK20261001001；申请总金额 100.00；计划日期 2026-11-25（R2）/);assert.match(reviews[0][5].v,/资金.xlsx\/对账明细-李奇!2/);assert.match(reviews[0][5].v,/OA 来源：OA.xlsx\/Sheet0!2/);
 assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未申请OA');assert.equal(negotiatedValue(r,'流程编号'),'');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未提交');assert.equal(r.stats.pending,0);assert.equal(r.stats.negotiatedRows,1);assert.equal(JSON.stringify(i),before);
 assert.equal(r.reviewHeaders.length,10);assert(!r.reviewHeaders.includes('候选付款申请'));assert(r.reviewRows.every(row=>row.length===10));
});
test('分批复核必须同时满足供应商、日期列金额、台账OA号及OA申请总额，异常值不猜测',()=>{
 for(const [role,field,val] of [['oa','流程编号','MDST-DGFK20990101999'],['oa','付款金额',99],['oa','收款单位','其他供应商'],['ledger','对公付款0A流程号','未申请OA'],['ledger','供应商名称','其他供应商'],['ledger','申请付款金额',''],['negotiated','申请付款金额',0],['negotiated','申请付款金额',10.001]]){
  const i=plannedOAFixture();set(i,role,field,2,val);assert.equal(plannedReviews(C.build(i)).length,0);
 }
 for(const plan of ['',0,-10,10.001,{t:'n',f:'A99'},{t:'e',v:23}]){const i=plannedOAFixture();set(i,'ledger','2026-11-25',2,plan);assert.equal(plannedReviews(C.build(i)).length,0);}
 for(const header of ['付款金额','11月25日','2026-02-30',{t:'e',v:23}]){const i=plannedOAFixture();i.ledger.book.Sheets['对账明细-李奇'].R1=typeof header==='object'?header:{t:'s',v:header};assert.equal(plannedReviews(C.build(i)).length,0);}
});
test('分批复核多日期及OA重复行只列一条，多个可验证流程都保留供核对，原有效OA不重复提示',()=>{
 const i=plannedOAFixture(),l=i.ledger.book.Sheets['对账明细-李奇'],o=i.oa.book.Sheets.Sheet0;
 X.utils.sheet_add_aoa(l,[['2026-12-25'],[10]],{origin:'S1'});const repeated=X.utils.sheet_to_json(l,{header:1})[1].slice();repeated[3]='MDST-DGFK20990101001';X.utils.sheet_add_aoa(l,[repeated],{origin:-1});
 const flow=X.utils.sheet_to_json(o,{header:1})[1].slice();X.utils.sheet_add_aoa(o,[flow],{origin:-1});flow[0]='MDST-DGFK20990101001';X.utils.sheet_add_aoa(o,[flow],{origin:-1});
 const r=C.build(i),reviews=plannedReviews(r);assert.equal(reviews.length,1);assert.match(reviews[0][5].v,/2026-11-25.*2026-12-25/);assert.match(reviews[0][5].v,/MDST-DGFK20261001001/);assert.match(reviews[0][5].v,/MDST-DGFK20990101001/);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未申请OA');
 set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20261001001');assert.equal(plannedReviews(C.build(i)).length,0);
});

test('台账名称金额相同且OA号金额已验证，OA收款名称不同仍列复核，禁止自动替换',()=>{
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未提交OA');set(i,'oa','收款单位',2,'测试供应商甲器械公司');
 const before=JSON.stringify(i),r=C.build(i),review=r.reviewRows.find(row=>row[5].v.startsWith('OA名称差异待核对'));assert(review);assert.match(review[5].v,/当前对公付款OA号：未提交OA\n查到的流程编码：MDST-DGFK20261001001/);assert.match(review[5].v,/流程来源：OA.xlsx \/ Sheet0 \/ 第2行/);assert.match(review[5].v,/台账供应商“测试供应商甲”/);assert.match(review[5].v,/OA收款单位“测试供应商甲器械公司”/);assert.match(review[5].v,/100.00/);assert.match(review[5].v,/资金.xlsx\/对账明细-李奇!2/);assert.match(review[5].v,/OA.xlsx\/Sheet0!2/);
 assert.equal(negotiatedValue(r,'对公付款0A流程号'),'未提交OA');assert.equal(negotiatedValue(r,'流程编号'),'');assert.equal(r.stats.pending,0);assert.equal(r.reviewHeaders.length,10);assert.equal(JSON.stringify(i),before);
});
test('名称差异复核不能只凭同金额或近似名称，需要资金台账OA号与对公付款金额一致',()=>{
 for(const [role,field,val] of [['ledger','供应商名称','其他供应商'],['ledger','申请付款金额',99],['ledger','对公付款0A流程号',''],['ledger','对公付款0A流程号','未申请OA'],['oa','流程编号','MDST-DGFK20990101999'],['oa','付款金额',99],['oa','付款金额',''],['oa','付款金额',{t:'e',v:23}],['oa','收款单位',''],['oa','收款单位',{t:'e',v:23}]]){
  const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'未提交OA');set(i,'oa','收款单位',2,'不同收款名称');set(i,role,field,2,val);
  assert(!C.build(i).reviewRows.some(row=>row[5].v.startsWith('OA名称差异待核对')));
 }
 const i=classifiedFixture();set(i,'oa','收款单位',2,'不同收款名称');assert(!C.build(i).reviewRows.some(row=>row[5].v.startsWith('OA名称差异待核对')),'已有对应OA号时不重复提示更新');
});

test('正常付款已有输出记录也可通过其他台账行发现名称差异，空OA正常记录仍不纳入',()=>{
 const i=classifiedFixture(),s=i.ledger.book.Sheets['对账明细-李奇'],reference=X.utils.sheet_to_json(s,{header:1})[2].slice();
 set(i,'ledger','对公付款0A流程号',3,'未提交OA');X.utils.sheet_add_aoa(s,[reference],{origin:-1});set(i,'oa','收款单位',3,'测试供应商乙器械公司');const r=C.build(i),reviews=r.reviewRows.filter(row=>row[5].v.startsWith('OA名称差异待核对'));
 assert.equal(reviews.length,1);assert.equal(reviews[0][0].v,'正常付款');assert.equal(reviews[0][8].v,3);assert.match(reviews[0][5].v,/MDST-DGFK20261001002/);assert.equal(r.normalRows[0][4].v,'未提交OA');assert.equal(r.normalRows[1][4].v,'MDST-DGFK20261001002');
 set(i,'ledger','对公付款0A流程号',3,'');const excluded=C.build(i);assert.equal(excluded.normalRows.length,1);assert(!excluded.reviewRows.some(row=>row[5].v.startsWith('OA名称差异待核对')));
});

test('通用OA复核覆盖两类付款和各种未提交写法，不限定供应商或流程号',()=>{
 for(const role of ['negotiated','ledger'])for(const old of ['未提交','未提交OA','未提交OAD','未申请OA','未交OA','MDST-DGFK20991231999']){
  const i=classifiedFixture(),row=role==='ledger'?3:2,id=role==='ledger'?'MDST-DGFK20261001002':'MDST-DGFK20261001001';set(i,role,'对公付款0A流程号',row,old);
  const before=JSON.stringify(i),r=C.build(i),review=r.reviewRows.find(line=>line[5].v.startsWith('OA号复核'));assert(review);assert(review[5].v.includes('“'+old+'”'));assert(review[5].v.includes(id));assert.equal(review[0].v,role==='ledger'?'正常付款':'谈判付款');assert.equal(role==='ledger'?r.normalRows[0][4].v:negotiatedValue(r,'对公付款0A流程号'),id);assert.equal(JSON.stringify(i),before);
 }
});

const missingFlowReviews=r=>r.reviewRows.filter(row=>row[5].v.includes('OA流程表未找到：'));
test('两类付款已有OA号未出现在上传流程表时列核对，保留状态金额公式且不阻止下载',()=>{
 const i=classifiedFixture();set(i,'oa','流程编号',2,'MDST-DGFK20990101001');set(i,'oa','收款单位',2,'其他供应商甲');set(i,'oa','流程编号',3,'MDST-DGFK20990101002');set(i,'oa','收款单位',3,'其他供应商乙');
 const s=i.negotiated.book.Sheets['10月谈判付款主表'];s.K2={t:'n',v:100,f:'G2-E2'};const before=JSON.stringify(i),r=C.build(i),reviews=missingFlowReviews(r);
 assert.equal(reviews.length,2);assert.deepEqual(new Set(reviews.map(row=>row[0].v)),new Set(['谈判付款','正常付款']));
 for(const row of reviews){assert.match(row[5].v,/核对范围：OA.xlsx \/ Sheet0（流程编号列）/);assert(row[5].v.endsWith('需要核对：'+row[0].v+'sheet的对公付款OA号是否正确'));assert(!/原对公付款OA号|当前对公付款OA号|OA 来源：/.test(row[5].v));assert.equal(C.formatReviewText(row[5].v).split('\n').length,3);assert(row[6].v);assert(row[7].v);assert(row[8].v);}
 assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已经提交全部');assert.equal(r.normalRows[0][4].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][17].v,'已提交');assert.equal(r.negotiated.rows[0][10].f,'G2-E2');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('OA覆盖逐号核对，大小写空格标准化，已存在但金额不同不报编号缺失',()=>{
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,' mdst-dgfk20261001001 、 MDST-DGFK20990101999、MDST-DGFK20990101999');set(i,'oa','付款金额',2,99);
 const r=C.build(i),reviews=missingFlowReviews(r);assert.equal(reviews.length,1);assert.match(reviews[0][5].v,/OA流程表未找到：MDST-DGFK20990101999。/);assert(!reviews[0][5].v.includes('OA流程表未找到：MDST-DGFK20261001001'));
 set(i,'negotiated','对公付款0A流程号',2,' mdst-dgfk20261001001 ');assert.equal(missingFlowReviews(C.build(i)).length,0);
});
test('OA未找到核对不依赖供应商金额，空值异常值保留，零金额仍检查',()=>{
 for(const amount of ['',0,'异常',{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=classifiedFixture();set(i,'negotiated','供应商',2,'');set(i,'negotiated','申请付款金额',2,amount);set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20990101999');
  const r=C.build(i),reviews=missingFlowReviews(r);assert.equal(reviews.length,1);assert.match(reviews[0][5].v,/MDST-DGFK20990101999/);assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20990101999');
  assert.deepEqual(reviews[0][3],typeof amount==='object'?{t:'s',v:amount.t==='e'?amount.w:''}:typeof amount==='number'?{t:'n',v:0,z:'#,##0.00'}:{t:'s',v:amount});
 }
});
test('原编号未找到但已唯一修正时合并到同一核对说明，不重复列付款',()=>{
 const i=classifiedFixture();set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20990101999');const r=C.build(i);assert.equal(r.reviewRows.length,1);assert.equal(missingFlowReviews(r).length,1);assert.match(r.reviewRows[0][5].v,/已更新为“MDST-DGFK20261001001”/);assert.match(r.reviewRows[0][5].v,/需要核对：谈判付款sheet的对公付款OA号是否正确/);
});

test('核对说明逐条编号且保留原标点、金额、编号和来源，空值不生成条目',()=>{
 const text='当前OA号：未提交OA\n流程编码：MDST-DGFK20261021013；金额：79500.00。\n来源：文件.xlsx / Sheet0 / 第6行';
 assert.equal(C.formatReviewText(text),'1、当前OA号：未提交OA\n2、流程编码：MDST-DGFK20261021013；\n3、金额：79500.00。\n4、来源：文件.xlsx / Sheet0 / 第6行');
 assert.equal(C.formatReviewText(''), '');assert.equal(C.formatReviewText(null), '');assert.equal(C.formatReviewText('原因\r\n\r\n来源'), '1、原因\n2、来源');
});

test('正常付款跟单员取五个人的台账sheet，兼容日期后缀且不采用OA申请人或人工同名字段',()=>{
 for(const [index,person] of C.PEOPLE.entries()){
  const i=classifiedFixture(),b=i.ledger.book,rows=X.utils.sheet_to_json(b.Sheets['对账明细-李奇'],{header:1}),name='对账明细-'+person;
  b.Sheets['对账明细-李奇']=X.utils.aoa_to_sheet(rows.slice(0,2));X.utils.sheet_add_aoa(b.Sheets[name],[rows[2]],{origin:-1});
  const renamed=name+(index%2?'1026':'20261026');b.Sheets[renamed]=b.Sheets[name];delete b.Sheets[name];b.SheetNames[b.SheetNames.indexOf(name)]=renamed;
  X.utils.sheet_add_aoa(b.Sheets[renamed],[['跟单员'],['手填其他人'],['手填其他人']],{origin:'R1'});set(i,'oa','申请人',3,'OA申请人');
  const before=JSON.stringify(i),r=C.build(i);assert.equal(r.normalHeaders[1],'跟单员');assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][1].v,person);assert.equal(r.normalRows[0][8].v,100);assert.equal(JSON.stringify(i),before);
 }
});
test('核对跟单员覆盖OA缺失和付款申请核对，保留来源列及谈判表原结构',()=>{
 const i=classifiedFixture();set(i,'oa','流程编号',2,'MDST-DGFK20990101001');set(i,'oa','收款单位',2,'其他甲');set(i,'oa','流程编号',3,'MDST-DGFK20990101002');set(i,'oa','收款单位',3,'其他乙');
 const before=JSON.stringify(i),r=C.build(i);assert.equal(r.reviewHeaders.at(-1),'跟单员');assert(!r.negotiated.headers.includes('跟单员'));assert.equal(missingFlowReviews(r).length,2);for(const row of missingFlowReviews(r)){assert.equal(row.at(-1).v,'李奇');assert(row[6].v);assert(row[7].v);assert(row[8].v);}assert.equal(JSON.stringify(i),before);
 const j=withoutMatchingOA(classifiedFixture());set(j,'negotiated','对公付款0A流程号',2,'未交OA');assert.equal(C.build(j).reviewRows.find(row=>row[5].v.includes('流程号仍为')).at(-1).v,'李奇');
});
test('核对跟单员多台账责任人冲突或没有匹配时留空，原谈判表已填跟单员优先',()=>{
 const i=withoutMatchingOA(classifiedFixture());set(i,'negotiated','对公付款0A流程号',2,'未交OA');const b=i.ledger.book,row=X.utils.sheet_to_json(b.Sheets['对账明细-李奇'],{header:1})[1];X.utils.sheet_add_aoa(b.Sheets['对账明细-徐英'],[row],{origin:-1});
 const review=r=>r.reviewRows.find(row=>row[5].v.includes('流程号仍为'));assert.equal(review(C.build(i)).at(-1).v,'');
 X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['跟单员'],['邢芳芳']],{origin:'N1'});assert.equal(review(C.build(i)).at(-1).v,'邢芳芳');
 const j=classifiedFixture(),s=j.applications.book.Sheets.付款申请汇总,app=X.utils.sheet_to_json(s,{header:1})[2].slice();app[1]='未关联供应商';X.utils.sheet_add_aoa(s,[app,app],{origin:-1});assert.equal(duplicateReviews(C.build(j))[0].at(-1).v,'');
});
test('重复提交核对按可靠台账识别跟单员，主体不一致不猜测',()=>{
 const i=classifiedFixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2].slice();X.utils.sheet_add_aoa(s,[row],{origin:-1});assert.equal(duplicateReviews(C.build(i))[0].at(-1).v,'李奇');
 row[2]='其他主体';X.utils.sheet_add_aoa(s,[row],{origin:-1});assert.equal(duplicateReviews(C.build(i))[0].at(-1).v,'');
});
