const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const H=['所属月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付时间','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单'];
const book=(name,rows)=>{const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet(rows),name);return b;};
function fixture(extra={},owner='李奇',reverse=false){
 const values={'所属月份':'2026年10月','供应商名称':'合成供应商','申请付款金额':100,'是否谈判付款':'否','账期':30,...extra};
 const headers=[...H,...Object.keys(values).filter(h=>!H.includes(h))];if(reverse)headers.reverse();
 const ledger=book('对账明细-'+owner,[headers,headers.map(h=>values[h]??'')]);
 for(const p of C.PEOPLE.filter(p=>p!==owner))X.utils.book_append_sheet(ledger,X.utils.aoa_to_sheet([H]),'对账明细-'+p);
 return {ledger:{book:ledger,name:'本次台账.xlsx'},negotiated:{book:book('谈判付款',[['供应商名称','申请付款金额','预估时间','账期']]),name:'谈判.xlsx'},applications:{book:book('付款申请汇总',[['供应商名称','申请付款金额']]),name:'申请.xlsx'},oa:{book:book('OA',[['流程编号','收款单位','付款金额','当前节点','未操作者','流程状态','申请人']]),name:'OA.xlsx'}};
}
const display=c=>c?.z?X.SSF.format(c.z,c.v):c?.v;
const result=i=>{const built=C.build(i);return {built,row:Object.fromEntries(C.NORMAL_HEADERS.map((h,c)=>[h,built.normalRows[0][c]]))};};
test('应付与预估均空必须人工补完整日期，不能保留空白或按账期自动放行',()=>{
 for(const extra of [{},{'预估时间':''},{'预估付款时间':'　 '},{'预估付款月份':''}]){
  const input=fixture(extra),before=JSON.stringify(input),built=C.build(input),pending=built.issues.find(i=>i.kind==='payment-date');
  assert(pending,'包括缺少预估列的记录');assert.equal(pending.resolved,false);assert.equal(pending.inputType,'date');assert.equal(built.stats.pending,1);
  assert.equal(built.normalRows[0][C.NORMAL_HEADERS.indexOf('预估时间')].v,'');
  for(const value of ['', 'keep', '2026-11', '2026-02-30', '2026-13-01', '1899-12-31', '未对账待定'])assert.equal(C.build(input,{[pending.id]:[value]}).stats.pending,1,value);
  const resolved=C.build(input,{[pending.id]:['2026-12-08']}),row=Object.fromEntries(C.NORMAL_HEADERS.map((h,c)=>[h,resolved.normalRows[0][c]]));
  assert.equal(resolved.stats.pending,0);assert.equal(row.应付时间.v,'');assert.equal(display(row.预估时间),'2026年12月8日');assert.equal(row['申请付款的日期'].v,'2026/11/30周');assert.equal(JSON.stringify(input),before);
 }
});
test('同供应商同金额两行分别补日期，改日期后排序、清空后重新拦截，原记录不改',()=>{
 const input=fixture({'预估时间':''}),sheet=input.ledger.book.Sheets['对账明细-李奇'];
 X.utils.sheet_add_aoa(sheet,[X.utils.sheet_to_json(sheet,{header:1})[1]],{origin:-1});
 const before=JSON.stringify(input),issues=C.build(input).issues.filter(i=>i.kind==='payment-date'),[first,second]=issues;
 assert.equal(issues.length,2);assert.notEqual(first.id,second.id);
 const decisions={[first.id]:['2026-12-15'],[second.id]:['2026-11-20']};
 assert.equal(C.build(input,{[first.id]:decisions[first.id]}).stats.pending,1);
 const complete=C.build(input,decisions),estimate=C.NORMAL_HEADERS.indexOf('预估时间');
 assert.equal(complete.stats.pending,0);assert.deepEqual(complete.normalRows.map(r=>display(r[estimate])),['2026年11月20日','2026年12月15日']);
 assert.equal(C.build(input,{...decisions,[first.id]:['']}).stats.pending,1);assert.equal(JSON.stringify(input),before);
});
test('导出函数再次拦截未补填日期，不依赖按钮禁用',async()=>{
 const vm=require('node:vm'),fs=require('node:fs'),context={window:{}};
 vm.runInNewContext(fs.readFileSync(require.resolve('../funds-export.js'),'utf8'),context);
 await assert.rejects(()=>context.window.FundsExport(new Uint8Array(),C.build(fixture())),/逐条补填并确认付款日期/);
});
test('正常付款月底应付日调整为同月25日，申请周使用调整日期且源数据不变',()=>{
 const serial=(Date.UTC(2026,9,31)-Date.UTC(1899,11,30))/86400000;
 for(const [raw,expected,week,date1904] of [
  ['2026-10-26','2026年10月25日','2026/10/12周'],['2026/11/28 23:59:59','2026年11月25日','2026/11/16周'],
  ['2028年2月29日','2028年2月25日','2028/02/14周'],['2026-04-30','2026年4月25日','2026/04/13周'],['2026-12-31','2026年12月25日','2026/12/14周'],
  ['2026-11-24','2026年11月24日','2026/11/16周'],['2026-11-25','2026年11月25日','2026/11/16周'],['2026-01-01','2026年1月1日','2025/12/22周'],
  [{t:'n',v:serial+.75,z:'yyyy-mm-dd'},'2026年10月25日','2026/10/12周'],
  [{t:'n',v:serial-1462,f:'DATE(2026,10,31)',z:'yyyy-mm-dd'},'2026年10月25日','2026/10/12周',true]
 ]){
  const i=fixture({'应付时间':raw,'预估时间':'2027-01-31'});if(date1904)i.ledger.book.Workbook={WBProps:{date1904:true}};
  const before=JSON.stringify(i),{row}=result(i);
  assert.equal(display(row.应付时间),expected,JSON.stringify(raw));assert.equal(row['申请付款的日期'].v,week);
  assert.equal(row.预估时间.v,'');assert(!row.应付时间.f,'输出日期不沿用源公式');assert.equal(JSON.stringify(i),before);
 }
 for(const raw of ['2026-11','2026年11月','2026-02-30','月底协商',{t:'e',v:23,w:'#REF!'},{t:'n',f:'DATE(2026,11,30)'}]){
  const i=fixture({'应付时间':raw,'预估时间':'2027-01-31'}),before=JSON.stringify(i),{row}=result(i);
  assert.equal(row['申请付款的日期'].v,'日期待确认');assert.equal(JSON.stringify(i),before);
  if(typeof raw==='string')assert.equal(row.应付时间.v,raw,'空白或不完整/无效日期不补值');
 }
});
test('正常付款具体预估按同月25日封顶计算上一周周一，预估显示及应付空白不变',()=>{
 const serial=(Date.UTC(2026,9,31)-Date.UTC(1899,11,30))/86400000;
 for(const [raw,expected,week,date1904] of [
  ['2026-11-18','2026年11月18日','2026/11/09周'],['2026-11-24','2026年11月24日','2026/11/16周'],
  ['2026-11-25','2026年11月25日','2026/11/16周'],['2026-11-26','2026年11月26日','2026/11/16周'],
  ['2026/11/30 23:59:59','2026年11月30日','2026/11/16周'],['2028年2月29日','2028年2月29日','2028/02/14周'],
  ['2027-01-01','2027年1月1日','2026/12/21周'],['2027-01-31','2027年1月31日','2027/01/18周'],
  [{t:'n',v:serial+.75,z:'yyyy-mm-dd'},'2026年10月31日','2026/10/12周'],
  [{t:'n',v:serial-1462,f:'DATE(2026,10,31)',z:'yyyy-mm-dd'},'2026年10月31日','2026/10/12周',true]
 ])for(const header of ['预估时间','预估付款时间','预估付款月份']){
  const i=fixture({[header]:raw});if(date1904)i.ledger.book.Workbook={WBProps:{date1904:true}};
  const before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(display(row.预估时间),expected,JSON.stringify(raw));assert.equal(row['申请付款的日期'].v,week);
  assert.equal(row.应付时间.v,'');assert(!row.预估时间.f);assert.equal(built.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('正常付款按调整后的申请周稳定排序，同周保持台账源顺序',()=>{
 const i=fixture({'应付时间':'2026-10-31','供应商名称':'月底先录入','预估时间':''}),sheet=i.ledger.book.Sheets['对账明细-李奇'];
 const rows=X.utils.sheet_to_json(sheet,{header:1}),headers=rows[0],base=rows[1];
 for(const [name,due] of [['同周25日','2026-10-25'],['同周20日','2026-10-20'],['更早一周','2026-10-18'],['次月','2026-11-02'],['日期待确认','']]){
  const row=base.slice();row[headers.indexOf('供应商名称')]=name;row[headers.indexOf('应付时间')]=due;X.utils.sheet_add_aoa(sheet,[row],{origin:-1});
 }
 const before=JSON.stringify(i),built=C.build(i),name=C.NORMAL_HEADERS.indexOf('供应商名称'),due=C.NORMAL_HEADERS.indexOf('应付时间');
 assert.deepEqual(built.normalRows.map(row=>row[name].v),['更早一周','月底先录入','同周25日','同周20日','次月','日期待确认']);
 assert.deepEqual(built.normalRows.map(row=>row.at(-1).v),['2026/10/05周','2026/10/12周','2026/10/12周','2026/10/12周','2026/10/26周','日期待确认']);
 assert.equal(display(built.normalRows[1][due]),'2026年10月25日');assert.equal(JSON.stringify(i),before);
});
test('正常付款日期互斥，应付优先，预估仅从本行表头读取，谈判手工列不受影响',()=>{
 for(const [index,owner] of C.PEOPLE.entries())for(const header of ['预估时间','预估付款时间','预估付款月份']){
  const i=fixture({'应付时间':'2026-11-18',[header]:'2027-01-01'},owner,index%2===0),before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(display(row.应付时间),'2026年11月18日');assert.equal(row.预估时间.v,'');assert.equal(built.stats.pending,0);assert.equal(JSON.stringify(i),before);
  assert(!built.negotiated.managedColumns.includes(2));assert(!built.negotiated.managedColumns.includes(3));assert(!built.negotiated.updates.some(c=>/^C|^D/.test(c.address)));
 }
 for(const header of ['预估时间','预估付款时间','预估付款月份']){
  const {row,built}=result(fixture({[header]:{t:'n',v:46344,f:'DATE(2026,11,18)',z:'yyyy-mm-dd'}}));
  assert.equal(row.应付时间.v,'');assert.equal(display(row.预估时间),'2026年11月18日');assert(!row.预估时间.f);assert.equal(built.stats.pending,0);assert.equal(row['申请付款的日期'].v,'2026/11/09周');
 }
});
test('人工补填跨月跨年及闰日，1904来源不偏移；原账期不再自动产生日期',()=>{
 for(const date1904 of [false,true])for(const [day,expected,week] of [
  ['2026-12-31','2026年12月31日','2026/12/14周'],['2027-01-01','2027年1月1日','2026/12/21周'],['2028-02-29','2028年2月29日','2028/02/14周']
 ]){
  const input=fixture({'申请付款金额':0,'是否结清':'是'});input.ledger.book.Workbook={WBProps:{date1904}};
  const before=JSON.stringify(input),id=C.build(input).issues.find(i=>i.kind==='payment-date').id,built=C.build(input,{[id]:[day]});
  const row=Object.fromEntries(C.NORMAL_HEADERS.map((h,c)=>[h,built.normalRows[0][c]]));
  assert.equal(display(row.预估时间),expected);assert.equal(row['申请付款的日期'].v,week);assert.equal(row.应付时间.v,'');assert.equal(row.本次申请付款金额.v,0);assert.equal(built.stats.pending,0);assert.equal(JSON.stringify(input),before);
  assert.equal(C.build(input,{[id]:['2027-02-29']}).stats.pending,1);
 }
});
test('已有预估列空白不推算，年月保持精度，错误与不确定账期保留核对来源',()=>{
 for(const month of ['2026年12月','2026-12',{t:'n',v:46371,z:'yyyy年m月'}]){
  const {row}=result(fixture({'预估付款时间':month}));assert.equal(display(row.预估时间),'2026年12月');
 }
 const blank=result(fixture({'预估时间':''}));assert.equal(blank.row.预估时间.v,'');assert.equal(blank.built.issues[0].kind,'payment-date');
 for(const terms of ['',-30,'30至60天','30/60天','30天或60天','预付50%','先款后货','7个工作日','约30天']){
  const {built,row}=result(fixture({'账期':terms}));assert.equal(row.预估时间.v,'');assert.equal(built.stats.pending,1);assert.equal(built.issues[0].kind,'payment-date');assert.deepEqual(built.issues[0].source,{file:'本次台账.xlsx',sheet:'对账明细-李奇',row:2});
 }
 for(const bad of [{t:'e',v:23,w:'#REF!'}, {t:'n',f:'A2+30'}, {t:'s',v:'',f:'A2+30'}])for(const h of ['预估时间']){
  const {built}=result(fixture({[h]:bad}));assert(built.issues.some(q=>q.id.startsWith('normal-estimate:')&&q.kind==='formula'));assert(built.stats.pending>0);
 }
 for(const month of ['', '11月', '2026-13']){
  const {built,row}=result(fixture({'所属月份':month}));assert.equal(row.预估时间.v,'');assert(built.issues.some(q=>q.kind==='payment-date'));
 }
 const source=result(fixture({'应付时间':'月底协商','预估时间':'2026-11-20'}));assert.equal(source.row.应付时间.v,'月底协商');assert.equal(source.row.预估时间.v,'');
 const bad=result(fixture({'预估时间':'2026-02-30'}));assert.equal(bad.row.预估时间.v,'2026-02-30');assert.equal(bad.built.stats.pending,1);
});
test('正常付款预估仅有年月时按该月25日计算上一周周一，保留月份显示及源数据',()=>{
 const serial=(Date.UTC(2026,10,1)-Date.UTC(1899,11,30))/86400000;
 for(const [raw,month,week,date1904] of [
  ['2026年11月','2026年11月','2026/11/16周'],['2026-11','2026年11月','2026/11/16周'],['2026/11','2026年11月','2026/11/16周'],
  [{t:'n',v:serial,z:'yyyy-mm'},'2026年11月','2026/11/16周'],
  [{t:'n',v:serial-1462,f:'DATE(2026,11,1)',z:'yyyy年m月'},'2026年11月','2026/11/16周',true],
  ['2026-12','2026年12月','2026/12/14周'],['2027-01','2027年1月','2027/01/18周'],
  ['2028-02','2028年2月','2028/02/14周'],['2027-02','2027年2月','2027/02/15周']
 ])for(const header of ['预估时间','预估付款时间','预估付款月份']){
  const i=fixture({[header]:raw});if(date1904)i.ledger.book.Workbook={WBProps:{date1904:true}};
  const before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(row['申请付款的日期'].v,week,JSON.stringify(raw));assert.equal(display(row.预估时间),month);
  assert.equal(row.应付时间.v,'');assert(!row.预估时间.f);assert.equal(built.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
 for(const [due,week] of [['2026-11-18','2026/11/09周'],['2026-11-30','2026/11/16周'],['月底协商','日期待确认']]){
  const i=fixture({'应付时间':due,'预估时间':'2027-01'}),before=JSON.stringify(i),{row}=result(i);
  assert.equal(row['申请付款的日期'].v,week);assert.equal(row.预估时间.v,'','非空应付时间仍优先，不用月份预估覆盖');assert.equal(JSON.stringify(i),before);
 }
});
test('空白、非法预估日期或月份和无效公式缓存不产生申请日期，仍保留核对提示',()=>{
 for(const raw of ['', '　 ', '2026-00', '2026-13', '11月', '2026-02-30', '月底协商', {t:'e',v:23,w:'#REF!'}, {t:'n',f:'DATE(2026,11,1)',z:'yyyy-mm'}, {t:'n',f:'DATE(2026,11,18)',z:'yyyy-mm-dd'}, {t:'s',v:'',f:'TEXT(A2,"yyyy-mm")'}]){
  const i=fixture({'预估时间':raw}),before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(row['申请付款的日期'].v,'日期待确认');assert.equal(built.stats.pending,1);
  assert(built.issues.some(q=>q.id.startsWith('normal-estimate:')||q.kind==='payment-date'));assert.equal(JSON.stringify(i),before);
  if(raw==='')assert.equal(row.预估时间.v,'','已有预估列空白仍不推算');
 }
});
test('月份及具体预估与有效应付日期共同按申请周排序，同周和待确认组均保持台账原序',()=>{
 const i=fixture({'供应商名称':'待确认先录','预估时间':'2026-13'}),sheet=i.ledger.book.Sheets['对账明细-李奇'];
 const [headers,base]=X.utils.sheet_to_json(sheet,{header:1});
 for(const [name,due,estimate] of [
  ['十一月先录','','2026年11月'],['具体预估','','2026-11-18'],['同周应付','2026-11-24','2027-01'],
  ['更早月份','','2026/10'],['十一月后录','','2026-11'],['更早应付','2026-11-18',''],['十二月','','2026-12'],['预估空白','','']
 ]){
  const row=base.slice();for(const [header,value] of [['供应商名称',name],['应付时间',due],['预估时间',estimate]])row[headers.indexOf(header)]=value;
  X.utils.sheet_add_aoa(sheet,[row],{origin:-1});
 }
 const before=JSON.stringify(i),built=C.build(i),name=C.NORMAL_HEADERS.indexOf('供应商名称');
 assert.deepEqual(built.normalRows.map(row=>row[name].v),['更早月份','具体预估','更早应付','十一月先录','同周应付','十一月后录','十二月','待确认先录','预估空白']);
 assert.deepEqual(built.normalRows.map(row=>row.at(-1).v),['2026/10/12周','2026/11/09周','2026/11/09周','2026/11/16周','2026/11/16周','2026/11/16周','2026/12/14周','日期待确认','日期待确认']);
 assert.equal(JSON.stringify(i),before);
});
