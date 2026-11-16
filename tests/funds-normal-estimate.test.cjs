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
 for(const raw of ['','2026-11','2026年11月','2026-02-30','月底协商',{t:'e',v:23,w:'#REF!'},{t:'n',f:'DATE(2026,11,30)'}]){
  const i=fixture({'应付时间':raw,'预估时间':'2027-01-31'}),before=JSON.stringify(i),{row}=result(i);
  assert.equal(row['申请付款的日期'].v,'日期待确认');assert.equal(JSON.stringify(i),before);
  if(typeof raw==='string')assert.equal(row.应付时间.v,raw,'空白或不完整/无效日期不补值');
 }
 const estimated=result(fixture({'预估时间':'2027-01-31'})).row;
 assert.equal(display(estimated.预估时间),'2027年1月31日');assert.equal(estimated['申请付款的日期'].v,'日期待确认');
});
test('正常付款按调整后的申请周稳定排序，同周保持台账源顺序',()=>{
 const i=fixture({'应付时间':'2026-10-31','供应商名称':'月底先录入'}),sheet=i.ledger.book.Sheets['对账明细-李奇'];
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
  assert.equal(row.应付时间.v,'');assert.equal(display(row.预估时间),'2026年11月18日');assert(!row.预估时间.f);assert.equal(built.stats.pending,0);assert.equal(row['申请付款的日期'].v,'日期待确认','不改变现有申请周计算来源');
 }
});
test('无预估列时从所属月最后一天加15天和账期，含跨年、闰年、零和月账期',()=>{
 for(const [month,terms,expected] of [
  ['2026年10月',30,'2026年12月15日'],['2026-10-31','票到7天','2026年11月22日'],['2026/9',30,'2026年11月14日'],
  ['2026-10',0,'2026年11月15日'],['2026-12',30,'2027年2月14日'],['2028-02',30,'2028年4月14日'],['2027-02',30,'2027年4月14日'],
  ['2026-10','5个月账期（月底付款）50%承兑','2027年4月14日']
 ]){
  const i=fixture({'所属月份':month,'账期':terms}),before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(display(row.预估时间),expected);assert.equal(row.应付时间.v,'');assert.equal(built.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
 const i=fixture({'所属月份':{t:'n',v:44864,f:'DATE(2026,10,31)',z:'yyyy-mm'},'账期':{t:'n',v:30,f:'15+15'}});i.ledger.book.Workbook={WBProps:{date1904:true}};
 assert.equal(display(result(i).row.预估时间),'2026年12月15日');
});
test('已有预估列空白不推算，年月保持精度，错误与不确定账期保留核对来源',()=>{
 for(const month of ['2026年12月','2026-12',{t:'n',v:46371,z:'yyyy年m月'}]){
  const {row}=result(fixture({'预估付款时间':month}));assert.equal(display(row.预估时间),'2026年12月');
 }
 const blank=result(fixture({'预估时间':''}));assert.equal(blank.row.预估时间.v,'');assert.match(blank.built.issues[0].message,/已有预估列不自动推算/);
 for(const terms of ['',-30,'30至60天','30/60天','30天或60天','预付50%','先款后货','7个工作日','约30天']){
  const {built,row}=result(fixture({'账期':terms}));assert.equal(row.预估时间.v,'');assert.equal(built.stats.pending,1);assert.match(built.issues[0].message,/账期/);assert.deepEqual(built.issues[0].source,{file:'本次台账.xlsx',sheet:'对账明细-李奇',row:2});
 }
 for(const bad of [{t:'e',v:23,w:'#REF!'}, {t:'n',f:'A2+30'}, {t:'s',v:'',f:'A2+30'}])for(const h of ['预估时间','所属月份','账期']){
  const {built}=result(fixture({[h]:bad}));assert(built.issues.some(q=>q.id.startsWith('normal-estimate:')&&q.kind==='formula'));assert(built.stats.pending>0);
 }
 for(const month of ['', '11月', '2026-13']){
  const {built,row}=result(fixture({'所属月份':month}));assert.equal(row.预估时间.v,'');assert.match(built.issues.find(q=>q.id.startsWith('normal-estimate:')).message,/所属月份/);
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
test('空白、非法预估月份和无效公式缓存不产生申请日期，仍保留核对提示',()=>{
 for(const raw of ['', '2026-00', '2026-13', '11月', {t:'e',v:23,w:'#REF!'}, {t:'n',f:'DATE(2026,11,1)',z:'yyyy-mm'}, {t:'s',v:'',f:'TEXT(A2,"yyyy-mm")'}]){
  const i=fixture({'预估时间':raw}),before=JSON.stringify(i),{built,row}=result(i);
  assert.equal(row['申请付款的日期'].v,'日期待确认');assert.equal(built.stats.pending,1);
  assert(built.issues.some(q=>q.id.startsWith('normal-estimate:')));assert.equal(JSON.stringify(i),before);
  if(raw==='')assert.equal(row.预估时间.v,'','已有预估列空白仍不推算');
 }
});
test('月份预估与有效应付日期共同按申请周排序，同周和待确认组均保持台账原序',()=>{
 const i=fixture({'供应商名称':'待确认先录','预估时间':'2026-13'}),sheet=i.ledger.book.Sheets['对账明细-李奇'];
 const [headers,base]=X.utils.sheet_to_json(sheet,{header:1});
 for(const [name,due,estimate] of [
  ['十一月先录','','2026年11月'],['具体预估保持待确认','','2026-11-18'],['同周应付','2026-11-24','2027-01'],
  ['更早月份','','2026/10'],['十一月后录','','2026-11'],['更早应付','2026-11-18',''],['十二月','','2026-12'],['预估空白','','']
 ]){
  const row=base.slice();for(const [header,value] of [['供应商名称',name],['应付时间',due],['预估时间',estimate]])row[headers.indexOf(header)]=value;
  X.utils.sheet_add_aoa(sheet,[row],{origin:-1});
 }
 const before=JSON.stringify(i),built=C.build(i),name=C.NORMAL_HEADERS.indexOf('供应商名称');
 assert.deepEqual(built.normalRows.map(row=>row[name].v),['更早月份','更早应付','十一月先录','同周应付','十一月后录','十二月','待确认先录','具体预估保持待确认','预估空白']);
 assert.deepEqual(built.normalRows.map(row=>row.at(-1).v),['2026/10/12周','2026/11/09周','2026/11/16周','2026/11/16周','2026/11/16周','2026/12/14周','日期待确认','日期待确认','日期待确认']);
 assert.equal(JSON.stringify(i),before);
});
