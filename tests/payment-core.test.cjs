const { test } = require('node:test');
const assert = require('node:assert/strict');
const C = require('../payment-core.js');
const row = (data = {}) => Object.assign(Object.fromEntries(C.required.map(k => [k, ''])), {
  _row: 2, 付款主体: '甲公司', 申请日期: '2026-09-25', 申请部门: '履约部', 经办人: '测试人', 本次申请付款金额: '2205', 付款总金额: '2205', 合同编号: '/', 合同总金额: '/', 已付金额: '/'
}, data);
test('sample maps the same sum to numeric and uppercase totals', () => {
  const a = C.analyze(C.required, [row()]);
  assert.deepEqual(a.errors, []); assert.equal(a.groups[0].total, 220500n);
  assert.equal(C.money(a.groups[0].total), '2,205.00');
  assert.equal(C.upper(a.groups[0].total), '贰仟贰佰零伍元整');
});
test('money uses cents, including 0.1 + 0.2', () => {
  assert.equal(C.cents(.1) + C.cents(.2), 30n);
  assert.equal(C.cents('￥1,234.56'), 123456n);
  assert.equal(C.cents(0), 0n); assert.equal(C.money(1n), '0.01');
});
test('invalid, negative, missing and over-precision amounts are exposed', () => {
  for (const v of ['', null, undefined, '-', '/', '-1', '1.001', '1,23', true, NaN, Infinity, '1e6', '9999999999999']) assert.throws(() => C.cents(v), String(v));
});
test('uppercase handles zero, jiao, fen, internal zero and large group boundaries', () => {
  const cases = { '0':'零元整', '0.01':'零元壹分', '0.10':'零元壹角', '1.01':'壹元零壹分', '10':'壹拾元整', '1001':'壹仟零壹元整', '10000':'壹万元整', '10001':'壹万零壹元整', '10010001':'壹仟零壹万零壹元整', '100000001':'壹亿零壹元整', '100010001.12':'壹亿零壹万零壹元壹角贰分' };
  for (const [v, expected] of Object.entries(cases)) assert.equal(C.upper(C.cents(v)), expected);
});
test('subjects and repeated supplier rows retain original order without deduplication', () => {
  const a = C.analyze(C.required, [row({ 本次申请付款金额: '.1' }), row({ 付款主体: '乙公司', _row: 3 }), row({ _row: 4 })]);
  assert.equal(a.errors.length, 1);
  assert.deepEqual(a.groups.map(g => g.subject), ['甲公司', '乙公司']);
  assert.deepEqual(a.groups[0].rows.map(r => r._row), [2, 4]);
});
test('repeated total column is checked, never added into sum', () => {
  const a = C.analyze(C.required, [row({本次申请付款金额:'1.10',付款总金额:'3.30'}), row({_row:3,本次申请付款金额:'2.20',付款总金额:'3.30'})]);
  assert.equal(a.groups[0].total,330n);assert.deepEqual(a.groups[0].checks,[]);
});
test('mismatching totals expose source row, zero check values are not skipped', () => {
  const a = C.analyze(C.required,[row({付款总金额:0})]);
  assert.equal(a.groups[0].checks.length,1);assert.match(a.groups[0].checks[0],/第 2 行/);
});
test('blank optional total is allowed but does not supply output total', () => {
  const a=C.analyze(C.required,[row({付款总金额:''})]);
  assert.equal(a.groups[0].total,220500n);assert.deepEqual(a.groups[0].checks,[]);
});
test('slash and blank fields stay distinct from zero; bank text stays exact', () => {
  const bank='收款银行\n账号 0012 3456\n行号 123456';
  const a=C.analyze(C.required,[row({开户银行及账号:bank})]);
  assert.equal(a.groups[0].rows[0].已付金额,'/');assert.equal(a.groups[0].rows[0].备注,'');assert.equal(a.groups[0].rows[0].开户银行及账号,bank);
});
test('conflicting subject information requires a new choice rather than first row wins', () => {
  const a=C.analyze(C.required,[row(),row({_row:3,经办人:'另一人'})]);
  assert.equal(a.groups[0].info.经办人,'');assert.equal(a.groups[0].choices.经办人.length,2);
});
test('missing subjects and amounts are reported with original row', () => {
  const a=C.analyze(C.required,[row({付款主体:''}),row({_row:4,本次申请付款金额:''})]);
  assert.equal(a.errors.length,2);assert.match(a.errors[0],/第 2 行/);assert.match(a.errors[1],/第 4 行/);
});
test('missing or duplicate column names stop processing', () => {
  assert.equal(C.analyze(C.required.slice(1),[row()]).groups.length,0);
  assert.equal(C.analyze([...C.required,'备注'],[row()]).errors.length,1);
});
test('real dates validate and leap-day errors are not silently normalized', () => {
  assert.equal(C.date('2024/2/29'),'2024-02-29');
  assert.equal(C.date('2026年9月25日'),'2026-09-25');
  assert.equal(C.date(''),'');
  for(const d of ['2026-02-29','2026-13-01','2026-00-02','9/25/26'])assert.throws(()=>C.date(d));
});
test('sum limit and formula error text block calculation',()=>{
  const a=C.analyze(C.required,[row({本次申请付款金额:'999999999999.99'}),row({本次申请付款金额:1})]);
  assert.match(a.errors.join(''),/合计超出/);
  assert.throws(()=>C.cents('#VALUE!'));
});
test('bank summary retains payee and spaced account only',()=>{
 const raw='收款单位 \n德州华彩纸制品有限公司\n收款银行名称 \n德州陵城农村商业银行边临镇支行\n收款账号 \n2730 0476 0420 50000 1025 6\n收款银行行号 \n402468300255';
 assert.deepEqual(C.bankSummary(raw),{value:'单位:德州华彩纸制品有限公司\n账号:2730 0476 0420 50000 1025 6',compact:true});
});
test('bank labels accept colon and inline forms, preserve leading zeros',()=>{
 assert.equal(C.bankSummary('单位：甲公司 账号：0012  3456 银行行号：999').value,'单位:甲公司\n账号:0012  3456');
 assert.equal(C.bankSummary('账号:\n0012 3456','甲公司').value,'单位:甲公司\n账号:0012 3456');
 assert.equal(C.bankSummary('单位:甲公司&#x20;\n账号:0012 3456&#x20;').value,'单位:甲公司\n账号:0012 3456');
});
test('bank ambiguity never selects routing number or discards unknown text',()=>{
 for(const raw of ['开户行:测试银行\n行号:123456','账号:123\n账号:456','单位:甲\n单位:乙\n账号:123','测试银行 123456','账号:123\n附言不明确']) {
  assert.deepEqual(C.bankSummary(raw,'甲公司'),{value:raw,compact:false});
 }
 assert.deepEqual(C.bankSummary(''),{value:'',compact:true});
 assert.deepEqual(C.bankSummary('/'),{value:'/',compact:true});
});
test('directory matches full names, fills accounts and preserves the original detail',()=>{
 const d=C.buildDirectory([{supplier:' 甲公司 ',bank:'0012  3456',row:2}]);
 const r={'供应商全称':'甲公司','开户银行及账号':''};
 assert.deepEqual(d.errors,[]);
 const result=C.resolveBank(r,d);
 assert.equal(result.value,'单位:甲公司\n账号:0012  3456'); assert.equal(result.source,'供应商名录第 2 行');
 assert.equal(r['开户银行及账号'],'');
 assert.match(C.resolveBank({'供应商全称':'甲公','开户银行及账号':''},d).error,/未匹配/);
});
test('directory priority reports differences; fill-only preserves populated detail',()=>{
 const d=C.buildDirectory([{supplier:'甲',bank:'222',row:2}]);
 const r={'供应商全称':'甲','开户银行及账号':'单位:甲\n账号:111'};
 assert.equal(C.resolveBank(r,d).conflict,true);
 assert.equal(C.resolveBank(r,d,'fill').value,r['开户银行及账号']);
 assert.equal(C.resolveBank({...r,'开户银行及账号':'/'},d,'fill').value,'单位:甲\n账号:222');
});
test('duplicate directory identities are accepted but conflicting accounts block',()=>{
 const same=C.buildDirectory([{supplier:'甲',bank:'0012 3456',row:2},{supplier:'甲',bank:'00123456',row:3}]);
 assert.equal(C.resolveBank({'供应商全称':'甲'},same).error,'');
 const conflict=C.buildDirectory([{supplier:'甲',bank:'111',row:2},{supplier:'甲',bank:'222',row:3}]);
 assert.match(C.resolveBank({'供应商全称':'甲'},conflict).error,/2、3.*不同/);
});
test('unmatched supplier explicitly falls back; missing bank blocks',()=>{
 const d=C.buildDirectory([{supplier:'甲',bank:'111',row:2}]);
 const result=C.resolveBank({'供应商全称':'乙','开户银行及账号':'原始银行信息'},d);
 assert.equal(result.value,'原始银行信息');assert.match(result.notice,/未匹配/);
 assert.match(C.resolveBank({'供应商全称':'乙'},d).error,/为空/);
});
test('directory payee mapping, blanks, precision and formula issues are explicit',()=>{
 const d=C.buildDirectory([{supplier:'甲',unit:'收款乙',bank:'单位:旧单位\n账号:000123',row:2}]);
 assert.equal(C.resolveBank({'供应商全称':'甲'},d).value,'单位:收款乙\n账号:000123');
 assert.match(C.buildDirectory([{supplier:'',bank:'123',row:4}]).errors[0],/第 4 行/);
 for(const item of [{bank:''},{bank:'测试银行'},{bank:'123',error:'精度丢失'},{bank:'123',error:'公式错误'}]){
 const index=C.buildDirectory([{supplier:'甲',row:2,...item}]);assert.ok(C.resolveBank({'供应商全称':'甲'},index).error);
 }
});
test('directory bank name and receiving account stay together; subjects support explicit alias',()=>{
 const d=C.buildDirectory([{supplier:'甲',subject:'浙江迈德斯特医疗器械科技有限公司',account:'001 234',bankName:'测试支行',row:2},{supplier:'甲',subject:'其他公司',account:'999',bankName:'其他支行',row:3}]);
 const r=C.resolveBank({'供应商全称':'甲','付款主体':'迈德斯特'},d);
 assert.equal(r.error,'');assert.equal(r.value,'银行名称:测试支行\n卡号:001 234');assert.equal(C.bankSummary(r.value).value,r.value);
 assert.equal(C.subjectName('浙江迈德斯特医疗器械科技有限公司'),'迈德斯特');
 assert.equal(C.subjectName('迈德斯特（宁波）医疗科技有限公司'),'迈德斯特（宁波）医疗科技有限公司');
 for(const row of [{account:'',bankName:'支行'},{account:'123',bankName:''},{account:'abc',bankName:'支行'}]) assert.ok(C.resolveBank({'供应商全称':'甲'},C.buildDirectory([{supplier:'甲',row:2,...row}])).error);
});
test('download name uses planned date, short subject and exact numeric total',()=>{
 const group={subject:'浙江迈德斯特医疗器械科技有限公司',total:220500n,info:{计划付款日期:'2026-09-25'}};
 assert.equal(C.downloadName([group],'docx'),'2026-09-25迈德斯特2205.00.docx');
 assert.equal(C.downloadName([group],'pdf'),'2026-09-25迈德斯特2205.00.pdf');
 assert.equal(C.downloadName([{...group,total:0n}],'pdf'),'2026-09-25迈德斯特0.00.pdf');
 assert.throws(()=>C.downloadName([{...group,info:{}}],'pdf'),/计划付款日期/);
 assert.throws(()=>C.downloadName([{...group,info:{计划付款日期:'2026-02-30'}}],'pdf'),/日期无效/);
 assert.throws(()=>C.downloadName([group,group],'pdf'),/分别下载/);
 assert.equal(C.downloadName([{...group,subject:'乙/公司',total:1n}],'pdf'),'2026-09-25乙_公司0.01.pdf');
});
