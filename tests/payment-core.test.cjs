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
