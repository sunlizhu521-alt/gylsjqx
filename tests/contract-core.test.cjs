const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../contract-core.js');

test('订单表头和明细解析保留原始顺序', () => {
  const result = C.analyzeMatrix([
    ['', ''],
    ['供应商', '数量', '备注'],
    ['甲公司', '0', ''],
    ['乙公司', '2.5', '保留'],
  ]);
  assert.deepEqual(result.headers, ['供应商', '数量', '备注']);
  assert.equal(result.rows.length, 2);
  assert.equal(result.rows[0].数量, '0');
  assert.equal(result.rows[1]._row, 4);
});

test('重复表头、空明细和超长明细均被拦截', () => {
  assert.throws(() => C.analyzeMatrix([['名称', '名称'], ['A', 'B']]), /重复字段/);
  assert.throws(() => C.analyzeMatrix([['名称']]), /只有表头/);
  assert.throws(() => C.analyzeMatrix([['名称'], ['A'], ['B']], 1), /超过 1 行/);
});

test('单值冲突必须明确策略，数字求和覆盖零值和小数', () => {
  const rows = [{ 金额: '0', 供应商: '甲' }, { 金额: '2.50', 供应商: '乙' }];
  assert.deepEqual(C.mappingIssues(rows, { p1: { field: '供应商', mode: 'single', strategy: '' } }, ''), ['供应商 有多个不同值，需要选择处理方式']);
  assert.equal(C.resolveField(rows, '金额', 'sum'), '2.5');
  assert.equal(C.resolveField(rows, '供应商', 'merge'), '甲、乙');
  assert.equal(C.resolveField(rows, '供应商', 'manual', '核实值'), '核实值');
});

test('非数字求和、空手工值和明细行缺失均被拦截', () => {
  const rows = [{ 金额: '1' }, { 金额: '异常' }];
  assert.throws(() => C.resolveField(rows, '金额', 'sum'), /非数字/);
  const issues = C.mappingIssues(rows, {
    a: { field: '金额', mode: 'single', strategy: 'sum' },
    b: { field: '金额', mode: 'single', strategy: 'manual', manual: '' },
    c: { field: '金额', mode: 'detail' },
  }, '');
  assert.ok(issues.some(item => item.includes('不能求和')));
  assert.ok(issues.some(item => item.includes('手工值不能为空')));
  assert.ok(issues.some(item => item.includes('尚未选择明细模板行')));
});

test('自定义合同文件名只过滤系统非法字符', () => {
  assert.equal(C.sanitizeFileName(' 杭州/采购:合同*2026 '), '杭州_采购_合同_2026');
  assert.equal(C.sanitizeFileName(''), '合同');
});
