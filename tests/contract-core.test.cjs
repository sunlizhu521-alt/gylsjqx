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

test('跳过采购合同标题行并识别真实订单表头', () => {
  const result = C.analyzeMatrix([
    ['采购合同', '', '', ''],
    ['物料编码', '物料名称', '规格型号', '数量'],
    ['MAT-001', '产品A', 'A型', '2'],
  ]);
  assert.equal(result.headerIndex, 1);
  assert.deepEqual(result.headers, ['物料编码', '物料名称', '规格型号', '数量']);
  assert.equal(result.rows[0].物料编码, 'MAT-001');
});

test('订单有序号列时按序号行生成明细并排除交货信息行', () => {
  const result = C.analyzeMatrix([
    ['合同编号：', '', '', '', '', '', '', '', '', 'CGDD015006'],
    ['序号', '物料编码', '物料名称', '', '规格型号', 'SKU', '单位', '数量', '含税单价\n(元)', '含税总金额\n(元)', '备注'],
    ['1', '2001020022', '产品A', '', '24V', '', '根', '200', '2.00', '400.00', ''],
    ['2', '2001020023', '产品B', '', '24V', '', '根', '400', '2.00', '800.00', ''],
    ['交货地点', '', '', '', '', '', '', '', '交货时间', '2026年9月30日', ''],
  ]);
  const detail = C.selectDetailRows(result.rows, result.headers, { materialCode: '错误字段' });
  assert.deepEqual(detail.map(row => row.序号), ['1', '2']);
  assert.deepEqual(detail.map(row => row._row), [3, 4]);
});

test('合同模板连续预留的空白序号行会作为一个明细区整体替换', () => {
  const rows = [
    { rowKey: 'word:0:0', cells: [{ value: '序号' }, { value: '物料名称' }] },
    { rowKey: 'word:0:1', cells: [{ value: '1' }, { value: '待填写' }] },
    { rowKey: 'word:0:2', cells: [{ value: '2' }, { value: '' }] },
    { rowKey: 'word:0:3', cells: [{ value: '3' }, { value: '空白位置' }] },
    { rowKey: 'word:0:4', cells: [{ value: '含税合计' }, { value: '' }] },
  ];
  assert.equal(C.reservedDetailRowCount(rows, 'word:0:1', 0), 3);
  rows[2].cells[1].value = '模板固定内容';
  assert.equal(C.reservedDetailRowCount(rows, 'word:0:1', 0), 1);
});

test('含税运总金额按整数分汇总并同步生成小写和人民币大写', () => {
  const total = C.sumAmountField([
    { _row: 3, 金额: '400.00' },
    { _row: 4, 金额: '800' },
    { _row: 5, 金额: '0.05' },
  ], '金额');
  assert.deepEqual(total, { cents: 120005n, lower: '1200.05', upper: '壹仟贰佰元零伍分' });
  assert.deepEqual(C.sumAmountField([{ _row: 3, 金额: '0' }], '金额'), { cents: 0n, lower: '0.00', upper: '零元整' });
  assert.throws(() => C.sumAmountField([{ _row: 8, 金额: '1.001' }], '金额'), /第 8 行金额须为非负数，且最多两位小数/);
  assert.throws(() => C.sumAmountField([{ _row: 9, 金额: '' }], '金额'), /第 9 行金额须为非负数/);
});

test('合同编号、交货地点和交货时间读取标签右侧内容', () => {
  const values = C.extractAdjacentLabelValues([
    ['合同编号：', 'HT-001', '', '交货地点', '杭州仓'],
    ['交货时间', '', '2026-10-01'],
    ['交货时间', '备注'],
  ], {
    contractNumber: ['合同编号', '合同号'],
    deliveryPlace: ['交货地点', '送货地址'],
    deliveryTime: ['交货时间', '交期'],
  }, [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }], [2]);
  assert.deepEqual(values, {
    contractNumber: 'HT-001',
    deliveryPlace: '杭州仓',
    deliveryTime: '2026-10-01',
  });
});

test('合同编号可从合并标签右侧读取', () => {
  const values = C.extractAdjacentLabelValues([
    ['                                     合同编号：', '', '', '', '', '', '', '', '', 'CGDD015036', ''],
    ['序号', '物料编码', '物料名称', '', '规格型号', 'SKU', '单位', '数量', '含税单价\n(元)', '含税总金额\n(元)', '备注'],
  ], {
    contractNumber: ['合同编号', '合同号'],
  }, [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } },
    { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } },
  ], [1]);
  assert.deepEqual(values, { contractNumber: 'CGDD015036' });
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

test('数量乘单价使用十进制定点运算并逐行四舍五入到分', () => {
  assert.equal(C.lineAmountCents('3','0.1'),30n);
  assert.equal(C.lineAmountCents('0','100'),0n);
  assert.equal(C.lineAmountCents('1','0.005'),1n);
  assert.equal(C.lineAmountCents('1','0.0049'),0n);
  assert.equal(C.lineAmountCents('1.25','10.50'),1313n);
  assert.equal(C.lineAmountCents('1,000','2.25'),225000n);
  for(const v of ['',null,'bad','-1','1,23','Infinity','0.000000001']) assert.throws(()=>C.lineAmountCents(v,'1'));
  assert.throws(()=>C.lineAmountCents('1',''));
  assert.throws(()=>C.lineAmountCents('999999999999999','999999999999999'));
});

test('原表合计仅核对，兼容圆元与系统点零零写法', () => {
 const calculated={cents:8761200n,lower:'87612.00',upper:C.amountUpper(8761200n)};
 assert.deepEqual(C.compareSourceTotals({lower:'87,612.00',upper:'捌万柒仟陆佰壹拾贰点零零'},calculated),[]);
 assert.deepEqual(C.compareSourceTotals({},calculated),[]);
 assert.equal(C.compareSourceTotals({lower:'0',upper:'零元整'},calculated).length,2);
 assert.equal(C.compareSourceTotals({lower:'错误'},calculated).length,1);
 assert.equal(calculated.cents,8761200n);
});

test('无序号订单排除合计及创建审核页脚，保留四条物料', () => {
  const headers = ['物料编码', '物料名称', '数量', '单价'];
  const rows = [
    ['1201', '螺丝A', '2500', '0.017699'], ['1202', '螺丝B', '1000', '0.028319'],
    ['1203', '螺丝C', '1000', '0.033628'], ['1204', '螺丝D', '3000', '0.019469'],
    ['合计', '', '7500', ''], ['创建人', '测试人员', '', ''], ['创建日期', '2026/9/29', '', ''],
  ].map(values => Object.fromEntries(headers.map((header, i) => [header, values[i]])));
  assert.equal(C.selectDetailRows(rows, headers, { materialCode: '物料编码', materialName: '物料名称' }).length, 4);
});

test('交货日期只去除时间且不跨时区换日', () => {
  for (const [input, expected] of [
    ['2026/10/6 23:59:59', '2026/10/6'], ['2026-10-06T23:59:59Z', '2026-10-06'],
    ['2026年10月6日 23:59', '2026年10月6日'], ['2026-10-06', '2026-10-06'],
    ['', ''], ['另行约定', '另行约定'],
  ]) assert.equal(C.dateOnly(input), expected);
});

test('价税合计换算含税单价精确四舍五入到三位，不截断', () => {
  assert.deepEqual([[50,2500],[32,1000],[38,1000],[66,3000]].map(([a,q]) => C.taxPriceFromAmount(a,q)), ['0.02','0.032','0.038','0.022']);
  assert.equal(C.taxPriceFromAmount('1.00', '3'), '0.333');
  assert.equal(C.taxPriceFromAmount('1.00', '16'), '0.063');
  assert.equal(C.taxPriceFromAmount('1.00', '0.5'), '2');
  assert.equal(C.taxPriceFromAmount('0.00', '10'), '0');
  for (const q of ['', '0', '-1', '无']) assert.throws(() => C.taxPriceFromAmount('1.00', q));
  for (const a of ['', '-1', '无']) assert.throws(() => C.taxPriceFromAmount(a, '2'));
  const cents = C.lineAmountCents('3', C.taxPriceFromAmount('1.00', '3'));
  assert.equal(cents, 100n);
  const roundedTotal = C.sumAmountField([{amount: C.formatAmountCents(C.lineAmountCents('16', C.taxPriceFromAmount('1.00','16')))}], 'amount');
  assert.equal(C.compareSourceTotals({lower:'1.00'}, roundedTotal).length, 1);
});
