const { test } = require('node:test');
const assert = require('node:assert/strict');
const Import = require('../quarter-budget-import.js');

const headers = ['预估付款月份', '供应商名称', '付款主体', '对公付款0A流程号', '摘要', '申请付款金额', '是否已支付', '备注', '分类', '跟单员', '是否剔除'];
function sheet(rows) {
  const data = {};
  rows.forEach((row, r) => row.forEach((v, c) => {
    const col = c < 26 ? String.fromCharCode(c + 65) : 'A' + String.fromCharCode(c - 26 + 65);
    if (v !== undefined) data[col + (r + 1)] = v && typeof v === 'object' && !(v instanceof Date) ? v : { v, t: typeof v === 'number' ? 'n' : 's' };
  }));
  return data;
}
function source(tabs, extra = {}) { return { book: { SheetNames: Object.keys(tabs), Sheets: tabs, ...extra }, filename: 'synthetic.xlsx', fingerprint: 'fixture-1' }; }
function one(values = [], role = 'manualBudget') { return Import.parse(source({ 明细: sheet([headers, values]) }), role); }
const red = v => ({ v, t: 'n', s: { patternType: 'solid', fgColor: { rgb: 'FF0000' } } });
const negotiationHeaders = ['跟单员', '供应商', '付款主体', '对公付款OA流程号', '摘要', '谈判明细', '2026-11-02', '2026-11-25', '2026-12-25'];

test('standardized 11 columns keep original values, order, zero and exclusion', () => {
  const input = source({ 明细: sheet([headers, [46327, ' 甲公司 ', '付款甲', 'OA-1', '款项', '0.00', '已支付', '人工备注', '自定分类', '张甲', '是'], [46357, '乙', '', '', '', null, '未支付']]) });
  const before = JSON.stringify(input);
  const result = Import.parse(input, 'manualBudget');
  assert.equal(result.records.length, 2);
  assert.equal(result.records[0].supplier, ' 甲公司 ');
  assert.equal(result.records[0].month, 46327);
  assert.equal(result.records[0].monthKey, '2026-11');
  assert.equal(result.records[0].amount, '0.00'); assert.equal(result.records[0].cents, 0);
  assert.equal(result.records[0].excluded, '是'); assert.equal(result.records[0].category, '自定分类');
  assert.match(result.records[0].notes, /人工备注/);
  assert.equal(result.records[1].amount, null); assert.equal(result.records[1].cents, null);
  assert.equal(result.evidence.length, 0, 'manual comparison status is not payment evidence');
  assert.equal(JSON.stringify(input), before);
});
test('a formatted million-row ref and blank styled cells never create phantom rows', () => {
  const s = sheet([headers, ['11月', '甲', '', '', '', 0]]);
  s['!ref'] = 'A1:XEZ1048576'; s.XEZ1048576 = { s: { fgColor: { rgb: 'FF0000' } } };
  const a = Import.parse(source({ 明细: s }), 'manualBudget');
  assert.equal(a.records.length, 1); assert.equal(a.records[0].monthKey, '');
  assert.ok(a.warnings.some(w => w.code === 'UNKNOWN_MONTH_YEAR'));
});
test('month-only values require an explicit year and never use the filename year', () => {
  const s = source({ 明细: sheet([headers, ['11月', '甲', '', '', '', 1]]) }); s.filename = '预算2026.xlsx';
  assert.equal(Import.parse(s, 'budget').records[0].monthKey, '');
  assert.equal(Import.parse(s, 'budget', { budgetYear: 2027 }).records[0].monthKey, '2027-11');
  assert.equal(one(['2026年11月', '甲', '', '', '', 1]).records[0].monthKey, '2026-11');
  assert.equal(one(['2026-02-30', '甲', '', '', '', 1]).records[0].monthKey, '');
});
test('original date values and 1904 epoch are supported without changing input', () => {
  const r = one([new Date('2026-11-01T00:00:00Z'), '甲', '', '', '', 1]).records[0];
  assert.equal(r.month, '2026-11-01T00:00:00.000Z'); assert.equal(r.monthKey, '2026-11');
  const input = source({ 明细: sheet([headers, [44865, '甲', '', '', '', 1]]) }, { Workbook: { WBProps: { date1904: true } } });
  assert.equal(Import.parse(input, 'manualBudget').records[0].monthKey, '2026-11');
});
test('money uses exact decimal rounding while preserving original overprecision', () => {
  for (const [v, expected] of [[0.1, 10], [0.2, 20], [1.005, 101], [-1.005, -101], ['￥1,234.56', 123456], ['1e-3', 0], ['90071992547409.91', Number.MAX_SAFE_INTEGER]]) {
    const a = one(['11月', '甲', '', '', '', v]);
    assert.equal(a.records[0].cents, expected, String(v)); assert.equal(a.records[0].amount, v);
  }
  assert.ok(one(['11月', '甲', '', '', '', 1.005]).warnings.some(w => w.code === 'AMOUNT_PRECISION'));
});
test('invalid, unsafe, error and uncached formula amounts are errors, never zero', () => {
  for (const v of ['1,23', '金额待定', Infinity, true, '90071992547409.92', { f: 'SUM(A1:A2)' }, { t: 'e', v: 7, w: '#DIV/0!' }]) {
    const a = one(['11月', '甲', '', '', '', v]);
    assert.equal(a.records[0].cents, null); assert.ok(a.errors.length, JSON.stringify(v));
  }
  assert.equal(one(['11月', '甲', '', '', '', { f: '1+1', v: 2, t: 'n' }]).records[0].cents, 200);
});
test('empty rows and trailing notes are excluded; personal budget sheets retain names', () => {
  const h = ['预估付款月份', '供应商名称', '付款主体', '摘要', '申请付款金额', '账期', '备注', '备注'];
  const input = source({ 张甲: sheet([h, [46327, '供应甲', '主体甲', '货款', 0, '票到30天', '原备注', '补充备注'], [], [undefined, undefined, undefined, undefined, undefined, undefined, '残留文字']]), 李乙: sheet([h]), '张甲-原始表': sheet([h, [46327, '供应甲', '主体甲', '货款', 0]]) });
  const a = Import.parse(input, 'budget');
  assert.equal(a.records.length, 1); assert.equal(a.records[0].owner, '张甲'); assert.equal(a.records[0].category, '正常付款-预估付款');
  for (const x of ['票到30天', '原备注', '补充备注']) assert.ok(a.records[0].notes.includes(x));
  assert.ok(a.sheets.find(s => s.name === '张甲-原始表').skipped);
});
test('explicit chosen backup sheet is honored and a missing sheet is an error', () => {
  const s = source({ '旧表': sheet([headers, [46327, '甲', '', '', '', 1]]) });
  assert.equal(Import.parse(s, 'manualBudget', { sheetName: '旧表' }).records.length, 1);
  assert.ok(Import.parse(s, 'manualBudget', { sheetName: '不存在' }).errors.length);
});
test('extracts only a unique OA and preserves historical-flow and other notes', () => {
  const a = one([46327, '甲', '', '', '', 1, '', '原流程编号：MDST-DGFK20260101001，本期未提交；2026/10/1已支付另一笔。'], 'budget');
  assert.equal(a.records[0].oa, 'MDST-DGFK20260101001'); assert.equal(a.records[0].status, '待核对');
  assert.match(a.records[0].notes, /原流程，本期未提交/);
  const b = one([46327, '甲', '', '', '', 1, '', 'MDST-DGFK20260101001、MDST-DGFK20260101002'], 'budget');
  assert.equal(b.records[0].oa, ''); assert.ok(b.warnings.some(w => w.code === 'MULTIPLE_OA'));
});
test('budget notes require explicit current-payment evidence; red font is not paid', () => {
  const h = ['预估付款月份', '供应商', '付款主体', '摘要', '申请付款金额', '备注'];
  const s = sheet([h, [46327, '甲', '主体甲', '货款', 1, '已付'], [46327, '乙', '', '', 2, '历史已付100元，本期未提交'], [46327, '丙', '', '', 3, { v: '已开票', s: { font: { color: { rgb: 'FF0000' } } } }]]);
  const a = Import.parse(source({ 张甲: s }), 'budget');
  assert.deepEqual(a.records.map(r => r.status), ['已支付', '待核对', '待核对']); assert.equal(a.evidence.length, 1);
});
test('whole OA archive status alone cannot turn the current installment into paid', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '申请付款金额', '流程状态', '当前节点'];
  const a = Import.parse(source({ 台账: sheet([h, ['甲', '主体甲', 'OA1', 100, '归档', '归档抄送']]) }), 'ledger');
  assert.equal(a.evidence.length, 0); assert.equal(a.records[0].status, '待核对');
});
test('ledger status evidence is tied to its own amount and source role', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '申请付款金额', '是否已付款', '摘要'];
  const a = Import.parse(source({ 进度: sheet([h, ['甲', '主体甲', 'OA1', 100, '已付款', '第一期'], ['甲', '主体甲', 'OA1', 200, '未付款', '第二期']]) }), 'ledger');
  assert.deepEqual(a.evidence.map(e => [e.cents, e.status]), [[10000, '已支付'], [20000, '未支付']]);
  assert.equal(a.evidence[0].summary, '第一期'); assert.equal(a.evidence[0].source.role, 'ledger');
  assert.equal(a.evidence[0].id, 'fixture-1:%E8%BF%9B%E5%BA%A6:D2');
});
test('same month and amount installments retain separate source cells and order', () => {
  const s = sheet([negotiationHeaders, ['张甲', '供应甲', '付款甲', 'OA1', '货款', '1、11月2日－100.00\n2、11月25日－100.00', red(100), 100]]);
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.deepEqual(a.records.map(r => r.status), ['已支付', '未支付']);
  assert.deepEqual(a.records.map(r => r.monthKey), ['2026-11', '2026-11']);
  assert.deepEqual(a.records.map(r => r.month), [46327, 46327]);
  assert.match(a.records[0].notes, /谈判明细：1、11月2日/);
  assert.deepEqual(a.evidence.map(e => e.source.cell), ['G2', 'H2']);
  assert.notEqual(a.evidence[0].id, a.evidence[1].id); assert.deepEqual(a.evidence.map(e => e.part), [1, 2]);
});
test('merged installments map one to each source row without copying OA or owner', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '第一笔', '11月25日-100\n12月25日-200', red(100)], ['乙', '供应甲', '主体乙', 'OA2', '第二笔', null, null, null, 200]]);
  s['!merges'] = [{ s: { r: 1, c: 5 }, e: { r: 2, c: 5 } }];
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.deepEqual(a.records.map(r => [r.oa, r.owner, r.amount, r.source.row, r.source.part]), [['OA1', '甲', 100, 2, 1], ['OA2', '乙', 200, 3, 2]]);
  assert.deepEqual(a.records.map(r => r.status), ['已支付', '未支付']);
});
test('ambiguous merged text preserves source rows but never replicates installment amounts', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', '11月25日-100\n12月25日-200\n1月25日-300', red(100)], ['乙', '供应甲', '主体乙', 'OA2', '', null, null, null, 200]]);
  s['!merges'] = [{ s: { r: 1, c: 5 }, e: { r: 2, c: 5 } }];
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.equal(a.records.length, 2); assert.ok(a.records.every(r => r.amount === null && r.status === '待核对'));
  assert.ok(a.warnings.some(w => w.code === 'MERGED_PARTS_AMBIGUOUS')); assert.equal(a.evidence.length, 2);
});
test('red font or colored text outside installment amount does not prove payment', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', red('11月25日-100'), { v: 100, s: { font: { color: { rgb: 'FF0000' } } } }]]);
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.equal(a.records[0].status, '未支付');
});
test('schedule differences keep original negotiated amount and warn without changing month', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', '10月30日-100', red(100.67)]]);
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.equal(a.records[0].amount, 100); assert.equal(a.records[0].cents, 10000); assert.equal(a.records[0].month, '10月');
  assert.equal(a.records[0].monthKey, ''); assert.equal(a.evidence[0].cents, 10067);
  assert.ok(a.warnings.some(w => w.code === 'SCHEDULE_AMOUNT_CONFLICT')); assert.ok(a.warnings.some(w => w.code === 'SCHEDULE_MONTH_CONFLICT'));
});
test('unmatched periods and blank details remain reviewable instead of disappearing', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', '11月25日-100'], ['乙', '供应乙', '主体乙', 'OA2', '', null]]);
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.equal(a.records.length, 2); assert.equal(a.records[0].amount, 100); assert.equal(a.records[1].amount, null);
  assert.ok(a.records.every(r => r.status === '待核对'));
});
test('zero installments are preserved, cached-error schedules are not paid', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', '11月25日-0', 0], ['乙', '供应乙', '主体乙', 'OA2', '', '11月25日-100', { f: '1/0', t: 'e', v: 7, w: '#DIV/0!', s: { fgColor: { rgb: 'FF0000' } } }]]);
  const a = Import.parse(source({ 谈判: s }), 'negotiated');
  assert.equal(a.records[0].amount, 0); assert.equal(a.records[0].cents, 0);
  assert.equal(a.records[1].status, '待核对'); assert.ok(a.errors.some(e => e.code === 'AMOUNT_FORMULA_ERROR'));
});
test('partial-payment states remain reviewable and are not emitted as paid evidence', () => {
  const a = one(['2026-11', '甲', '', '', '', 200, '部分支付'], 'ledger');
  assert.equal(a.records[0].status, '待核对'); assert.equal(a.evidence.length, 0); assert.match(a.records[0].notes, /部分支付/);
});
test('unknown layouts fail clearly, even when a workbook only contains decorative text', () => {
  const a = Import.parse(source({ 首页: sheet([['年度报告'], ['本页无明细']]) }), 'manualBudget');
  assert.equal(a.records.length, 0); assert.ok(a.errors.some(e => e.code === 'NO_PAYMENT_SHEET'));
  assert.throws(() => Import.parse({}, 'manualBudget'), /工作簿/); assert.throws(() => Import.parse(source({}), 'wrong'), /角色/);
  for (const role of ['existing', 'evidence']) assert.throws(() => Import.parse(source({}), role), /角色/);
});
test('roles select the intended detail sheet and never merge auxiliary progress rows', () => {
  const schedule = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '', '11月25日-100', red(100)]]);
  const generic = sheet([headers, [46327, '供应乙', '主体乙', 'OA2', '', 200, '已支付']]);
  const input = source({ 谈判付款: schedule, 正常付款: generic, 付款汇总明细: generic, 付款进度整理: generic });
  assert.equal(Import.parse(input, 'negotiated').records.length, 1);
  assert.equal(Import.parse(input, 'negotiated').records[0].supplier, '供应甲');
  assert.equal(Import.parse(input, 'ledger').records[0].source.sheet, '正常付款');
  assert.equal(Import.parse(input, 'manualBudget').records.length, 0, 'multiple old summaries do not silently choose a manual comparison source');
});
test('multiple fallback detail sheets require selection instead of mixing datasets', () => {
  const s = sheet([headers, [46327, '甲', '', '', '', 1]]);
  for (const role of ['negotiated', 'ledger', 'manualBudget', 'budget']) {
    const a = Import.parse(source({ 明细甲: s, 明细乙: s }), role);
    assert.equal(a.records.length, 0); assert.ok(a.errors.some(e => e.code === 'AMBIGUOUS_SHEETS'), role);
  }
});
test('budget chooses a unique standardized budget instead of mixing it with personal inputs', () => {
  const raw = sheet([['预估付款月份', '供应商', '申请付款金额'], [46327, '甲', 1]]);
  const standard = sheet([headers, [46327, '乙', '', '', '', 2]]);
  const a = Import.parse(source({ 张甲: raw, 未预估付款明细: standard }), 'budget');
  assert.equal(a.records.length, 1); assert.equal(a.records[0].supplier, '乙');
});
test('dated paid-only budget notes apply to this row; installment history stays unconfirmed', () => {
  const h = ['预估付款月份', '供应商', '申请付款金额', '备注'];
  const a = Import.parse(source({ 张甲: sheet([h, [46327, '甲', 100, '2026/11/1，2026/11/5号已支付'], [46327, '乙', 100, '2026/11/1已支付另一笔100元，本期未付']]) }), 'budget');
  assert.deepEqual(a.records.map(r => r.status), ['已支付', '待核对']);
});
test('historical progress summary cannot add a second paid proof to identical installments', () => {
  const s = sheet([negotiationHeaders, ['甲', '供应甲', '主体甲', 'OA1', '货款', '11月2日-100\n11月25日-100', red(100), 100]]);
  const h = ['供应商', '付款主体', '对公付款OA流程号', '申请付款金额', '是否已付款'];
  const a = Import.parse(source({ 谈判付款: s, 付款进度整理: sheet([h, ['供应甲', '主体甲', 'OA1', 100, '已付款']]) }), 'negotiated');
  assert.equal(a.evidence.length, 2); assert.deepEqual(a.evidence.map(e => e.status), ['已支付', '未支付']);
  assert.ok(a.sheets.find(s => s.name === '付款进度整理').skipped);
});
test('application-week headers do not reorder equal payments in the next calendar month', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '谈判明细', '2026-11-30', '2026-12-14'];
  const a = Import.parse(source({ 谈判付款: sheet([h, ['甲', '付款甲', 'OA1', '12月10日-100\n12月25日-100', red(100), 100]]) }), 'negotiated');
  assert.deepEqual(a.evidence.map(e => e.source.cell), ['E2', 'F2']);
  assert.deepEqual(a.records.map(r => r.status), ['已支付', '未支付']);
  assert.deepEqual(a.records.map(r => r.monthKey), ['2026-12', '2026-12']);
  assert.deepEqual(a.evidence.map(r => r.monthKey), ['2026-12', '2026-12']);
  assert.ok(!a.warnings.some(w => w.code === 'UNMATCHED_INSTALLMENT'));
});
test('year inheritance stays within the same payment month and original source row', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '谈判明细', '2026-11-30', '2027-01-18'];
  const a = Import.parse(source({ 谈判付款: sheet([h, ['甲', '付款甲', 'OA1', '12月10日-100\n1月25日-200', 100, 200], ['乙', '付款乙', 'OA2', '12月10日-100', 100]]) }), 'negotiated');
  assert.deepEqual(a.records.map(r => r.monthKey), ['', '2027-01', '']);
  assert.deepEqual(a.records.map(r => r.month), ['12月', 46388, '12月']);
});
test('conflicting years for one payment month are never propagated to an undated installment', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '谈判明细', '2026-11-30', '2026-12-14', '2027-12-14'];
  const a = Import.parse(source({ 谈判付款: sheet([h, ['甲', '付款甲', 'OA1', '12月10日-100\n12月25日-200\n12月25日-300', 100, 200, 300]]) }), 'negotiated');
  assert.deepEqual(a.records.map(r => r.monthKey), ['', '2026-12', '2027-12']);
  assert.equal(a.records[0].month, '12月');
});
test('November December January installments retain the actual cross-year header association', () => {
  const h = ['供应商', '付款主体', '对公付款OA流程号', '谈判明细', '2026-11-16', '2026-12-14', '2027-01-18'];
  const a = Import.parse(source({ 谈判付款: sheet([h, ['甲', '付款甲', 'OA1', '11月25日-100\n12月25日-100\n1月25日-100', 100, 100, 100]]) }), 'negotiated');
  assert.deepEqual(a.records.map(r => r.monthKey), ['2026-11', '2026-12', '2027-01']);
  assert.deepEqual(a.records.map(r => r.month), [46327, 46357, 46388]);
});

const manualHeaders = ['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','付款时间','备注'];
const summaryHeaders = [...manualHeaders.slice(0, 7),'付款日期','申请纸质单日期','是否申请纸质单','是否已付款'];
test('new installment summary is authoritative, uses payment date and explicit batch status, and excludes only explicit nonpayment', () => {
  const input = source({
    谈判付款: sheet([negotiationHeaders, ['甲', '旧来源', '主体', 'OA1', '', '12月25日-999', 999]]),
    谈判付款付款表付款节奏: sheet([summaryHeaders,
      ['谈判付款-谈判付款表','李奇','甲','主体','OA1','',100,'2026-12-10','2026-11-30','是','否'],
      ['正常付款-谈判付款表','徐英','乙','主体','OA2','',0,'2027-01-10','2026-12-28','否','是'],
      ['谈判付款-谈判付款表','徐英','丙','主体','OA3','',10,'2026-12-25','','是','待核对'],
      ['未知','','不支付供应商','','不支付','',999,'2026-12-25'],
      ['谈判付款-谈判付款表','','保留错误OA','',{t:'e',v:'不支付'},'',5,'2026-12-25']]),
    付款进度整理: sheet([headers, ['2026-12','甲','','OA1','',100,'已支付']])
  });
  const before = JSON.stringify(input), result = Import.parse(input, 'negotiated', {budgetStartMonth:'2026-12'});
  assert.equal(result.records.length, 4); assert.deepEqual(result.records.map(r=>r.supplier),['甲','乙','丙','保留错误OA']);
  assert.deepEqual(result.records.map(r=>r.monthKey),['2026-12','2027-01','2026-12','2026-12']);
  assert.deepEqual(result.records.map(r=>r.status),['未支付','已支付','待核对','待核对']);
  assert.equal(result.records[1].category,'正常付款-谈判付款表'); assert.equal(result.records[1].cents,0);
  assert.deepEqual(result.evidence.map(e=>e.status),['未支付','已支付']);
  assert.equal(result.sheets.filter(s=>!s.skipped).length,1); assert.equal(result.records[0].source.sheet,'谈判付款付款表付款节奏');
  assert.equal(JSON.stringify(input),before);
});
test('legacy two-row headers retain original schedule dates and classify normal negotiated-source records', () => {
  const s = sheet([['供应商','付款主体','对公付款OA流程号','谈判明细','2026年12月付款','是否谈判分批'],
    [null,null,null,null,'2026-11-30'], ['甲','主体','OA1','2026年12月10日-100',red(100),'否']]);
  const a = Import.parse(source({谈判付款:s}), 'negotiated');
  assert.equal(a.records[0].monthKey,'2026-12'); assert.equal(a.records[0].status,'已支付');
  assert.equal(a.records[0].category,'正常付款-谈判付款表'); assert.equal(a.evidence[0].source.cell,'E3');
});
test('raw ledger selects the five personal sources, honors merged negotiated flags, and ignores historic summaries', () => {
  const h=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','是否谈判付款','应付时间','预估付款时间','已付款金额','是否结清'];
  const merged=sheet([h,['谈判甲','主体','OA1',100,'是','2026-12-10'],['谈判乙','主体','OA2',100,null,'2026-12-10'],['正常甲','主体','OA3',0,'否','2026-12-10'],['正常乙','主体','OA4',50,'','2026-12-10',null,50,'是']]);
  merged['!merges']=[{s:{r:1,c:4},e:{r:2,c:4}}];
  const input=source({'对账明细-李奇':merged,...Object.fromEntries(['徐英','邢芳芳','操思敏','孙立柱'].map(p=>['对账明细-'+p,sheet([h,[p,'主体','',10,'否','2026-12-01']])])),旧月份汇总:sheet([h,['不能导入','主体','',999,'否','2026-12-01']])});
  const before=JSON.stringify(input),a=Import.parse(input,'ledger');
  assert.equal(a.records.length,6); assert.equal(a.sheets.filter(s=>!s.skipped).length,5);assert.equal(a.errors.length,0);
  assert.deepEqual(a.records.slice(0,2).map(r=>r.supplier),['正常甲','正常乙']); assert.equal(a.records[0].cents,0);
  assert.equal(a.records[1].status,'已支付'); assert.equal(a.records[0].owner,'李奇');assert.equal(a.records[0].category,'正常付款-正常对账表');
  assert.equal(JSON.stringify(input),before);
});
test('raw ledger invalid, missing and uncached classification is never silently normal', () => {
  const h=['供应商','申请付款金额','是否谈判付款','应付时间'];
  for(const flag of ['异常',{t:'e',v:'否'},{f:'A1',t:'z',v:'否'},{f:'A1'}]) {
    const a=Import.parse(source({'对账明细-徐英':sheet([h,['甲',100,flag,'2026-12-01']])}),'ledger');
    assert.equal(a.records.length,0);assert.equal(a.errors[0].code,'INVALID_NEGOTIATED_CLASSIFICATION');assert.equal(a.errors[0].cell,'C2');
  }
  const a=Import.parse(source({'对账明细-徐英':sheet([['供应商','申请付款金额','应付时间'],['甲',100,'2026-12-01']])}),'ledger');
  assert.equal(a.records.length,0);assert.equal(a.errors[0].code,'INVALID_NEGOTIATED_CLASSIFICATION');
});
test('ledger due date takes precedence and only truly blank due dates use estimates', () => {
  const h=['供应商','申请付款金额','是否谈判付款','应付款时间','预估付款月份','已付金额','是否结清'];
  const rows=[['甲',100,'否','2026-12-10','2027年1月',100,'是'],['乙',100,'否',' ','1月',0,'否'],['丙',100,'否','未发货','2027年1月'],['丁',100,'否',{f:'A1',t:'s',v:''},'2027年1月'],['戊',100,'否',{t:'e',v:7,w:'#DIV/0!'},'2027年1月']];
  const a=Import.parse(source({'对账明细-李奇':sheet([h,...rows])}),'ledger',{budgetStartMonth:'2026-12'});
  assert.deepEqual(a.records.map(r=>r.monthKey),['2026-12','2027-01','','','']);assert.equal(a.records[0].status,'已支付');
  assert.equal(a.records[2].month,'未发货');assert.equal(a.warnings.filter(w=>w.code==='UNKNOWN_MONTH_YEAR').length,3);
});
test('four forecast sheets with empty classification are merged, preserve explicit years and fill owners', () => {
  const h=['预估付款月份','供应商名称','付款主体','申请付款金额','已支付未支付','分类','跟单员'];
  const input=source(Object.fromEntries(['李奇','徐英','操思敏','邢芳芳'].map((p,i)=>['已发货未登记到资金台账-'+p,sheet([h,[["1月","5月","","2026.4月"][i],p,'主体',10,'未支付','','']])])));
  const a=Import.parse(input,'budget',{budgetStartMonth:'2026-12'});
  assert.equal(a.records.length,4);assert.equal(a.errors.length,0);assert.deepEqual(a.records.map(r=>r.monthKey),['2027-01','2027-05','','2026-04']);
  assert.deepEqual(a.records.map(r=>r.owner),['李奇','徐英','操思敏','邢芳芳']);assert(a.records.every(r=>r.status==='未支付'&&r.category==='正常付款-预估付款'));
});
test('month-only values roll within twelve months from the budget start without replacing explicit dates', () => {
  const input=source({预估预算:sheet([headers,...['12月','1月','5月','11月','2026.4月',0,'2027-02-30'].map(m=>[m,'甲','','','',10])])});
  const a=Import.parse(input,'budget',{budgetStartMonth:'2026-12'});
  assert.deepEqual(a.records.map(r=>r.monthKey),['2026-12','2027-01','2027-05','2027-11','2026-04','','']);
});
test('manual budget retains comparison rows, identifies only explicit forecasts and normalizes their classification on both sides', () => {
  const input=source({Sheet1:sheet([manualHeaders,
    ['正常付款-正常对账表','甲','供应甲','主体','','原文',100,'2026-12-10','上周五没登记，此款为预估'],
    ['正常付款-预估付款','乙','多家供应商','主体','','',200,'2026-12-31','新增'],
    ['谈判付款-谈判付款表','丙','供应丙','主体','OA3','',300,'2026-12-15','原11月支付改12月'],
    ['非预估付款','丁','供应丁','主体','','',0,'2026-12-15','预估时间待确认']]),
    其他明细:sheet([manualHeaders,['正常付款-正常对账表','','不纳入','','','',999,'2026-12-01']])});
  const before=JSON.stringify(input),a=Import.parse(input,'manualBudget');
  assert.equal(a.records.length,4);assert.deepEqual(a.records.map(r=>r.isForecast),[true,true,false,false]);
  assert.deepEqual(a.records.map(r=>r.category),['正常付款-预估付款','正常付款-预估付款','谈判付款-谈判付款表','非预估付款']);
  assert.match(a.records[0].notes,/原分类：正常付款-正常对账表/);assert.match(a.records[0].notes,/此款为预估/);assert.equal(a.evidence.length,0);
  assert.equal(JSON.stringify(input),before);assert.equal(Import.inferManualMonth(input),'2026-12');
});
test('manual month inference handles numeric dates and requires one valid month and one eligible source', () => {
  const make=(dates,extra={})=>source({Sheet1:sheet([manualHeaders,...dates.map(d=>['正常付款-正常对账表','','甲','','','',1,d])])},extra);
  assert.equal(Import.inferManualMonth(make([46357,46387])),'2026-12');
  assert.equal(Import.inferManualMonth(make([44895],{Workbook:{WBProps:{date1904:true}}})),'2026-12');
  assert.equal(Import.inferManualMonth(make(['12月',''])),'');assert.equal(Import.inferManualMonth(make(['2026-12-01','2027-01-01'])),'');
  assert.equal(Import.inferManualMonth(make([{t:'e',v:42}])),'');assert.equal(Import.inferManualMonth(make([{f:'A1',t:'z',v:46357}])),'');
  assert.equal(Import.inferManualMonth(source({甲:sheet([manualHeaders,['','','甲','','','',1,'2026-12-01']]),乙:sheet([manualHeaders,['','','乙','','','',1,'2026-12-01']])})), '');
});
test('amount errors remain tied to record and cell so excluding only that source can resolve them', () => {
  for(const amount of [null,{t:'n',f:'SUM(Z1:Z2)'},{t:'e',v:7,w:'#DIV/0!'},'金额待定']) {
    const input=source({Sheet1:sheet([manualHeaders,['正常付款-预估付款','','甲','','','',amount,'2026-12-01']])}),a=Import.parse(input,'manualBudget');
    assert.equal(a.records.length,1);assert.equal(a.records[0].cents,null);assert.equal(a.records[0].invalidAmount,true);
    assert.equal(a.errors[0].source.role,'manualBudget');assert.equal(a.errors[0].source.sheet,'Sheet1');assert.equal(a.errors[0].source.row,2);assert.equal(a.errors[0].cell,'G2');assert.equal(a.errors[0].recordId,a.records[0].id);
  }
  const missingBoth=Import.parse(source({Sheet1:sheet([manualHeaders,['正常付款-预估付款','','甲','','','',null,'']])}),'manualBudget');
  assert.equal(missingBoth.records.length,1);assert.equal(missingBoth.records[0].monthKey,'');assert.equal(missingBoth.errors[0].code,'MISSING_AMOUNT');
});
test('same physical workbook in different roles has distinct records but one payment evidence identity', () => {
  const input=source({明细:sheet([headers,['2026-12','甲','主体','OA1','',100,'已支付']])});
  const budget=Import.parse(input,'budget'),ledger=Import.parse(input,'ledger');
  assert.notEqual(budget.records[0].id,ledger.records[0].id);assert.equal(budget.evidence[0].id,ledger.evidence[0].id);
});
test('legacy installment header years take precedence over rolling budget year and conflicts stay unresolved', () => {
  const h=['供应商','付款主体','对公付款OA流程号','谈判明细','2026-11-16','2026-12-14'];
  const input=source({谈判付款:sheet([h,['甲','主体','OA1','11月25日-100\n12月25日-200',100,200]])});
  assert.deepEqual(Import.parse(input,'negotiated',{budgetStartMonth:'2026-12'}).records.map(r=>r.monthKey),['2026-11','2026-12']);
  const conflict=['供应商','付款主体','对公付款OA流程号','谈判明细','2026-11-30','2026-12-14','2027-12-14'];
  const other=source({谈判付款:sheet([conflict,['甲','主体','OA1','12月10日-100\n12月25日-200\n12月25日-300',100,200,300]])});
  assert.deepEqual(Import.parse(other,'negotiated',{budgetStartMonth:'2026-12'}).records.map(r=>r.monthKey),['','2026-12','2027-12']);
});
test('forecast notes must be affirmative, while an explicit forecast category remains authoritative', () => {
  const notes=['并非此款为预估','不是此款为预估','此款为预估已取消','可能此款为预估','上周五未登记，此款为预估','备注：此款为预估金额'];
  const input=source({Sheet1:sheet([manualHeaders,...notes.map(note=>['正常付款-正常对账表','','甲','','','',100,'2026-12-01',note]),['正常付款-预估付款','','乙','','','',100,'2026-12-01','并非此款为预估']])});
  assert.deepEqual(Import.parse(input,'manualBudget').records.map(r=>r.isForecast),[false,false,false,false,true,true,true]);
});
