const test = require('node:test');
const assert = require('node:assert/strict');
const X = require('../vendor/xlsx.full.min.js');
const C = require('../product-dimension-core.js');
const D = ['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组','人工分类'];
const M = ['使用组织','编码','领星SKU','SKU','名称','创建日期','物料分组','备案产品型号'];
function book(headers, rows, name='明细', prefix=[]) {
  const result = X.utils.book_new();
  X.utils.book_append_sheet(result, X.utils.aoa_to_sheet([...prefix, headers, ...rows]), name);
  return result;
}
const dimension = rows => book(D, rows);
const material = rows => book(M, rows);
const drow = (code, extra={}) => [code,'已有SKU',0,'原商品','2026/11/01','原分组',extra.manual ?? '人工保留'];
const mrow = (code, extra={}) => [extra.org ?? '组织一',code,extra.sku ?? '新领星',0,extra.name ?? '新商品','2026/11/14','新分组','不可自动映射'];

test('appends only absent codes, retains source values and never fills existing blanks', () => {
  const d = dimension([drow('001'), ['002','','','','','','手填']]);
  const m = material([mrow('001'),mrow('002'),mrow('003')]);
  const before = JSON.stringify(d);
  const analysis = C.analyze(d, m), result = C.resolve(analysis);
  assert.equal(result.canExport, true);
  assert.deepEqual(result.rows.map(row => row.code), ['001','002','003']);
  assert.deepEqual(result.rows[1].values, ['002','','','','','','手填']);
  assert.deepEqual(result.rows[2].values, ['003','新领星','0','新商品','2026/11/14','新分组','']);
  assert.equal(result.rows[0].cells[2].t, 'n');
  assert.equal(JSON.stringify(d), before);
});

test('zero is a valid code and leading zeros are never discarded', () => {
  const analysis = C.analyze(dimension([drow(0),drow('1')]), material([mrow('0'),mrow('0001'),mrow(' 002 ')]));
  const result = C.resolve(analysis);
  assert.deepEqual(result.rows.map(row => row.code), ['0','1','0001','002']);
  assert.equal(result.rows[0].values[0], '0');
  assert.equal(result.rows[0].cells[0].t, 's');
  assert.equal(result.rows[2].cells[0].t, 's');
});

test('existing duplicate codes always retain the first row and conflicts identify discarded rows', () => {
  const analysis = C.analyze(dimension([drow('1'),drow('1'),drow('2'),drow('2',{manual:'另一个人工值'})]), material([]));
  assert.equal(analysis.counts.dimensionDuplicateRows, 2);
  assert.equal(analysis.counts.dimensionConflictGroups, 1);
  const result = C.resolve(analysis);
  assert.equal(result.canExport, true);
  assert.equal(result.rows[1].sourceRow, 4);
  assert.equal(result.rows[1].values[6], '人工保留');
  assert.equal(analysis.issues[0].severity, 'warning');
  assert.equal(analysis.issues[0].keptRow, 4);
  assert.deepEqual(analysis.issues[0].droppedRows, [5]);
});

test('numeric and text values in existing duplicate rows are distinct', () => {
  const analysis = C.analyze(dimension([drow(123),drow('123')]), material([]));
  assert.equal(analysis.groups.length, 1);
  assert.equal(analysis.groups[0].conflict, true);
});

test('new materials compare only mapped fields across organizations and expose conflicts', () => {
  const analysis = C.analyze(dimension([]), material([mrow('1'),mrow('1',{org:'组织二'}),mrow('2'),mrow('2',{name:'不同名称'})]));
  assert.equal(analysis.added[0].conflict, false);
  assert.deepEqual(analysis.added[0].sourceRows, [2,3]);
  assert.equal(analysis.added[1].conflict, true);
  const result = C.resolve(analysis);
  assert.equal(result.canExport, false);
  assert.deepEqual(result.rows, []);
  assert.equal(analysis.issues[0].code, 'new-material-conflict');
});

test('reads raw text despite date number formats and compares formulas by cached values', () => {
  const d = dimension([drow('1'),drow('1')]);
  d.Sheets.明细.G2 = {t:'s',v:'手动轮椅',z:'m/d/yy',f:'"手动轮椅"'};
  d.Sheets.明细.G3 = {t:'s',v:'手动轮椅',z:'m/d/yy'};
  const analysis = C.analyze(d,material([]));
  assert.equal(analysis.groups[0].conflict, false);
  assert.equal(C.resolve(analysis).rows[0].values[6], '手动轮椅');
  assert.equal(C.resolve(analysis).rows[0].cells[6].f, '"手动轮椅"');
});

test('hidden and filtered rows are included, all-empty rows are skipped', () => {
  const d = dimension([drow('1'),[],drow('2')]);
  d.Sheets.明细['!rows'] = [null,null,null,{hidden:true}];
  d.Sheets.明细['!autofilter'] = {ref:'A1:G4'};
  assert.deepEqual(C.resolve(C.analyze(d,material([]))).rows.map(row => row.code), ['1','2']);
});

test('detects headers within first twenty rows and refuses missing, repeated or ambiguous headers', () => {
  const shifted = book(D,[drow('1')],'表一',[['说明'],[]]);
  assert.equal(C.analyze(shifted,material([])).headerRow, 3);
  assert.equal(C.analyze(book(['物料编码'],[['1']]),material([])).issues[0].code, 'missing-headers');
  assert.equal(C.analyze(book([...D,'SKU'],[drow('1')]),material([])).issues[0].code, 'duplicate-header');
  X.utils.book_append_sheet(shifted, X.utils.aoa_to_sheet([D,drow('2')]), '表二');
  assert.equal(C.analyze(shifted,material([])).issues[0].code, 'ambiguous-sheet');
  assert.equal(C.resolve(C.analyze(shifted,material([]))).canExport, false);
});

test('nonempty rows without a code block output with source location', () => {
  for (const role of ['dimension','material']) {
    const analysis = C.analyze(dimension(role === 'dimension' ? [drow('')] : []),material(role === 'material' ? [mrow('')] : []));
    const problem = analysis.issues.find(issue => issue.code === 'empty-code');
    assert.equal(problem.row, 2);
    assert.equal(problem.role, role);
    assert.equal(C.resolve(analysis).canExport, false);
    assert.deepEqual(C.resolve(analysis).rows, []);
  }
});

test('errors, invalid codes and missing formula results block output; empty cached formula is valid', () => {
  for (const [cell, expected] of [
    [{t:'e',v:7},'cell-error'], [{t:'n',f:'1+2'},'missing-formula-cache'],
    [{t:'n',v:Number.MAX_SAFE_INTEGER+1},'invalid-code'], [{t:'n',v:1.5},'invalid-code'],
    [{t:'b',v:true},'invalid-code'], [{t:'n',v:Infinity},'invalid-number']
  ]) {
    const d = dimension([drow('1')]); d.Sheets.明细.A2 = cell;
    const analysis = C.analyze(d,material([]));
    assert.ok(analysis.issues.some(issue => issue.code === expected));
    assert.equal(C.resolve(analysis).canExport, false);
  }
  const d = dimension([drow('1')]);
  d.Sheets.明细.G2 = {t:'s',f:'IF(A2="1","",1)',v:''};
  assert.equal(C.resolve(C.analyze(d,material([]))).canExport, true);
  d.Sheets.明细.G2 = {t:'e',v:23,f:'1/0'};
  assert.equal(C.analyze(d,material([])).issues[0].address, 'G2');
});

test('material errors block only code or mapped values that will actually be appended', () => {
  const d = dimension([drow('existing')]);
  const m = book([...M,'UPC'],[mrow('new'),mrow('existing')]);
  m.Sheets.明细.I2 = {t:'e',v:23};
  m.Sheets.明细.E3 = {t:'n',f:'1+2'};
  assert.equal(C.resolve(C.analyze(d,m)).canExport, true);
  m.Sheets.明细.E2 = {t:'e',v:23};
  const analysis = C.analyze(d,m);
  assert.equal(C.resolve(analysis).canExport, false);
  assert.deepEqual(analysis.issues.map(issue => issue.address), ['E2']);
});

test('new numeric creation dates retain the calendar date across 1900 and 1904 workbooks', () => {
  for (const source1904 of [false,true]) {
    const d = dimension([drow('existing')]), m = material([mrow('new')]);
    d.Workbook = {WBProps:{date1904:!source1904}};
    m.Workbook = {WBProps:{date1904:source1904}};
    m.Sheets.明细.F2 = {t:'n',v:45000.5,z:'yyyy/mm/dd hh:mm',w:'旧缓存'};
    const result = C.resolve(C.analyze(d,m));
    assert.equal(result.canExport, true);
    assert.equal(result.rows[1].values[4], 45000.5 + (source1904 ? 1462 : -1462));
    const before = X.SSF.parse_date_code(45000.5,{date1904:source1904});
    const after = X.SSF.parse_date_code(result.rows[1].values[4],{date1904:!source1904});
    for (const part of ['y','m','d','H','M','S']) assert.equal(before[part],after[part]);
    assert.equal(result.rows[1].cells[4].z, 'yyyy/mm/dd hh:mm');
    assert.equal(result.rows[1].cells[4].w, undefined);
    assert.equal(m.Sheets.明细.F2.v, 45000.5);
    assert.equal(result.rows[0].values[4], '2026/11/01');
  }
  const d = dimension([]), m = material([mrow('new')]);
  d.Workbook = {WBProps:{date1904:true}};
  m.Sheets.明细.F2 = {t:'n',v:1,z:'yyyy/mm/dd'};
  const analysis = C.analyze(d,m);
  assert.equal(C.resolve(analysis).canExport, false);
  assert.equal(analysis.issues[0].code, 'incompatible-date-system');
  assert.equal(analysis.issues[0].address, 'F2');
});
