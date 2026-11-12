const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const XLSX = require('../vendor/xlsx.full.min.js');

function loadApp() {
  const elements = new Map();
  const context = vm.createContext({
    console,
    document: { querySelector(selector) {
      if (!elements.has(selector)) elements.set(selector, {});
      return elements.get(selector);
    } },
    window: { XLSX: { ...XLSX } },
  });
  const source = fs.readFileSync(path.join(__dirname, '../file-library.js'), 'utf8');
  vm.runInContext(source.replace(/\ninit\(\)\.catch\([\s\S]*$/, ''), context);
  // Unit checks inspect the workbook passed to the styled writer; browser checks verify the actual ZIP/styles.
  context.window.JSZip = {};
  context.writeStyledDetailWorkbook = async (workbook, shapes, filename) => context.window.XLSX.writeFile(workbook, filename);
  return { context, elements, run: (code) => vm.runInContext(code, context) };
}

function allocationRecord(rows, sheetName = '需求') {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  return { name: `${sheetName}.xlsx`, file: { name: `${sheetName}.xlsx`, arrayBuffer: async () => bytes } };
}

test('division OA-to-correct mapping drives allocation and exports without mixing suppliers of the corrected code', async () => {
  const { context, run } = loadApp();
  context.records = {
    category: allocationRecord([
      ['物料编码', '', 'SKU', '物料名称', '销售系列', '采购组'],
      ['1935010018', '', 'SKU-1', '正确物料一', '系列一', '组一'],
      ['0002', '', 'SKU-2', '正确物料二', '系列二', '组二'],
    ], '商品'),
    singapore: allocationRecord([['供应商名称', '系列'], ['供应商甲', '系列一']], '新加坡'),
  };
  const division = [
    ['采购组', '备货OA物料编码', ' 正确物料编码 ', '采购下单人', '供应商简称', '供应商名称', '起订量', '物料编码'],
    ['组一', '1935010017', '1935010018', '徐英', '供应商甲', '供应商甲', 5, '错误旧编码'],
    ['组一', '1935010007', '1935010018', '徐英', '供应商甲', '供应商甲', 5, '错误旧编码'],
    ['组二', '0088', '0002', '李奇', '供应商乙', '供应商乙', 6, '错误旧编码'],
    ['其他组', '1935010018', '1935010018', '孙立柱', '其他供应商', '其他供应商', 99, '错误旧编码'],
  ];
  const demand = [
    ['备货OA物料编码', '正确物料编码', '物料编码', '数量', '流程号'],
    ['1935010017', '需求表错误编码', '错误旧编码', 2, 'OA-1'],
    ['1935010007', '需求表错误编码', '错误旧编码', 3.5, 'OA-2'],
    ['1935010017', '需求表错误编码', '错误旧编码', 0, 'OA-3'],
    ['0088', '需求表错误编码', '错误旧编码', 1, 'OA-4'],
  ];
  for (const reverse of [false, true]) {
    context.records.purchaseDivision = allocationRecord(division.map(row => reverse ? [...row].reverse() : row), '产品线明细');
    context.records.demand = allocationRecord(demand.map(row => reverse ? [...row].reverse() : row));
    const rows = await run('buildDemandAllocationRows(records)');
    assert.deepEqual(Array.from(rows, r => [r.materialCode, r.oaMaterialCode, r.quantityNumber, r.buyer, r.sku, r.supplierShortName, r.minimumOrderQuantity, r.minimumOrderStatus, r.singaporeOrder]), [
      ['1935010018', '1935010017、1935010007', 5.5, '徐英', 'SKU-1', '供应商甲', '5', '满足', '是'],
      ['0002', '0088', 1, '李奇', 'SKU-2', '供应商乙', '6', '不满足', '否'],
    ]);
    assert.deepEqual(Array.from(rows[0].sourceRows, r => r.oaMaterialCode), ['1935010017', '1935010007', '1935010017']);
    context.rows = rows;
    for (const combined of [false, true]) {
      const { workbook } = run(`buildDetailWorkbook(rows, ${combined})`);
      const reopened = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
      const output = reopened.SheetNames.flatMap(name => XLSX.utils.sheet_to_json(reopened.Sheets[name], { header: 1 }).slice(1));
      assert.deepEqual(output.map(row => [row[6], row[9]]).sort(), [['0002', 1], ['1935010018', 5.5]]);
      for (const sheet of Object.values(reopened.Sheets)) {
        assert.equal(sheet.G2.t, 's');
        const labels = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
        assert.equal(labels.includes('备货OA的物料编码'), combined);
        if (combined) {
          assert.equal(labels[labels.indexOf('采购单订单号') + 1], '备货OA的物料编码');
          assert.equal(sheet.P2.t, 's');
        }
      }
      if (combined) assert.deepEqual(output.map(row => [row[6], row[15]]).sort(), [['0002', '0088'], ['1935010018', '1935010017、1935010007']]);
    }
  }
});

test('missing, blank, conflicting division mappings retain rows and report demand and division locations', async () => {
  const { context, elements, run } = loadApp();
  context.records = {
    purchaseDivision: allocationRecord([[], ['备货OA物料编码', '正确物料编码', '采购下单人'], [], ['0099', '', '旧采购人']], '产品线明细'),
    demand: allocationRecord([[], ['物料编码', '数量', '申请人'], [], ['0099', 0, '采购甲'], ['0099', 1.5, '采购甲'], ['0088', '', '']]),
  };
  let rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows.length, 3);
  assert.deepEqual(Array.from(rows, r => [r.materialCode, r.oaMaterialCode, r.buyer, r.sourceRowNumber, r.materialCodeValid]), [
    ['', '0099', '', 4, '待核对'], ['', '0099', '', 5, '待核对'], ['', '0088', '', 6, '待核对'],
  ]);
  assert.equal(rows[0].quantityNumber, 0);
  assert.equal(rows[1].quantityNumber, 1.5);
  assert(Number.isNaN(rows[2].quantityNumber));
  assert.match(rows[2].materialCodeIssue, /未在.*找到.*0088/);
  context.rows = rows;
  run('state.demandRows = rows; renderMaterialCodeWarnings()');
  const warnings = elements.get('#materialCodeWarnings');
  assert.equal(warnings.hidden, false);
  assert.match(warnings.innerHTML, /需求.xlsx \/ 需求 第 4 行.*分工表“正确物料编码”为空/);
  assert.match(warnings.innerHTML, /产品线明细.xlsx \/ 产品线明细 第 4 行/);
  assert.match(warnings.innerHTML, /3 条物料编码待核对/);
  const { workbook } = run('buildDetailWorkbook(rows, true)');
  assert.equal(XLSX.utils.decode_range(workbook.Sheets.未填写下单人['!ref']).e.r, 3);
  context.records.purchaseDivision = allocationRecord([['备货OA物料编码', '采购下单人'], ['0099', '采购甲']], '产品线明细');
  rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows[0].materialCode, '');
  assert.match(rows[0].materialCodeIssue, /分工表缺少“正确物料编码”列/);
  context.records.purchaseDivision = allocationRecord([['备货OA物料编码', '正确物料编码'], ['0099', '0001'], ['0099', '0002']], '产品线明细');
  rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows[0].materialCode, '');
  assert.match(rows[0].materialCodeIssue, /对应多个正确物料编码.*0001.*0002/);
  context.records.demand = allocationRecord([['正确物料编码', '数量'], ['0001', 2]]);
  rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows[0].materialCode, '');
  assert.equal(rows[0].oaMaterialCode, '', 'demand corrected-code column must not become the original-code key');
  assert.match(rows[0].materialCodeIssue, /缺少原编码列/);
  context.records.demand = allocationRecord([['识别码', '数量'], ['0099', 2]]);
  context.records.purchaseDivision = allocationRecord([['正确物料编码', '备货OA物料编码'], ['0001', '0099'], ['0001', '0099']], '产品线明细');
  rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows[0].materialCode, '0001', 'identical mapping duplicates do not create a conflict');
  assert.equal(rows[0].quantityNumber, 2);
  context.rows = rows;
  run('state.demandRows = rows; renderMaterialCodeWarnings()');
  assert.equal(warnings.hidden, true);
});

test('original purchase group resolves division candidates and merged codes retain all matched suppliers', async () => {
  const { context, run } = loadApp();
  context.records = {
    category: allocationRecord([['物料编码', '采购组'], ['OA-X', '组二']], '商品'),
    purchaseDivision: allocationRecord([
      ['备货OA物料编码', '正确物料编码', '采购组', '采购下单人', '供应商'],
      ['OA-X', '0001', '组一', '采购一', '供应商甲'],
      ['OA-X', '0002', '组二', '采购二', '供应商乙'],
      ['0002', '0002', '其他组', '无关采购', '无关供应商'],
    ], '产品线明细'),
    demand: allocationRecord([['物料编码', '数量'], ['OA-X', 2]]),
  };
  const rows = await run('buildDemandAllocationRows(records)');
  assert.deepEqual(Array.from(rows, row => [row.materialCode, row.buyer, row.supplierShortName, row.quantityNumber]), [['0002', '采购二', '供应商乙', 2]]);
  context.rows = [{ materialCode: '0002', quantityNumber: 2, supplierNames: ['供应商甲'], salesSeries: '系列甲' }, { materialCode: '0002', quantityNumber: 3, supplierNames: ['供应商乙'], salesSeries: '系列甲' }];
  const merged = run('mergeDemandRowsByMaterial(rows)');
  assert.deepEqual(Array.from(merged[0].supplierNames), ['供应商甲', '供应商乙']);
  context.row = merged[0];
  context.index = new Map([['供应商甲', new Set(['系列甲'])]]);
  assert.equal(run('getSingaporeOrderStatus(row, index)'), '待核对');
});

test('legacy material headers retain original codes for combined download', async () => {
  const { context, run } = loadApp();
  context.records = {};
  for (const header of ['物料编码', '识别码', '物料编号']) {
    context.records.demand = allocationRecord([[header, '数量'], ['00001', 0]]);
    const rows = await run('buildDemandAllocationRows(records)');
    assert.deepEqual(Array.from(rows, r => [r.materialCode, r.oaMaterialCode, r.quantityNumber, r.materialCodeIssue]), [['00001', '00001', 0, '']]);
  }
});

test('新加坡下单按供应商与系列成对匹配，空白及多供应商不猜测', () => {
  const { context, run } = loadApp();
  context.tables = [{ rows: [['系列', '供应商名称'], ['Z12', '供应商甲'], ['AY03', '供应商乙'], ['Z12', '供应商甲'], ['', '']] }];
  run('globalThis.singaporeIndex = buildSingaporeIndex(tables)');
  const status = (suppliers, series) => {
    context.row = { supplierNames: suppliers, salesSeries: series };
    return run('getSingaporeOrderStatus(row, singaporeIndex)');
  };
  assert.equal(status([' 供应商甲 '], ' z12 '), '是');
  assert.equal(status(['供应商甲'], 'AY03'), '否');
  assert.equal(status(['供应商乙'], 'Z12'), '否');
  assert.equal(status(['供应商丙'], 'Z12'), '否');
  assert.equal(status(['供应商甲', '供应商乙'], 'Z12'), '待核对');
  assert.equal(status(['供应商乙', '供应商丙'], 'Z12'), '否');
  assert.equal(status([], 'Z12'), '待核对');
  assert.equal(status(['供应商甲'], ''), '待核对');
  assert.equal(run('getSingaporeOrderStatus(row, null)'), '');
  run('singaporeIndex.get("供应商乙").add("z12")');
  assert.equal(status(['供应商甲', '供应商乙'], 'Z12'), '是');
});

test('新加坡名录缺表头或配对字段不完整时明确报错，不输出旧判断', () => {
  const { context, run } = loadApp();
  context.tables = [{ rows: [['供应商名称', '备注'], ['供应商甲', 'Z12']] }];
  assert.throws(() => run('buildSingaporeIndex(tables)'), /未找到.*供应商名称.*系列/);
  context.tables = [{ rows: [['供应商名称', '系列'], ['供应商甲', '']], source: { fileName: '新加坡.xlsx', sheetName: '明细' } }];
  assert.throws(() => run('buildSingaporeIndex(tables)'), /新加坡.xlsx.*明细.*第 2 行.*供应商名称或系列为空/);
  context.tables = [{ rows: [['供应商名称', '系列']] }];
  run('globalThis.emptyIndex = buildSingaporeIndex(tables)');
  context.row = { supplierNames: ['供应商甲'], salesSeries: 'Z12' };
  assert.equal(run('getSingaporeOrderStatus(row, emptyIndex)'), '否');
});

test('新加坡文件内容异常时不替换原文件、不丢手填天数和筛选', async () => {
  const { context, elements, run } = loadApp();
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['供应商名称', '系列'], ['供应商甲', '']]), '明细');
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  context.badFile = { name: '错误新加坡.xlsx', size: bytes.length, arrayBuffer: async () => bytes };
  run('state.records.set("file-5", { name: "正确新加坡.xlsx", applied: true }); state.demandRows = [{ materialCode: "001", deliveryDays: 9, singaporeOrder: "是" }]; state.filters.get("buyer").add("李奇")');
  await run('savePendingFile("file-5", badFile)');
  assert.equal(run('state.records.get("file-5").name'), '正确新加坡.xlsx');
  assert.equal(run('state.demandRows[0].deliveryDays'), 9);
  assert.equal(run('state.demandRows[0].singaporeOrder'), '是');
  assert.equal(run('state.filters.get("buyer").has("李奇")'), true);
  assert.match(elements.get('#libraryState').textContent, /新加坡公司明细未更新，原文件保留.*供应商名称或系列为空/);
});

test('新加坡判定读取分工全称和商品系列，合并及两种导出一致且不改采购主体', async () => {
  const { context, run } = loadApp();
  const record = (rows, sheetName) => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
    const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    return { name: `${sheetName}.xlsx`, file: { name: `${sheetName}.xlsx`, arrayBuffer: async () => bytes } };
  };
  context.records = {
    category: record([['物料编码', '销售系列'], ['001', 'Z12'], ['002', 'AY03']], '商品'),
    purchaseDivision: record([['物料编码', '供应商简称', '采购下单人', '供应商名称'], ['001', '甲简称', '李奇', '供应商甲'], ['002', '甲简称', '徐英', '供应商甲']], '产品线明细'),
    demand: record([['物料编码', '数量', '采购主体'], ['001', 2, '原采购主体'], ['001', 0, '原采购主体'], ['002', 3.5, '原采购主体']], '需求'),
    singapore: record([['供应商名称', '系列'], ['供应商甲', 'Z12']], 'Dim-新加坡公司下单'),
  };
  const rows = await run('buildDemandAllocationRows(records)');
  assert.deepEqual(Array.from(rows, row => [row.materialCode, row.quantityNumber, row.singaporeOrder, row.purchaseEntity]), [['001', 2, '是', '原采购主体'], ['002', 3.5, '否', '原采购主体']]);
  context.rows = rows;
  run('state.demandRows = rows; state.purchaseEntityOverride = "手填主体"');
  assert.deepEqual(Array.from(run('getFilteredDemandRows()'), row => [row.singaporeOrder, row.purchaseEntity]), [['是', '原采购主体'], ['否', '原采购主体']]);
  for (const split of [false, true]) {
    const { workbook } = run(`buildDetailWorkbook(getFilteredDemandRows(), ${split})`);
    const reopened = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
    const data = reopened.SheetNames.flatMap(name => XLSX.utils.sheet_to_json(reopened.Sheets[name], { header: 1 }).slice(1));
    assert.deepEqual(data.map(row => [row[6], row[4], row[1]]), [['001', '原采购主体', '是'], ['002', '原采购主体', '否']]);
  }
  context.records.singapore = null;
  assert.deepEqual(Array.from(await run('buildDemandAllocationRows(records)'), row => row.singaporeOrder), ['', '']);
});

test('preview and export table headers, rendered values and exported cells follow the requested order', async () => {
  const { context, run } = loadApp();
  const expected = [
    ['采购单订单下单人', 'buyer'], ['是否新加坡公司下单', 'singaporeOrder'], ['事业部', 'businessUnit'], ['申请人', 'applicant'],
    ['采购主体', 'purchaseEntity'], ['供应商简称', 'supplierShortName'], ['正确物料编码', 'materialCode'], ['SKU', 'sku'],
    ['物料名称', 'materialName'], ['数量', 'quantity'], ['OA备货流程号', 'oaProcessNo'],
    ['采购分工明细是否存在', 'materialCodeValid'],
    ['今天日期', 'todayDate'], ['天数', 'deliveryDays'], ['供应商反馈交期', 'supplierDeliveryDate'], ['采购单订单号', 'purchaseOrderNo'], ['要求货好时间', 'requiredReadyDate'], ['起订量', 'minimumOrderQuantity'],
    ['起订量是否满足', 'minimumOrderStatus'], ['备货原因', 'stockReason'],
  ];
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const header = html.match(/<table class="detail-table">([\s\S]*?)<\/thead>/)[1];
  assert.deepEqual([...header.matchAll(/<th>(.*?)<\/th>/g)].map(match => match[1]), expected.filter(item => item[1] !== 'stockReason').map(item => item[0]));
  context.row = Object.fromEntries(expected.map(([, key], index) => [key, `value-${index + 1}`]));
  const cells = [...run('renderDemandTableRow(row)').matchAll(/<td>(.*?)<\/td>/g)].map(match => match[1].startsWith('<input') ? match[1].match(/value="([^"]*)"/)[1] : match[1]);
  assert.deepEqual(cells, expected.map(([, key], index) => key === 'stockReason' ? null : `value-${index + 1}`).filter(value => value !== null));
  let exported;
  context.window.XLSX.writeFile = (output) => { exported = output; };
  run('state.filteredRows = [row]');
  await run('downloadDetailWorkbook()');
  const output = XLSX.utils.sheet_to_json(exported.Sheets[exported.SheetNames[0]], { header: 1 });
  assert.deepEqual(output[0], expected.filter(item => item[1] !== "materialCodeValid").map(item => item[0]));
  assert.deepEqual(output[1], expected.map(([,key],i) => key === 'quantity' ? '' : `value-${i+1}`).filter((_,i) => i !== 11));
});

test('per-buyer workbooks preserve rows, reasons, formulas and WPS highlight',()=>{
 const {context,run}=loadApp();
 context.rows=[{buyer:'甲',quantity:0,stockReason:'需要保留',todayDate:'2026-10-21',deliveryDays:2,supplierDeliveryDate:'2026-10-23'},{buyer:'甲',quantity:3,stockReason:'另一个原因'}];
 const output=run('buildDetailWorkbook(rows)');
 const w=XLSX.read(XLSX.write(output.workbook,{type:'buffer',bookType:'xlsx'}),{type:'buffer'});
 assert.equal(w.Custprops.KSOReadingLayout,true);assert.equal(w.Sheets.备货需求分配明细.S2.v,'需要保留');assert.equal(w.Sheets.备货需求分配明细.J2.v,0);assert.equal(w.Sheets.备货需求分配明细.N2.f,'IF(M2="","",TEXT(DATE(VALUE(LEFT(L2,4)),VALUE(MID(L2,6,2)),VALUE(RIGHT(L2,2)))+M2,"yyyy-mm-dd"))');assert.equal(w.Sheets.备货需求分配明细['!ref'],'A1:S3');
});

test('all exported sheets use single-line text, full-text widths and no-wrap centered styles', () => {
  const { context, run } = loadApp();
  const reason = '甲'.repeat(60) + '\r\n' + '乙'.repeat(50);
  context.rows = [
    { buyer: '李奇', materialCode: '0001', materialName: '第一段\n第二段', sku: 'SKU\r型号', stockReason: reason, quantityNumber: 0 },
    { buyer: '徐英', materialCode: '0002', materialName: '一\u2028二\u2029三', stockReason: '', quantityNumber: 1.5 },
  ];
  for (const split of [false, true]) {
    const { workbook } = run(`buildDetailWorkbook(rows, ${split})`);
    assert.equal(workbook.Sheets[workbook.SheetNames[0]]['!cols'][split ? 19 : 18].wch, 223);
    const reopened = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellStyles: true });
    const data = reopened.SheetNames.flatMap(name => XLSX.utils.sheet_to_json(reopened.Sheets[name], { header: 1 }).slice(1));
    assert.deepEqual(data.map(row => [row[6], row[7], row[8], row[9], row[split ? 19 : 18]]), [
      ['0001', 'SKU 型号', '第一段 第二段', 0, reason.replace(/\r\n/, ' ')],
      ['0002', '', '一 二 三', 1.5, ''],
    ]);
    for (const name of reopened.SheetNames) {
      for (const cell of Object.values(reopened.Sheets[name])) {
        if (cell?.t === 's') assert(!/[\r\n\u2028\u2029]/.test(cell.v));
      }
    }
  }
  assert.equal(context.rows[0].stockReason, reason, 'source values remain unchanged');
  const alignments = [...run('buildDetailWorkbookStylesXml()').matchAll(/<alignment\b[^>]*\/>/g)].map(match => match[0]);
  assert.equal(alignments.length, 5);
  for (const alignment of alignments) {
    assert.match(alignment, /wrapText="0"/);
    assert.match(alignment, /horizontal="center"/);
    assert.match(alignment, /vertical="center"/);
  }
  const styles=run('buildDetailWorkbookStylesXml()');
  const borders=[...styles.matchAll(/<border>([\s\S]*?)<\/border>/g)].map(match=>match[1]);
  assert.equal(borders.length,3);
  for(const [index,line] of [[1,'thin'],[2,'hair']]) {
    for(const side of ['left','right','top','bottom']) assert.match(borders[index],new RegExp(`<${side} style="${line}"><color rgb="FF404040"/>`));
  }
  const xfs=[...styles.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf\b[^>]*>/g)].map(match=>match[0]);
  assert.match(xfs[1],/borderId="1"/);
  for(const xf of xfs.slice(2)) assert.match(xf,/borderId="2"/);
});

test('one workbook splits buyers into sheets with complete rows and sheet-local formulas', () => {
  const { context, run } = loadApp();
  context.rows = [
    { buyer: '李奇', materialCode: '0001', sku: '001', quantityNumber: 0, stockReason: '原因一', todayDate: '2026-11-05', deliveryDays: 2, supplierDeliveryDate: '2026-11-07' },
    { buyer: '徐英', materialCode: '0002', quantityNumber: 1.5, stockReason: '原因二', todayDate: '2026-11-05', deliveryDays: 0, supplierDeliveryDate: '2026-11-05' },
    { buyer: '李奇', materialCode: '0003', quantity: '', stockReason: '原因三', todayDate: '2026-11-05' },
    { buyer: '', materialCode: '0004', quantityNumber: 4 },
  ];
  const { workbook, sheetShapes } = run('buildDetailWorkbook(rows, true)');
  const reopened = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellNF: true, cellStyles: true });
  assert.equal(reopened.SheetNames.length, 3);
  const li = reopened.Sheets.李奇, xu = reopened.Sheets.徐英;
  assert.equal(reopened.Custprops.KSOReadingLayout, true);
  assert.deepEqual([li.G2.v, li.G3.v, xu.G2.v, reopened.Sheets.未填写下单人.G2.v], ['0001', '0003', '0002', '0004']);
  assert.equal(li.H2.v, '001');
  assert.deepEqual([li.J2.t, li.J2.v, li.J2.z, xu.J2.v, li.J3.v], ['n', 0, '0', 1.5, '']);
  assert.equal(li.T3.v, '原因三');
  assert.equal(li.L2.f, 'TEXT(TODAY(),"yyyy-mm-dd")');
  assert.equal(li.N2.f, 'IF(M2="","",TEXT(DATE(VALUE(LEFT(L2,4)),VALUE(MID(L2,6,2)),VALUE(RIGHT(L2,2)))+M2,"yyyy-mm-dd"))');
  assert.equal(li.N3.f, 'IF(M3="","",TEXT(DATE(VALUE(LEFT(L3,4)),VALUE(MID(L3,6,2)),VALUE(RIGHT(L3,2)))+M3,"yyyy-mm-dd"))');
  assert.equal(xu.N2.f, 'IF(M2="","",TEXT(DATE(VALUE(LEFT(L2,4)),VALUE(MID(L2,6,2)),VALUE(RIGHT(L2,2)))+M2,"yyyy-mm-dd"))');
  assert.equal(xu.N2.z, '@');
  assert.equal(xu.N2.t, 's');assert.equal(li.L2.t, 's');
  reopened.SheetNames.forEach((name, index) => {
    const sheet = reopened.Sheets[name];
    assert.equal(sheet['!cols'].length, 20);
    assert.equal(sheetShapes[index].rowCount, XLSX.utils.decode_range(sheet['!ref']).e.r + 1);
    assert.equal(sheetShapes[index].columnCount, 20);
  });
});

test('buyer sheet names stay valid and unique without merging distinct buyers', () => {
  const { context, run } = loadApp();
  context.rows = ['', '未填写下单人', '甲/乙', '甲:乙', 'Buyer', 'buyer', '长'.repeat(40) + '甲', '长'.repeat(40) + '乙', "'名字'", '[]:?*\\/'].map((buyer, index) => ({ buyer, materialCode: String(index) }));
  const output = run('buildDetailWorkbook(rows, true)');
  const reopened = XLSX.read(XLSX.write(output.workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
  assert.equal(reopened.SheetNames.length, context.rows.length);
  assert.equal(new Set(reopened.SheetNames.map(name => name.toLowerCase())).size, context.rows.length);
  for (const name of reopened.SheetNames) {
    assert(name.length <= 31);
    assert(!/[\[\]:*?/\\\x00-\x1f']/.test(name));
  }
  assert.deepEqual(reopened.SheetNames.map(name => reopened.Sheets[name].G2.v).sort(), context.rows.map(row => row.materialCode).sort());
});

test('combined download exports only filtered rows and guards empty selection', async () => {
  const { context, run } = loadApp();
  let exported, fileName, alert;
  context.window.XLSX.writeFile = (workbook, name) => { exported = workbook; fileName = name; };
  context.window.alert = message => { alert = message; };
  await run('downloadDetailWorkbook(true)');
  assert.equal(alert, '暂无可下载的明细数据。');
  assert.equal(exported, undefined);
  context.rows = [{ buyer: '李奇', materialCode: '保留' }, { buyer: '徐英', materialCode: '排除' }];
  run('state.demandRows = rows; state.filteredRows = rows.slice(0, 1)');
  await run('downloadDetailWorkbook(true)');
  assert.deepEqual(exported.SheetNames, ['李奇']);
  assert.equal(exported.Sheets.李奇.G2.v, '保留');
  assert.match(fileName, /^\d{4}-\d{2}-\d{2}-备货分配分配\.xlsx$/);
});

test('missing ZIP exporter blocks every download mode instead of emitting an unstyled workbook',async()=>{
  for(const [combined,buyers] of [[false,['甲']],[true,['甲','乙']],[false,['甲','乙']]]) {
    const {context,run}=loadApp();let message='',writes=0;
    delete context.window.JSZip;
    context.window.alert=text=>{message=text;};
    context.window.XLSX.writeFile=()=>{writes++;};
    context.writeStyledDetailWorkbook=async()=>{writes++;};
    context.rows=buyers.map(buyer=>({buyer,quantityNumber:1}));run('state.filteredRows=rows');
    await run(`downloadDetailWorkbook(${combined})`);
    assert.equal(message,'Excel 导出组件未加载，请刷新页面后重试。');
    assert.equal(writes,0);
  }
});

test('quantity exports as numeric cells with zero-decimal formatting and preserves missing values', async () => {
  const { context, run } = loadApp();
  context.rows = [7, 0, -2, 1.5, Number.MAX_SAFE_INTEGER].map(quantityNumber => ({ quantityNumber, quantity: String(quantityNumber) }));
  context.rows.push({ quantity: '' }, { quantity: 'bad' }, { quantity: '1,234' });
  let exported;
  context.window.XLSX.writeFile = workbook => { exported = workbook; };
  run('state.filteredRows = rows');
  await run('downloadDetailWorkbook()');
  const reopened = XLSX.read(XLSX.write(exported, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellNF: true });
  const sheet = reopened.Sheets[reopened.SheetNames[0]];
  for (const [i, value] of [7, 0, -2, 1.5, Number.MAX_SAFE_INTEGER].entries()) {
    assert.equal(sheet[`J${i + 2}`].t, 'n');
    assert.equal(sheet[`J${i + 2}`].v, value);
    assert.equal(sheet[`J${i + 2}`].z, '0');
    assert(!sheet[`J${i + 2}`].w.includes('.'));
  }
  assert.equal(sheet.J7.v, '');
  assert.equal(sheet.J8.v, '');
  assert.equal(sheet.J9.v, 1234);
  assert.match(run('buildDetailWorkbookStylesXml()'), /numFmtId="1"[^>]+applyNumberFormat="1"/);
});

test('same material merges across sheets with unique OA numbers, summed quantity and matching export', async () => {
  const { context, elements, run } = loadApp();
  const workbook = XLSX.utils.book_new();
  const header = ['创建人', '数量', '要求货好时间', '识别码', '流程号', '备货原因', '采购主体'];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    header,
    ['甲', 3, '2026-10-01', 'TEST-001', 'OA001', '原因甲', '主体甲'],
    ['甲', 2, '2026-10-01', 'TEST-001', 'OA001', '原因甲', '主体甲'],
    ['甲', 4, '2026-10-01', 'TEST-002', 'OA003', '', ''],
  ]), '需求一');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    header, ['乙', 5, '2026-10-02', 'TEST-001', 'OA002', '原因乙', '主体乙'],
  ]), '需求二');
  const division = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(division, XLSX.utils.aoa_to_sheet([
    ['物料编码', '采购单订单下单人', '供应商简称', '起订量'],
    ['TEST-001', '采购甲', '供应商甲', 8],
  ]), '产品线明细');
  const record = (wb) => {
    const bytes = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    return { name: 'synthetic.xlsx', file: { name: 'synthetic.xlsx', arrayBuffer: async () => bytes } };
  };
  context.records = { demand: record(workbook), purchaseDivision: record(division) };
  const rows = await run('buildDemandAllocationRows(records)');
  assert.equal(rows.length, 2);
  assert.equal(rows[0].oaProcessNo, 'OA001/OA002');
  assert.equal(rows[0].quantityNumber, 10);
  assert.equal(rows[0].minimumOrderStatus, '满足');
  assert.equal(rows[0].applicant, '甲、乙');
  assert.equal(rows[0].requiredReadyDate, '2026-10-01、2026-10-02');
  assert.equal(rows[0].stockReason, '原因甲、原因乙');
  assert.equal(rows[0].purchaseEntity, '主体甲、主体乙');
  assert.equal(rows[0].sourceRows.length, 3);
  assert.equal(rows[0].supplierDeliveryDate, '');
  assert.equal(rows[0].purchaseOrderNo, '');
  assert.equal(rows[1].quantityNumber, 4);
  context.rows = rows;
  run('state.demandRows = rows; state.filteredRows = getFilteredDemandRows(); updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '14');
  assert.match(run('renderDemandTableRow(rows[0])'), /OA001\/OA002/);
  run('state.filters.set("buyer", new Set(["采购甲"])); state.filteredRows = getFilteredDemandRows(); updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '10');
  let exported;
  context.window.XLSX.writeFile = (output) => { exported = output; };
  await run('downloadDetailWorkbook()');
  assert.deepEqual(exported.SheetNames, ['备货需求分配明细']);
  const output = XLSX.utils.sheet_to_json(exported.Sheets['备货需求分配明细'], { header: 1 });
  assert.equal(output.length, 2);
  assert.equal(output[1][output[0].indexOf('数量')], 10);
  assert.equal(output[1][output[0].indexOf('起订量')], '8');
  assert.equal(output[1][output[0].indexOf('起订量是否满足')], '满足');
  assert.equal(output[1][output[0].indexOf('OA备货流程号')], 'OA001/OA002');
  assert.equal(output[1][output[0].indexOf('采购单订单号')], '');
});

test('business unit and corrected code jointly determine rows, quantities, MOQ and both exports', async () => {
  const { context, run } = loadApp();
  context.records = {
    purchaseDivision: allocationRecord([
      ['备货OA物料编码', '正确物料编码', '采购下单人', '供应商简称', '起订量'],
      ['OA-A', '0001', '李奇', '供应商甲', 10],
      ['OA-B', '0001', '李奇', '供应商甲', 10],
      ['OA-C', '0002', '李奇', '供应商甲', 10],
    ], '产品线明细'),
    roster: allocationRecord([['姓名', '一级部门'], ['甲', '海外事业一部'], ['乙', '国内事业部'], ['丙', '海外事业一部']], '花名册'),
    demand: allocationRecord([
      ['物料编码', '数量', '申请人', '流程号'],
      ['OA-A', 3, '甲', '流程一'], ['OA-B', 4, '丙', '流程二'],
      ['OA-A', 8, '乙', '流程三'], ['OA-A', 0, '乙', '流程四'],
      ['OA-A', 1.5, '未知人员', '流程五'], ['OA-C', 10, '甲', '流程六'],
    ]),
  };
  const rows = await run('buildDemandAllocationRows(records)');
  const expected = [['海外事业一部', '0001', 7], ['国内事业部', '0001', 8], ['', '0001', 1.5], ['海外事业一部', '0002', 10]];
  assert.deepEqual(Array.from(rows, r => [r.businessUnit, r.materialCode, r.quantityNumber]), expected);
  assert.deepEqual(Array.from(rows, r => r.oaProcessNo), ['流程一/流程二', '流程三/流程四', '流程五', '流程六']);
  assert.deepEqual(Array.from(rows, r => r.minimumOrderStatus), ['不满足', '不满足', '不满足', '满足']);
  assert.deepEqual(Array.from(rows, r => r.sourceRows.length), [2, 2, 1, 1]);
  assert.deepEqual(Array.from(rows, r => r.oaMaterialCode), ['OA-A、OA-B', 'OA-A', 'OA-A', 'OA-C']);
  context.rows = rows;
  for (const combined of [false, true]) {
    const { workbook } = run(`buildDetailWorkbook(rows, ${combined})`);
    const reopened = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
    const data = reopened.SheetNames.flatMap(name => XLSX.utils.sheet_to_json(reopened.Sheets[name], { header: 1, defval: '' }).slice(1));
    assert.deepEqual(data.map(r => [r[2], r[6], r[9]]), expected);
  }
  run('state.demandRows = rows; state.filters.set("businessUnit", new Set(["国内事业部"]))');
  assert.deepEqual(Array.from(run('getFilteredDemandRows()'), r => [r.businessUnit, r.quantityNumber]), [['国内事业部', 8]]);
});

test('merge preserves blank codes and handles unknown, zero, signed decimal and overflowing quantities', () => {
  const { context, elements, run } = loadApp();
  const row = (materialCode, quantityNumber, oaProcessNo = '') => ({ materialCode, quantityNumber, oaProcessNo, minimumOrderQuantity: '1' });
  context.rows = [
    row('A', 2, 'OA1'), row('A', NaN, 'OA2'),
    row('B', 0, 'OA3'), row('B', 0, 'OA3'),
    row('C', -2), row('C', 2.5),
    row('', 1, 'OA4'), row(' ', 2, 'OA5'),
    row('D', Number.MAX_VALUE), row('D', Number.MAX_VALUE),
    row('E', NaN), row('E', NaN),
  ];
  const rows = run('mergeDemandRowsByMaterial(rows)');
  assert.equal(rows.length, 7);
  assert(Number.isNaN(rows[0].quantityNumber));
  assert.equal(rows[0].quantity, '');
  assert.equal(rows[0].minimumOrderStatus, '');
  assert.equal(rows[1].quantity, '0');
  assert.equal(rows[1].oaProcessNo, 'OA3');
  assert.equal(rows[2].quantity, '0.5');
  assert.equal(rows[2].minimumOrderStatus, '不满足');
  assert.equal(rows[3].oaProcessNo, 'OA4');
  assert.equal(rows[4].oaProcessNo, 'OA5');
  assert(Number.isNaN(rows[5].quantityNumber));
  assert(Number.isNaN(rows[6].quantityNumber));
  context.merged = rows;
  run('state.filteredRows = merged; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '待核对');
});

test('duplicate quantity headers select demand detail, never SKU digits or transfer quantity', () => {
  const { context, run } = loadApp();
  context.headers = ['创建人', '数量', '品名', '数量', '要求货好时间', 'sku', '识别码'];
  context.row = ['测试申请人', 99, '测试产品', '7', '2026-09-25', 'YT02-WT-C', 'TEST-001'];
  assert.equal(run('findDemandQuantityColumnIndex(headers, { requiredReadyDate: 4 })'), 3);
  assert.equal(run('getDemandRowQuantity(row, { quantity: 3, materialCode: 6 })'), 7);
  assert.equal(run('findDemandQuantityColumnIndex(["识别码", "数量"], {})'), 1);
  assert.equal(run('findDemandQuantityColumnIndex(["数量", "数量", "识别码"], {})'), undefined);
  assert.equal(run('findDemandQuantityColumnIndex(["识别码", "sku"], {})'), undefined);
});

test('missing and invalid quantities never fall back; zero and valid signed decimals survive', () => {
  const { context, run } = loadApp();
  for (const [input, expected] of [[0, 0], ['0', 0], [' 7 ', 7], ['1,234.5', 1234.5], [-2, -2], ['+0.5', 0.5], [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]]) {
    context.input = input;
    assert.equal(run('getDemandRowQuantity([input, "YT02-WT-C", "2026-09-25"], { quantity: 0 })'), expected);
  }
  for (const input of ['', ' ', null, undefined, NaN, Infinity, 'YT02-WT-C', '2026-09-25', '7个', '1,2', '#N/A', 'Infinity']) {
    context.input = input;
    assert(Number.isNaN(run('getDemandRowQuantity([input, 7, "YT02-WT-C"], { quantity: 0 })')));
  }
  assert(Number.isNaN(run('getDemandRowQuantity([7, "YT02-WT-C"], {})')));
});

test('workbook import, rows, metrics, minimum-order comparison and export use detail quantities', async () => {
  const { context, elements, run } = loadApp();
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['创建人', '数量', '品名', '数量', '要求货好时间', 'sku', '识别码'],
    ['测试申请人', 99, '产品A', 7, '2026-09-25', 'YT02-WT-C', 'TEST-001'],
    ['测试申请人', 99, '产品B', 3, '2026-09-25', 'YT02-WT-PRO-C', 'TEST-002'],
    ['测试申请人', 99, '产品C', 0, '2026-09-25', 'SKU09', 'TEST-003'],
  ]), '需要备货');
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  context.records = { demand: { name: 'synthetic.xlsx', file: { name: 'synthetic.xlsx', arrayBuffer: async () => bytes } } };
  const rows = await run('buildDemandAllocationRows(records)');
  assert.deepEqual(Array.from(rows, row => row.quantityNumber), [7, 3, 0]);
  context.rows = rows;
  run('state.filteredRows = rows; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '10');
  assert.equal(run('getMinimumOrderStatus(rows[0].quantityNumber, 5)'), '满足');
  assert.match(run('renderDemandTableRow(rows[0])'), /<td>7<\/td>/);
  let exported;
  context.window.XLSX.writeFile = (output) => { exported = output; };
  await run('downloadDetailWorkbook()');
  const output = XLSX.utils.sheet_to_json(exported.Sheets[exported.SheetNames[0]], { header: 1 });
  const index = output[0].indexOf('数量');
  assert.deepEqual(output.slice(1).map(row => row[index]), [7, 3, 0]);
  run('state.filteredRows = [{ quantityNumber: NaN }]; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '待核对');
  assert.equal(elements.get('#quantityTotal').textContent, '数量合计：0（1 条数量待核对）');
  run('state.filteredRows = []; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '0');
});

test('采购分工只读取产品线明细，不混入整理或同名副本，缺表明确报错', async () => {
  const {context,run}=loadApp();
  const division=XLSX.utils.book_new();
  for(const [name,supplier] of [['整理','错误供应商'],['产品线明细','正确供应商'],['产品线明细旧版','旧供应商']]){
    XLSX.utils.book_append_sheet(division,XLSX.utils.aoa_to_sheet([['物料编码','供应商简称','采购下单人'],['TEST-ONLY',supplier,'采购员']]),name);
  }
  const demand=XLSX.utils.book_new();XLSX.utils.book_append_sheet(demand,XLSX.utils.aoa_to_sheet([['物料编码','数量','申请人'],['TEST-ONLY',5,'申请人']]),'需求');
  const record=wb=>({name:'测试.xlsx',file:{name:'测试.xlsx',arrayBuffer:async()=>XLSX.write(wb,{type:'buffer',bookType:'xlsx'})}});
  context.records={purchaseDivision:record(division),demand:record(demand)};
  const rows=await run('buildDemandAllocationRows(records)');assert.equal(rows[0].supplierShortName,'正确供应商');
  division.SheetNames=division.SheetNames.filter(n=>n!=='产品线明细');delete division.Sheets.产品线明细;
  context.records.purchaseDivision=record(division);
  await assert.rejects(run('buildDemandAllocationRows(records)'),/缺少“产品线明细”/);
});

test('供应商按中文名称排序，空值最后，同名稳定，导出与页面一致', async () => {
 const {context,run}=loadApp();
 context.rows=['中山','鸿诚','阿明','鸿诚','供应商10','供应商2',''].map((supplierShortName,i)=>({supplierShortName,materialCode:String(i),quantityNumber:i}));
 run('state.demandRows=rows; state.filteredRows=getFilteredDemandRows()');
 const sorted=run('state.filteredRows');assert.deepEqual(Array.from(sorted,r=>r.supplierShortName),['阿明','供应商2','供应商10','鸿诚','鸿诚','中山','']);
 assert.deepEqual(Array.from(sorted.filter(r=>r.supplierShortName==='鸿诚'),r=>r.materialCode),['1','3']);
 assert.equal(context.rows[0].supplierShortName,'中山');
 let book;context.window.XLSX.writeFile=w=>{book=w;};await run('downloadDetailWorkbook()');
 const data=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1});
 assert.deepEqual(data.slice(1).map(r=>r[5]),Array.from(sorted,r=>r.supplierShortName));
 run('state.filters.get("supplierShortName").add("鸿诚")');assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>r.materialCode),['1','3']);
});

test('需求表采购主体优先，手填只补空白，清空后恢复空白且不修改原始数据', async () => {
 const {context,run}=loadApp();context.rows=[{materialCode:'A',purchaseEntity:'原主体甲'},{materialCode:'B',purchaseEntity:''}];
 run('state.demandRows=rows;state.purchaseEntityOverride="手填主体";state.filteredRows=getFilteredDemandRows()');
 assert.deepEqual(Array.from(run('state.filteredRows'),r=>r.purchaseEntity),['原主体甲','手填主体']);
 let book;context.window.XLSX.writeFile=w=>{book=w;};await run('downloadDetailWorkbook()');
 const data=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1});const col=data[0].indexOf('采购主体');assert.deepEqual(data.slice(1).map(r=>r[col]),['原主体甲','手填主体']);
 run('state.purchaseEntityOverride=""');assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>r.purchaseEntity),['原主体甲','']);assert.equal(context.rows[0].purchaseEntity,'原主体甲');
});

test('采购主体按明确表头读取，兼容旧别名并保留合并后的来源主体', async () => {
 const {context,run}=loadApp();
 for (const header of ['采购主体','采购公司','采购组织']) {
  context.records={demand:allocationRecord([[header,'物料编码','数量'],[' 主体甲 ','A',1],['主体乙','A',2],['  ','B',0]])};
  const rows=await run('buildDemandAllocationRows(records)');
  context.rows=rows;run('state.demandRows=rows;state.purchaseEntityOverride="补充主体"');
  assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>[r.purchaseEntity,r.quantityNumber]),[['主体甲、主体乙',3],['补充主体',0]]);
 }
 for (const reverse of [false,true]) {
  const rows=[['采购组织','物料编码','数量','采购主体'],['其他组织','A',1,'明细主体'],['其他组织','B',2,'']];
  context.records={demand:allocationRecord(rows.map(row=>reverse?[...row].reverse():row))};
  context.rows=await run('buildDemandAllocationRows(records)');
  run('state.demandRows=rows');
  assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>r.purchaseEntity),['明细主体','补充主体']);
 }
});

test('交期按北京时间加自然日，覆盖零、空值、跨月跨年闰日及异常',()=>{
 const {context,run}=loadApp();
 const calculate=(days,date)=>{context.days=days;context.now=new Date(date);return run('deliveryFields(days,now)');};
 assert.equal(calculate(0,'2026-10-14T16:01:00Z').supplierDeliveryDate,'2026-10-15');
 assert.equal(calculate(20,'2026-10-15T00:00:00Z').supplierDeliveryDate,'2026-11-04');
 assert.equal(calculate(2,'2026-12-31T00:00:00Z').supplierDeliveryDate,'2027-01-02');
 assert.equal(calculate(1,'2028-02-28T00:00:00Z').supplierDeliveryDate,'2028-02-29');
 for(const v of ['',undefined,-1,1.5,'abc',36501])assert.equal(calculate(v,'2026-10-15T00:00:00Z').supplierDeliveryDate,'');
});

test('下载日期为文字且保留TODAY及加天数公式，空值和0均保留，重开不丢失',()=>{
 const {context,run}=loadApp();context.sheet={};context.rows=[{todayDate:'2026-10-15',deliveryDays:20,supplierDeliveryDate:'2026-11-04'},{todayDate:'2026-10-15',deliveryDays:'',supplierDeliveryDate:''},{todayDate:'2026-10-15',deliveryDays:0,supplierDeliveryDate:'2026-10-15'}];
 run('addDeliveryFormulas(sheet,rows)');context.sheet['!ref']='A1:S4';
 const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,context.sheet,'测试');const sheet=XLSX.read(XLSX.write(book,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}).Sheets.测试;
 assert.equal(sheet.L2.f,'TEXT(TODAY(),"yyyy-mm-dd")');assert.equal(sheet.N2.f,'IF(M2="","",TEXT(DATE(VALUE(LEFT(L2,4)),VALUE(MID(L2,6,2)),VALUE(RIGHT(L2,2)))+M2,"yyyy-mm-dd"))');assert.equal(sheet.N3.f,'IF(M3="","",TEXT(DATE(VALUE(LEFT(L3,4)),VALUE(MID(L3,6,2)),VALUE(RIGHT(L3,2)))+M3,"yyyy-mm-dd"))');assert.equal(sheet.N3.v,'');assert.equal(sheet.N4.v,sheet.L4.v);assert.equal(sheet.N2.v,'2026-11-04');assert.equal(sheet.N2.t,'s');assert.equal(sheet.L2.t,'s');assert.equal(sheet.N2.z,'@');assert.equal(sheet.L2.z,'@');
});


test('起订量按表头识别，列移动、零值、空白及异常不会误判', () => {
 const {context,run}=loadApp();
 context.tables=[
  {source:{sheetName:'产品线明细'},rows:[
   ['起订量','供应商简称','物料编码'],
   [100,'甲','A'],[0,'甲','B'],['','甲','C'],['/','甲','D'],['无','甲','E'],['1,000','甲','F']
  ]},
  {source:{sheetName:'产品线明细备份'},rows:[['物料编码','起订量'],['A',1]]}
 ];
 const index=run('buildMinimumOrderIndex(tables).byMaterial');
 assert.equal(index.get('a'),100);assert.equal(index.get('b'),0);assert.equal(index.get('f'),1000);
 for(const code of ['c','d','e'])assert(!index.has(code));
 for(const [quantity,minimum,expected] of [[99,100,'不满足'],[100,100,'满足'],[101,100,'满足'],[0,100,'不满足'],[0,0,'满足'],[NaN,100,''],[100,NaN,'']]){
  context.quantity=quantity;context.minimum=minimum;
  assert.equal(run('getMinimumOrderStatus(quantity,minimum)'),expected);
 }
 context.tables=[{source:{sheetName:'产品线明细'},rows:[['物料编码','','','','','','','','','价格'],['A','','','','','','','','',100]]}];
 assert.equal(run('buildMinimumOrderIndex(tables).byMaterial.size'),0);
});
