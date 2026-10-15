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
  return { context, elements, run: (code) => vm.runInContext(code, context) };
}

test('all 19 table headers, rendered values and exported cells follow the requested order', async () => {
  const { context, run } = loadApp();
  const expected = [
    ['采购单订单下单人', 'buyer'], ['事业部', 'businessUnit'], ['申请人', 'applicant'],
    ['供应商简称', 'supplierShortName'], ['物料编码', 'materialCode'], ['SKU', 'sku'],
    ['物料名称', 'materialName'], ['数量', 'quantity'], ['OA备货流程号', 'oaProcessNo'],
    ['采购主体', 'purchaseEntity'], ['采购分工明细是否存在', 'materialCodeValid'],
    ['今天日期', 'todayDate'], ['天数', 'deliveryDays'], ['供应商反馈交期', 'supplierDeliveryDate'], ['采购单订单号', 'purchaseOrderNo'], ['要求货好时间', 'requiredReadyDate'], ['起订量', 'minimumOrderQuantity'],
    ['起订量是否满足', 'minimumOrderStatus'], ['备货原因', 'stockReason'],
  ];
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const header = html.match(/<table class="detail-table">([\s\S]*?)<\/thead>/)[1];
  assert.deepEqual([...header.matchAll(/<th>(.*?)<\/th>/g)].map(match => match[1]), expected.map(item => item[0]));
  context.row = Object.fromEntries(expected.map(([, key], index) => [key, `value-${index + 1}`]));
  const cells = [...run('renderDemandTableRow(row)').matchAll(/<td>(.*?)<\/td>/g)].map(match => match[1].startsWith('<input') ? match[1].match(/value="([^"]*)"/)[1] : match[1]);
  assert.deepEqual(cells, expected.map((_, index) => `value-${index + 1}`));
  let exported;
  context.window.XLSX.writeFile = (output) => { exported = output; };
  run('state.filteredRows = [row]');
  await run('downloadDetailWorkbook()');
  const output = XLSX.utils.sheet_to_json(exported.Sheets[exported.SheetNames[0]], { header: 1 });
  assert.deepEqual(output[0], expected.map(item => item[0]));
  assert.deepEqual(output[1], cells.map((value, i) => expected[i][1] === 'quantity' ? '' : value));
});

test('all buyers and pages export into one sheet in display order on both export paths', async () => {
  const { context, run } = loadApp();
  context.rows = Array.from({ length: 23 }, (_, i) => ({
    buyer: ['采购乙', '采购甲', ''][i % 3],
    materialCode: `TEST-${i}`, quantity: String(i),
    purchaseOrderNo: i === 0 ? 'PO001' : '', supplierDeliveryDate: '',
  }));
  run('state.filteredRows = rows; state.currentPage = 2');
  let exported;
  context.window.XLSX.writeFile = workbook => { exported = workbook; };
  await run('downloadDetailWorkbook()');
  const verify = workbook => {
    assert.deepEqual(Array.from(workbook.SheetNames), ['备货需求分配明细']);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });
    assert.equal(data.length, 24);
    for (const key of ['buyer', 'materialCode', 'quantity', 'purchaseOrderNo', 'supplierDeliveryDate']) {
      const labels = { buyer: '采购单订单下单人', materialCode: '物料编码', quantity: '数量', purchaseOrderNo: '采购单订单号', supplierDeliveryDate: '供应商反馈交期' };
      assert.deepEqual(data.slice(1).map(row => row[data[0].indexOf(labels[key])]), context.rows.map(row => key === 'quantity' ? Number(row[key]) : row[key]));
    }
  };
  verify(exported);
  context.window.JSZip = {};
  let styled;
  context.captureStyled = (workbook, shapes) => { styled = { workbook, shapes }; };
  run('writeStyledDetailWorkbook = captureStyled');
  await run('downloadDetailWorkbook()');
  assert(styled, 'styled export must receive the single-sheet workbook');
  assert.deepEqual(JSON.parse(JSON.stringify(styled.shapes)), [{ rowCount: 24, columnCount: 19 }]);
  verify(XLSX.read(XLSX.write(styled.workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' }));
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
    assert.equal(sheet[`H${i + 2}`].t, 'n');
    assert.equal(sheet[`H${i + 2}`].v, value);
    assert.equal(sheet[`H${i + 2}`].z, '0');
    assert(!sheet[`H${i + 2}`].w.includes('.'));
  }
  assert.equal(sheet.H7.v, '');
  assert.equal(sheet.H8.v, '');
  assert.equal(sheet.H9.v, 1234);
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
    ['物料编码', '采购单订单下单人', '供应商简称', '', '', '', '', '', '', '起订量'],
    ['TEST-001', '采购甲', '供应商甲', '', '', '', '', '', '', 8],
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
  assert.equal(output[1][output[0].indexOf('OA备货流程号')], 'OA001/OA002');
  assert.equal(output[1][output[0].indexOf('采购单订单号')], '');
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
 assert.deepEqual(data.slice(1).map(r=>r[3]),Array.from(sorted,r=>r.supplierShortName));
 run('state.filters.get("supplierShortName").add("鸿诚")');assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>r.materialCode),['1','3']);
});

test('手动采购主体统一覆盖页面和导出，清空后恢复来源且不修改原始数据', async () => {
 const {context,run}=loadApp();context.rows=[{materialCode:'A',purchaseEntity:'原主体甲'},{materialCode:'B',purchaseEntity:''}];
 run('state.demandRows=rows;state.purchaseEntityOverride="手填主体";state.filteredRows=getFilteredDemandRows()');
 assert.deepEqual(Array.from(run('state.filteredRows'),r=>r.purchaseEntity),['手填主体','手填主体']);
 let book;context.window.XLSX.writeFile=w=>{book=w;};await run('downloadDetailWorkbook()');
 const data=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1});const col=data[0].indexOf('采购主体');assert.deepEqual(data.slice(1).map(r=>r[col]),['手填主体','手填主体']);
 run('state.purchaseEntityOverride=""');assert.deepEqual(Array.from(run('getFilteredDemandRows()'),r=>r.purchaseEntity),['原主体甲','']);assert.equal(context.rows[0].purchaseEntity,'原主体甲');
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

test('下载日期保留TODAY及加天数公式，空值和0均保留，重开不丢失',()=>{
 const {context,run}=loadApp();context.sheet={};context.rows=[{todayDate:'2026-10-15',deliveryDays:20,supplierDeliveryDate:'2026-11-04'},{todayDate:'2026-10-15',deliveryDays:'',supplierDeliveryDate:''},{todayDate:'2026-10-15',deliveryDays:0,supplierDeliveryDate:'2026-10-15'}];
 run('addDeliveryFormulas(sheet,rows)');context.sheet['!ref']='A1:S4';
 const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,context.sheet,'测试');const sheet=XLSX.read(XLSX.write(book,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}).Sheets.测试;
 assert.equal(sheet.L2.f,'TODAY()');assert.equal(sheet.N2.f,'IF(M2="","",L2+M2)');assert.equal(sheet.N3.f,'IF(M3="","",L3+M3)');assert.equal(sheet.N3.v,'');assert.equal(sheet.N4.v,sheet.L4.v);assert.equal(XLSX.SSF.format(sheet.N2.z,sheet.N2.v),'2026-11-04');
});
