const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../fulfillment-core.js');

// Synthetic files only: this same acceptance test can run against GitHub Pages.
const divisions = ['海外事业一部', '海外事业二部', '国内事业部'];
const months = ['历史订单', '/', '2025-11', 46318]; // Last value is 2026-10-23, not a month-only source.
const displays = ['2025年1月', '2026年4月', '2025年11月', '2026年10月'];
const monthFormat = 'yyyy"年"m"月"';
const quantities = C.QUANTITY_FIELDS;
const out = process.env.QA_OUT || '/private/tmp/fulfillment-format-qa';

function fixture(delta = 0) {
  const book = X.utils.book_new(), expected = [];
  for (const [person, name] of C.SHEETS.entries()) {
    const columns = C.COLUMNS.slice(person * 5).concat(C.COLUMNS.slice(0, person * 5));
    if (person % 2) columns.reverse();
    const rows = [0, 1].map(row => {
      const values = {
        '事业部': divisions[person % divisions.length],
        '采购下单人': name,
        '供应商简称': '合成供应商' + person,
        '物料编码': '000' + person + row,
        'SKU': '000-SKU-' + person + row,
        '物料名称': '格式测试',
        'OA备货单号': '0000' + person + row,
        '采购订单号': '000-ORDER-' + person + row,
        '下单日期': '2024年1月12日',
        '下单月份': months[person],
        '合同约定交期': '2026年11月2日',
        '生产中交付时间': '2026年11月3日',
        '未生产预计交付时间': '2026年11月4日',
        '未交付数量': row ? 0 : 1234 + (person === 0 ? delta : 0),
        '已下单未备料未生产': row ? 12 : '约0.125件',
        '已备料未生产': row ? 0.0000001 : 0,
        '生产中产品': row ? '' : '物料通用   474套左右',
        '完工未发产品': '',
        '已发货数量': row ? 3.0000001 : 2.5
      };
      expected.push({ ...values, '已下单未备料未生产': row ? 12 : 0.125, '生产中产品': row ? 0 : 474, month: displays[person] });
      return columns.map(field => values[field] ?? '');
    });
    const sheet = X.utils.aoa_to_sheet([columns.map(h => ' ' + h + ' '), ...rows]);
    const quantityCol = columns.indexOf('生产中产品');
    sheet['!merges'] = [{ s: { r: 1, c: quantityCol }, e: { r: 2, c: quantityCol } }];
    sheet['!cols'] = columns.map((_, col) => ({ wch: 11 + col }));
    sheet['!rows'] = [{ hpt: 29 }, { hpt: 36 }, { hpt: 41 }];
    for (let row = 1; row <= 2; row++) {
      const month = sheet[X.utils.encode_cell({ r: row, c: columns.indexOf('下单月份') })];
      if (month.t === 'n') month.z = 'yyyy-mm-dd';
      // Source number formats intentionally differ from the new summary format.
      sheet[X.utils.encode_cell({ r: row, c: columns.indexOf('未交付数量') })].z = '0.00';
    }
    X.utils.book_append_sheet(book, sheet, name === '邢芳芳' ? '邢芳芳1017' : name);
  }
  return { bytes: Buffer.from(X.write(book, { type: 'buffer', bookType: 'xlsx' })), expected };
}

function checkSummary(sheet, expected) {
  assert.deepEqual(X.utils.sheet_to_json(sheet, { header: 1 })[0], C.SUMMARY_COLUMNS);
  assert.equal(sheet.D1.v, '下单月份');
  assert.equal(C.SUMMARY_COLUMNS.length, 21);
  assert.equal(X.utils.decode_range(sheet['!ref']).e.r, expected.length);
  for (const [index, row] of expected.entries()) {
    const cell = field => sheet[X.utils.encode_cell({ r: index + 1, c: C.SUMMARY_COLUMNS.indexOf(field) })];
    assert.equal(cell('下单月份').t, 'n');
    assert.equal(cell('下单月份').z, monthFormat);
    assert.equal(cell('下单月份').w, row.month);
    for (const field of ['物料编码', 'SKU']) {
      assert.equal(cell(field).t, 's');
      assert.equal(cell(field).v, row[field]);
    }
    assert.equal(cell('OA流程号').t, 's');
    assert.equal(cell('OA流程号').v, row.OA备货单号);
    for (const [field, text] of [['合同约定交期', '2026年11月2日'], ['生产中交付时间', '2026年11月3日'], ['为生产预计交付时间', '2026年11月4日']]) {
      assert.equal(cell(field).t, 'n');
      assert.equal(cell(field).w, text);
    }
    for (const field of quantities) {
      const actual = cell(field);
      if (row[field] === '') { assert.equal(actual?.v ?? '', '', field + '空白不能变零'); continue; }
      assert.equal(actual.t, 'n', field);
      assert.equal(actual.v, row[field], field + '数值不能因格式改变');
      assert.match(actual.z, /^0(?:\.0+)?$/, field + '必须有明确数值格式');
      // The bundled SSF misformats 1e-7 in its read-back .w; verify stored v/z and the browser preview instead.
      if (Math.abs(row[field]) >= 0.000001 || row[field] === 0) assert.equal(actual.w, Number(row[field]).toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 12 }));
      else assert.equal(actual.z, '0.0000000');
      if (Number.isInteger(row[field])) assert.equal(actual.z, '0');
    }
  }
  assert(sheet['!cols'][3].wch >= 12, '年月与表头完整显示');
  assert.equal(sheet['!autofilter'].ref, `A1:U${expected.length + 1}`);
}

(async () => {
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
  try {
    const page = await browser.newPage(), errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const url = new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8940/').href;
    await page.goto(url);
    const ready = () => page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);
    await ready();
    async function upload(bytes, name) {
      await page.locator('#file').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes });
      await page.waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
      assert.equal(await page.locator('#download').isDisabled(), false, await page.locator('#quantityIssues').innerText());
    }
    async function download(selector, filename) {
      const wait = page.waitForEvent('download');
      await page.locator(selector).click();
      const result = await wait, bytes = await fs.readFile(await result.path());
      await fs.writeFile(path.join(out, filename), bytes);
      return { bytes, book: X.read(bytes, { type: 'buffer', cellNF: true, cellStyles: true }) };
    }
    const history = () => page.evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(), (_, v) => v instanceof ArrayBuffer ? Array.from(new Uint8Array(v)) : v)));
    const baseline = fixture(-1), current = fixture();
    await upload(baseline.bytes, '基准2026.10.16更新版.xlsx');
    await upload(current.bytes, '格式2026.10.23更新版.xlsx');
    assert.equal(await page.locator('#preview thead th').nth(3).innerText(), '下单月份');
    assert.deepEqual(await page.locator('#preview tbody tr td:nth-child(4)').allTextContents(), current.expected.map(r => r.month));
    for (const field of quantities) {
      const column = C.SUMMARY_COLUMNS.indexOf(field) + 1;
      assert.deepEqual(await page.locator('#preview tbody tr td:nth-child(' + column + ')').allTextContents(), current.expected.map(r => r[field] === '' ? '' : Number(r[field]).toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 12 })));
    }
    const full = await download('#download', '完整汇总.xlsx');
    checkSummary(full.book.Sheets.汇总表, current.expected);
    assert.deepEqual(full.book.SheetNames, ['汇总表', '差异', '徐英', '操思敏', '邢芳芳1017', '李奇']);
    const diff = full.book.Sheets.差异;
    assert.deepEqual(X.utils.sheet_to_json(diff, { header: 1 })[0], C.DIFF_COLUMNS);
    assert.equal(diff.F2.v, 1);
    assert.equal(diff.F2.z, 'General', '差异表不套用汇总数字格式');
    assert.equal(diff.A1.s.fgColor.rgb, '315B9D');
    const inputZip = await Zip.loadAsync(current.bytes), outputZip = await Zip.loadAsync(full.bytes);
    for (const name of Object.keys(inputZip.files).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))) {
      assert.equal(await outputZip.file(name).async('string'), await inputZip.file(name).async('string'), '保留原表 XML ' + name);
    }
    const diffXml = await outputZip.file('xl/worksheets/fulfillment2.xml').async('string');
    assert.match(diffXml, /ht="22"/);
    assert.match(diffXml, /ht="18"/);
    for (const division of divisions) {
      const divided = await download('[data-download-division="' + division + '"]', division + '.xlsx');
      assert.deepEqual(divided.book.SheetNames, ['汇总表']);
      checkSummary(divided.book.Sheets.汇总表, current.expected.filter(r => r.事业部 === division));
    }

    // Simulate cached snapshots from the preceding version, where summary cells had day or General formats.
    await page.evaluate(async () => {
      const history = await FulfillmentCache.get(), C = FulfillmentCore;
      for (const row of history.current.result.rows) {
        row[C.COLUMNS.indexOf('下单月份')].z = 'yyyy"年"m"月"d"日"';
        for (const field of C.QUANTITY_FIELDS) delete row[C.COLUMNS.indexOf(field)].z;
      }
      await FulfillmentCache.replace(history);
    });
    const before = await history();
    await page.reload(); await ready();
    const after = await history();
    assert.equal(after.current.savedAt, before.current.savedAt);
    assert.equal(after.current.filename, before.current.filename);
    assert.deepEqual(after.previous, before.previous, '显示格式升级不推进历史基准');
    assert.deepEqual(after.current.result.rows.map(r => quantities.map(k => r[C.COLUMNS.indexOf(k)].v)), before.current.result.rows.map(r => quantities.map(k => r[C.COLUMNS.indexOf(k)].v)));
    assert.deepEqual(await page.locator('#preview tbody tr td:nth-child(4)').allTextContents(), current.expected.map(r => r.month));
    const restored = await download('#download', '缓存恢复汇总.xlsx');
    checkSummary(restored.book.Sheets.汇总表, current.expected);
    assert.equal(await (await Zip.loadAsync(restored.bytes)).file('xl/worksheets/fulfillment2.xml').async('string'), diffXml, '缓存格式升级不改变差异表内容和样式');
    await page.screenshot({ path: path.join(out, 'preview.png'), fullPage: true });
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ status: 'PASS', url, rows: current.expected.length, summaryColumns: 21, downloads: 4, monthDates: 'passed', quantityFormats: 'passed', sourceXml: 'unchanged', differences: 'unchanged', previousSnapshot: 'unchanged', output: out }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
