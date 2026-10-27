const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../fulfillment-core.js');

// Synthetic workbooks only: the same suite can run against the published page.
const people = ['徐英', '操思敏', '邢芳芳', '李奇'];
const sheetNames = ['徐英', '操思敏', '邢芳芳1017', '李奇'];
const quantities = ['未交付数量', '已下单未备料未生产', '已备料未生产', '生产中产品', '完工未发产品', '已发货数量'];
const divisions = ['海外事业一部', '海外事业二部', '国内事业部'];
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function fixture(change) {
  const book = X.utils.book_new();
  for (const [person, name] of sheetNames.entries()) {
    const rows = [0, 1].map(row => C.COLUMNS.map(key => {
      const q = quantities.indexOf(key);
      if (q >= 0) return (person + 1) * 10 + (row + 1) * 2 + q;
      return ({
        '采购下单人': person === 1 ? '陈燕' : people[person],
        '事业部': divisions[(person + row) % divisions.length],
        '供应商简称': '合成供应商' + person,
        '物料编码': 'SYN-' + person + '-' + row,
        'SKU': 'SKU-' + person + '-' + row,
        '物料名称': '合成测试物料',
        '下单日期': '2026年10月23日'
      })[key] ?? '';
    }));
    X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([C.COLUMNS, ...rows]), name);
  }
  X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([C.COLUMNS, C.COLUMNS.map(key =>
    quantities.includes(key) ? 99999 : ({ '采购下单人': '孙立柱', '事业部': divisions[0], '物料编码': 'EXCLUDED' })[key] ?? ''
  )]), '孙立柱');
  const note = X.utils.aoa_to_sheet([['保留原始工作表', ''], [2, 3], [5]]);
  note.A3.f = 'SUM(A2:B2)';
  note['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
  X.utils.book_append_sheet(book, note, '说明');
  if (change) change(book);
  return Buffer.from(X.write(book, { type: 'buffer', bookType: 'xlsx' }));
}

function staleFormula(field, formula = 'SUM(AA2:AA3)') {
  const address = X.utils.encode_cell({ r: 1, c: C.COLUMNS.indexOf(field) });
  return {
    address,
    buffer: fixture(book => {
      const sheet = book.Sheets.李奇;
      sheet[address] = { t: 'n', v: 0, f: formula };
      sheet.AA1 = { t: 's', v: '计算依据' };
      sheet.AA2 = { t: 'n', v: 10 };
      sheet.AA3 = { t: 'n', v: 20 };
      sheet['!ref'] = 'A1:AA3';
    })
  };
}

(async () => {
  const browser = await chromium.launch({
    executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true
  });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const url = new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8940/').href;
    await page.goto(url);
    await page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);

    async function upload(buffer, name) {
      await page.locator('#file').setInputFiles({ name, mimeType, buffer });
      await page.waitForFunction(() => {
        const input = document.getElementById('file');
        return !input.disabled && !input.value;
      });
    }
    async function history() {
      return page.evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(),
        (_, value) => value instanceof ArrayBuffer ? Array.from(new Uint8Array(value)) : value)));
    }
    async function assertBlocked(field, address, person = '李奇') {
      assert.equal(await page.locator('#quantityValidation').isVisible(), true);
      assert.match(await page.locator('#quantityStatus').innerText(), /未通过|不通过|失败|需核对|差异/);
      const issues = await page.locator('#quantityIssues').innerText();
      assert(issues.includes(person), '问题必须指出来源人员：' + person);
      if (field) assert(issues.includes(field), '问题必须指出数量字段：' + field);
      if (address) assert(issues.includes(address), '问题必须指出具体单元格：' + address);
      assert.equal(await page.locator('#download').isDisabled(), true);
      for (const division of divisions) assert.equal(await page.locator('[data-download-division="' + division + '"]').isDisabled(), true);
    }
    async function putLegacyCache(buffer, cachedRowsWrong = false) {
      await page.evaluate(async ({ bytes, cachedRowsWrong }) => {
        const sourceBytes = Uint8Array.from(bytes).buffer;
        const result = FulfillmentCore.merge(XLSX.read(sourceBytes, { type: 'array', cellNF: true }), '旧缓存更新版.xlsx');
        if (cachedRowsWrong) result.rows[0][FulfillmentCore.COLUMNS.indexOf('未交付数量')] = { t: 'n', v: 987654321 };
        const current = { filename: '旧缓存更新版.xlsx', savedAt: '2026-10-20T00:00:00.000Z', sourceBytes, result };
        await new Promise((resolve, reject) => {
          const open = indexedDB.open('gylsjqx-fulfillment-history', 1);
          open.onerror = () => reject(open.error);
          open.onsuccess = () => {
            const db = open.result, tx = db.transaction('snapshots', 'readwrite');
            tx.objectStore('snapshots').put({ version: 1, previous: null, current }, 'history');
            tx.oncomplete = () => { db.close(); resolve(); };
            tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
          };
        });
      }, { bytes: [...buffer], cachedRowsWrong });
      await page.reload();
      await page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);
    }

    const valid = fixture();
    await upload(valid, '合成订单2026.10.23更新版.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false);
    assert.match(await page.locator('#quantityStatus').innerText(), /通过/);
    const checkRows = page.locator('#quantityChecks tbody tr');
    assert.equal(await checkRows.count(), 30, '四人各六项，加六项合计');
    const checks = await checkRows.allTextContents();
    for (const person of people) assert.equal(checks.filter(text => text.includes(person)).length, 6);
    assert.equal(checks.some(text => text.includes('陈燕')), false, '归属按来源工作表，不按采购下单人字段');
    assert.equal(checks.some(text => text.includes('孙立柱')), false, '孙立柱不参与校验汇总');

    const downloaded = page.waitForEvent('download');
    await page.locator('#download').click();
    const download = await downloaded;
    assert.equal(download.suggestedFilename(), '合成订单2026.10.23汇总版.xlsx');
    const output = await fs.readFile(await download.path());
    const exported = X.read(output, { type: 'buffer', cellNF: true });
    assert.deepEqual(exported.SheetNames, ['汇总表', '差异', ...sheetNames, '孙立柱', '说明']);
    const summary = X.utils.sheet_to_json(exported.Sheets.汇总表, { header: 1, defval: '' });
    assert.deepEqual(summary[0], C.SUMMARY_COLUMNS);
    assert.equal(summary[0].length, 21);
    assert.equal(summary.length - 1, 8);
    assert.deepEqual(summary.slice(1).map(row => row[8]), people.flatMap((_, person) => [0, 1].map(row => 'SYN-' + person + '-' + row)));
    for (const [q, field] of quantities.entries()) {
      const col = summary[0].indexOf(field);
      assert(summary.slice(1).every(row => typeof row[col] === 'number'), field + '必须导出数字');
      assert.equal(summary.slice(1).reduce((sum, row) => sum + row[col], 0), 224 + q * 8);
    }
    const [originalZip, outputZip] = await Promise.all([Zip.loadAsync(valid), Zip.loadAsync(output)]);
    for (let n = 1; n <= sheetNames.length + 2; n++) {
      const name = 'xl/worksheets/sheet' + n + '.xml';
      assert.equal(await outputZip.file(name).async('string'), await originalZip.file(name).async('string'), '保留原表 XML：' + name);
    }

    const goodHistory = await history();
    for (const formula of ['SUM(AA2:AA3)', 'SUBTOTAL(9,AA2:AA3)']) {
      const invalid = staleFormula(quantities[0], formula);
      await upload(invalid.buffer, '旧公式结果.xlsx');
      await assertBlocked(quantities[0], invalid.address);
      assert.deepEqual(await history(), goodHistory, '校验失败不推进历史基准');
    }
    for (const field of quantities.slice(1)) {
      const invalid = staleFormula(field);
      await upload(invalid.buffer, '其他数量字段公式异常.xlsx');
      await assertBlocked(field, invalid.address);
      assert.deepEqual(await history(), goodHistory);
      const nonnumeric = fixture(book => { book.Sheets.李奇[invalid.address] = { t: 's', v: '错误数量' }; });
      await upload(nonnumeric, '其他数量字段文本异常.xlsx');
      await assertBlocked(field, invalid.address);
      assert.deepEqual(await history(), goodHistory);
    }

    await upload(valid, '恢复有效上传.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false);
    assert.match(await page.locator('#quantityStatus').innerText(), /通过/);
    assert.equal((await history()).previous, null, '恢复同一有效数据不推进基准');

    await putLegacyCache(valid);
    assert.equal(await page.locator('#download').isDisabled(), false, '旧缓存须从原文件重新核验');
    assert.match(await page.locator('#quantityStatus').innerText(), /通过/);

    await putLegacyCache(valid, true);
    await assertBlocked(quantities[0], X.utils.encode_cell({ r: 1, c: C.COLUMNS.indexOf(quantities[0]) }), '徐英');

    const cachedInvalid = staleFormula(quantities[0], 'SUBTOTAL(9,AA2:AA3)');
    await putLegacyCache(cachedInvalid.buffer);
    await assertBlocked(quantities[0], cachedInvalid.address);
    await page.reload();
    await page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);
    await assertBlocked(quantities[0], cachedInvalid.address);
    await upload(valid, '校验后有效上传.xlsx');
    assert.equal(await page.locator('#download').isDisabled(), false);
    assert.equal(await page.locator('#quantityIssues').innerText(), '', '有效上传应清除旧差异');
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ url, checks: 30, fields: quantities, rows: 8, sourcePreservation: 'passed', formulaCache: 'passed', blockedDownloads: 'passed', historyRollback: 'passed', legacyCache: 'passed' }));
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
