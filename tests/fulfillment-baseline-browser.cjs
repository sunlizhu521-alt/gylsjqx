const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), C = require('../fulfillment-core.js');
const url = new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8960/').href;
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function fixture(items, {supplierB = 5, badQuantity = false, badFormula = false} = {}) {
  const book = X.utils.book_new();
  C.SHEETS.forEach((sheetName, person) => {
    const entries = person ? [[`FIX${person}`, person, `固定供应商${person}`]] : [...items, ['SUP-B', supplierB, '供应商乙']];
    const records = entries.map(([code, quantity, supplier = '供应商甲']) => ({
      '采购下单人': person ? sheetName : '共同采购员', '事业部': '国内事业部',
      '销售分渠道': '国内事业部*TEMU业务部', '供应商简称': supplier, '系列': '共同系列',
      '运营': '王志豪', '物料编码': code, 'SKU': 'SKU-' + code, '物料名称': '物料-' + code,
      '采购订单号': 'PO-' + code, 'OA备货单号': 'OA-' + code,
      '下单月份': '2026-11', '下单日期': '2026-11-01', '未交付数量': quantity,
      '已下单未备料未生产': 0, '已备料未生产': 0, '生产中产品': 0, '完工未发产品': 0, '已发货数量': 0
    }));
    const sheet = X.utils.aoa_to_sheet([C.COLUMNS, ...records.map(row => C.COLUMNS.map(key => row[key] ?? ''))]);
    const quantityAddress = X.utils.encode_cell({r: 1, c: C.COLUMNS.indexOf('未交付数量')});
    if (!person && badQuantity) sheet[quantityAddress] = {t: 's', v: '100或200'};
    if (!person && badFormula) sheet[quantityAddress] = {t: 'n', v: 0, f: '#REF!'};
    X.utils.book_append_sheet(book, sheet, sheetName);
  });
  return Buffer.from(X.write(book, {type: 'buffer', bookType: 'xlsx'}));
}

(async () => {
  const browser = await chromium.launch({headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
  const errors = [], temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'fulfillment-baseline-'));
  try {
    const ready = page => page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled && !document.getElementById('previous-file').value);
    async function open() {
      const page = await browser.newPage();
      page.on('pageerror', error => errors.push(error.message));
      await page.goto(url); await ready(page);
      assert.equal(await page.locator('#replace-previous').isDisabled(), true, '没有本次文件时不能替换基准');
      return page;
    }
    const history = page => page.evaluate(async () => JSON.parse(JSON.stringify(await FulfillmentCache.get(), (_, value) => value instanceof ArrayBuffer ? Array.from(new Uint8Array(value)) : value)));
    async function upload(page, buffer, name) {
      await page.locator('#file').setInputFiles({name, mimeType, buffer});
      await page.waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
      assert.equal(await page.locator('#download').isDisabled(), false);
      assert.equal(await page.locator('#replace-previous').isDisabled(), false);
    }
    async function choosePrevious(page, file, checkBusy = false) {
      if (checkBusy) await page.evaluate(() => {
        const prepare = FulfillmentValues.prepare;
        window.baselineOriginalPrepare = prepare;
        FulfillmentValues.prepare = (...args) => new Promise(resolve => {window.continueBaselinePrepare = () => resolve(prepare(...args));});
      });
      const event = page.waitForEvent('filechooser');
      await page.locator('#replace-previous').click();
      await (await event).setFiles(file);
      if (checkBusy) {
        await page.waitForFunction(() => typeof window.continueBaselinePrepare === 'function');
        for (const selector of ['#file', '#replace-previous', '#previous-file', '#download', '[data-download-division="国内事业部"]']) assert.equal(await page.locator(selector).isDisabled(), true, selector + '替换处理中禁用');
        await page.evaluate(() => {
          FulfillmentValues.prepare = window.baselineOriginalPrepare;
          window.continueBaselinePrepare(); delete window.continueBaselinePrepare; delete window.baselineOriginalPrepare;
        });
      }
      await ready(page);
    }
    async function exportSeries(page) {
      const pending = page.waitForEvent('download'); await page.locator('#download').click();
      const file = await pending, book = X.read(await fs.readFile(await file.path()), {type: 'buffer', cellNF: true});
      assert.match(file.suggestedFilename(), /_V\d+\.xlsx$/);
      assert.deepEqual(X.utils.sheet_to_json(book.Sheets.系列对比, {header: 1})[0], C.SERIES_COLUMNS);
      return X.utils.sheet_to_json(book.Sheets.系列对比);
    }
    const bySupplier = (rows, supplier) => rows.find(row => row.采购下单人 === '共同采购员' && row.供应商简称 === supplier);
    function checkTotals(rows, supplier, before, current, difference, judgment) {
      const row = bySupplier(rows, supplier); assert(row, supplier + '应有对比记录');
      assert.deepEqual([row.上期未交付数据, row.本期未交付数据, row.差异数量], [before, current, difference]);
      assert.equal(row.备注 ?? '', ''); assert.equal(row.判断 ?? '', judgment);
    }

    const page = await open(), week = fixture([['A', 10]]);
    await upload(page, week, '上周固定版.xlsx');
    const first = await history(page); assert.equal(first.previous, null);
    await upload(page, fixture([['A', 14], ['B', 4]]), '本周一.xlsx');
    assert.deepEqual((await history(page)).previous, first.current);
    await upload(page, fixture([['A', 15], ['B', 5], ['C', 6]]), '本周二.xlsx');
    assert.deepEqual((await history(page)).previous, first.current, '第三次上传不能滚动上一版');
    const currentFile = fixture([['A', 16], ['B', 5], ['C', 6], ['D', 3]]);
    await upload(page, currentFile, '本周三.xlsx');
    assert.deepEqual((await history(page)).previous, first.current, '第四次上传仍使用上周版');
    checkTotals(await exportSeries(page), '供应商甲', 10, 30, 20, '2026年11月净增3行（上期1行，本期4行）；涉及行：徐英第3-5行');
    const beforeReload = await history(page);
    await page.reload(); await ready(page);
    assert.deepEqual(await history(page), beforeReload, '刷新不变动两期');
    await upload(page, currentFile, '重复本周三.xlsx');
    assert.deepEqual((await history(page)).previous, first.current, '重复上传不变动固定基准');
    const stable = await history(page);
    await choosePrevious(page, []);
    assert.deepEqual(await history(page), stable, '取消文件选择不变动两期');

    await choosePrevious(page, {name: '上周更正版.xlsx', mimeType, buffer: fixture([['A', 10], ['B', 2]], {supplierB: 8})}, true);
    const replaced = await history(page);
    assert.equal(replaced.previous.filename, '上周更正版.xlsx');
    assert.deepEqual(replaced.current, stable.current, '替换上一版不能修改本次文件、时间或数据');
    assert.match(await page.locator('#history').innerText(), /上次：上周更正版.xlsx/);
    const replacementOutput = await exportSeries(page);
    checkTotals(replacementOutput, '供应商甲', 12, 30, 18, '2026年11月净增2行（上期2行，本期4行）；涉及行：徐英第4-5行');
    checkTotals(replacementOutput, '供应商乙', 8, 5, -3, '');
    assert.deepEqual((await history(page)).previous, replaced.previous, '下载不变动替换后的基准');
    await page.reload(); await ready(page);
    assert.deepEqual(await history(page), replaced, '刷新保存替换后的基准');
    await upload(page, fixture([['A', 20], ['B', 5], ['C', 6], ['D', 3], ['E', 2]]), '本周四.xlsx');
    assert.deepEqual((await history(page)).previous, replaced.previous, '后续普通上传继续使用显式替换的上一版');
    checkTotals(await exportSeries(page), '供应商甲', 12, 36, 24, '2026年11月净增3行（上期2行，本期5行）；涉及行：徐英第4-6行');

    const stableAfter = await history(page), preview = await page.locator('#preview').textContent();
    for (const [name, buffer] of [
      ['上一版数量异常.xlsx', fixture([['A', 1]], {badQuantity: true})],
      ['上一版公式异常.xlsx', fixture([['A', 1]], {badFormula: true})],
      ['损坏上一版.xlsx', Buffer.from('This is not an Excel workbook')]
    ]) {
      await choosePrevious(page, {name, mimeType, buffer});
      assert.match(await page.locator('#status').innerText(), /上一版替换失败/);
      assert.deepEqual(await history(page), stableAfter, name + '不得覆盖两期');
      assert.equal(await page.locator('#download').isDisabled(), false);
      assert.equal(await page.locator('#preview').textContent(), preview);
    }
    const oversized = path.join(temporary, '超过60MB上一版.xlsx'), handle = await fs.open(oversized, 'w');
    await handle.truncate(60 * 1024 * 1024 + 1); await handle.close();
    await choosePrevious(page, oversized);
    assert.match(await page.locator('#status').innerText(), /60MB/);
    assert.deepEqual(await history(page), stableAfter, '超限文件不覆盖两期');
    assert.equal(await page.locator('#download').isDisabled(), false);

    await page.evaluate(() => {
      const put = IDBObjectStore.prototype.put;
      window.baselineOriginalPut = put;
      IDBObjectStore.prototype.put = function (value, key) {
        if (this.name === 'snapshots' && key === 'history') throw new DOMException('合成缓存写入失败', 'QuotaExceededError');
        return put.call(this, value, key);
      };
    });
    try {
      await choosePrevious(page, {name: '缓存失败上一版.xlsx', mimeType, buffer: fixture([['A', 3]])});
      assert.match(await page.locator('#status').innerText(), /上一版替换失败/);
      assert.deepEqual(await history(page), stableAfter, '缓存写入失败原子回滚，不变动两期');
      assert.equal(await page.locator('#download').isDisabled(), false);
    } finally {
      await page.evaluate(() => {IDBObjectStore.prototype.put = window.baselineOriginalPut; delete window.baselineOriginalPut;});
    }
    checkTotals(await exportSeries(page), '供应商甲', 12, 36, 24, '2026年11月净增3行（上期2行，本期5行）；涉及行：徐英第4-6行');

    const single = await open();
    await upload(single, fixture([['A', 20]]), '仅本次.xlsx');
    const singleBefore = await history(single); assert.equal(singleBefore.previous, null);
    await choosePrevious(single, {name: '直接指定上周.xlsx', mimeType, buffer: week});
    const singleAfter = await history(single);
    assert.equal(singleAfter.previous.filename, '直接指定上周.xlsx');
    assert.deepEqual(singleAfter.current, singleBefore.current, '首次本次文件可直接指定上一版');
    checkTotals(await exportSeries(single), '供应商甲', 10, 20, 10, '');
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({status: 'PASS', url, ordinaryUploads: 'fixed baseline', replacement: 'explicit picker only; current unchanged', persistence: 'reload/duplicate/download', comparison: 'supplier totals and new-row judgment recalculated', failures: 'invalid quantity/formula/file, oversized file and IndexedDB rollback', firstCurrent: 'can explicitly choose baseline', busyControls: 'disabled'}));
  } finally {await browser.close(); await fs.rm(temporary, {recursive: true, force: true});}
})().catch(error => {console.error(error); process.exitCode = 1;});
