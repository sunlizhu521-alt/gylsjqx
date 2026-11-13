const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), C = require('../fulfillment-core.js');
const key = 'gylsjqx-fulfillment-export-versions';
const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const url = new URL('fulfillment.html', process.env.QA_URL || 'http://127.0.0.1:8957/').href;
function fixture() {
  const book = X.utils.book_new();
  C.SHEETS.forEach((name, i) => {
    const row = {'采购下单人': name, '事业部': '国内事业部', '销售分渠道': '合成渠道',
      '系列': '合成系列', '物料编码': 'M' + i, 'SKU': 'SKU-' + i, '未交付数量': i,
      '已下单未备料未生产': 0, '已备料未生产': '', '生产中产品': 2.5, '完工未发产品': 3, '已发货数量': 4};
    X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([C.COLUMNS, C.COLUMNS.map(field => row[field] ?? '')]), name);
  });
  return Buffer.from(X.write(book, {type: 'buffer', bookType: 'xlsx'}));
}
(async () => {
  const browser = await chromium.launch({headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
  const errors = [];
  try {
    async function open(init, arg) {
      const page = await browser.newPage();
      page.on('pageerror', error => errors.push(error.message));
      if (init) await page.addInitScript(init, arg);
      await page.goto(url);
      await page.waitForFunction(() => window.FulfillmentToolReady && !document.getElementById('file').disabled);
      assert.equal(await page.locator('main > section').last().getAttribute('id'), 'usage-notes');
      assert.equal(await page.locator('#usage-notes #status').count(), 1);
      assert.equal(await page.locator('#history').evaluate(node => node.nextElementSibling?.id), 'quantityValidation');
      assert.equal(await page.locator('#download-review').count(), 0);
      return page;
    }
    async function upload(page, name = '版本测试更新版.xlsx', buffer = fixture()) {
      await page.locator('#file').setInputFiles({name, mimeType, buffer});
      await page.waitForFunction(() => !document.getElementById('file').disabled && !document.getElementById('file').value);
      assert.equal(await page.locator('#download').isDisabled(), false);
    }
    async function download(page, expected, selector = '#download') {
      const event = page.waitForEvent('download');
      await page.locator(selector).click();
      const file = await event;
      assert.equal(file.suggestedFilename(), expected);
      return fs.readFile(await file.path());
    }
    const page = await open();
    await page.evaluate(key => localStorage.setItem(key, '{broken json'), key);
    await upload(page);
    const first = await download(page, '版本测试汇总版_V1.xlsx');
    const second = await download(page, '版本测试汇总版_V2.xlsx');
    const read = bytes => X.read(bytes, {type: 'buffer'});
    const book = read(first);
    assert.deepEqual(book.SheetNames, ['汇总表', '合并核对', '系列对比', ...C.SHEETS]);
    assert.equal(X.utils.sheet_to_json(book.Sheets.汇总表).length, 4);
    assert.deepEqual(X.utils.sheet_to_json(book.Sheets.系列对比, {header: 1})[0], C.SERIES_COLUMNS);
    const firstSeries = X.utils.sheet_to_json(book.Sheets.系列对比, {header: 1, defval: ''}).slice(1);
    assert.equal(firstSeries.length, 4, '无上期时含零数量的全部分组保留');
    assert(firstSeries.every(row => row[8] === '' && row[9] === ''), '无上期时备注和判断均为空');
    for (const sheet of book.SheetNames) assert.deepEqual(X.utils.sheet_to_json(read(second).Sheets[sheet], {header: 1}), X.utils.sheet_to_json(book.Sheets[sheet], {header: 1}));
    await download(page, '版本测试汇总版-国内事业部_V1.xlsx', '[data-download-division="国内事业部"]');
    await download(page, '版本测试汇总版-国内事业部_V2.xlsx', '[data-download-division="国内事业部"]');
    await page.evaluate(key => localStorage.setItem(key, '{broken again'), key);
    await download(page, '版本测试汇总版_V3.xlsx');
    await page.evaluate(key => localStorage.removeItem(key), key);
    await download(page, '版本测试汇总版_V4.xlsx');
    await page.reload();
    await page.waitForFunction(() => !document.getElementById('download').disabled);
    await download(page, '版本测试汇总版_V5.xlsx');
    assert.equal(await page.evaluate(key => JSON.parse(localStorage.getItem(key))['版本测试汇总版.xlsx'], key), 5);

    let attempts = 0;
    page.on('download', () => attempts++);
    await page.evaluate(() => {
      window.savedVersionTestExport = window.FulfillmentExport;
      window.FulfillmentExport = async () => {throw new Error('版本回归：合成导出失败');};
    });
    await page.locator('#download').click();
    await page.waitForFunction(() => document.getElementById('status').textContent.includes('版本回归：合成导出失败'));
    assert.equal(attempts, 0);
    assert.equal(await page.evaluate(key => JSON.parse(localStorage.getItem(key))['版本测试汇总版.xlsx'], key), 5);
    await page.evaluate(() => {window.FulfillmentExport = window.savedVersionTestExport; delete window.savedVersionTestExport;});
    await download(page, '版本测试汇总版_V6.xlsx');

    const inherited = await open();
    await upload(inherited, '继承汇总版_V5.xlsx', first);
    await download(inherited, '继承汇总版_V6.xlsx');
    await download(inherited, '继承汇总版_V7.xlsx');
    await download(inherited, '继承汇总版-国内事业部_V1.xlsx', '[data-download-division="国内事业部"]');
    await upload(inherited, '其他更新版_V50.xlsx');
    await download(inherited, '其他汇总版_V1.xlsx');

    for (const method of ['getItem', 'setItem']) {
      const unavailable = await open(method => {
        const original = Storage.prototype[method];
        Storage.prototype[method] = function (name, ...args) {
          if (name === 'gylsjqx-fulfillment-export-versions') throw new DOMException('合成存储不可用', 'SecurityError');
          return original.call(this, name, ...args);
        };
      }, method);
      await upload(unavailable, '存储不可用.xlsx');
      await download(unavailable, '存储不可用汇总版_V1.xlsx');
      assert.match(await unavailable.locator('body').innerText(), /版本.{0,40}(保存|记录|浏览器|刷新)/);
      await download(unavailable, '存储不可用汇总版_V2.xlsx');
      await unavailable.close();
    }

    for (const [name, stored] of [
      ['负数', {'负数汇总版.xlsx': -1}],
      ['字符串', {'字符串汇总版.xlsx': '9'}],
      ['数组', []]
    ]) {
      const invalid = await open();
      await invalid.evaluate(([key, stored]) => localStorage.setItem(key, JSON.stringify(stored)), [key, stored]);
      await upload(invalid, name + '.xlsx');
      await download(invalid, name + '汇总版_V1.xlsx');
      await invalid.close();
    }
    const invalid = await open();
    await invalid.evaluate(key => localStorage.setItem(key, JSON.stringify({'上限汇总版.xlsx': Number.MAX_SAFE_INTEGER})), key);
    await upload(invalid, '上限.xlsx');
    let unsafeDownloads = 0;
    invalid.on('download', () => unsafeDownloads++);
    await invalid.locator('#download').click();
    await invalid.waitForFunction(() => document.body.innerText.includes('版本号已达到安全上限'));
    assert.equal(unsafeDownloads, 0);
    assert.equal(await invalid.evaluate(key => JSON.parse(localStorage.getItem(key))['上限汇总版.xlsx'], key), Number.MAX_SAFE_INTEGER);
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({status: 'PASS', url, downloads: 'full and division independently versioned', reviewDownload: 'removed', persistence: 'reload, malformed/cleared/unavailable storage', failedExport: 'no number consumed', sourceVersion: 'same final base only', invalidCounters: 'ignored', safeIntegerLimit: 'blocked', workbookContents: 'unchanged'}));
  } finally {await browser.close();}
})().catch(error => {console.error(error); process.exitCode = 1;});
