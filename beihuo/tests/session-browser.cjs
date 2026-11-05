const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');

const items = [
  ['TEST-001', '供应商甲', '系列甲', '采购甲'],
  ['TEST-002', '供应商甲', '系列乙', '采购甲'],
  ['TEST-003', '供应商乙', '系列甲', '采购乙'],
  ['TEST-004', '供应商乙', '系列乙', '采购乙'],
];
function fileForSlot(slot, value) {
  const rows = {
    1: [['物料编码', '', 'SKU', '物料名称', '销售系列'], ...items.map(([code, , series]) => [code, '', `SKU-${code}`, code, series])],
    2: [['物料编码', '采购下单人', '供应商简称', '供应商名称'], ...items.map(([code, supplier, , buyer]) => [code, buyer, supplier, supplier])],
    3: [['姓名', '一级部门'], ['测试采购', '测试事业部']],
    4: [['物料编码', '数量', '申请人'], ...items.map(([code], index) => [code, index ? index + 2 : value, '测试采购'])],
    5: value === 444 ? [['供应商名称'], ['供应商甲']] : [['供应商名称', '系列'], ['供应商甲', value === 111 ? '系列甲' : '系列乙'], ['供应商乙', value === 111 ? '系列乙' : '系列甲']],
  }[slot];
  const workbook = X.utils.book_new();
  X.utils.book_append_sheet(workbook, X.utils.aoa_to_sheet(rows), slot === 2 ? '产品线明细' : '数据');
  return { name: slot === 5 ? '新加坡公司明细.xlsx' : `同名最新版-${slot}.xlsx`, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(X.write(workbook, { type: 'buffer', bookType: 'xlsx' })) };
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto((process.env.QA_URL || 'http://127.0.0.1:8940/') + 'beihuo/index.html');
    const ready = () => page.waitForFunction(() => document.querySelectorAll('[data-upload]').length === 5 && document.querySelector('#libraryState').textContent === '已恢复最近上传文件');
    await ready();
    async function upload(slot, value, expectedCount) {
      const previous = await page.evaluate(slot => state.records.get(`file-${slot}`)?.savedAt, slot);
      await page.locator(`[data-upload="file-${slot}"]`).setInputFiles(fileForSlot(slot, value));
      await page.waitForFunction(({ slot, previous }) => {
        const record = state.records.get(`file-${slot}`);
        return record?.file && record.savedAt !== previous;
      }, { slot, previous });
      await page.waitForFunction(count => document.querySelector('#uploadedCount').textContent === String(count), expectedCount);
      await ready();
    }
    async function savedSeries() {
      return page.evaluate(async () => {
        const record = state.records.get('file-5');
        const workbook = XLSX.read(await record.file.arrayBuffer(), { type: 'array' });
        return workbook.Sheets['数据'].B2.v;
      });
    }
    const snapshot = () => page.evaluate(() => {
      const withoutSingapore = ({ singaporeOrder, ...row }) => row;
      return {
        rows: state.demandRows.map(withoutSingapore),
        filtered: state.filteredRows.map(withoutSingapore),
        filters: [...state.filters].map(([key, values]) => [key, [...values]]),
        entity: state.purchaseEntityOverride,
        total: document.querySelector('#beihuoTotal').textContent,
      };
    });
    const statuses = () => page.evaluate(() => state.demandRows.map(row => row.singaporeOrder));
    async function failNextSave() {
      await page.evaluate(() => {
        const original = IDBObjectStore.prototype.put;
        IDBObjectStore.prototype.put = function (...args) {
          IDBObjectStore.prototype.put = original;
          const request = original.apply(this, args);
          this.transaction.abort();
          return request;
        };
      });
    }

    assert.equal(await page.locator('[data-drop="file-5"] h2').innerText(), '新加坡公司明细');
    for (let slot = 1; slot <= 4; slot += 1) await upload(slot, 10, slot);
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '22');
    assert.equal(await page.locator('#appliedCount').innerText(), '4');
    assert.equal(await page.locator('#downloadDetailButton').isDisabled(), false);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), false);
    assert.deepEqual(await statuses(), ['', '', '', '']);
    await upload(4, 25, 4);
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '37');
    await page.reload();
    await ready();
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '37');
    assert.equal(await page.locator('#uploadedCount').innerText(), '4');

    await page.locator('#purchaseEntityInput').fill('手填采购主体');
    await page.locator('[data-delivery-days="0"]').fill('9');
    await page.locator('[data-delivery-days="0"]').press('Tab');
    await page.locator('[data-filter-button="buyer"]').click();
    await page.locator('[data-filter-key="buyer"][value="采购甲"]').check();
    await page.waitForFunction(() => state.filteredRows.length === 2);
    const beforeSingapore = await snapshot();

    await upload(5, 111, 5);
    assert.equal(await savedSeries(), '系列甲');
    assert.equal(await page.locator('[data-drop="file-5"] .slot-status').innerText(), '已引用');
    assert.equal(await page.locator('[data-apply="file-5"]').count(), 1);
    assert.equal(await page.locator('#appliedCount').innerText(), '5');
    assert.deepEqual(await statuses(), ['是', '否', '否', '是'], 'supplier and series must match the same source row, not a cross product');
    assert.equal(await page.locator('.detail-table th').last().innerText(), '是否新加坡公司下单');
    assert.deepEqual(await page.locator('#detailTableBody tr td:last-child').allTextContents(), ['是', '否']);
    assert.deepEqual(await snapshot(), beforeSingapore, 'Singapore upload preserves quantities, manual entity, days and filters');
    if (process.env.QA_ARTIFACT_DIR) {
      await fs.mkdir(process.env.QA_ARTIFACT_DIR, { recursive: true });
      await page.locator('.detail-panel h2').click();
      await page.locator('[data-drop="file-5"]').screenshot({ path: `${process.env.QA_ARTIFACT_DIR}/singapore-slot.png` });
      await page.locator('.detail-panel .table-wrap').evaluate(element => { element.scrollLeft = element.scrollWidth; });
      await page.locator('.detail-panel .table-wrap').screenshot({ path: `${process.env.QA_ARTIFACT_DIR}/singapore-status.png` });
    }
    const downloadEvent = page.waitForEvent('download');
    await page.locator('#downloadBuyerWorkbookButton').click();
    const download = await downloadEvent;
    const exported = X.read(await fs.readFile(await download.path()), { type: 'buffer' });
    const sheet = exported.Sheets['采购甲'];
    assert.equal(sheet.S1.v, '是否新加坡公司下单');
    assert.deepEqual([sheet.S2.v, sheet.S3.v], ['是', '否']);
    assert.equal(sheet.J2.v, '手填采购主体');
    assert.equal(sheet.L2.v, 9);

    await page.locator('[data-upload="file-5"]').setInputFiles(fileForSlot(5, 444));
    await page.waitForFunction(() => document.querySelector('#libraryState').textContent.includes('新加坡公司明细未更新，原文件保留'));
    assert.equal(await savedSeries(), '系列甲', 'bad headers cannot replace the last valid file');
    assert.deepEqual(await statuses(), ['是', '否', '否', '是']);
    assert.deepEqual(await snapshot(), beforeSingapore, 'invalid upload preserves allocation, manual days and filters');
    await upload(5, 222, 5);
    assert.equal(await savedSeries(), '系列乙');
    assert.deepEqual(await statuses(), ['否', '是', '是', '否']);
    assert.deepEqual(await snapshot(), beforeSingapore, 'same-name replacement only updates Singapore matching');
    await failNextSave();
    await page.locator('[data-upload="file-5"]').setInputFiles(fileForSlot(5, 111));
    await page.waitForFunction(() => document.querySelector('#libraryState').textContent.includes('文件保存失败'));
    assert.equal(await savedSeries(), '系列乙');
    assert.deepEqual(await statuses(), ['否', '是', '是', '否']);
    assert.deepEqual(await snapshot(), beforeSingapore, 'failed save keeps last file and allocation');
    await page.reload();
    await ready();
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '37');
    assert.equal(await savedSeries(), '系列乙', 'last successful file survives reload');
    assert.deepEqual(await statuses(), ['否', '是', '是', '否']);
    assert.equal(await page.locator('#uploadedCount').innerText(), '5');
    assert.equal(await page.locator('#appliedCount').innerText(), '5');

    await page.locator('[data-delivery-days="0"]').fill('7');
    await page.locator('[data-delivery-days="0"]').press('Tab');
    await page.locator('[data-filter-button="buyer"]').click();
    await page.locator('[data-filter-key="buyer"][value="采购甲"]').check();
    await page.waitForFunction(() => state.filteredRows.length === 2);
    const beforeDelete = await snapshot();
    await page.locator('[data-delete="file-5"]').click();
    await page.waitForFunction(() => document.querySelector('#uploadedCount').textContent === '4');
    assert.deepEqual(await statuses(), ['', '', '', '']);
    assert.deepEqual(await snapshot(), beforeDelete, 'deletion clears only Singapore matching, not edits or selected rows');
    assert.equal(await page.locator('#appliedCount').innerText(), '4');
    await page.reload();
    await ready();
    assert.equal(await page.evaluate(() => Boolean(state.records.get('file-5')?.file)), false);
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '37');

    await failNextSave();
    await page.locator('[data-upload="file-4"]').setInputFiles(fileForSlot(4, 99));
    await page.waitForFunction(() => document.querySelector('#libraryState').textContent.includes('文件保存失败'));
    await page.reload();
    await ready();
    await page.waitForFunction(() => document.querySelector('#beihuoTotal').textContent === '37');
    await page.locator('[data-delete="file-4"]').click();
    await page.waitForFunction(() => document.querySelector('#uploadedCount').textContent === '3');
    await page.reload();
    await ready();
    assert.equal(await page.locator('#uploadedCount').innerText(), '3');
    assert.equal(await page.locator('#downloadDetailButton').isDisabled(), true);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), true);

    await upload(5, 111, 4);
    assert.equal(await page.locator('#appliedCount').innerText(), '4');
    assert.equal(await page.locator('#downloadDetailButton').isDisabled(), true, 'Singapore file does not replace the required demand file');
    page.once('dialog', dialog => dialog.accept());
    await page.locator('#clearCacheButton').click();
    await page.waitForFunction(() => document.querySelector('#uploadedCount').textContent === '0');
    await page.reload();
    await ready();
    assert.equal(await page.locator('#uploadedCount').innerText(), '0');
    assert.equal(await page.locator('#appliedCount').innerText(), '0');
    assert.deepEqual(errors, []);
    console.log('PASS optional Singapore upload, exact supplier+series pairs, preview/export, latest replacement, reload, failed-save rollback, deletion, clear cache and unchanged quantities/manual entity/days/filters');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
