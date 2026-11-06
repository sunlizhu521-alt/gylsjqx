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
    2: [['正确物料编码', '采购下单人', '备货OA物料编码', '供应商简称', '供应商名称'], ...items.map(([code, supplier, , buyer], index) => [code, buyer, `000${index + 1}`, supplier, supplier])],
    3: [['姓名', '一级部门'], ['测试采购', '测试事业部']],
    4: [['物料编码', '数量', '申请人', '正确物料编码'], ...items.map((_, index) => [`000${index + 1}`, index ? index + 2 : value, '测试采购', '需求表错误编码'])],
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
    assert.equal(await page.locator('.detail-table th').nth(1).innerText(), '是否新加坡公司下单');
    assert.deepEqual(await page.locator('#detailTableBody tr td:nth-child(2)').allTextContents(), ['是', '否']);
    assert.deepEqual(await snapshot(), beforeSingapore, 'Singapore upload preserves quantities, manual entity, days and filters');
    if (process.env.QA_ARTIFACT_DIR) {
      await fs.mkdir(process.env.QA_ARTIFACT_DIR, { recursive: true });
      await page.locator('.detail-panel h2').click();
      await page.locator('[data-drop="file-5"]').screenshot({ path: `${process.env.QA_ARTIFACT_DIR}/singapore-slot.png` });
      await page.locator('.detail-panel .table-wrap').evaluate(element => { element.scrollLeft = 0; });
      await page.locator('.detail-panel .table-wrap').screenshot({ path: `${process.env.QA_ARTIFACT_DIR}/singapore-status.png` });
    }
    const downloadEvent = page.waitForEvent('download');
    await page.locator('#downloadBuyerWorkbookButton').click();
    const download = await downloadEvent;
    const exported = X.read(await fs.readFile(await download.path()), { type: 'buffer' });
    const sheet = exported.Sheets['采购甲'];
    assert.equal(sheet.B1.v, '是否新加坡公司下单');
    assert.deepEqual([sheet.B2.v, sheet.B3.v], ['是', '否']);
    assert.equal(sheet.K2.v, '手填采购主体');
    assert.equal(sheet.M2.v, 9);
    assert.equal(sheet.O1.v, '采购单订单号');
    assert.equal(sheet.P1.v, '备货OA的物料编码');
    assert.deepEqual([sheet.P2.v, sheet.P3.v], ['0001', '0002']);
    assert.equal(sheet.P2.t, 's');
    assert.equal(sheet.F1.v, '正确物料编码');
    assert.equal(await page.locator('.detail-table th').nth(5).innerText(), '正确物料编码');
    assert.deepEqual([sheet.F2.v, sheet.F3.v], ['TEST-001', 'TEST-002']);
    assert.equal(await page.locator('#materialCodeWarnings').isHidden(), true);

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

    function missingMappingFile(withCorrectColumn = true) {
      const rows = [['备货OA物料编码', '正确物料编码', '采购下单人', '供应商名称', '供应商简称'],
        ['0001', 'TEST-001', '采购甲', '供应商甲', '供应商甲'], [], ['0002', '', '采购甲', '供应商甲', '供应商甲']];
      const w = X.utils.book_new();
      X.utils.book_append_sheet(w, X.utils.aoa_to_sheet(withCorrectColumn ? rows : rows.map(row => row.filter((_, index) => index !== 1))), '产品线明细');
      return { name: '分工编码待补充.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(X.write(w, { type: 'buffer', bookType: 'xlsx' })) };
    }
    const blankDemand = X.utils.book_new();
    X.utils.book_append_sheet(blankDemand, X.utils.aoa_to_sheet([[], ['物料编码', '数量', '申请人'], [],
      ['0001', 2, '测试采购'], ['0002', 0, '测试采购'], ['0002', 1.5, '测试采购']]), '待补充');
    await page.locator('[data-upload="file-2"]').setInputFiles(missingMappingFile());
    await page.waitForFunction(() => state.records.get('file-2')?.name === '分工编码待补充.xlsx');
    await page.locator('[data-upload="file-4"]').setInputFiles({name: '待补充编码.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(X.write(blankDemand, { type: 'buffer', bookType: 'xlsx' }))});
    await page.waitForFunction(() => state.demandRows.length === 3 && state.demandRows[2].materialCodeIssue);
    await page.locator('#materialCodeWarnings summary').click();
    assert.match(await page.locator('#materialCodeWarnings').innerText(), /待补充编码.xlsx \/ 待补充 第 5 行.*请在分工表补充正确物料编码/);
    assert.match(await page.locator('#materialCodeWarnings').innerText(), /分工编码待补充.xlsx \/ 产品线明细 第 4 行/);
    assert.match(await page.locator('#materialCodeWarnings').innerText(), /第 6 行/);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), false);
    await page.reload();
    await ready();
    await page.waitForFunction(() => state.demandRows.length === 3);
    assert.equal(await page.locator('#materialCodeWarnings').isHidden(), false);
    const blankDownloadEvent = page.waitForEvent('download');
    await page.locator('#downloadBuyerWorkbookButton').click();
    const blankDownload = await blankDownloadEvent;
    const blankWorkbook = X.read(await fs.readFile(await blankDownload.path()), { type: 'buffer' });
    const blanks = X.utils.sheet_to_json(blankWorkbook.Sheets.未填写下单人, { header: 1, defval: '' });
    assert.deepEqual(blanks.slice(1).map(row => [row[5], row[8], row[15]]), [['', 0, '0002'], ['', 1.5, '0002']]);
    assert.equal(blankWorkbook.Sheets.采购甲.F2.v, 'TEST-001');
    await page.locator('[data-upload="file-2"]').setInputFiles(missingMappingFile(false));
    await page.waitForFunction(() => state.demandRows.length === 3 && state.demandRows.every(row => !row.materialCode && row.materialCodeIssue.includes('缺少')));
    assert.match(await page.locator('#materialCodeWarnings').textContent(), /分工表缺少“正确物料编码”列/);
    await upload(2, 25, 5);
    await upload(4, 25, 5);
    await page.waitForFunction(() => state.demandRows.length === 4 && state.demandRows.every(row => !row.materialCodeIssue));
    assert.equal(await page.locator('#materialCodeWarnings').isHidden(), true);

    page.once('dialog', dialog => dialog.accept());
    await page.locator('#clearCacheButton').click();
    await page.waitForFunction(() => document.querySelector('#uploadedCount').textContent === '0');
    await page.reload();
    await ready();
    assert.equal(await page.locator('#uploadedCount').innerText(), '0');
    assert.equal(await page.locator('#appliedCount').innerText(), '0');
    assert.deepEqual(errors, []);
    console.log('PASS division-only corrected-code mapping, ignored demand corrected codes, combined original-code column, blank/missing division mappings with source rows, cache restore/replacement, optional Singapore matching, rollback and unchanged quantities/manual entity/days/filters');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
