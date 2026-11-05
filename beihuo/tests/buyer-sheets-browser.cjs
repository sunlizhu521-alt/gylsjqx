const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const XLSX = require('../vendor/xlsx.full.min.js');

function uploadFile(rows, sheetName, name) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })) };
}

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, timezoneId: 'America/Los_Angeles' });
    await page.clock.install({ time: new Date('2026-11-05T16:30:00Z') });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto((process.env.QA_URL || 'http://127.0.0.1:8940/') + 'beihuo/index.html');
    await page.waitForFunction(() => document.querySelectorAll('[data-upload]').length === 5);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), true);

    const longPrefix = '非常长的采购下单人名称用于验证工作表名称截断后仍然不会发生重名';
    const buyers = [...Array(23).fill('李奇'), '徐英', '', '甲/乙', '甲:乙', longPrefix + '甲', longPrefix + '乙'];
    const records = buyers.map((buyer, index) => ({ buyer, code: `00${String(index + 1).padStart(4, '0')}`, quantity: index === 0 ? 0 : index === 1 ? 1.5 : index === 2 ? '' : index + 1, reason: `合成原因-${index + 1}` }));
    const categoryHeader = Array(22).fill('');
    for (const [index, label] of [[0, '物料编码'], [2, 'SKU'], [3, '物料名称'], [6, '销售产品线'], [7, '销售系列'], [21, '采购组']]) categoryHeader[index] = label;
    const categoryRows = records.map((record, index) => {
      const row = Array(22).fill('');
      for (const [column, value] of [[0, record.code], [2, `SKU-${index + 1}`], [3, `测试物料-${index + 1}`], [6, '测试产品线'], [7, '测试系列'], [21, '测试组']]) row[column] = value;
      return row;
    });
    const uploads = [
      uploadFile([categoryHeader, ...categoryRows], 'Dim-YL医疗器械商品分类', '合成商品分类.xlsx'),
      uploadFile([['物料编码', '采购下单人', '供应商简称', '采购组', '起订量'], ...records.map(record => [record.code, record.buyer, '测试供应商', '测试组', 10])], '产品线明细', '合成采购分工.xlsx'),
      uploadFile([['姓名', '一级部门'], ['测试申请人', '测试事业部']], '花名册', '合成花名册.xlsx'),
      uploadFile([['识别码', '数量', '创建人', '流程号', '备货原因', '采购主体', '要求货好时间'], ...records.map((record, index) => [record.code, record.quantity, '测试申请人', `TEST-OA-${index + 1}`, record.reason, '原采购主体', '2026-12-01'])], '备货需求', '合成备货需求.xlsx'),
    ];
    for (const [index, file] of uploads.entries()) {
      await page.locator(`[data-upload="file-${index + 1}"]`).setInputFiles(file);
      await page.waitForFunction(count => document.querySelector('#uploadedCount').textContent === String(count), index + 1);
    }
    await page.waitForFunction(count => state.filteredRows.length === count, records.length);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), false);
    assert.equal(await page.locator('#detailTableBody tr').count(), 20);
    assert(!(await page.locator('.detail-table').innerText()).includes('合成原因-'));
    await page.locator('#purchaseEntityInput').fill('手填采购主体');
    await page.locator('[data-delivery-days]').first().fill('2');
    await page.locator('[data-delivery-days]').first().press('Tab');

    async function downloadCombined() {
      const event = page.waitForEvent('download');
      await page.locator('#downloadBuyerWorkbookButton').click();
      const download = await event;
      assert.equal(download.suggestedFilename(), '2026-11-06-备货分配分配.xlsx');
      const bytes = await fs.readFile(await download.path());
      const workbook = XLSX.read(bytes, { type: 'buffer', cellNF: true, cellStyles: true });
      const xmls = await page.evaluate(async bytes => {
        const zip = await JSZip.loadAsync(new Uint8Array(bytes));
        return Promise.all(Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/.test(path)).map(async path => ({ path, text: await zip.file(path).async('string') })));
      }, [...bytes]);
      return { workbook, xmls };
    }

    const { workbook, xmls } = await downloadCombined();
    assert.equal(workbook.SheetNames.length, 7);
    assert.equal(new Set(workbook.SheetNames.map(name => name.toLowerCase())).size, 7);
    assert(workbook.SheetNames.includes('未填写下单人'));
    assert(workbook.SheetNames.every(name => name.length <= 31 && !/[\[\]:*?/\\]/.test(name)));
    assert.equal(workbook.Custprops.KSOReadingLayout, true);
    const seen = [];
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      assert.equal(rows[0].length, 19);
      assert.equal(rows[0][18], '是否新加坡公司下单');
      assert.equal(rows[0][0], '采购单订单下单人');
      assert.equal(rows[0][17], '备货原因');
      assert(!rows[0].includes('采购分工明细是否存在'));
      assert.equal(new Set(rows.slice(1).map(row => row[0])).size, 1, `one buyer per sheet: ${name}`);
      assert(sheet['!cols'].every(column => column.width > 0));
      for (const [index, row] of rows.slice(1).entries()) {
        const source = records.find(record => record.code === row[4]);
        assert(source, `known material ${row[4]}`);
        assert.equal(row[0], source.buyer);
        assert.equal(row[7], source.quantity);
        assert.equal(row[9], '手填采购主体');
        assert.equal(row[15], '10');
        assert.equal(row[17], source.reason);
        assert.equal(row[18], '', 'optional Singapore source is not uploaded');
        const r = index + 2;
        assert.equal(sheet[`K${r}`].f, 'TODAY()');
        assert.equal(sheet[`M${r}`].f, `IF(L${r}="","",K${r}+L${r})`);
        assert.equal(sheet[`M${r}`].z, 'yyyy-mm-dd');
        assert.equal(sheet[`H${r}`].z, '0');
        if (source.quantity !== '') assert.equal(sheet[`H${r}`].t, 'n');
        seen.push(row[4]);
      }
    }
    assert.deepEqual([...seen].sort(), records.map(record => record.code).sort());
    assert.equal(XLSX.utils.decode_range(workbook.Sheets['李奇']['!ref']).e.r, 23, 'include rows beyond the first preview page');
    assert.equal(workbook.Sheets['李奇'].L2.v, 2);
    assert.equal(xmls.length, 7);
    for (const { text } of xmls) {
      assert.match(text, /<c[^>]*r="A1"[^>]*s="1"/);
      assert.match(text, /<c[^>]*r="R2"[^>]*s="3"/);
      assert.match(text, /<c[^>]*r="H2"[^>]*s="4"/);
    }

    await page.locator('[data-filter-button="buyer"]').click();
    await page.locator('[data-filter-key="buyer"][value="徐英"]').check();
    await page.waitForFunction(() => state.filteredRows.length === 1);
    const filtered = await downloadCombined();
    assert.deepEqual(filtered.workbook.SheetNames, ['徐英']);
    assert.equal(filtered.workbook.Sheets['徐英'].E2.v, records[23].code);
    await page.locator('#clearFiltersButton').click();
    await page.reload();
    await page.waitForFunction(count => state.filteredRows.length === count, records.length);
    assert.equal(await page.locator('#downloadBuyerWorkbookButton').isDisabled(), false);
    await page.setViewportSize({ width: 390, height: 844 });
    for (const selector of ['#downloadDetailButton', '#downloadBuyerWorkbookButton']) {
      const bounds = await page.locator(selector).boundingBox();
      assert(bounds.x >= 0 && bounds.x + bounds.width <= 391, `button fits phone: ${selector}`);
    }
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8940/');
    await page.locator('.tool-tab[data-tool="beihuo"]').click();
    const embeddedButton = page.frameLocator('iframe[data-tool="beihuo"]').locator('#downloadBuyerWorkbookButton:not([disabled])');
    await embeddedButton.waitFor();
    const homepageEvent = page.waitForEvent('download');
    await embeddedButton.click();
    const homepageDownload = await homepageEvent;
    const homepageWorkbook = XLSX.read(await fs.readFile(await homepageDownload.path()), { type: 'buffer' });
    assert.equal(homepageWorkbook.SheetNames.length, 7);
    assert.equal(homepageWorkbook.SheetNames.reduce((count, name) => count + XLSX.utils.decode_range(homepageWorkbook.Sheets[name]['!ref']).e.r, 0), records.length);
    assert.deepEqual(errors, []);
    console.log('PASS four uploads, one XLSX with seven buyer sheets, 29 rows beyond pagination, blank/invalid/long names, formulas/styles/MOQ/reasons, manual entity, filtering, cache restoration, phone buttons and homepage embedded download');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
