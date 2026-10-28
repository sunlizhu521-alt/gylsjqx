// Opt-in integration check: uses the live conversion engine with synthetic data only.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');

(async () => {
  const out = await fs.mkdtemp(path.join(os.tmpdir(), 'contract-converter-'));
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8940/', { waitUntil: 'networkidle' });
    await page.locator('button[data-tool="contract"]').click();
    const tool = await (await page.locator('iframe[data-tool="contract"]').elementHandle()).contentFrame();
    await tool.waitForFunction(() => window.XLSX && window.ContractCore);
    const files = await tool.evaluate(() => {
      const write = matrix => {
        const book = XLSX.utils.book_new(), sheet = XLSX.utils.aoa_to_sheet(matrix);
        sheet['!cols'] = matrix[1].map(() => ({ wch: 16 }));
        XLSX.utils.book_append_sheet(book, sheet, 'Sheet1');
        return Array.from(new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })));
      };
      const headers = ['序号', '物料编码', '物料名称', 'SKU', '单位', '数量', '含税单价', '含税总金额', '备注'];
      return {
        order: write([['合同编号：', 'TEST-001'], headers, [1, 'MAT-001', '虚构产品A', 'SKU-A', '件', 2, 10.5, 21, ''], [2, 'MAT-002', '虚构产品B', 'SKU-B', '件', 3, 20, 60, '']]),
        template: write([['采购合同'], headers, headers.map(() => '待填写'), ['含税合计（小写）', '', '含税合计（大写）', ''], ['甲方签署日期：', '']]),
      };
    });
    for (const key of ['order', 'template']) {
      await tool.locator(`#${key}File`).setInputFiles({ name: `${key}-synthetic.xlsx`, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(files[key]) });
    }
    await tool.locator('#confirmGenerate').check();
    assert.equal(await tool.locator('#generateContract').isEnabled(), true, await tool.locator('#contractIssues').innerText());
    for (let attempt = 1; attempt <= 2; attempt++) {
      const start = Date.now();
      await tool.locator('#generateContract').click();
      await tool.waitForFunction(() => !document.querySelector('#generateContract').disabled, null, { timeout: 360000 });
      assert.match(await tool.locator('#contractStatus').innerText(), /PDF预览已生成/);
      await tool.waitForFunction(() => document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
      assert.equal(await tool.locator('.onlyoffice-converter-frame').count(), 1);
      await tool.locator('#confirmExport').check();
      const downloadPromise = page.waitForEvent('download');
      await tool.locator('#downloadPdf').click();
      const download = await downloadPromise, pdfPath = path.join(out, `synthetic-${attempt}.pdf`);
      await download.saveAs(pdfPath);
      const bytes = await fs.readFile(pdfPath);
      assert.equal(bytes.subarray(0, 5).toString(), '%PDF-');
      const content = await tool.evaluate(async data => {
        const pdfjs = await import('./vendor/pdf.min.mjs');
        const task = pdfjs.getDocument({ data: new Uint8Array(data), useSystemFonts: true });
        const pdf = await task.promise, text = [];
        for (let i = 1; i <= pdf.numPages; i++) text.push((await (await pdf.getPage(i)).getTextContent()).items.map(item => item.str).join(' '));
        await pdf.destroy();
        return text.join('\n');
      }, Array.from(bytes));
      assert.match(content, /MAT-001/); assert.match(content, /MAT-002/); assert.match(content, /81\.00/);
      console.log(JSON.stringify({ attempt, seconds: Math.round((Date.now() - start) / 1000), bytes: bytes.length, status: await tool.locator('#contractStatus').innerText() }));
    }
    await tool.locator('#resultStage').screenshot({ path: path.join(out, 'preview.png') });
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ ok: true, output: out }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
