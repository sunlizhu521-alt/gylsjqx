const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const runtimeNodeModules = process.env.PLAYWRIGHT_MODULE ? path.dirname(process.env.PLAYWRIGHT_MODULE) : '';
const { PDFDocument, rgb } = require(runtimeNodeModules ? path.join(runtimeNodeModules, 'pdf-lib') : 'pdf-lib');

const QA_URL = process.env.QA_URL || 'http://127.0.0.1:8899/contract.html';
const HOME_URL = new URL('.', QA_URL).toString();

(async () => {
  const out = process.env.QA_OUTPUT || await fs.mkdtemp(path.join(os.tmpdir(), 'contract-browser-'));
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
    const page = await context.newPage();
    const errors = [], consoleErrors = [], httpErrors = [], mutatingRequests = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text()); });
    page.on('response', response => { if (response.status() >= 400 && !/favicon\.ico(?:\?|$)/.test(response.url())) httpErrors.push(`${response.status()} ${response.url()}`); });
    page.on('request', request => { if (!['GET', 'HEAD'].includes(request.method())) mutatingRequests.push(`${request.method()} ${request.url()}`); });
    await page.goto(HOME_URL, { waitUntil: 'networkidle' });
    await page.locator('[data-tool="contract"]').click();
    await page.waitForURL(/contract\.html$/);
    assert.equal(await page.title(), '合同生成｜供应链数据清洗');
    assert.match(await page.locator('h1').innerText(), /合同生成/);
    assert.doesNotMatch(await page.locator('body').innerText(), /登录/);

    const orderBytes = await page.evaluate(() => {
      const sheet = XLSX.utils.aoa_to_sheet([
        ['供应商', '物料', '数量'],
        ['甲公司', '产品A', '2'],
        ['甲公司', '产品B', '3'],
      ]);
      const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '订单');
      return Array.from(new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })));
    });
    const previewPdf = await PDFDocument.create();
    for (const [index, title] of ['GENERATED CONTRACT PAGE 1', 'GENERATED CONTRACT PAGE 2'].entries()) {
      const sheet = previewPdf.addPage([595.28, 841.89]);
      sheet.drawText(title, { x: 92, y: 730, size: 24, color: rgb(.08, .16, .3) });
      sheet.drawText(`Generated PDF page ${index + 1} - this is the file available for download.`, { x: 92, y: 690, size: 12 });
    }
    const previewPdfBytes = await previewPdf.save();
    const docxBytes = await page.evaluate(async () => {
      const zip = new JSZip();
      const extraParagraphs = Array.from({ length: 50 }, (_, index) => `<w:p><w:r><w:t>附加条款 ${index + 1}：本条为A4分页预览测试内容。</w:t></w:r></w:p>`).join('');
      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
      zip.folder('_rels').file('.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
      zip.folder('word').file('document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>采购合同</w:t></w:r></w:p><w:p><w:r><w:t>供应商名称</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>物料</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>待填写</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>待填写</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${extraParagraphs}<w:sectPr/></w:body></w:document>`);
      return Array.from(await zip.generateAsync({ type: 'uint8array' }));
    });
    await page.locator('#orderFile').setInputFiles({ name: '虚构订单.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(orderBytes) });
    await page.locator('#templateFile').setInputFiles({ name: '虚构合同模板.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(docxBytes) });
    await page.waitForFunction(() => !document.querySelector('#mappingStage').hidden);
    assert.equal(await page.locator('#previewFile').count(), 0);
    assert.equal(await page.locator('.contract-upload').count(), 2);
    assert.match(await page.locator('.mapping-structure-note').innerText(), /生成PDF预览/);
    await page.locator('#mappingStage').screenshot({ path: path.join(out, 'contract-mapping-structure.png') });

    await page.locator('[data-target="p:1"]').click();
    await page.locator('#mapField').selectOption('供应商');
    assert.equal(await page.locator('#mapMode').inputValue(), 'single');
    await page.locator('[data-detail-row="word:0:1"]').click();
    await page.locator('[data-target="t:0:r:1:c:0"]').click();
    await page.locator('#mapField').selectOption('物料');
    await page.locator('[data-target="t:0:r:1:c:1"]').click();
    await page.locator('#mapField').selectOption('数量');
    await page.locator('#outputName').fill('虚构采购合同');
    await page.locator('#confirmGenerate').check();
    assert.equal(await page.locator('#generateContract').isDisabled(), false);
    await page.evaluate(bytes => {
      window.__CONTRACT_PDF_CONVERTER__ = async () => new File([new Uint8Array(bytes)], '虚构采购合同.pdf', { type: 'application/pdf' });
    }, Array.from(previewPdfBytes));
    await page.locator('#generateContract').click();
    await page.waitForFunction(() => !document.querySelector('#resultStage').hidden && document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
    assert.equal(await page.locator('#resultPagination').getAttribute('data-total-pages'), '2');
    assert.equal(await page.locator('#resultPrevious').isDisabled(), true);
    assert.equal(await page.locator('#downloadContract').isDisabled(), true);
    assert.equal(await page.locator('#downloadPdf').isDisabled(), true);
    await page.locator('#resultNext').click();
    await page.waitForFunction(() => document.querySelector('.pdf-preview-canvas')?.dataset.page === '2' && document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
    assert.equal(await page.locator('#resultPageLabel').innerText(), '第 2 / 2 页');
    await page.locator('#resultStage').screenshot({ path: path.join(out, 'contract-pdf-page2.png') });
    await page.locator('#confirmExport').check();
    assert.equal(await page.locator('#downloadContract').isDisabled(), false);
    assert.equal(await page.locator('#downloadPdf').isDisabled(), false);
    const downloadPromise = page.waitForEvent('download'); await page.locator('#downloadContract').click(); const download = await downloadPromise;
    assert.equal(download.suggestedFilename(), '虚构采购合同.docx');
    const docxPath = path.join(out, '虚构采购合同.docx'); await download.saveAs(docxPath);
    const generatedXml = await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')), [...await fs.readFile(docxPath)]);
    for (const value of ['甲公司', '产品A', '产品B', '>2<', '>3<']) assert.ok(generatedXml.includes(value), value);
    assert.equal((generatedXml.match(/<w:tr>/g) || []).length, 3);
    const pdfDownloadPromise = page.waitForEvent('download'); await page.locator('#downloadPdf').click(); const pdfDownload = await pdfDownloadPromise;
    assert.equal(pdfDownload.suggestedFilename(), '虚构采购合同.pdf');
    const pdfPath = path.join(out, '虚构采购合同.pdf'); await pdfDownload.saveAs(pdfPath);
    assert.deepEqual(await fs.readFile(pdfPath), Buffer.from(previewPdfBytes));

    await page.screenshot({ path: path.join(out, 'contract-desktop.png'), fullPage: true });

    // Excel template: repeat one detail row while keeping the original template file untouched.
    await page.reload({ waitUntil: 'networkidle' });
    await page.locator('#orderFile').setInputFiles({ name: '虚构订单.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(orderBytes) });
    const templateXlsx = await page.evaluate(async () => {
      const matrix = [['采购合同', ''], ['物料', '数量'], ['待填写', '待填写']];
      for (let row = 4; row <= 30; row += 1) matrix.push([`附注${row}`, '']);
      const sheet = XLSX.utils.aoa_to_sheet(matrix);
      sheet['!cols'] = [{ wch: 22 }, { wch: 12 }];
      const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '合同');
      const zip = await JSZip.loadAsync(XLSX.write(book, { type: 'array', bookType: 'xlsx' }));
      const sheetPath = 'xl/worksheets/sheet1.xml';
      let sheetXml = await zip.file(sheetPath).async('string');
      sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
        .replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts></worksheet>');
      zip.file(sheetPath, sheetXml);
      zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
      zip.file('xl/tables/table1.xml', '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="ContractTable" displayName="ContractTable" ref="A1:B3" totalsRowShown="0"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="物料"/><tableColumn id="2" name="数量"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>');
      zip.file('xl/drawings/drawing-contract.xml', '<drawing-preserve>DRAWING-MARKER</drawing-preserve>');
      zip.file('xl/pivotTables/pivot-contract.xml', '<pivot-preserve>PIVOT-MARKER</pivot-preserve>');
      let types = await zip.file('[Content_Types].xml').async('string');
      types = types.replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
      zip.file('[Content_Types].xml', types);
      return Array.from(await zip.generateAsync({ type: 'uint8array' }));
    });
    await page.locator('#templateFile').setInputFiles({ name: '虚构合同模板.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(templateXlsx) });
    await page.waitForFunction(() => !document.querySelector('#mappingStage').hidden);
    assert.match(await page.locator('#contractStatus').innerText(), /已识别并保留图形、数据透视表、结构化表/);
    assert.match(await page.locator('#templatePreview').innerText(), /附注23/);
    await page.evaluate(() => scrollTo(0, 0));
    await page.screenshot({ path: path.join(out, 'contract-complex-upload.png'), fullPage: false });
    await page.locator('[data-detail-row="excel:3"]').click();
    await page.locator('[data-target="x:A3"]').click(); await page.locator('#mapField').selectOption('物料');
    await page.locator('[data-target="x:B3"]').click(); await page.locator('#mapField').selectOption('数量');
    await page.locator('#outputName').fill('虚构Excel合同'); await page.locator('#confirmGenerate').check();
    await page.evaluate(bytes => {
      window.__CONTRACT_PDF_CONVERTER__ = async () => new File([new Uint8Array(bytes)], '虚构Excel合同.pdf', { type: 'application/pdf' });
    }, Array.from(previewPdfBytes));
    await page.locator('#generateContract').click();
    await page.waitForFunction(() => !document.querySelector('#resultStage').hidden && document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
    await page.locator('#confirmExport').check();
    const xlsxDownloadPromise = page.waitForEvent('download'); await page.locator('#downloadContract').click(); const xlsxDownload = await xlsxDownloadPromise;
    assert.equal(xlsxDownload.suggestedFilename(), '虚构Excel合同.xlsx');
    const xlsxPath = path.join(out, '虚构Excel合同.xlsx'); await xlsxDownload.saveAs(xlsxPath);
    const generatedCells = await page.evaluate(bytes => {
      const book = XLSX.read(new Uint8Array(bytes), { type: 'array' }), sheet = book.Sheets['合同'];
      return ['A3', 'B3', 'A4', 'B4'].map(address => sheet[address]?.v);
    }, [...await fs.readFile(xlsxPath)]);
    assert.deepEqual(generatedCells, ['产品A', '2', '产品B', '3']);
    const preservedParts = await page.evaluate(async bytes => {
      const zip = await JSZip.loadAsync(new Uint8Array(bytes));
      return {
        table: await zip.file('xl/tables/table1.xml').async('string'),
        drawing: await zip.file('xl/drawings/drawing-contract.xml').async('string'),
        pivot: await zip.file('xl/pivotTables/pivot-contract.xml').async('string'),
      };
    }, [...await fs.readFile(xlsxPath)]);
    assert.match(preservedParts.table, /ref="A1:B4"/);
    assert.match(preservedParts.drawing, /DRAWING-MARKER/);
    assert.match(preservedParts.pivot, /PIVOT-MARKER/);
    await page.screenshot({ path: path.join(out, 'contract-excel.png'), fullPage: false });

    await page.setViewportSize({ width: 390, height: 844 });
    await page.locator('#resultStage').scrollIntoViewIfNeeded();
    await page.waitForFunction(() => {
      const canvas = document.querySelector('.pdf-preview-canvas'), host = document.querySelector('#generatedPreview');
      return canvas?.dataset.rendered === 'true' && canvas.getBoundingClientRect().width <= host.clientWidth - 18;
    });
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth);
    assert.equal(overflow, false);
    await page.screenshot({ path: path.join(out, 'contract-mobile.png'), fullPage: false });
    assert.deepEqual(errors, []);
    assert.deepEqual(consoleErrors.filter(message => !message.includes('Failed to load resource: the server responded with a status of 404')), []);
    assert.deepEqual(httpErrors, []);
    assert.deepEqual(mutatingRequests, []);
    console.log(JSON.stringify({ status: 'PASS', output: out, screenshots: ['contract-pdf-page2.png', 'contract-mapping-structure.png', 'contract-desktop.png', 'contract-complex-upload.png', 'contract-excel.png', 'contract-mobile.png'] }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
