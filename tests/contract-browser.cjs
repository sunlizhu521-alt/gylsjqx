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
        ['合同编号：', 'HT-TEST-001', '', '交货地点：', '虚构交货地点', '', '', '', '', ''],
        ['交货时间：', '', '2026-10-01', '', '', '', '', '', '', ''],
        ['采购合同', '', '', '', '', '', '', '', '', ''],
        ['供应商', '物料编码', '物料名称', '规格型号', 'SKU', '单位', '数量', '含税运单价（元）', '含税运总金额（元）', '备注'],
        ['甲公司', 'MAT-001', '产品A', 'A型', 'SKU-A', '件', '2', '10.5', '21', '首批'],
        ['甲公司', 'MAT-002', '产品B', 'B型', 'SKU-B', '件', '3', '20', '60', ''],
      ]);
      sheet['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
      const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '订单');
      return Array.from(new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })));
    });
    const previewPdf = await PDFDocument.create();
    for (const [index, title] of ['GENERATED CONTRACT PAGE 1', 'GENERATED CONTRACT PAGE 2'].entries()) {
      const sheet = previewPdf.addPage([595.28, 841.89]);
      sheet.drawText(title, { x: 92, y: 730, size: 24, color: rgb(.08, .16, .3) });
      sheet.drawText(`Generated PDF page ${index + 1} - this is the file available for download.`, { x: 92, y: 690, size: 12 });
      sheet.drawText('END OF CONTRACT TERMS', { x: 92, y: 28, size: 10, color: rgb(.08, .16, .3) });
    }
    const previewPdfBytes = await previewPdf.save();
    const docxBytes = await page.evaluate(async () => {
      const zip = new JSZip();
      const extraParagraphs = Array.from({ length: 50 }, (_, index) => `<w:p><w:r><w:t>附加条款 ${index + 1}：本条为A4分页预览测试内容。</w:t></w:r></w:p>`).join('');
      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
      zip.folder('_rels').file('.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
      const headers = ['序号', '物料编码', '物料名称', '规格型号', 'SKU', '单位', '数量', '含税运单价（元）', '含税运总金额（元）', '备注'];
      const headerRow = headers.map(value => `<w:tc><w:p><w:r><w:t>${value}</w:t></w:r></w:p></w:tc>`).join('');
      const detailRow = headers.map(() => '<w:tc><w:p><w:r><w:t>待填写</w:t></w:r></w:p></w:tc>').join('');
      const deliveryRow = headers.map((_, index) => `<w:tc><w:p><w:r><w:t>${index === 7 ? '交货时间：' : (index === 8 ? '待填写' : '')}</w:t></w:r></w:p></w:tc>`).join('');
      const labeledRow = (label, placeholder = '待填写') => `<w:tr><w:tc><w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${placeholder}</w:t></w:r></w:p></w:tc></w:tr>`;
      zip.folder('word').file('document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>采购合同</w:t></w:r></w:p><w:p><w:r><w:t>供应商名称：</w:t></w:r></w:p><w:tbl><w:tr>${headerRow}</w:tr><w:tr>${detailRow}</w:tr><w:tr>${deliveryRow}</w:tr>${labeledRow('合同编号：')}${labeledRow('交货地点：')}</w:tbl>${extraParagraphs}<w:sectPr/></w:body></w:document>`);
      return Array.from(await zip.generateAsync({ type: 'uint8array' }));
    });
    await page.locator('#orderFile').setInputFiles({ name: '虚构订单.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(orderBytes) });
    await page.locator('#templateFile').setInputFiles({ name: '虚构合同模板.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(docxBytes) });
    await page.waitForFunction(() => !document.querySelector('#mappingStage').hidden);
    assert.equal(await page.locator('#previewFile').count(), 0);
    assert.equal(await page.locator('.contract-upload').count(), 2);
    assert.equal(await page.locator('#templatePreview').count(), 0);
    assert.equal(await page.locator('#mappingInspector').count(), 0);
    assert.equal(await page.locator('.business-mapping-row').count(), 21);
    assert.deepEqual(await page.locator('.business-mapping-head [role="columnheader"]').allTextContents(), ['映射字段', '订单明细', '合同模板', '是否识别', '写入方式']);
    assert.equal(await page.locator('[data-field-key="materialCode"] option').allTextContents().then(items => items.includes('采购合同')), false);
    assert.equal(await page.locator('[data-field-key="sequence"]').inputValue(), '@sequence');
    assert.equal(await page.locator('[data-field-key="sequence"]').isDisabled(), true);
    for (const [key, field] of [['materialCode', '物料编码'], ['materialName', '物料名称'], ['specification', '规格型号'], ['sku', 'SKU'], ['unit', '单位'], ['quantity', '数量'], ['taxUnitPrice', '含税运单价（元）'], ['taxAmount', '含税运总金额（元）'], ['remark', '备注'], ['deliveryTime', '交货时间（右侧内容）'], ['contractNumber', '合同编号（右侧内容）'], ['deliveryPlace', '交货地点（右侧内容）'], ['supplier', '供应商']]) {
      assert.equal(await page.locator(`[data-field-key="${key}"]`).inputValue(), field);
    }
    assert.match(await page.locator('#templateDetectionSummary').innerText(), /序号按 2 条物料明细自动生成/);
    assert.equal(await page.locator('[data-business-key="materialCode"] .business-template-field').innerText(), '物料编码');
    assert.equal(await page.locator('[data-business-key="materialCode"] .template-detection').innerText(), '识别成功');
    assert.equal(await page.locator('[data-business-key="materialCode"] .business-write-mode').innerText(), '按订单逐行写入');
    assert.equal(await page.locator('[data-business-key="deliveryTime"] .business-template-field').innerText(), '交货时间： → 右侧填写位置');
    assert.equal(await page.locator('[data-business-key="deliveryTime"] .template-detection').innerText(), '识别成功');
    assert.equal(await page.locator('[data-business-key="deliveryTime"] .business-write-mode').innerText(), '填充内容');
    assert.match(await page.locator('[data-business-key="contractNumber"] .business-template-field').innerText(), /右侧填写位置/);
    assert.match(await page.locator('[data-business-key="deliveryPlace"] .business-template-field').innerText(), /右侧填写位置/);
    await page.locator('#mappingStage').screenshot({ path: path.join(out, 'contract-field-mapping.png') });

    await page.locator('#outputName').fill('虚构采购合同');
    assert.match(await page.locator('#confirmSummary').innerText(), /识别 2 条物料明细/);
    await page.locator('#confirmGenerate').check();
    assert.equal(await page.locator('#generateContract').isDisabled(), false);
    await page.evaluate(bytes => {
      let calls = 0;
      window.__CONTRACT_PDF_CONVERTER__ = async () => {
        calls += 1;
        if (calls === 1) return new File([new Uint8Array(120)], '错误.pdf', { type: 'application/pdf' });
        if (calls === 2) return new File(['%PDF-1.4\n', new Uint8Array(160)], '截断.pdf', { type: 'application/pdf' });
        return new File([new Uint8Array([0, 1, 2]), new Uint8Array(bytes)], '虚构采购合同.pdf', { type: 'application/pdf' });
      };
    }, Array.from(previewPdfBytes));
    await page.locator('#generateContract').click();
    await page.waitForFunction(() => !document.querySelector('#resultStage').hidden && document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
    const previewLayout = await page.evaluate(() => {
      const host = document.querySelector('#generatedPreview'), shell = host.querySelector('.pdf-preview-shell'), canvas = host.querySelector('.pdf-preview-canvas');
      const hostRect = host.getBoundingClientRect(), shellRect = shell.getBoundingClientRect(), canvasRect = canvas.getBoundingClientRect();
      const nestedScrollers = [...document.querySelector('#resultStage').querySelectorAll('*')].filter(element => {
        const style = getComputedStyle(element);
        return /(auto|scroll)/.test(style.overflowY) && element.scrollHeight > element.clientHeight + 1;
      }).length;
      return {
        nestedScrollers,
        hostOverflow: getComputedStyle(host).overflowY,
        hostHeight: hostRect.height,
        shellHeight: shellRect.height,
        canvasHeight: canvasRect.height,
        verticallyContained: shellRect.bottom <= hostRect.bottom + 1,
        horizontallyContained: shellRect.left >= hostRect.left - 1 && shellRect.right <= hostRect.right + 1,
      };
    });
    assert.equal(previewLayout.nestedScrollers, 0);
    assert.equal(previewLayout.hostOverflow, 'visible');
    assert.ok(previewLayout.canvasHeight > 900);
    assert.ok(Math.abs(previewLayout.shellHeight - previewLayout.canvasHeight) <= 1);
    assert.equal(previewLayout.verticallyContained, true);
    assert.equal(previewLayout.horizontallyContained, true);
    assert.equal(await page.locator('.native-pdf-preview').count(), 0);
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
    for (const value of ['供应商名称：甲公司', '交货时间：', '>2026-10-01<', '合同编号：', '>HT-TEST-001<', '交货地点：', '>虚构交货地点<', 'MAT-001', 'MAT-002', '产品A', '产品B', 'SKU-A', 'SKU-B', '>1<', '>2<', '>3<']) assert.ok(generatedXml.includes(value), value);
    assert.doesNotMatch(generatedXml, /合同编号：HT-TEST-001|交货地点：虚构交货地点|交货时间：2026-10-01/);
    assert.equal((generatedXml.match(/<w:tr>/g) || []).length, 6);
    const pdfDownloadPromise = page.waitForEvent('download'); await page.locator('#downloadPdf').click(); const pdfDownload = await pdfDownloadPromise;
    assert.equal(pdfDownload.suggestedFilename(), '虚构采购合同.pdf');
    const pdfPath = path.join(out, '虚构采购合同.pdf'); await pdfDownload.saveAs(pdfPath);
    assert.deepEqual(await fs.readFile(pdfPath), Buffer.from(previewPdfBytes));

    await page.screenshot({ path: path.join(out, 'contract-desktop.png'), fullPage: true });

    // Excel template: repeat one detail row while keeping the original template file untouched.
    await page.reload({ waitUntil: 'networkidle' });
    await page.locator('#orderFile').setInputFiles({ name: '虚构订单.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(orderBytes) });
    const templateXlsx = await page.evaluate(async () => {
      const matrix = [['采购合同', '', '', ''], ['序号', '物料名称', '数量', '行金额'], ['待填写', '待填写', '待填写', ''], ['合计', '', '', '']];
      for (let row = 5; row <= 30; row += 1) matrix.push([`附注${row}`, '', '', '']);
      const sheet = XLSX.utils.aoa_to_sheet(matrix);
      sheet.D1 = { t: 'n', f: 'SUM(D3:D3)', v: 0 };
      sheet.D3 = { t: 'n', f: 'C3*10+$C$1+参考!A3+IF(A3="A1",0,0)', v: 0 };
      sheet.D4 = { t: 'n', f: 'SUM(D3:D3)+参考!A4+LOG10(100)', v: 0 };
      sheet['!cols'] = [{ wch: 10 }, { wch: 22 }, { wch: 12 }, { wch: 14 }];
      const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '合同');
      const referenceSheet = XLSX.utils.aoa_to_sheet([['参考数据', ''], [1, ''], [2, ''], [3, '']]);
      referenceSheet.B1 = { t: 'n', f: 'SUM(合同!D3:D3)', v: 0 };
      XLSX.utils.book_append_sheet(book, referenceSheet, '参考');
      const zip = await JSZip.loadAsync(XLSX.write(book, { type: 'array', bookType: 'xlsx' }));
      const sheetPath = 'xl/worksheets/sheet1.xml';
      let sheetXml = await zip.file(sheetPath).async('string');
      sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
        .replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts></worksheet>');
      zip.file(sheetPath, sheetXml);
      zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
      zip.file('xl/tables/table1.xml', '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="ContractTable" displayName="ContractTable" ref="A1:D4" totalsRowShown="0"><autoFilter ref="A1:D4"/><tableColumns count="4"><tableColumn id="1" name="序号"/><tableColumn id="2" name="物料名称"/><tableColumn id="3" name="数量"/><tableColumn id="4" name="行金额"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>');
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
    assert.equal(await page.locator('#templatePreview').count(), 0);
    assert.match(await page.locator('#templateDetectionSummary').innerText(), /序号按 2 条物料明细自动生成/);
    await page.evaluate(() => scrollTo(0, 0));
    await page.screenshot({ path: path.join(out, 'contract-complex-upload.png'), fullPage: false });
    await page.locator('#outputName').fill('虚构Excel合同'); await page.locator('#confirmGenerate').check();
    await page.evaluate(bytes => {
      window.__CONTRACT_PDF_CONVERTER__ = async () => new File([new Uint8Array(bytes)], '虚构Excel合同.pdf', { type: 'application/pdf' });
    }, Array.from(previewPdfBytes));
    await page.locator('#generateContract').click();
    await page.waitForFunction(() => !document.querySelector('#resultStage').hidden && document.querySelector('.pdf-preview-canvas')?.dataset.rendered === 'true');
    assert.equal(await page.locator('.native-pdf-preview').count(), 0);
    assert.equal(await page.locator('#resultPagination').getAttribute('data-total-pages'), '2');
    assert.match(await page.locator('#contractStatus').innerText(), /整页显示/);
    await page.locator('#confirmExport').check();
    const xlsxDownloadPromise = page.waitForEvent('download'); await page.locator('#downloadContract').click(); const xlsxDownload = await xlsxDownloadPromise;
    assert.equal(xlsxDownload.suggestedFilename(), '虚构Excel合同.xlsx');
    const xlsxPath = path.join(out, '虚构Excel合同.xlsx'); await xlsxDownload.saveAs(xlsxPath);
    const generatedCells = await page.evaluate(bytes => {
      const book = XLSX.read(new Uint8Array(bytes), { type: 'array' }), sheet = book.Sheets['合同'];
      return ['A3', 'B3', 'C3', 'A4', 'B4', 'C4', 'A5'].map(address => sheet[address]?.v);
    }, [...await fs.readFile(xlsxPath)]);
    assert.deepEqual(generatedCells, [1, '产品A', 2, 2, '产品B', 3, '合计']);
    const preservedParts = await page.evaluate(async bytes => {
      const zip = await JSZip.loadAsync(new Uint8Array(bytes));
      return {
        sheet: await zip.file('xl/worksheets/sheet1.xml').async('string'),
        referenceSheet: await zip.file('xl/worksheets/sheet2.xml').async('string'),
        table: await zip.file('xl/tables/table1.xml').async('string'),
        drawing: await zip.file('xl/drawings/drawing-contract.xml').async('string'),
        pivot: await zip.file('xl/pivotTables/pivot-contract.xml').async('string'),
      };
    }, [...await fs.readFile(xlsxPath)]);
    assert.match(preservedParts.sheet, /<c r="D1"><f>SUM\(D3:D4\)<\/f><\/c>/);
    assert.match(preservedParts.sheet, /<c r="D3"><f>C3\*10\+\$C\$1\+参考!A3\+IF\(A3="A1",0,0\)<\/f><\/c>/);
    assert.match(preservedParts.sheet, /<c r="D4"><f>C4\*10\+\$C\$1\+参考!A4\+IF\(A4="A1",0,0\)<\/f><\/c>/);
    assert.match(preservedParts.sheet, /<c r="D5"><f>SUM\(D3:D4\)\+参考!A4\+LOG10\(100\)<\/f><\/c>/);
    assert.match(preservedParts.referenceSheet, /<c r="B1"><f>SUM\(合同!D3:D4\)<\/f><\/c>/);
    assert.match(preservedParts.table, /ref="A1:D5"/);
    assert.match(preservedParts.drawing, /DRAWING-MARKER/);
    assert.match(preservedParts.pivot, /PIVOT-MARKER/);
    await page.screenshot({ path: path.join(out, 'contract-excel.png'), fullPage: false });

    assert.deepEqual(errors, []);
    assert.deepEqual(consoleErrors.filter(message => !message.includes('Failed to load resource: the server responded with a status of 404')), []);
    assert.deepEqual(httpErrors, []);
    assert.deepEqual(mutatingRequests, []);
    console.log(JSON.stringify({ status: 'PASS', output: out, screenshots: ['contract-pdf-page2.png', 'contract-field-mapping.png', 'contract-desktop.png', 'contract-complex-upload.png', 'contract-excel.png'] }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
