// Synthetic files only, including the opt-in live PDF conversion check.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const live = process.env.QA_LIVE_CONVERTER === '1';
const url = process.env.QA_URL || 'http://127.0.0.1:8973/contract.html';

(async () => {
  const out = process.env.QA_OUTPUT || await fs.mkdtemp(path.join(os.tmpdir(), 'contract-a4-'));
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ acceptDownloads: true });
    if (!live) await page.route('**/*', route => new URL(route.request().url()).origin === new URL(url).origin ? route.continue() : route.abort());
    const errors = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto(url);
    await page.waitForFunction(() => window.XLSX && window.JSZip && window.PDFLib);
    if (!live) await page.evaluate(() => {
      window.__CONTRACT_PDF_CONVERTER__ = async () => {
        const pdf = await PDFLib.PDFDocument.create(); pdf.addPage([595.28, 841.89]);
        return new File([await pdf.save()], 'synthetic-preview.pdf', { type: 'application/pdf' });
      };
    });
    for (const count of live ? [11] : [1, 11, 50]) {
      const files = await page.evaluate(async count => {
        const headers = ['序号', '物料编码', '物料名称', '规格型号', 'SKU', '单位', '数量', '含税运单价（元）', '含税运总金额（元）', '备注'];
        const write = sheet => {
          const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '合同');
          return XLSX.write(book, { type: 'array', bookType: 'xlsx' });
        };
        const order = write(XLSX.utils.aoa_to_sheet([headers, ...Array.from({ length: count }, (_, i) => [i + 1, `MAT-${i + 1}`, `测试控制器连接线${i + 1}`, '/', '/', '根', 20, 10, 200, '2026/12/18'])]));
        const terms = ('乙方应根据甲方要求进行包装，产品所有权及风险转移到甲方。本合同一式二份。\n').repeat(6);
        const sheet = XLSX.utils.aoa_to_sheet([
          ['采购合同'], [], ['', '', '', '', '', '', '', '合同编号：'],
          ['采购方（甲方）', '', '虚构采购公司', '', '', '供应方（乙方）', '', '虚构供应公司'],
          ['地址', '', '测试地址一号楼二层', '', '', '地址', '', '测试工业园区'],
          ['授权代表人', '', '测试甲', '', '', '授权代表人', '', '测试乙'], headers,
          headers.map(() => '待填写'),
          ['含税运合计（小写）', '', '0元', '', '', '含税运合计（大写）', '', '零元整'],
          ['交货地点', '', '测试收货地址', '', '', '交货时间', '', '详见备注'],
          ['收货人姓名', '', '测试收货人', '', '', '联系方式', '', '/'],
          ['账期', '', '有账期', '', '', '预付款比例', '', 0],
          ['付款方式', '', '验收完成后按约付款。'], ['发票类型', '', '增值税专用发票（税率13%）'],
          ['乙方收款信息', '', '户名：虚构供应公司\n开户行：测试银行\n账号：测试账号\n行号：测试行号'],
          ['合同条款', '', terms], ['甲方（盖章）：', '', '', '', '', '乙方（盖章）：'],
          ['授权代表人及联系电话：测试甲', '', '', '', '', '授权代表人及联系电话：测试乙'],
          ['签署日期：', '', '', '', '', '签署日期：'],
        ]);
        sheet['!cols'] = [4.5, 10, 14.25, 9.125, 5.75, 6.75, 6.125, 9.75, 12.125, 7.75].map(width => ({ width }));
        sheet['!rows'] = [18,18,18,18,27,18,40,32.1,30,35,18,23.1,18,50,75,309,18,18,18].map(hpt => ({ hpt }));
        const ranges = ['A1:J2', 'A3:G3', 'H3:J3'];
        for (const r of [4,5,6,9,10,11,12]) ranges.push(`A${r}:B${r}`, `C${r}:E${r}`, `F${r}:G${r}`, `H${r}:J${r}`);
        for (const r of [13,14,15,16]) ranges.push(`A${r}:B${r}`, `C${r}:J${r}`);
        for (const r of [17,18,19]) ranges.push(`A${r}:E${r}`, `F${r}:J${r}`);
        sheet['!merges'] = ranges.map(XLSX.utils.decode_range);
        const zip = await JSZip.loadAsync(write(sheet));
        zip.file('xl/styles.xml', (await zip.file('xl/styles.xml').async('string')).replace('<sz val="12"/>', '<sz val="10"/>').replace('<name val="Calibri"/>', '<name val="宋体"/>'));
        return { order: Array.from(new Uint8Array(order)), template: Array.from(await zip.generateAsync({ type: 'uint8array' })) };
      }, count);
      for (const key of ['order', 'template']) {
        await page.locator(`#${key}File`).setInputFiles({ name: `synthetic-${key}-${count}.xlsx`, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(files[key]) });
        await page.waitForFunction(({ key, count }) => document.querySelector(`#${key}FileName`).textContent === `synthetic-${key}-${count}.xlsx`, { key, count });
      }
      await page.locator('[data-manual-key="signDate"]').fill('2026-11-24');
      await page.locator('#confirmGenerate').check();
      assert.equal(await page.locator('#generateContract').isEnabled(), true, await page.locator('#contractIssues').innerText());
      await page.locator('#generateContract').click();
      await page.waitForFunction(() => !document.querySelector('#resultStage').hidden || document.querySelector('#contractStatus').classList.contains('error'), null, { timeout: 360000 });
      assert.match(await page.locator('#contractStatus').innerText(), /PDF预览已生成/);
      await page.locator('#confirmExport').check();
      let pending = page.waitForEvent('download'); await page.locator('#downloadContract').click();
      const download = await pending, file = path.join(out, `contract-${count}.xlsx`); await download.saveAs(file);
      const result = await page.evaluate(async bytes => {
        const zip = await JSZip.loadAsync(new Uint8Array(bytes)), xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        const doc = new DOMParser().parseFromString(xml, 'application/xml'), ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
        const all = tag => [...doc.getElementsByTagNameNS(ns, tag)], attrs = node => Object.fromEntries([...node.attributes].map(a => [a.name, a.value]));
        const sheet = XLSX.read(new Uint8Array(bytes), { type: 'array' }).Sheets['合同'];
        return { setup: attrs(all('pageSetup')[0]), properties: attrs(all('pageSetUpPr')[0]), breaks: all('brk').map(attrs),
          firstWidth: Number(all('col')[0].getAttribute('width')), values: Object.values(sheet).map(cell => cell?.v).filter(v => v !== undefined) };
      }, Array.from(await fs.readFile(file)));
      assert.equal(result.setup.paperSize, '9');
      assert.ok(result.firstWidth > 4.5, '窄模板应自动扩宽');
      for (let i = 1; i <= count; i++) assert.equal(result.values.filter(v => v === `MAT-${i}`).length, 1);
      assert.ok(result.values.some(v => String(v).includes((count * 200).toFixed(2))), '合计金额保持正确');
      assert.ok(result.values.some(v => String(v).includes('9、本合同一式二份')), '九条合同条款保持完整');
      assert.ok(result.values.some(v => String(v).includes('甲方（盖章）')), '签章区保留');
      if (count <= 11) {
        assert.equal(result.properties.fitToPage, '1'); assert.equal(result.setup.fitToWidth, '1'); assert.equal(result.setup.fitToHeight, '1');
        assert.deepEqual(result.breaks, [], '签章区不得被强制拆页');
      } else {
        assert.equal(result.properties.fitToPage, '0'); assert.ok(Number(result.setup.scale) >= 60, '大订单保持可读分页');
      }
      if (live) {
        assert.match(await page.locator('#contractStatus').innerText(), /（1页/);
        pending = page.waitForEvent('download'); await page.locator('#downloadPdf').click();
        const pdf = path.join(out, 'contract-11.pdf'); await (await pending).saveAs(pdf);
        const geometry = await page.evaluate(async bytes => {
          const pdfjs = await import('./vendor/pdf.min.mjs'); const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes) }).promise;
          const page = await doc.getPage(1), items = (await page.getTextContent()).items.filter(item => item.str.trim());
          const result = { pages: doc.numPages, width: page.view[2], span: Math.max(...items.map(i => i.transform[4] + i.width)) - Math.min(...items.map(i => i.transform[4])), text: items.map(i => i.str).join('') };
          await doc.destroy(); return result;
        }, Array.from(await fs.readFile(pdf)));
        assert.equal(geometry.pages, 1); assert.ok(geometry.span / geometry.width > 0.85, 'PDF内容应利用A4可打印宽度');
        assert.match(geometry.text, /MAT-11/); assert.match(geometry.text, /甲方（盖章）/);
        assert.match(geometry.text, /扫描件、复印件与本合同原件具有同等法律效力/);
        console.log('PDF', { pages: geometry.pages, widthUsage: geometry.span / geometry.width });
      }
      console.log('PASS', { count, setup: result.setup, out });
    }
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
