// Start a static server first. All fixtures are synthetic; artifacts default to a temporary directory.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');

function rows(count, subject, prefix, long = false) {
  return Array.from({ length: count }, (_, i) => ({
    _row: i + 2, 申请日期: '2026-11-06', 付款主体: subject, 申请部门: '采购履约部', 经办人: '排版测试人',
    供应商全称: `${prefix}${String(i + 1).padStart(3, '0')}测试供应商有限公司`,
    开户银行及账号: long
      ? '收款单位\n排版测试供应商有限公司\n收款账号\n0012 3456 7890 1234'
      : '单位:测试供应商\n账号:0012345678901234',
    合同编号: long ? `HT-${prefix}-${String(i + 1).padStart(4, '0')}` : '/',
    合同总金额: '/', 已付金额: '/', 本次申请付款金额: '1.10', 票据状态: '已开票',
    付款事由: long ? '采购配件及运输费用，按已确认对账单支付本批货款。' : '货款',
    付款分类: long ? `正常付款 AP${String(i + 1).padStart(8, '0')}（￥1.10）` : '',
    对接人: `${prefix}-记录-${String(i + 1).padStart(3, '0')}`, 备注: '',
  }));
}

(async () => {
  const out = process.env.QA_OUTPUT || await fs.mkdtemp(path.join(os.tmpdir(), 'payment-layout-'));
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ viewport: { width: 1100, height: 850 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8899/');
    await page.waitForFunction(() => window.PaymentDocument && window.PaymentCore && window.PDFLib && window.JSZip);
    await page.evaluate(async () => {
      const response = await fetch('assets/payment-template.docx');
      if (!response.ok) throw new Error('Layout test could not load the built-in template');
      const template = await PaymentDocument.readTemplate(await response.arrayBuffer());
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const children = (element, name) => [...element.children].filter(e => e.namespaceURI === ns && e.localName === name);
      const text = element => [...element.getElementsByTagNameNS(ns, 't')].map(e => e.textContent).join('').replace(/\s/g, '');
      const sizes = element => [...element.getElementsByTagNameNS(ns, 'sz')]
        .filter(e => e.parentNode.parentNode.localName === 'r').map(e => Number(e.getAttributeNS(ns, 'val')) / 2);
      const original = new DOMParser().parseFromString(template.xml, 'application/xml');
      const originalTables = [...original.getElementsByTagNameNS(ns, 'tbl')];
      function inspectPage(blocks) {
        const tables = blocks.filter(e => e.localName === 'tbl');
        const paragraphs = blocks.filter(e => e.localName === 'p');
        if (tables.length < 2) throw new Error('A generated page is missing its metadata or detail table');
        const detailRows = children(tables[1], 'tr');
        const dataRows = detailRows.slice(1).filter(row => children(row, 'tc').length === 10);
        const totalRows = detailRows.slice(1).filter(row => children(row, 'tc').length !== 10);
        return {
          title: text(paragraphs[0]), date: text(paragraphs[1]), metadata: text(tables[0]),
          headers: children(detailRows[0], 'tc').map(text),
          rows: dataRows.map(row => children(row, 'tc').map(text)),
          detailSizes: sizes(tables[1]), totalRows: totalRows.map(text),
          footerTables: tables.slice(2).map(text), footerSizes: tables.slice(2).flatMap(sizes),
          footerTitles: paragraphs.filter(p => /审批意见|财务办结记录/.test(text(p))).map(p => ({ text: text(p), sizes: sizes(p) })),
        };
      }
      async function inspectDocx(blob) {
        const zip = await JSZip.loadAsync(await blob.arrayBuffer());
        const document = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
        const body = document.getElementsByTagNameNS(ns, 'body')[0];
        const pages = [], blocks = [];
        for (const block of body.children) {
          if (block.localName === 'sectPr' || (block.localName === 'p' && block.getElementsByTagNameNS(ns, 'sectPr').length)) {
            pages.push(inspectPage(blocks.splice(0)));
          } else blocks.push(block);
        }
        if (blocks.length) throw new Error('The final DOCX page has no section properties');
        const sections = [...document.getElementsByTagNameNS(ns, 'sectPr')];
        return {
          pages, sections: sections.length,
          landscape: sections.every(section => {
            const size = section.getElementsByTagNameNS(ns, 'pgSz')[0];
            return size && size.getAttributeNS(ns, 'orient') === 'landscape' && size.getAttributeNS(ns, 'w') === '16838' && size.getAttributeNS(ns, 'h') === '11906';
          }),
        };
      }
      async function base64(blob) {
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result.split(',')[1]);
          reader.onerror = () => reject(reader.error);
          reader.readAsDataURL(blob);
        });
      }
      window.paymentLayoutQa = {
        template, inspectDocx, base64, cases: {},
        originalFooter: originalTables.slice(2).map(text),
        normalize: value => String(value).replace(/\s/g, ''),
      };
    });

    const fixtures = [
      { name: 'single', rows: rows(1, '单条主体', 'S'), singlePage: true, targetSize: 10.5 },
      { name: 'medium', rows: rows(8, '中等明细主体', 'M'), singlePage: true },
      { name: 'long-30', rows: rows(30, '长明细主体', 'L', true), singlePage: false },
      { name: 'short-45', rows: rows(45, '短明细主体', 'T'), singlePage: false },
    ];
    const reports = [], validatedPages = new Map();
    for (const fixture of fixtures) {
      const result = await page.evaluate(async fixture => {
        const qa = window.paymentLayoutQa, C = PaymentCore, D = PaymentDocument;
        const analysis = C.analyze([...C.required, '付款分类', '对接人'], fixture.rows);
        if (analysis.errors.length || analysis.groups.length !== 1) throw new Error(JSON.stringify(analysis.errors));
        const group = analysis.groups[0];
        group.info['计划付款日期'] = '2026-11-06';
        const layouts = await D.layout(qa.template, group);
        if (!Array.isArray(layouts) || !layouts.length) throw new Error('PaymentDocument.layout must return a nonempty array');
        qa.cases[fixture.name] = { group, layouts };
        const docx = await D.docx(qa.template, layouts), pdf = await D.pdf(layouts);
        const doc = await qa.inspectDocx(docx), parsedPdf = await PDFLib.PDFDocument.load(await pdf.arrayBuffer());
        document.getElementById('layoutQaOutput')?.remove();
        const output = document.createElement('div'); output.id = 'layoutQaOutput';
        output.style.cssText = 'position:relative;background:white;padding:16px';
        for (const layout of layouts) {
          const canvas = D.canvas(layout);
          canvas.style.cssText = 'display:block;width:842px;height:auto;margin:12px 0';
          output.append(canvas);
        }
        document.body.append(output);
        return {
          ...doc, pdfPages: parsedPdf.getPageCount(), pdfSizes: parsedPdf.getPages().map(p => p.getSize()),
          expectedRows: group.rows.map((row, i) => [String(i + 1), ...C.fields.map(field => {
            const value = field === '本次申请付款金额' ? C.money(row._cents)
              : field === '开户银行及账号' ? C.bankSummary(row._resolvedBank ?? row[field], row['供应商全称']).value : C.text(row[field]);
            return qa.normalize(value);
          })]),
          expectedHeader: ['序号', ...C.fields], subject: group.subject, total: C.money(group.total), upper: C.upper(group.total),
          originalFooter: qa.originalFooter, docx: await qa.base64(docx), pdf: await qa.base64(pdf),
        };
      }, fixture);
      await fs.writeFile(path.join(out, `${fixture.name}.docx`), Buffer.from(result.docx, 'base64'));
      await fs.writeFile(path.join(out, `${fixture.name}.pdf`), Buffer.from(result.pdf, 'base64'));
      for (let i = 0; i < result.pages.length; i++) {
        await page.locator('#layoutQaOutput canvas').nth(i).screenshot({ path: path.join(out, `${fixture.name}-page-${i + 1}.png`) });
      }
      assert.equal(result.sections, result.pages.length, fixture.name);
      assert.equal(result.pdfPages, result.pages.length, `${fixture.name}: planned DOCX pages and PDF pages agree`);
      assert.equal(result.landscape, true, `${fixture.name}: every DOCX section uses A4 landscape`);
      assert.ok(result.pdfSizes.every(size => Math.abs(size.width - 841.9) < .1 && Math.abs(size.height - 595.3) < .1));
      assert.ok(fixture.singlePage ? result.pages.length === 1 : result.pages.length > 1, `${fixture.name}: unexpected pagination (${result.pages.length})`);
      assert.deepEqual(result.pages.flatMap(p => p.rows), result.expectedRows, `${fixture.name}: all columns and continuous row numbers survive without omissions or duplication`);
      for (const [i, current] of result.pages.entries()) {
        const last = i === result.pages.length - 1;
        assert.equal(current.title, `${result.subject}付款申请单`);
        assert.ok(current.date.includes('申请日期：2026年11月06日'));
        if (result.pages.length > 1) assert.match(current.date, new RegExp(`第${i + 1}[/／]${result.pages.length}页`));
        for (const value of ['采购履约部', '排版测试人', result.total, result.upper]) assert.ok(current.metadata.includes(value), `${fixture.name} page ${i + 1}: missing metadata ${value}`);
        assert.deepEqual(current.headers, result.expectedHeader, `${fixture.name}: detail headers repeat on every page`);
        assert.ok(current.rows.length > 0, `${fixture.name}: no empty trailing approval page`);
        assert.ok(current.detailSizes.length > 0 && current.detailSizes.every(size => size >= 9 && size <= 10.5), `${fixture.name}: detail text must stay between 9 and 10.5 pt`);
        if (fixture.targetSize) assert.ok(current.detailSizes.every(size => size === fixture.targetSize), `${fixture.name}: use the clear target font while space is available`);
        assert.equal(current.totalRows.length, last ? 1 : 0, `${fixture.name}: detail total only belongs on the final page`);
        assert.deepEqual(current.footerTables, last ? result.originalFooter : [], `${fixture.name}: original approval and settlement text only belong on the final page`);
        assert.equal(current.footerTitles.length, last ? 2 : 0, `${fixture.name}: footer headings only belong on the final page`);
        if (last) {
          assert.ok(current.totalRows[0].includes(result.total));
          assert.ok(current.footerSizes.length > 0 && current.footerSizes.every(size => size >= 6.5 && size <= 7.5 && size < Math.min(...current.detailSizes)), `${fixture.name}: footer uses smaller readable text`);
          assert.ok(current.footerTitles.every(title => title.sizes.length && title.sizes.every(size => size < Math.min(...current.detailSizes))));
        }
      }
      validatedPages.set(fixture.name, result.pages);
      reports.push({ name: fixture.name, rows: result.expectedRows.length, pages: result.pages.length, detailSizes: [...new Set(result.pages.flatMap(p => p.detailSizes))] });
    }

    const combined = await page.evaluate(async () => {
      const qa = window.paymentLayoutQa;
      const layouts = ['single', 'long-30', 'short-45'].flatMap(name => qa.cases[name].layouts);
      const docx = await PaymentDocument.docx(qa.template, layouts), pdf = await PaymentDocument.pdf(layouts);
      const doc = await qa.inspectDocx(docx);
      return { ...doc, pdfPages: (await PDFLib.PDFDocument.load(await pdf.arrayBuffer())).getPageCount(), docx: await qa.base64(docx), pdf: await qa.base64(pdf) };
    });
    await fs.writeFile(path.join(out, 'combined-76-rows.docx'), Buffer.from(combined.docx, 'base64'));
    await fs.writeFile(path.join(out, 'combined-76-rows.pdf'), Buffer.from(combined.pdf, 'base64'));
    assert.deepEqual(combined.pages, ['single', 'long-30', 'short-45'].flatMap(name => validatedPages.get(name)), 'combined Word preserves all 76 rows and every validated page in subject order');
    assert.equal(combined.pages.flatMap(page => page.rows).length, 76);
    assert.equal(combined.sections, reports.filter(r => r.name !== 'medium').reduce((sum, r) => sum + r.pages, 0));
    assert.equal(combined.pdfPages, combined.sections);
    assert.equal(combined.pages.filter(page => page.footerTables.length).length, 3);

    const oversized = rows(1, '超长单条主体', 'O');
    oversized[0].付款事由 = '这一条明细必须完整显示\n'.repeat(600);
    const failure = await page.evaluate(async input => {
      const analysis = PaymentCore.analyze([...PaymentCore.required, '付款分类', '对接人'], input);
      if (analysis.errors.length) throw new Error(JSON.stringify(analysis.errors));
      analysis.groups[0].info['计划付款日期'] = '2026-11-06';
      try { await PaymentDocument.layout(window.paymentLayoutQa.template, analysis.groups[0]); return ''; }
      catch (error) { return error.message; }
    }, oversized);
    assert.match(failure, /单条|单行|第\s*2\s*行/, 'an oversized individual row must be identified explicitly');
    assert.match(failure, /无法|过长|过多|超|容纳|拆分/, 'an oversized row must fail instead of shrinking below the readable minimum');
    assert.deepEqual(errors, []);
    // DOCX sections encode intended pages; render the saved Word files to verify physical pagination.
    const report = { status: 'PASS', cases: reports, combinedRows: 76, combinedDocxSections: combined.sections, combinedPdfPages: combined.pdfPages, oversizedRowError: failure, output: out };
    await fs.writeFile(path.join(out, 'layout-report.json'), JSON.stringify(report, null, 2));
    console.log(JSON.stringify(report, null, 2));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
