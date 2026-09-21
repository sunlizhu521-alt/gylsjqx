// Start a static server first. PLAYWRIGHT_MODULE / CHROME_PATH can select an existing runtime.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const C = require('../payment-core.js');
const base = (extra = {}) => ({
  申请日期: '2026-09-25', 付款主体: '甲公司', 申请部门: '履约部', 经办人: '测试人', 付款总金额: '',
  供应商全称: '测试供应商有限公司', 开户银行及账号: '收款单位\n测试供应商有限公司\n收款银行名称\n测试银行支行\n收款账号\n0012 3456 7890 1234\n收款银行行号\n123456789012',
  合同编号: '/', 合同总金额: '/', 已付金额: '/', 本次申请付款金额: '2205', 票据状态: '已开票', 付款事由: '货款', 备注: '', ...extra,
});
(async () => {
  const out = process.env.QA_OUTPUT || await fs.mkdtemp(path.join(os.tmpdir(), 'payment-browser-'));
  await fs.mkdir(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    const errors = []; page.on('pageerror', e => errors.push(e.message));
    // Existing OCR CDN is unrelated to this tool. Payment flow must work without it.
    await page.route('https://cdn.jsdelivr.net/**', route => route.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8899/');
    assert.equal(await page.title(), '供应链数据清洗');
    assert.match(await page.locator('#toolTitle').innerText(), /报表合并/);
    await page.locator('[data-tool="payment"]').click();
    assert.equal(await page.locator('.result-layout').isVisible(), false);
    async function upload(rows, options = {}) {
      const array = await page.evaluate(({ rows, headers, options }) => {
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows.map(r => headers.map(h => r[h] ?? ''))]);
        if (options.formulaError) ws.L2 = { t: 'e', v: 15, f: '1/0' };
        const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '付款明细');
        return Array.from(new Uint8Array(XLSX.write(wb, { type: 'array', bookType: options.xls ? 'biff8' : 'xlsx' })));
      }, { rows, headers: ['填写时间', ...C.required], options });
      await page.locator('#paymentExcel').setInputFiles({ name: options.xls ? '测试.xls' : '测试.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from(array) });
      await page.waitForFunction(() => document.querySelector('#paymentGroups').textContent.length > 0);
    }
    async function generate() {
      await page.locator('#paymentGenerate').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
    }
    async function save(kind, name) {
      assert.equal(await page.locator(`[data-payment-download="${kind}"]`).isDisabled(), false);
      const wait = page.waitForEvent('download');
      await page.locator(`[data-payment-download="${kind}"]`).click();
      const p = path.join(out, `${name}.${kind}`); await (await wait).saveAs(p); return p;
    }
    await upload([base()]); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成 1 页/);
    const sample = await save('docx', 'synthetic-single'); await save('pdf', 'synthetic-single');
    const docStats = await page.evaluate(async bytes => {
      const z = await JSZip.loadAsync(new Uint8Array(bytes));
      const d = new DOMParser().parseFromString(await z.file('word/document.xml').async('string'), 'application/xml');
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      return { text: [...d.getElementsByTagNameNS(ns, 't')].map(e => e.textContent).join(''), sections: d.getElementsByTagNameNS(ns, 'sectPr').length };
    }, [...await fs.readFile(sample)]);
    assert.equal(docStats.sections, 1);
    for (const text of ['贰仟贰佰零伍元整', '☑ 采购货款', '☑ 对公转账', '三、审批意见', '四、财务办结记录', '整体付款凭证：详见银行付款单', '0012 3456 7890 1234']) assert.ok(docStats.text.includes(text), text);
    assert.equal((docStats.text.match(/2,205\.00/g) || []).length, 3);
    await page.locator('[data-field="计划付款日期"]').fill('2026-09-26');
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(), true);
    await page.locator('[data-tool="merge"]').click();
    assert.equal(await page.locator('.result-layout').isVisible(), true);
    await page.locator('[data-tool="payment"]').click();
    assert.equal(await page.locator('[data-field="计划付款日期"]').inputValue(), '2026-09-26');
    await upload([base({ 本次申请付款金额: '' })]); await generate();
    assert.match(await page.locator('#paymentGroups').innerText(), /第 2 行/);
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(), true);
    await upload([base({ 付款主体: '' })]);
    assert.match(await page.locator('#paymentGroups').innerText(), /缺少付款主体/);
    await upload([base(), base({ 经办人: '另一人' })]); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /请填写经办人/);
    await page.locator('[data-field="经办人"]').fill('核实人'); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await upload([base({ 付款总金额: '0' })]); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /核实付款总金额差异/);
    await page.locator('[data-confirm]').check(); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    // Template replacement changes preserved approval content, never payment constants.
    const custom = await page.evaluate(async () => {
      const z = await JSZip.loadAsync(await (await fetch('assets/payment-template.docx')).arrayBuffer());
      z.file('word/document.xml', (await z.file('word/document.xml').async('string')).replace('部门负责人审核', '部门主管审核'));
      return Array.from(await z.generateAsync({ type: 'uint8array' }));
    });
    await page.locator('#paymentTemplate').setInputFiles({ name: '自定义模板.docx', mimeType: 'application/octet-stream', buffer: Buffer.from(custom) });
    await page.waitForFunction(() => document.querySelector('#paymentStatus').textContent.includes('模板结构检查通过'));
    await generate();
    const customPath = await save('docx', 'custom');
    assert.equal(await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')).includes('部门主管审核'), [...await fs.readFile(customPath)]), true);
    await page.locator('#paymentTemplate').setInputFiles({ name: '损坏模板.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('broken') });
    await page.waitForFunction(() => document.querySelector('#paymentStatus').classList.contains('payment-error'));
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(), true);
    await page.locator('#paymentResetTemplate').click();
    await page.waitForFunction(() => document.querySelector('#paymentStatus').textContent.includes('模板结构检查通过'));
    await upload([base()], { xls: true }); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await upload([base()], { formulaError: true });
    assert.match(await page.locator('#paymentGroups').innerText(), /公式错误/);
    const stress = [base({ 付款主体: '单笔主体' }), ...Array.from({ length: 30 }, (_, i) => base({ 付款主体: '多笔主体', 备注: `记录${i+1}`, 本次申请付款金额: '0.01' })), ...Array.from({ length: 45 }, (_, i) => base({ 付款主体: '短明细主体', 开户银行及账号: '测试银行 0012 3456', 本次申请付款金额: '1.10', 备注: `记录${i+1}` }))];
    await upload(stress); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成 3 页/);
    assert.match(await page.locator('#paymentPages').innerText(), /继续压缩/);
    await save('docx', 'stress'); const pdfPath = await save('pdf', 'stress');
    const pageCount = await page.evaluate(async bytes => (await PDFLib.PDFDocument.load(new Uint8Array(bytes))).getPageCount(), [...await fs.readFile(pdfPath)]);
    assert.equal(pageCount, 3);
    await page.screenshot({ path: path.join(out, 'desktop.png'), fullPage: false });
    await page.setViewportSize({ width: 390, height: 844 });
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), true);
    await page.locator('#paymentGroups').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(out, 'mobile.png'), fullPage: false });
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ status: 'PASS', scenarios: ['identity', 'single/export/uppercase', 'stale-download', 'switch-state', 'invalid-amount', 'missing-subject', 'conflict', 'total-confirmation', 'custom-template', 'invalid-template', 'xls', 'formula-error', 'three-subjects', '30-long-rows', '45-short-rows', 'mobile', 'console'], output: out }, null, 2));
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });
