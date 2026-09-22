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
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    const errors = []; page.on('pageerror', e => errors.push(e.message));
    // Existing OCR CDN is unrelated to this tool. Payment flow must work without it.
    await page.route('https://cdn.jsdelivr.net/**', route => route.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8899/');
    assert.equal(await page.title(), '供应链数据清洗');
    assert.match(await page.locator('#toolTitle').innerText(), /报表合并/);
    await page.locator('[data-tool="payment"]').click();
    assert.equal(await page.locator('.result-layout').isVisible(), false);
    async function upload(rows, options = {}) {
      await page.waitForFunction(() => !document.querySelector('#paymentExcel').disabled);
      const array = await page.evaluate(({ rows, headers, options }) => {
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows.map(r => headers.map(h => r[h] ?? ''))]);
        if (options.formulaError) ws.L2 = { t: 'e', v: 15, f: '1/0' };
        const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '付款明细');
        return Array.from(new Uint8Array(XLSX.write(wb, { type: 'array', bookType: options.xls ? 'biff8' : 'xlsx' })));
      }, { rows, headers: ['填写时间', ...C.required.filter(h => !options.noBank || h !== '开户银行及账号')], options });
      await page.locator('#paymentExcel').setInputFiles({ name: options.xls ? '测试.xls' : '测试.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from(array) });
      await page.waitForFunction(() => !document.querySelector('[data-reference="detail"]').textContent.includes('正在读取') && document.querySelector('#paymentGroups').textContent.length > 0);
      for (const input of await page.locator('[data-field="计划付款日期"]').all()) await input.fill('2026-09-25');
    }
    async function generate() {
      await page.waitForFunction(() => [...document.querySelectorAll('[data-reference]')].every(e => !e.textContent.includes('正在读取')));
      await page.locator('#paymentGenerate').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
    }
    async function save(kind, name, index = 0) {
      assert.equal(await page.locator(`[data-payment-download="${kind}"]:visible`).nth(index).isDisabled(), false);
      const wait = page.waitForEvent('download');
      await page.locator(`[data-payment-download="${kind}"]:visible`).nth(index).click();
      const download = await wait;
      assert.match(download.suggestedFilename(), new RegExp('^2026-09-25.+[0-9]+\\.[0-9]{2}\\.' + kind + '$'));
      if (name === 'synthetic-single') assert.equal(download.suggestedFilename(), `2026-09-25甲公司2205.00.${kind}`);
      const p = path.join(out, `${name}.${kind}`); await download.saveAs(p); return p;
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
    assert.ok(docStats.text.includes('单位:测试供应商有限公司'));
    assert.ok(docStats.text.includes('账号:0012 3456 7890 1234'));
    assert.ok(!docStats.text.includes('测试银行支行'));
    assert.ok(!docStats.text.includes('123456789012'));
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
    assert.equal(await page.locator('[data-confirm]').count(), 0);
    assert.doesNotMatch(await page.locator('#paymentGroups').innerText(), /行付款总金额/);
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await upload([base({本次申请付款金额:0,备注:'ZERO-OMIT'}),base({本次申请付款金额:'0.01',备注:'KEEP-ONE',开户银行及账号:'未知银行原文'}),base({本次申请付款金额:'2.20',备注:'KEEP-TWO'})]); await generate();
    assert.doesNotMatch(await page.locator('#paymentGroups').innerText(), /银行信息无法明确识别|ZERO-OMIT/);
    assert.match(await page.locator('#paymentGroups').innerText(), /2 笔/);
    const filtered = await save('docx','zero-filtered'); await save('pdf','zero-filtered');
    const filteredText = await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')), [...await fs.readFile(filtered)]);
    assert.ok(!filteredText.includes('ZERO-OMIT'));assert.ok(filteredText.includes('KEEP-ONE'));assert.ok(filteredText.includes('KEEP-TWO'));assert.ok(filteredText.includes('未知银行原文'));
    await upload([base({本次申请付款金额:0})]); await generate();
    assert.match(await page.locator('#paymentGroups').innerText(), /均为 0/);
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(), true);
    await upload([base()]); await generate();
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
    // Directory enrichment, timestamps, explicit account conflict handling.
    assert.match(await page.locator('label.payment-upload').nth(1).innerText(), /付款申请模板/);
    assert.doesNotMatch(await page.locator('[data-reference="template"]').innerText(), /未引用|正在读取/);
    const templateTime = await page.locator('[data-reference="template"]').innerText();
    async function uploadDirectory(rows, headers = ['供应商全称', '收款账号', '收款银行名称']) {
      const bytes = await page.evaluate(({rows,headers}) => {
        const book = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([headers,...rows.map(r => headers[2] === '收款银行名称' && r.length === 2 ? [...r, '测试银行支行'] : r)]), '名录');
        return Array.from(new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})));
      }, {rows,headers});
      await page.locator('#paymentDirectory').setInputFiles({name:'供应商名录.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from(bytes)});
      await page.waitForFunction(()=>!document.querySelector('[data-reference="directory"]').textContent.includes('正在读取'));
    }
    await upload([base({'开户银行及账号':''})]);
    await uploadDirectory([['测试供应商有限公司','0099  0011']]);
    assert.doesNotMatch(await page.locator('[data-reference="detail"]').innerText(), /未引用|正在读取/);
    assert.doesNotMatch(await page.locator('[data-reference="directory"]').innerText(), /未引用|正在读取/);
    const directoryTime = await page.locator('[data-reference="directory"]').innerText();
    await generate(); assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    const directoryDoc = await save('docx', 'directory'); await save('pdf','directory');
    const directoryXml = await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')), [...await fs.readFile(directoryDoc)]);
    assert.ok(directoryXml.includes('卡号:0099  0011'));
    await page.screenshot({path:path.join(out,'directory-desktop.png'),fullPage:false});
    await page.setViewportSize({width:390,height:844});
    await page.locator('#paymentDirectorySettings').scrollIntoViewIfNeeded();
    assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth + 1),true);
    await page.screenshot({path:path.join(out,'directory-mobile.png'),fullPage:false});
    await page.setViewportSize({width:1440,height:1000});
    await page.locator('[data-tool="merge"]').click(); await page.locator('[data-tool="payment"]').click();
    assert.equal(await page.locator('[data-reference="directory"]').innerText(), directoryTime);
    assert.equal(await page.locator('[data-reference="template"]').innerText(), templateTime);
    await upload([base()]); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /银行信息差异/);
    await page.locator('[data-bank-confirm]').check(); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await page.locator('#paymentBankPolicy').selectOption('fill'); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    assert.equal(await page.locator('[data-bank-confirm]').count(),0);
    await page.locator('#paymentBankPolicy').selectOption('directory');
    await uploadDirectory([['测试供应商有限公司','111'],['测试供应商有限公司','222']]);
    assert.match(await page.locator('#paymentGroups').innerText(), /不同的收款单位或账号/);
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(),true);
    await uploadDirectory([['其他供应商','222']]);
    assert.match(await page.locator('#paymentGroups').innerText(), /未匹配，沿用/);
    await upload([base({'开户银行及账号':''})]);
    assert.match(await page.locator('#paymentGroups').innerText(), /未匹配且付款明细银行信息为空/);
    await uploadDirectory([['测试供应商有限公司','000123','收款公司']],['公司名称','收款账号','收款银行']);
    await page.locator('[data-directory-field="supplier"]').selectOption('公司名称');
    await generate(); assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await upload([base()],{noBank:true}); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(),/已生成/);
    await uploadDirectory([['测试供应商有限公司',27300476042050000]]);
    assert.match(await page.locator('#paymentGroups').innerText(),/精度/);
    await page.locator('#paymentDirectory').setInputFiles({name:'错误.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from([0,1,2,3])});
    await page.waitForFunction(()=>!document.querySelector('[data-reference="directory"]').textContent.includes('正在读取'));
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(),true);
    await page.locator('#paymentClearDirectory').click();
    assert.match(await page.locator('[data-reference="directory"]').innerText(), /未引用/);
    assert.equal(await page.locator('[data-payment-download="docx"]').isDisabled(),true);
    const stress = [base({ 付款主体: '单笔主体' }), ...Array.from({ length: 30 }, (_, i) => base({ 付款主体: '多笔主体', 备注: `记录${i+1}`, 本次申请付款金额: '0.01' })), ...Array.from({ length: 45 }, (_, i) => base({ 付款主体: '短明细主体', 开户银行及账号: '测试银行 0012 3456', 本次申请付款金额: '1.10', 备注: `记录${i+1}` }))];
    await upload(stress); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成 3 页/);
    assert.match(await page.locator('#paymentPages').innerText(), /继续压缩/);
    await save('docx', 'stress'); const pdfPath = await save('pdf', 'stress');
    assert.equal(await page.locator('[data-payment-download="pdf"]:visible').count(), 3);
    for (let i = 1; i < 3; i++) {
      await save('docx', `stress-${i}`, i);
      const file = await save('pdf', `stress-${i}`, i);
      assert.equal(await page.evaluate(async bytes => (await PDFLib.PDFDocument.load(new Uint8Array(bytes))).getPageCount(), [...await fs.readFile(file)]), 1);
    }
    const pageCount = await page.evaluate(async bytes => (await PDFLib.PDFDocument.load(new Uint8Array(bytes))).getPageCount(), [...await fs.readFile(pdfPath)]);
    assert.equal(pageCount, 1);
    await page.screenshot({ path: path.join(out, 'desktop.png'), fullPage: false });
    await page.setViewportSize({ width: 390, height: 844 });
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), true);
    await page.locator('#paymentGroups').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(out, 'mobile.png'), fullPage: false });
    if (process.env.QA_DIRECTORY_FILE) {
      await page.setViewportSize({width:1440,height:1000});
      await page.locator('#paymentDirectory').setInputFiles(process.env.QA_DIRECTORY_FILE);
      await page.waitForFunction(()=>!document.querySelector('[data-reference="directory"]').textContent.includes('正在读取'));
      await upload([base({付款主体:'浙江迈德斯特医疗器械科技有限公司',供应商全称:'浙江耐心医疗器械有限公司',开户银行及账号:''})]);
      await generate();
      assert.match(await page.locator('#paymentStatus').innerText(), /已生成 1 页/);
      const doc = await save('docx','actual-directory'); await save('pdf','actual-directory');
      const xml = await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')), [...await fs.readFile(doc)]);
      assert.ok(xml.includes('迈德斯特付款申请单'));
      assert.ok(xml.includes('银行名称:农业银行永康市支行'));
      assert.ok(xml.includes('卡号:19627201040060557'));
      assert.ok(!xml.includes('103338262728'));
      await upload([base({付款主体:'迈德斯特',供应商全称:'浙江耐心医疗器械有限公司',开户银行及账号:''})]);
      await generate(); assert.match(await page.locator('#paymentStatus').innerText(), /已生成 1 页/);
      await page.locator('#paymentPages').scrollIntoViewIfNeeded();
      await page.screenshot({path:path.join(out,'actual-directory.png'),fullPage:false});
    }
    await upload([base({'开户银行及账号':''})]);
    await uploadDirectory([['测试供应商有限公司','0011 2233']]);
    await page.locator('#paymentTemplate').setInputFiles({ name: '缓存模板.docx', mimeType: 'application/octet-stream', buffer: Buffer.from(custom) });
    await page.waitForFunction(() => !document.querySelector('[data-reference="template"]').textContent.includes('正在读取'));
    // Wait for all queued IndexedDB writes, then restore in a fresh document.
    const savedNames = await page.evaluate(async () => {
      const records = await Promise.all(['detail','template','directory'].map(k => PaymentCache.get(k)));
      return records.map(r => ({name:r.name,time:r.time}));
    });
    await page.reload(); await page.locator('[data-tool="payment"]').click();
    await page.waitForFunction(() => !document.querySelector('#paymentExcel').disabled);
    for (const [i,kind] of ['detail','template','directory'].entries()) assert.equal(await page.locator(`[data-reference="${kind}"]`).innerText(), '引用时间：'+savedNames[i].time);
    assert.match(await page.locator('label.payment-upload').nth(1).innerText(), /缓存模板.docx/);
    await page.locator('[data-field="计划付款日期"]').fill('2026-09-25'); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    const restored = await save('docx','cache-restored');
    const restoredXml = await page.evaluate(async bytes => (await (await JSZip.loadAsync(new Uint8Array(bytes))).file('word/document.xml').async('string')), [...await fs.readFile(restored)]);
    assert.ok(restoredXml.includes('部门主管审核')); assert.ok(restoredXml.includes('卡号:0011 2233'));
    const fresh = await page.context().newPage();
    await fresh.route('https://cdn.jsdelivr.net/**', route => route.fulfill({status:200,contentType:'text/javascript',body:''}));
    await fresh.goto(process.env.QA_URL || 'http://127.0.0.1:8899/'); await fresh.locator('[data-tool="payment"]').click();
    await fresh.waitForFunction(() => !document.querySelector('#paymentExcel').disabled);
    assert.match(await fresh.locator('#paymentGroups').innerText(), /甲公司/); await fresh.close();
    await page.locator('#paymentClearDirectory').click();
    assert.equal(await page.evaluate(async()=>!!(await PaymentCache.get('directory'))),false);
    await page.locator('#paymentResetTemplate').click();
    await page.waitForFunction(() => !document.querySelector('[data-reference="template"]').textContent.includes('正在读取'));
    assert.equal(await page.evaluate(async()=>!!(await PaymentCache.get('template')).blob),false);
    await page.locator('#paymentClearCache').click();
    await page.waitForFunction(() => !document.querySelector('#paymentExcel').disabled);
    assert.equal(await page.evaluate(async()=>!!(await PaymentCache.get('detail'))),false);
    await page.reload(); await page.locator('[data-tool="payment"]').click();
    await page.waitForFunction(() => !document.querySelector('#paymentExcel').disabled);
    assert.match(await page.locator('[data-reference="detail"]').innerText(), /未引用/);
    assert.match(await page.locator('[data-reference="directory"]').innerText(), /未引用/);
    // Failure to persist must be visible and must not break the current upload.
    await page.evaluate(() => { PaymentCache.put = async()=>{throw new Error('quota');}; });
    await upload([base()]); await generate();
    assert.match(await page.locator('#paymentCacheMessage').innerText(), /未能保存/);
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await page.screenshot({path:path.join(out,'cache-desktop.png'),fullPage:false});
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ status: 'PASS', scenarios: ['identity', 'single/export/uppercase', 'stale-download', 'switch-state', 'invalid-amount', 'missing-subject', 'conflict', 'total-confirmation', 'custom-template', 'invalid-template', 'xls', 'formula-error', 'directory-match/export', 'three-reference-times', 'directory-conflict', 'directory-duplicate', 'directory-unmatched', 'directory-mapping', 'directory-fill-only', 'optional-bank-column', 'directory-precision', 'directory-failed-upload', 'directory-removal', 'three-subjects', '30-long-rows', '45-short-rows', 'mobile', 'console'], output: out }, null, 2));
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });
