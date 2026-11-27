// Run against a static server or QA_URL. Uses synthetic files only.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../payment-core.js');
const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const budgetHeaders = ['分类', '创跟单员', '供应商', '付款主体', '对公付款0A流程号', '摘要', 'OA-申请付款金额', '付款时间', '备注'];
const detail = (extra = {}) => ({
  申请日期: '2026-11-27', 付款主体: '合成甲公司', 申请部门: '履约部', 经办人: '测试人',
  供应商全称: '合成供应商甲', 开户银行及账号: '单位:合成供应商甲\n账号:000123456789',
  合同编号: 'MDST-DGFK20261127001', 合同总金额: '/', 已付金额: '/', 本次申请付款金额: '100.00',
  票据状态: '已开票', 付款事由: '合成货款', 备注: '', ...extra,
});
const first = detail(), second = detail({供应商全称:'合成供应商乙',开户银行及账号:'单位:合成供应商乙\n账号:000123456788',合同编号:'MDST-DGFK20261127002',本次申请付款金额:'200.00'});
const budgetRow = (row, amount = row.本次申请付款金额) => ['正常付款', '测试人', row.供应商全称, row.付款主体, row.合同编号, '合成预算', amount, '2027年01月25日', 'BUDGET-ONLY-NOTE'];
function bytes(sheets) {
  const book = X.utils.book_new();
  for (const [name, rows] of sheets) X.utils.book_append_sheet(book, X.utils.aoa_to_sheet(rows), name);
  return Buffer.from(X.write(book, {type:'buffer',bookType:'xlsx'}));
}
const detailBytes = rows => bytes([['付款明细', [C.required, ...rows.map(row => C.required.map(h => row[h] ?? ''))]]]);
const budgetBytes = rows => bytes([['预算明细', [budgetHeaders, ...rows]]]);

(async () => {
  const base = process.env.QA_URL || 'http://127.0.0.1:8899/';
  const out = process.env.OUT_DIR || process.env.QA_OUTPUT || '/private/tmp/payment-budget-browser';
  await fs.mkdir(out, {recursive:true});
  const browser = await chromium.launch({headless:true, ...(process.env.CHROME_PATH ? {executablePath:process.env.CHROME_PATH} : {})});
  try {
    const page = await browser.newPage({viewport:{width:1440,height:1000},timezoneId:'Asia/Shanghai'});
    await page.clock.setFixedTime(new Date('2026-11-27T02:00:00Z'));
    const errors = [], downloads = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('download', download => downloads.push(download.suggestedFilename()));
    await page.goto(base);
    await page.locator('button[data-tool="payment"]').click();
    await page.locator('#paymentExcel:not([disabled])').waitFor();
    async function settled(kind) {
      await page.waitForFunction(kind => {
        const reference = document.querySelector(`[data-reference="${kind}"]`);
        return reference && !reference.textContent.includes('正在读取');
      }, kind);
    }
    async function upload(role, buffer, name) {
      const id = role === 'detail' ? 'paymentExcel' : 'paymentBudget';
      await page.locator('#' + id).setInputFiles({name,mimeType:MIME,buffer});
      await settled(role);
    }
    async function generate() {
      await page.locator('#paymentGenerate').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
    }
    async function blocked() {
      assert.equal(await page.locator('#paymentPages canvas').count(), 0, 'invalidated results have no old preview');
      assert(await page.locator('[data-payment-download="docx"]').isDisabled());
      assert(await page.locator('[data-payment-download="pdf"]').isDisabled());
    }
    async function blockedGeneration() {
      await generate();
      await blocked();
      assert.match(await page.locator('#paymentStatus').innerText(), /预算/);
      const before = downloads.length;
      await page.locator('#paymentDownloadWord').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
      assert.equal(downloads.length, before, 'one-click cannot bypass budget review');
      await blocked();
    }
    async function save(kind, label) {
      const event = page.waitForEvent('download');
      await page.locator(`[data-payment-download="${kind}"]:visible`).click();
      const file = await event, target = path.join(out, label + '.' + kind);
      await file.saveAs(target);
      return fs.readFile(target);
    }
    const checks = () => page.locator('#paymentBudgetChecks tbody > tr.payment-budget-check');
    const results = () => page.locator('#paymentBudgetChecks .payment-budget-result').allTextContents();
    const previewImages = () => page.locator('#paymentPages canvas').evaluateAll(canvases => canvases.map(canvas => canvas.toDataURL()));
    async function assertBudgetTable(sourceRows) {
      const tables = page.locator('#paymentBudgetChecks table');
      assert.equal(await tables.count(), 1, 'budget checks use one semantic table');
      const snapshot = await tables.evaluate(table => ({
        headings: [...table.tHead.rows[0].cells].map(cell => ({tag:cell.tagName,text:cell.textContent.trim()})),
        bodies: table.tBodies.length,
        rows: [...table.tBodies[0].rows].map(row => ({tag:row.tagName,source:Number(row.dataset.budgetRow),cells:row.cells.length})),
      }));
      assert.equal(snapshot.headings.length, 6);
      [/核对结果/,/付款申请/,/申请金额/,/预算信息/,/核对说明/,/确认/].forEach((label, index) => {
        assert.equal(snapshot.headings[index].tag, 'TH'); assert.match(snapshot.headings[index].text, label);
      });
      assert.equal(snapshot.bodies, 1);
      assert.deepEqual(snapshot.rows, sourceRows.map(source => ({tag:'TR',source,cells:6})), 'each application remains one six-column row');
      assert.equal(await checks().count(), sourceRows.length);
    }
    async function confirmAll() {
      for (const input of await page.locator('[data-budget-confirm]').all()) await input.check();
    }

    // No optional budget preserves the existing workflow.
    await upload('detail', detailBytes([first]), '合成付款明细.xlsx');
    assert.match(await page.locator('#paymentBudgetChecks').innerText(), /未进行预算核对/);
    await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    const noBudgetImage = await previewImages(), noBudgetDoc = await save('docx', 'without-budget');
    await save('pdf', 'without-budget');

    // A later year's budget is still the selected budget; no implicit date filter.
    await upload('budget', budgetBytes([budgetRow(first)]), '2027年1月预算.xlsx');
    await blocked();
    await assertBudgetTable([2]);
    assert.deepEqual(await results(), ['预算内']);
    assert.equal(await page.locator('[data-budget-confirm]').count(), 0);
    assert.match(await page.locator('#paymentBudgetSettings').innerText(), /2027年1月预算.xlsx|预算明细/);
    await generate();
    assert.deepEqual(await previewImages(), noBudgetImage, 'budget matching must not change printable content');
    const matchedDoc = await save('docx', 'matched-budget'), matchedPdf = await save('pdf', 'matched-budget');
    const docXml = async data => (await Zip.loadAsync(data)).file('word/document.xml').async('string');
    assert.equal(await docXml(matchedDoc), await docXml(noBudgetDoc));
    assert.doesNotMatch(await docXml(matchedDoc), /预算内|BUDGET-ONLY-NOTE|确认继续/);
    assert.equal(await page.evaluate(async array => (await PDFLib.PDFDocument.load(Uint8Array.from(array))).getPageCount(), [...matchedPdf]), 1);

    // Both missing budget and unequal amount need their own explicit confirmation.
    await upload('detail', detailBytes([first, second]), '合成付款明细.xlsx');
    await upload('budget', budgetBytes([budgetRow(first, '99.00')]), '合成月份预算.xlsx');
    await assertBudgetTable([2,3]);
    assert.deepEqual(await results(), ['待核对', '待核对']);
    assert.equal(await page.locator('[data-budget-confirm]').count(), 2);
    assert.match(await checks().first().innerText(), /99\.00|99/);
    assert.match(await checks().first().innerText(), /100\.00|100/);
    assert.match(await checks().first().innerText(), /合成月份预算.xlsx/);
    assert.match(await checks().first().innerText(), /付款明细.*2/);
    await blockedGeneration();
    await page.locator('[data-budget-confirm]').first().check();
    await blockedGeneration();
    await confirmAll();
    assert.deepEqual(await results(), ['待核对', '待核对'], 'confirmation does not relabel a mismatch as passed');
    await generate();
    assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    const confirmedImages = await previewImages(), confirmedDoc = await save('docx', 'confirmed-budget');
    const confirmedPdf = await save('pdf', 'confirmed-budget');
    assert.doesNotMatch(await docXml(confirmedDoc), /预算|BUDGET-ONLY-NOTE|确认继续/);
    assert.match(await docXml(confirmedDoc), /合成供应商甲/);
    assert.match(await docXml(confirmedDoc), /合成供应商乙/);

    // Re-uploading the same filename is a new version, not an old confirmed result.
    await upload('budget', budgetBytes([budgetRow(first, '98.00')]), '合成月份预算.xlsx');
    await blocked();
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 0);
    assert.match(await checks().first().innerText(), /98\.00|98/);
    await blockedGeneration();
    await confirmAll();
    await upload('detail', detailBytes([first, second]), '合成付款明细.xlsx');
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 0, 'new detail file resets budget confirmations');

    // Multiple candidate budgets stay inside the same application row and need one confirmation.
    await upload('budget', budgetBytes([budgetRow(first),budgetRow(first,'101.00'),budgetRow(second)]), '合成多候选预算.xlsx');
    await assertBudgetTable([2,3]);
    assert.deepEqual(await results(), ['待核对','预算内']);
    const candidateCell = checks().first().locator('td').nth(3);
    assert.equal(await candidateCell.locator('.payment-budget-candidate').count(), 2);
    assert.equal(await checks().nth(1).locator('td').nth(3).locator('.payment-budget-candidate').count(), 1);
    assert.match(await candidateCell.innerText(), /合成多候选预算.xlsx/);
    assert.match(await candidateCell.innerText(), /预算明细.*第\s*2\s*行/);
    assert.match(await candidateCell.innerText(), /预算明细.*第\s*3\s*行/);
    assert.match(await candidateCell.innerText(), /100\.00/);
    assert.match(await candidateCell.innerText(), /101\.00/);
    assert.equal(await checks().first().locator('[data-budget-confirm]').count(), 1);
    assert.equal(await checks().nth(1).locator('[data-budget-confirm]').count(), 0);
    await blockedGeneration();
    await checks().first().locator('[data-budget-confirm]').check();
    await generate(); assert.match(await page.locator('#paymentStatus').innerText(), /已生成/);
    await assertBudgetTable([2,3]);

    // Multiple valid sheets require a deliberate selection and switching clears approval.
    const multi = bytes([
      ['一月预算', [budgetHeaders, budgetRow(first, '99.00'), budgetRow(second)]],
      ['二月预算', [budgetHeaders, budgetRow(first), budgetRow(second)]],
      ['说明', [['此表没有预算表头']]],
    ]);
    await upload('budget', multi, '多月份预算.xlsx');
    assert.equal(await page.locator('#paymentBudgetSheet').inputValue(), '');
    await blockedGeneration();
    await page.locator('#paymentBudgetSheet').selectOption('一月预算');
    assert.deepEqual(await results(), ['待核对', '预算内']);
    await confirmAll(); await generate();
    await page.locator('#paymentBudgetSheet').selectOption('二月预算');
    await blocked(); assert.deepEqual(await results(), ['预算内', '预算内']);
    await page.locator('#paymentBudgetSheet').selectOption('一月预算');
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 0);

    // An old delayed read must not overwrite a later budget or its provenance.
    await page.evaluate(() => {
      window.__budgetRead = File.prototype.arrayBuffer;
      File.prototype.arrayBuffer = function () {
        const reading = window.__budgetRead.call(this);
        if (this.name !== '延迟旧预算.xlsx') return reading;
        return new Promise((resolve, reject) => { window.__releaseBudget = async () => { try { resolve(await reading); } catch (error) { reject(error); } }; });
      };
    });
    await page.locator('#paymentBudget').setInputFiles({name:'延迟旧预算.xlsx',mimeType:MIME,buffer:budgetBytes([budgetRow(first, '1.00')])});
    await page.waitForFunction(() => !!window.__releaseBudget);
    await blocked();
    await upload('budget', budgetBytes([budgetRow(first),budgetRow(second)]), '最新跨年预算.xlsx');
    const newestText = await page.locator('#paymentBudgetChecks').innerText(), newestReference = await page.locator('[data-reference="budget"]').innerText();
    await page.evaluate(async () => {
      await window.__releaseBudget();
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      File.prototype.arrayBuffer = window.__budgetRead;
      delete window.__releaseBudget; delete window.__budgetRead;
    });
    assert.equal(await page.locator('#paymentBudgetChecks').innerText(), newestText);
    assert.equal(await page.locator('[data-reference="budget"]').innerText(), newestReference);
    assert.deepEqual(await results(), ['预算内', '预算内']);

    // Invalid replacement clears old output and cannot be treated as optional absence.
    await generate();
    await upload('budget', Buffer.from([0,1,2,3]), '损坏预算.xlsx');
    await blockedGeneration();
    assert.equal(await page.evaluate(async () => (await PaymentCache.get('budget'))?.name || ''), '');
    const errorBudget = budgetBytes([budgetRow(first)]), damaged = X.read(errorBudget, {type:'buffer'});
    damaged.Sheets.预算明细.G2 = {t:'e',v:7,f:'1/0'};
    await upload('budget', Buffer.from(X.write(damaged,{type:'buffer',bookType:'xlsx'})), '公式错误预算.xlsx');
    await blockedGeneration();
    assert.match(await page.locator('#paymentBudgetChecks').innerText(), /公式|Excel.*错误|计算结果/);
    await page.locator('#paymentClearBudget').click();
    assert.match(await page.locator('#paymentBudgetChecks').innerText(), /未进行预算核对/);
    await generate(); assert.deepEqual(await previewImages(), confirmedImages);
    const removedPdf = await save('pdf', 'budget-removed');
    assert.deepEqual(removedPdf, confirmedPdf, 'budget exception information must not leak into the PDF');

    // The file persists across tool changes/reload; confirmations do not persist.
    await upload('budget', budgetBytes([budgetRow(first,'99.00'),budgetRow(second)]), '缓存预算.xlsx');
    await confirmAll();
    await page.evaluate(async () => { await PaymentCache.get('budget'); });
    const reference = await page.locator('[data-reference="budget"]').innerText();
    await page.locator('button[data-tool="funds"]').click();
    await page.locator('button[data-tool="payment"]').click();
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 1);
    assert.equal(await page.locator('[data-reference="budget"]').innerText(), reference);
    await page.reload(); await page.locator('button[data-tool="payment"]').click();
    await page.locator('#paymentExcel:not([disabled])').waitFor();
    assert.equal(await page.locator('[data-reference="budget"]').innerText(), reference);
    assert.deepEqual(await results(), ['待核对','预算内']);
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 0);
    await blockedGeneration();
    await page.setViewportSize({width:390,height:844});
    await page.locator('#paymentBudget').scrollIntoViewIfNeeded();
    assert(await page.locator('#paymentBudget').isVisible());
    assert(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
    const box = await page.locator('#paymentBudget').boundingBox();
    assert(box && box.width > 0 && box.x >= -1 && box.x + box.width <= 391, 'budget upload control remains usable on mobile');
    await page.screenshot({path:path.join(out,'budget-mobile.png'),fullPage:false});
    const tableScroll = page.locator('.payment-budget-table-scroll');
    await tableScroll.scrollIntoViewIfNeeded();
    assert(await tableScroll.evaluate(el => el.scrollWidth > el.clientWidth), 'wide table scrolls inside the mobile page');
    await page.screenshot({path:path.join(out,'budget-table-mobile.png'),fullPage:false});
    await page.locator('[data-budget-confirm]').first().scrollIntoViewIfNeeded();
    const confirmationBox = await page.locator('[data-budget-confirm]').first().boundingBox();
    assert(confirmationBox && confirmationBox.x >= 0 && confirmationBox.x + confirmationBox.width <= 391, 'confirmation is reachable by horizontal scroll');
    await page.setViewportSize({width:1440,height:1000});
    await tableScroll.evaluate(el => { el.scrollLeft = 0; });
    await page.locator('#paymentBudgetChecks').scrollIntoViewIfNeeded();
    await page.screenshot({path:path.join(out,'budget-review.png'),fullPage:false});
    // Clearing cache invalidates an in-flight generation before the database clear finishes.
    await confirmAll();
    await page.evaluate(() => {
      window.__budgetOriginalLayout = PaymentDocument.layout;
      window.__budgetOriginalClear = PaymentCache.clear;
      PaymentDocument.layout = (...args) => new Promise((resolve, reject) => {
        window.__releaseBudgetLayout = async () => {
          try { resolve(await window.__budgetOriginalLayout.apply(PaymentDocument, args)); } catch (error) { reject(error); }
        };
      });
      PaymentCache.clear = () => new Promise((resolve, reject) => {
        window.__releaseBudgetClear = async () => {
          try { await window.__budgetOriginalClear.call(PaymentCache); resolve(); } catch (error) { reject(error); }
        };
      });
    });
    await page.locator('#paymentGenerate').click();
    await page.waitForFunction(() => !!window.__releaseBudgetLayout);
    assert(await page.locator('#paymentGenerate').isDisabled());
    await page.locator('#paymentClearCache').click();
    await page.waitForFunction(() => !!window.__releaseBudgetClear);
    const beforeClearDownloads = downloads.length;
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(), 0, 'cache clearing immediately revokes confirmations');
    await page.evaluate(async () => { await window.__releaseBudgetLayout(); });
    await page.waitForFunction(() => document.querySelector('#paymentGenerate').textContent === '核对并生成预览');
    await blocked();
    assert(await page.locator('#paymentGenerate').isDisabled(), 'old generation finally cannot enable controls while cache clear is pending');
    assert(await page.locator('#paymentDownloadWord').isDisabled());
    assert(await page.locator('#paymentExcel').isDisabled());
    assert.equal(downloads.length, beforeClearDownloads);
    assert(await page.evaluate(async () => !!(await PaymentCache.get('budget'))), 'database clear remains gated until the explicit release');
    await page.evaluate(async () => { await window.__releaseBudgetClear(); });
    await page.locator('#paymentExcel:not([disabled])').waitFor();
    await page.evaluate(() => {
      PaymentDocument.layout = window.__budgetOriginalLayout; PaymentCache.clear = window.__budgetOriginalClear;
      delete window.__budgetOriginalLayout; delete window.__budgetOriginalClear;
      delete window.__releaseBudgetLayout; delete window.__releaseBudgetClear;
    });
    assert.deepEqual(await page.evaluate(async () => Promise.all(['detail','directory','budget'].map(async kind => !!(await PaymentCache.get(kind))))), [false,false,false]);
    assert.match(await page.locator('#paymentBudgetChecks').innerText(), /未进行预算核对/);
    assert.equal(await checks().count(), 0);
    await blocked();
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({status:'PASS',base,out,optionalBudget:true,exactMatch:true,budgetTable:true,multipleCandidatesOneRow:true,individualReview:true,allDownloadGates:true,unchangedPrint:true,sameNameReplacement:true,detailReset:true,dynamicMonths:true,sheetSelection:true,latestBudgetWins:true,invalidRecovery:true,formulaGuard:true,cacheRestore:true,clearCache:true,clearDuringGeneration:true,mobile:true,errors}));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
