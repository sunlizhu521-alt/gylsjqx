// Synthetic browser acceptance for the optional, memory-only OA source.
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const C = require('../payment-core.js');
const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const oaHeaders = ['流程编号','付款金额','未操作者','收款单位','付款主体'];
const budgetHeaders = ['供应商','付款主体','对公付款0A流程号','OA-申请付款金额','付款时间'];
const detail = (extra = {}) => ({
  申请日期:'2026-11-27',付款主体:'合成甲公司',申请部门:'履约部',经办人:'测试人',
  供应商全称:'合成供应商甲',开户银行及账号:'单位:合成供应商甲\n账号:000123456789',
  合同编号:'MDST-DGFK20261127001',合同总金额:'/',已付金额:'/',本次申请付款金额:'100.00',
  票据状态:'已开票',付款事由:'合成货款',备注:'',...extra,
});
const first = detail(), second = detail({供应商全称:'合成供应商乙',开户银行及账号:'单位:合成供应商乙\n账号:000123456788',合同编号:'MDST-DGFK20261127002',本次申请付款金额:'200.00'});
const oaRow = (row, amount = row.本次申请付款金额, pending = '合成办理甲') => [row.合同编号,amount,pending,row.供应商全称,row.付款主体];
const budgetRow = (row, amount = row.本次申请付款金额) => [row.供应商全称,row.付款主体,row.合同编号,amount,'2027年01月25日'];
function bytes(sheets) {
  const book = X.utils.book_new();
  for (const [name, rows] of sheets) X.utils.book_append_sheet(book, X.utils.aoa_to_sheet(rows), name);
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
const detailBytes = rows => bytes([['付款明细',[C.required,...rows.map(row => C.required.map(h => row[h] ?? ''))]]]);
const oaBytes = (rows, headers = oaHeaders) => bytes([['流程导出',[headers,...rows]]]);
const budgetBytes = rows => bytes([['月份预算',[budgetHeaders,...rows]]]);
const docXml = async data => (await Zip.loadAsync(data)).file('word/document.xml').async('string');

(async () => {
  const base = process.env.QA_URL || 'http://127.0.0.1:8899/';
  const out = process.env.OUT_DIR || process.env.QA_OUTPUT || '/private/tmp/payment-oa-browser';
  await fs.mkdir(out,{recursive:true});
  const browser = await chromium.launch({headless:true,...(process.env.CHROME_PATH ? {executablePath:process.env.CHROME_PATH} : {})});
  try {
    const context = await browser.newContext({viewport:{width:1440,height:1000},timezoneId:'Asia/Shanghai'});
    const page = await context.newPage(), errors = [], downloads = [];
    await page.clock.setFixedTime(new Date('2026-11-27T02:00:00Z'));
    page.on('pageerror', error => errors.push(error.message));
    page.on('download', download => downloads.push(download.suggestedFilename()));
    await page.goto(base); await page.locator('button[data-tool="payment"]').click();
    await page.locator('#paymentExcel:not([disabled])').waitFor();
    await page.evaluate(() => {
      window.__oaCacheWrites = [];
      const put = PaymentCache.put;
      PaymentCache.put = function (key, ...args) { window.__oaCacheWrites.push(key); return put.call(this,key,...args); };
    });
    async function upload(role, buffer, name) {
      const id = {detail:'paymentExcel',budget:'paymentBudget',oa:'paymentOA'}[role];
      await page.locator('#'+id).setInputFiles({name,mimeType:MIME,buffer});
      await page.waitForFunction(kind => {
        const reference = document.querySelector(`[data-reference="${kind}"]`);
        return reference && !reference.textContent.includes('正在读取');
      },role);
    }
    async function generate() {
      await page.locator('#paymentGenerate').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
    }
    async function blocked() {
      assert.equal(await page.locator('#paymentPages canvas').count(),0);
      assert(await page.locator('[data-payment-download="docx"]').isDisabled());
      assert(await page.locator('[data-payment-download="pdf"]').isDisabled());
    }
    async function blockedGeneration() {
      await generate(); await blocked();
      assert.match(await page.locator('#paymentStatus').innerText(),/OA|核对/);
      const before = downloads.length;
      await page.locator('#paymentDownloadWord').click();
      await page.waitForFunction(() => !document.querySelector('#paymentGenerate').disabled);
      assert.equal(downloads.length,before,'one-click Word must respect OA review');
      await blocked();
    }
    const rows = () => page.locator('#paymentBudgetChecks tbody > tr.payment-budget-check');
    const previews = () => page.locator('#paymentPages canvas').evaluateAll(canvases => canvases.map(canvas => canvas.toDataURL()));
    async function table(hasOA) {
      const headers = await page.locator('#paymentBudgetChecks thead th').allTextContents();
      assert.equal(headers.length,hasOA ? 8 : 6);
      if (hasOA) assert.deepEqual(headers.slice(4,7),['OA付款金额','未操作者','核对说明']);
      else { assert(!headers.includes('OA付款金额')); assert(!headers.includes('未操作者')); }
      assert.deepEqual(await rows().evaluateAll(records => records.map(row => Number(row.dataset.budgetRow))),[2,3]);
      assert(await rows().evaluateAll((records, length) => records.every(row => row.cells.length === length),hasOA ? 8 : 6));
      if (hasOA) for (const value of await page.locator('.payment-oa-amount').allTextContents()) {
        assert.doesNotMatch(value, /OA来源|第[ ]*[0-9]+[ ]*行|[.]xlsx?/, 'OA amount cells omit source files and rows');
      }
    }
    async function confirmAll() { for (const input of await page.locator('[data-budget-confirm]').all()) await input.check(); }
    async function save(kind, label) {
      const event = page.waitForEvent('download');
      await page.locator(`[data-payment-download="${kind}"]:visible`).click();
      const download = await event, target = path.join(out,label+'.'+kind);
      await download.saveAs(target); return fs.readFile(target);
    }
    async function printedContract(data) {
      return page.evaluate(xml => {
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const doc = new DOMParser().parseFromString(xml,'application/xml');
        const cell = doc.getElementsByTagNameNS(ns,'tbl')[1].getElementsByTagNameNS(ns,'tr')[1].getElementsByTagNameNS(ns,'tc')[3];
        return [...cell.getElementsByTagNameNS(ns,'t')].map(node => node.textContent).join('');
      },await docXml(data));
    }
    async function noOACandidate() {
      assert.equal(await rows().first().locator('.payment-oa-amount').innerText(),'未找到对应OA');
      assert.equal(await rows().first().locator('.payment-oa-amount .payment-budget-candidate').count(),0);
      assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'—');
      assert.equal(await rows().first().locator('[data-budget-confirm]').count(),1);
    }
    async function assertNoOACache() {
      assert.equal(await page.evaluate(async () => !!(await PaymentCache.get('oa'))),false);
      assert(await page.evaluate(() => window.__oaCacheWrites.every(key => ['detail','template','directory','budget'].includes(key))),'OA uploads must not use a different persistent cache key');
    }

    await upload('detail',detailBytes([first,second]),'合成付款明细.xlsx');
    assert.match(await page.locator('.payment-tool').innerText(),/未进行OA核对/);
    await generate(); assert.match(await page.locator('#paymentStatus').innerText(),/已生成/);
    const initialImages = await previews(), initialDoc = await save('docx','without-oa'), initialPdf = await save('pdf','without-oa');

    // OA is independently usable without a budget and must not change the print form.
    await upload('oa',oaBytes([oaRow(first),oaRow(second,'200.00','合成办理乙')]),'合成OA流程.xlsx');
    await blocked(); await table(true);
    assert.deepEqual(await page.locator('.payment-oa-amount .payment-budget-amount').allTextContents(),['￥100.00','￥200.00']);
    assert.deepEqual(await page.locator('.payment-oa-unoperated').allTextContents(),['合成办理甲','合成办理乙']);
    assert.equal(await page.locator('[data-budget-confirm]').count(),0);
    await generate(); assert.deepEqual(await previews(),initialImages);
    const matchedDoc = await save('docx','matched-oa'), matchedPdf = await save('pdf','matched-oa');
    assert.equal(await docXml(matchedDoc),await docXml(initialDoc));
    assert.deepEqual(matchedPdf,initialPdf);
    assert.doesNotMatch(await docXml(matchedDoc),/OA付款金额|未操作者|合成办理甲|合成OA流程/);
    await assertNoOACache();

    // Numeric OA cells must retain leading zeros supplied by the Excel number format.
    const paddedBook = X.read(detailBytes([first,second]),{type:'buffer'});
    const paddedAddress = X.utils.encode_cell({r:1,c:C.required.indexOf('合同编号')});
    paddedBook.Sheets.付款明细[paddedAddress] = {t:'n',v:123,z:'00000000'};
    await upload('detail',Buffer.from(X.write(paddedBook,{type:'buffer',bookType:'xlsx'})),'数字格式补零付款明细.xlsx');
    await upload('oa',oaBytes([oaRow({...first,合同编号:'00000123'}),oaRow(second)]),'补零编号OA.xlsx');
    await table(true);
    assert.equal(await rows().first().locator('.payment-budget-result').innerText(),'核对通过');
    assert.match(await rows().first().locator('td').nth(1).innerText(),/OA号：00000123/);
    assert.match(await rows().first().locator('.payment-oa-amount').innerText(),/流程编码：00000123/);
    assert.equal(await page.locator('[data-budget-confirm]').count(),0);
    await generate();
    assert.equal(await printedContract(await save('docx','numeric-oa-leading-zeros')),'00000123');
    await upload('detail',detailBytes([first,second]),'合成付款明细.xlsx');
    await upload('oa',oaBytes([oaRow(first),oaRow(second,'200.00','合成办理乙')]),'合成OA流程.xlsx');

    // Existing budget checks and the new OA checks share one row and confirmation.
    await upload('budget',budgetBytes([budgetRow(first),budgetRow(second)]),'合成月份预算.xlsx');
    await table(true); await generate();
    const suggestedOA = 'MDST-DGFK20261127999';
    await upload('oa',oaBytes([oaRow({...first,合同编号:suggestedOA},'100.00','建议办理人'),oaRow(second)]),'唯一建议OA.xlsx');
    assert.equal(await rows().first().locator('[data-budget-confirm]').count(),1);
    assert.equal(await rows().first().locator('.payment-budget-result').innerText(),'待核对');
    assert.match(await rows().first().locator('.payment-oa-amount').innerText(),new RegExp(suggestedOA));
    assert.match(await rows().first().locator('.payment-oa-amount').innerText(),new RegExp(first.供应商全称));
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'建议办理人');
    await blockedGeneration(); await confirmAll(); await generate();
    const suggestedDoc = await save('docx','confirmed-suggested-oa');
    assert.equal(await printedContract(suggestedDoc),first.合同编号,'confirmation does not replace the original OA in the print form');
    assert.equal(await docXml(suggestedDoc),await docXml(initialDoc));
    assert.deepEqual(await previews(),initialImages);

    // An original OA does not let another amount or supplier enter the candidate list.
    await upload('oa',oaBytes([oaRow(first,'99.00'),oaRow(second)]),'合成OA流程.xlsx');
    await noOACandidate();
    assert.equal(await page.locator('[data-budget-confirm]').count(),1);
    await blockedGeneration(); await confirmAll(); await generate();
    assert.match(await page.locator('#paymentStatus').innerText(),/已生成/);
    const confirmedDoc = await save('docx','confirmed-oa'), confirmedPdf = await save('pdf','confirmed-oa');
    assert.equal(await docXml(confirmedDoc),await docXml(initialDoc)); assert.deepEqual(confirmedPdf,initialPdf);

    await upload('budget',budgetBytes([budgetRow(first,'90.00'),budgetRow(second)]),'合成月份预算.xlsx');
    assert.equal(await rows().first().locator('[data-budget-confirm]').count(),1,'a budget and OA mismatch share one confirmation per application');
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(),0);
    assert.match(await rows().first().innerText(),/90\.00/); await noOACandidate();
    await blockedGeneration(); await confirmAll(); await generate();
    await upload('budget',budgetBytes([budgetRow(first),budgetRow(second)]),'合成月份预算.xlsx');
    await upload('oa',oaBytes([oaRow({...first,供应商全称:'合成其他收款单位'}),oaRow(second)]),'同号同额其他供应商.xlsx');
    await noOACandidate(); await blockedGeneration();

    // Missing original OA still has a supplier/amount suggestion, which remains review-only.
    await upload('detail',detailBytes([{...first,合同编号:'/'},second]),'原OA空白付款明细.xlsx');
    await upload('oa',oaBytes([oaRow(first,'100.00','缺号建议办理人'),oaRow(second)]),'原号缺失建议.xlsx');
    await table(true);
    assert.equal(await rows().first().locator('[data-budget-confirm]').count(),1);
    assert.match(await rows().first().locator('.payment-oa-amount').innerText(),new RegExp(first.合同编号));
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'缺号建议办理人');
    await blockedGeneration(); await confirmAll(); await generate();
    const missingOADoc = await save('docx','confirmed-missing-original-oa');
    assert.equal(await printedContract(missingOADoc),'/','suggested OA is never filled into the print form');
    await upload('detail',detailBytes([first,second]),'合成付款明细.xlsx');

    // Matching the original OA must not hide other same-supplier/same-amount processes.
    await upload('oa',oaBytes([
      oaRow({...first,合同编号:suggestedOA},'100.00','候选办理甲'),
      oaRow(first,'100.00','候选办理乙'),oaRow(second),
    ]),'同供应商同额多个OA.xlsx');
    await table(true);
    const candidates = rows().first().locator('.payment-oa-amount .payment-budget-candidate');
    assert.equal(await candidates.count(),2);
    assert.deepEqual(await candidates.locator('.payment-budget-amount').allTextContents(),['￥100.00','￥100.00']);
    assert.match(await candidates.nth(0).innerText(),new RegExp(suggestedOA));
    assert.match(await candidates.nth(1).innerText(),new RegExp(first.合同编号));
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'记录 1：候选办理甲\n记录 2：候选办理乙');
    assert.equal(await rows().first().locator('.payment-budget-result').innerText(),'待核对');
    assert.equal(await rows().first().locator('[data-budget-confirm]').count(),1);
    await blockedGeneration(); await confirmAll(); await generate();
    assert.equal(await printedContract(await save('docx','confirmed-multiple-oa')),first.合同编号);

    // Repeated consistent OA entries are read once; conflicting duplicates need review.
    await upload('oa',oaBytes([oaRow(first),oaRow(first),oaRow(second,'200.00','')]),'重复一致OA.xlsx');
    await table(true); assert.equal(await page.locator('[data-budget-confirm]').count(),0);
    assert.equal(await rows().first().locator('.payment-oa-amount .payment-budget-amount').innerText(),'￥100.00');
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'合成办理甲');
    const emptyPending = (await rows().nth(1).locator('.payment-oa-unoperated').innerText()).trim();
    assert(/^$|^(?:[—/-]|（?空白）?)$/.test(emptyPending),'a blank pending operator must stay blank or be explicitly displayed as blank');
    await generate(); assert.match(await page.locator('#paymentStatus').innerText(),/已生成/);
    for (const [label, conflict] of [
      ['金额',oaRow(first,'101.00')],
      ['供应商',oaRow({...first,供应商全称:'合成其他收款单位'})],
      ['未操作者',oaRow(first,'100.00','合成冲突办理人')],
    ]) {
      await upload('oa',oaBytes([oaRow(first),conflict,oaRow(second)]),`同流程${label}冲突OA.xlsx`);
      assert.equal(await page.locator('[data-budget-confirm]').count(),1);
      assert.match(await rows().first().innerText(),/冲突|不一致/,'conflicts anywhere in the same process must be reviewed');
      await blockedGeneration(); await confirmAll(); await generate();
    }

    // Even an identical filename is a new upload and resets all confirmations/results.
    await upload('oa',oaBytes([oaRow({...first,合同编号:suggestedOA},'100.00','旧版办理人'),oaRow(second)]),'合成OA流程.xlsx');
    await confirmAll(); await generate();
    await upload('oa',oaBytes([oaRow({...first,合同编号:'MDST-DGFK20261127998'},'100.00','新版办理人'),oaRow(second)]),'合成OA流程.xlsx');
    await blocked(); assert.equal(await page.locator('[data-budget-confirm]:checked').count(),0);
    assert.match(await rows().first().locator('.payment-oa-amount').innerText(),/MDST-DGFK20261127998/);
    assert.doesNotMatch(await rows().first().locator('.payment-oa-amount').innerText(),new RegExp(suggestedOA));
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'新版办理人');
    await confirmAll();
    await upload('detail',detailBytes([first,second]),'合成付款明细.xlsx');
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(),0);

    const multi = bytes([
      ['流程导出一',[oaHeaders,oaRow(first,'99.00'),oaRow(second)]],
      ['流程导出二',[oaHeaders,oaRow(first),oaRow(second)]],
    ]);
    await upload('oa',multi,'多工作表OA.xlsx');
    assert.equal(await page.locator('#paymentOASheet').inputValue(),'');
    await blockedGeneration();
    await page.locator('#paymentOASheet').selectOption('流程导出一'); await confirmAll(); await generate();
    await page.locator('#paymentOASheet').selectOption('流程导出二'); await blocked();
    assert.equal(await page.locator('[data-budget-confirm]').count(),0);
    await page.locator('#paymentOASheet').selectOption('流程导出一');
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(),0);

    // An old slow read cannot replace the latest OA or revive one that was removed.
    await page.evaluate(() => {
      window.__originalOARead = File.prototype.arrayBuffer;
      File.prototype.arrayBuffer = function () {
        const reading = window.__originalOARead.call(this);
        if (!this.name.startsWith('延迟OA')) return reading;
        return new Promise((resolve,reject) => { window.__releaseOARead = async () => { try { resolve(await reading); } catch (error) { reject(error); } }; });
      };
    });
    async function delayedOA(name) {
      await page.locator('#paymentOA').setInputFiles({name,mimeType:MIME,buffer:oaBytes([oaRow(first,'1.00','旧办理人'),oaRow(second)])});
      await page.waitForFunction(() => !!window.__releaseOARead); await blocked();
    }
    async function releaseOA() {
      await page.evaluate(async () => { await window.__releaseOARead(); delete window.__releaseOARead; await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))); });
    }
    await delayedOA('延迟OA旧版本.xlsx');
    await upload('oa',oaBytes([oaRow(first,'100.00','最新办理人'),oaRow(second)],['流程编码',...oaHeaders.slice(1)]),'最新OA.xlsx');
    const latestText = await page.locator('#paymentBudgetChecks').innerText(), latestReference = await page.locator('[data-reference="oa"]').innerText();
    await releaseOA();
    assert.equal(await page.locator('#paymentBudgetChecks').innerText(),latestText);
    assert.equal(await page.locator('[data-reference="oa"]').innerText(),latestReference);
    assert.equal(await rows().first().locator('.payment-oa-unoperated').innerText(),'最新办理人');
    await delayedOA('延迟OA待移除.xlsx'); await page.locator('#paymentClearOA').click(); await releaseOA();
    assert.match(await page.locator('.payment-tool').innerText(),/未进行OA核对/); await table(false);
    await page.evaluate(() => { File.prototype.arrayBuffer = window.__originalOARead; delete window.__originalOARead; });

    await upload('oa',oaBytes([oaRow(first),oaRow(second)]),'有效OA.xlsx'); await generate();
    await upload('oa',Buffer.from([0,1,2,3]),'损坏OA.xlsx'); await blockedGeneration();
    await upload('oa',oaBytes([oaRow(first).slice(0,3),oaRow(second).slice(0,3)],oaHeaders.slice(0,3)),'缺少收款单位OA.xlsx');
    await blockedGeneration(); assert.match(await page.locator('#paymentBudgetChecks').innerText(),/收款单位/);
    const badBook = X.read(oaBytes([oaRow(first),oaRow(second)]),{type:'buffer'});
    badBook.Sheets.流程导出.B2 = {t:'e',v:7,f:'1/0'};
    await upload('oa',Buffer.from(X.write(badBook,{type:'buffer',bookType:'xlsx'})),'公式错误OA.xlsx');
    await blockedGeneration(); assert.match(await page.locator('#paymentBudgetChecks').innerText(),/公式|Excel.*错误|计算结果/);
    await page.locator('#paymentClearOA').click(); await blocked(); await table(false);
    await generate(); assert.deepEqual(await previews(),initialImages);

    // Budget/detail persist; OA exists only in the current document's memory.
    await upload('oa',oaBytes([oaRow(first,'99.00'),oaRow(second)]),'本次内存OA.xlsx'); await confirmAll();
    await assertNoOACache();
    const cache = await page.evaluate(async () => Object.fromEntries(await Promise.all(['detail','budget'].map(async key => [key,(await PaymentCache.get(key)).name]))));
    const reference = await page.locator('[data-reference="oa"]').innerText();
    await page.locator('button[data-tool="funds"]').click(); await page.locator('button[data-tool="payment"]').click();
    assert.equal(await page.locator('[data-reference="oa"]').innerText(),reference);
    assert.equal(await page.locator('[data-budget-confirm]:checked').count(),1);
    const fresh = await context.newPage(); fresh.on('pageerror', error => errors.push(error.message));
    await fresh.goto(base); await fresh.locator('button[data-tool="payment"]').click(); await fresh.locator('#paymentExcel:not([disabled])').waitFor();
    assert.match(await fresh.locator('.payment-tool').innerText(),/未进行OA核对/);
    assert.doesNotMatch(await fresh.locator('.payment-tool').innerText(),/本次内存OA.xlsx/);
    assert.equal(await fresh.locator('#paymentBudgetChecks thead th').count(),6);
    await fresh.close();
    await page.reload(); await page.locator('button[data-tool="payment"]').click(); await page.locator('#paymentExcel:not([disabled])').waitFor();
    assert.match(await page.locator('.payment-tool').innerText(),/未进行OA核对/);
    assert.doesNotMatch(await page.locator('.payment-tool').innerText(),/本次内存OA.xlsx/);
    assert.equal(await page.evaluate(async () => !!(await PaymentCache.get('oa'))),false);
    assert.deepEqual(await page.evaluate(async () => Object.fromEntries(await Promise.all(['detail','budget'].map(async key => [key,(await PaymentCache.get(key)).name])))),cache);
    await table(false); await generate(); assert.match(await page.locator('#paymentStatus').innerText(),/已生成/);

    await upload('oa',oaBytes([oaRow({...first,合同编号:suggestedOA}),oaRow(second)]),'移动端OA.xlsx');
    await page.setViewportSize({width:390,height:844});
    const uploads = page.locator('.payment-upload-grid');
    const tops = await uploads.locator('.payment-upload').evaluateAll(slots => slots.map(slot => slot.getBoundingClientRect().top));
    assert.equal(tops.length,5); assert(tops.every(top => Math.abs(top-tops[0]) < 1),'five source slots stay in one scrollable row');
    assert(await uploads.evaluate(el => el.scrollWidth > el.clientWidth));
    await page.locator('#paymentOA').scrollIntoViewIfNeeded();
    assert(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
    const box = await page.locator('#paymentOA').boundingBox();
    assert(box && box.width > 0 && box.x >= -1 && box.x+box.width <= 391,'OA upload remains reachable by horizontal scrolling');
    await page.screenshot({path:path.join(out,'oa-slots-mobile.png'),fullPage:false});
    await page.locator('[data-budget-confirm]').first().scrollIntoViewIfNeeded();
    await page.screenshot({path:path.join(out,'oa-review-mobile.png'),fullPage:false});
    await page.setViewportSize({width:1440,height:1000});
    await page.locator('.payment-budget-table-scroll').evaluate(el => { el.scrollLeft = 0; });
    await page.locator('#paymentBudgetChecks').scrollIntoViewIfNeeded();
    await page.screenshot({path:path.join(out,'oa-review.png'),fullPage:false});
    assert.deepEqual(errors,[]);
    console.log(JSON.stringify({status:'PASS',base,out,optionalOA:true,oaOnly:true,combinedReview:true,supplierAmountMatch:true,suggestedOA:true,missingOASuggestion:true,multipleCandidates:true,noUnrelatedCandidates:true,originalOAPreserved:true,duplicateConsistency:true,duplicateConflict:true,blankOperator:true,sameNameReplacement:true,detailReset:true,sheetSelection:true,latestOA:true,removeDuringRead:true,invalidRecovery:true,memoryOnly:true,toolSwitch:true,refreshAndNewPage:true,otherCachesPreserved:true,unchangedPrint:true,mobile:true,errors}));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
