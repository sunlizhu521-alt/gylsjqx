const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');
const C = require('../fulfillment-core.js');
(async () => {
  const browser = await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH});
  try {
    const page = await browser.newPage();
    const errors=[]; page.on('pageerror',e=>errors.push(e.message));
    const base=process.env.QA_URL || 'http://127.0.0.1:8940/';
    await page.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
    // Simulate stale HTML at the previously unversioned child URL.
    await page.route('**/receipt-sign.html',r=>r.fulfill({contentType:'text/html',body:'<input id="signature">旧版签字页面'}));
    await page.goto(base,{waitUntil:'networkidle'});
    const url=page.url();
    async function select(id) {
      await page.locator(`[data-tool="${id}"].tool-tab`).click();
      assert.equal(page.url(),url);
      return page.frameLocator(`#embeddedTools iframe[data-tool="${id}"]`);
    }
    let receipt=await select('receipt-sign');
    await receipt.locator('#signatureWidth').fill('42');
    assert.match(await page.locator('iframe[data-tool="receipt-sign"]').getAttribute('src'), /[?&]v=/);
    assert.equal(await receipt.locator('#signature,#clearSignature').count(),0);
    const receiptBook=X.utils.book_new();
    X.utils.book_append_sheet(receiptBook,X.utils.aoa_to_sheet([['采购入库单'],['物料编码','数量'],['TEST',2]]),'Sheet1');
    await receipt.locator('#file:not([disabled])').waitFor();
    await receipt.locator('#file').setInputFiles({name:'自动签字.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(X.write(receiptBook,{type:'buffer',bookType:'xlsx'}))});
    await receipt.locator('#download:not([disabled])').waitFor();
    assert.match(await receipt.locator('#status').innerText(),/签名已添加/);
    const signedDownload=page.waitForEvent('download');
    await receipt.locator('#download').click();
    assert.equal((await signedDownload).suggestedFilename(),'自动签字-签字.pdf');
    await select('beihuo');
    await page.frameLocator('iframe[data-tool="beihuo"]').locator('#applyAllButton').waitFor();
    await select('contract');
    await page.frameLocator('iframe[data-tool="contract"]').locator('#orderFile').waitFor();
    const fulfillment=await select('fulfillment');
    const wb=X.utils.book_new();
    for (const name of C.SHEETS) X.utils.book_append_sheet(wb,X.utils.aoa_to_sheet([C.COLUMNS,C.COLUMNS.map(k=>k==='未交付数量'?2:k==='物料编码'?'123':k==='事业部'?'国内事业部':'测试')]),name);
    await fulfillment.locator('#file:not([disabled])').waitFor();
    await fulfillment.locator('#file').setInputFiles({name:'测试2026.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(X.write(wb,{type:'buffer',bookType:'xlsx'}))});
    await fulfillment.locator('#download:not([disabled])').waitFor();
    const before=await fulfillment.locator('#count').innerText();
    await select('payment');
    assert.equal(await page.locator('#toolPanel').isVisible(),true);
    await select('receipt-sign');
    assert.equal(await receipt.locator('#signatureWidth').inputValue(),'42');
    await select('fulfillment');
    assert.equal(await fulfillment.locator('#count').innerText(),before);
    const pending=page.waitForEvent('download');
    await fulfillment.locator('#download').click();
    assert.match((await pending).suggestedFilename(),/汇总版.xlsx$/);
    assert.equal(await page.locator('#embeddedTools iframe:visible').count(),1);
    await page.setViewportSize({width:390,height:844});
    assert.equal(await fulfillment.locator('#download').isVisible(),true);
    assert.deepEqual(errors,[]);
    console.log('PASS: inline tools, retained state, upload and download, mobile');
  } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exit(1);});
