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
    await page.goto(base,{waitUntil:'networkidle'});
    const url=page.url();
    async function select(id) {
      await page.locator(`[data-tool="${id}"].tool-tab`).click();
      assert.equal(page.url(),url);
      return page.frameLocator(`#embeddedTools iframe[data-tool="${id}"]`);
    }
    let receipt=await select('receipt-sign');
    await receipt.locator('#signatureWidth').fill('42');
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
