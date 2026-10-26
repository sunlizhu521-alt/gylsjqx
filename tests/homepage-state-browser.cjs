const assert = require('node:assert/strict');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');

const url = process.env.QA_URL || 'http://127.0.0.1:8940/';
const key = 'gylsjqx-last-tool';
const toolIds = ['funds','receipt-sign','beihuo','fulfillment','contract','payment'];
const retiredIds = ['merge','compare','seal','ocr','match','convert','clean','diff','mask','split','formula','pivot','chart'];
const workbook = X.utils.book_new();
X.utils.book_append_sheet(workbook, X.utils.aoa_to_sheet([['物料编码', '数量'], ['TEST-001', 12]]), '明细');
const upload = {name:'工具切换测试.xlsx', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer:Buffer.from(X.write(workbook,{type:'buffer',bookType:'xlsx'}))};
const active = async (page, id) => page.waitForFunction(id => document.querySelector('.tool-tab.active')?.dataset.tool === id, id);
const select = async (page, id) => { await page.locator(`.tool-tab[data-tool="${id}"]`).click(); await active(page,id); };

(async () => {
  const browser = await chromium.launch({headless:true, executablePath:process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
  const errors = [];
  const open = async context => {const page = await context.newPage();page.on('pageerror',e=>errors.push(e.message));await page.goto(url,{waitUntil:'networkidle'});return page;};
  try {
    const context = await browser.newContext();
    const page = await open(context);
    assert.deepEqual(await page.locator('.tool-tab').evaluateAll(tabs=>tabs.map(tab=>tab.dataset.tool)),toolIds);
    assert.equal(await page.locator('#previewTable, #chartCanvas, #downloadResult, #toastArea').count(),0);
    assert.deepEqual(await page.evaluate(()=>[
      'Chart','Tesseract','renderToolPanel','runMatch','runConvert','runClean','runMerge','runDiff',
      'runMask','runSplit','runFormula','runPivot','runCompare','runSealStamp','runOcr','runChart'
    ].filter(name=>typeof window[name]!=='undefined')),[]);
    assert.deepEqual(await page.locator('script[src]').evaluateAll(scripts=>scripts.map(script=>script.src).filter(src=>/chart|tesseract|ocr/i.test(src))),[]);
    await active(page,'funds');
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'funds');
    let frame = page.frameLocator('iframe[data-tool="funds"]');
    await frame.locator('#ledgerFile:not([disabled])').setInputFiles(upload);
    await frame.locator('#ledgerInfo').filter({hasText:upload.name}).waitFor();
    await frame.locator('body').evaluate(el=>el.dataset.retained='yes');
    await select(page,'receipt-sign');
    await page.frameLocator('iframe[data-tool="receipt-sign"]').locator('body').waitFor();
    await select(page,'payment');
    await select(page,'funds');
    assert.equal(await page.locator('iframe[data-tool="funds"]').count(),1);
    assert.equal(await frame.locator('body').getAttribute('data-retained'),'yes');
    assert.equal(await frame.locator('#ledgerFile').evaluate(el=>el.files.length),1);
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'funds');

    await select(page,'receipt-sign');
    const second = await open(context);
    await active(second,'receipt-sign');
    await select(second,'funds');
    assert.equal(await second.frameLocator('iframe[data-tool="funds"]').locator('#ledgerFile').evaluate(el=>el.files.length),0);
    await select(second,'receipt-sign');
    await second.close();
    await page.reload({waitUntil:'networkidle'});
    await active(page,'receipt-sign');
    await select(page,'funds');
    frame = page.frameLocator('iframe[data-tool="funds"]');
    assert.equal(await frame.locator('#ledgerFile').evaluate(el=>el.files.length),0);
    await select(page,'payment');
    await page.reload({waitUntil:'networkidle'});
    await active(page,'payment');
    assert.equal(await page.locator('iframe').count(),0);
    await page.evaluate(()=>switchTool('removed-tool'));
    await active(page,'payment');
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'payment');

    for (const invalid of [...retiredIds,'removed-tool','{"id":"funds"}','../../funds.html','']) {
      await page.evaluate(({key,invalid})=>localStorage.setItem(key,invalid),{key,invalid});
      await page.reload({waitUntil:'networkidle'});
      await active(page,'funds');
      assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'funds');
      assert.deepEqual(await page.locator('.tool-tab').evaluateAll(tabs=>tabs.map(tab=>tab.dataset.tool)),toolIds);
    }
    await page.evaluate(key=>{localStorage.removeItem(key);localStorage.setItem('dpc-last-tool','payment');},key);
    await page.reload({waitUntil:'networkidle'});
    await active(page,'funds');
    await context.close();

    for (const mode of ['unavailable','read-error','write-error']) {
      const context = await browser.newContext({storageState:{cookies:[],origins:[{origin:new URL(url).origin,localStorage:[{name:key,value:'payment'}]}]}});
      await context.addInitScript(mode=>{
        const denied=()=>{throw new DOMException('Storage disabled','SecurityError');};
        if(mode==='unavailable')Object.defineProperty(window,'localStorage',{get:denied});
        else Object.defineProperty(Storage.prototype,mode==='read-error'?'getItem':'setItem',{value:denied});
      },mode);
      const page = await open(context);
      await active(page,mode==='write-error'?'payment':'funds');
      await select(page,'payment');
      await page.locator('#themeToggle').click();
      assert.equal(await page.locator('html').getAttribute('data-theme'),'dark');
      await select(page,'funds');
      await context.close();
    }
    assert.deepEqual(errors,[]);
    console.log('PASS homepage state: exactly six tools, retired UI/scripts/global functions removed, funds default, last tool on refresh/new tab, embedded upload state retained while switching, uploads not persisted by homepage, all retired tool IDs ignored, disabled storage safe');
  } finally { await browser.close(); }
})().catch(error=>{console.error(error);process.exit(1);});
