const assert = require('node:assert/strict');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js');

const url = process.env.QA_URL || 'http://127.0.0.1:8940/';
const key = 'gylsjqx-last-tool';
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
    await active(page,'merge');
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'merge');
    await page.locator('input[data-input="merge1"]').setInputFiles(upload);
    await page.waitForFunction(()=>document.querySelector('#merge1Name')?.textContent==='工具切换测试.xlsx');
    await page.locator('[data-merge-map="merge1"]').selectOption('物料编码');
    await page.locator('#runAction').click();
    await page.locator('#downloadResult:not([disabled])').waitFor();
    await select(page,'funds');
    let frame = page.frameLocator('iframe[data-tool="funds"]');
    await frame.locator('#ledgerFile:not([disabled])').setInputFiles(upload);
    await frame.locator('#ledgerInfo').filter({hasText:upload.name}).waitFor();
    await frame.locator('body').evaluate(el=>el.dataset.retained='yes');
    await select(page,'merge');
    assert.equal(await page.locator('#merge1Name').innerText(),upload.name);
    assert.equal(await page.locator('#downloadResult').isEnabled(),true);
    await select(page,'funds');
    assert.equal(await page.locator('iframe[data-tool="funds"]').count(),1);
    assert.equal(await frame.locator('body').getAttribute('data-retained'),'yes');
    assert.equal(await frame.locator('#ledgerFile').evaluate(el=>el.files.length),1);
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'funds');

    const second = await open(context);
    await active(second,'funds');
    assert.equal(await second.frameLocator('iframe[data-tool="funds"]').locator('#ledgerFile').evaluate(el=>el.files.length),0);
    await second.close();
    await page.reload({waitUntil:'networkidle'});
    await active(page,'funds');
    frame = page.frameLocator('iframe[data-tool="funds"]');
    assert.equal(await frame.locator('#ledgerFile').evaluate(el=>el.files.length),0);
    await select(page,'compare');
    await page.reload({waitUntil:'networkidle'});
    await active(page,'compare');
    assert.equal(await page.locator('iframe').count(),0);
    await page.evaluate(()=>switchTool('removed-tool'));
    await active(page,'compare');
    assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'compare');

    for (const invalid of ['removed-tool','{"id":"funds"}','../../funds.html','']) {
      await page.evaluate(({key,invalid})=>localStorage.setItem(key,invalid),{key,invalid});
      await page.reload({waitUntil:'networkidle'});
      await active(page,'merge');
      assert.equal(await page.evaluate(key=>localStorage.getItem(key),key),'merge');
    }
    await page.evaluate(key=>{localStorage.removeItem(key);localStorage.setItem('dpc-last-tool','funds');},key);
    await page.reload({waitUntil:'networkidle'});
    await active(page,'merge');
    await context.close();

    for (const mode of ['unavailable','read-error','write-error']) {
      const context = await browser.newContext({storageState:{cookies:[],origins:[{origin:new URL(url).origin,localStorage:[{name:key,value:'compare'}]}]}});
      await context.addInitScript(mode=>{
        const denied=()=>{throw new DOMException('Storage disabled','SecurityError');};
        if(mode==='unavailable')Object.defineProperty(window,'localStorage',{get:denied});
        else Object.defineProperty(Storage.prototype,mode==='read-error'?'getItem':'setItem',{value:denied});
      },mode);
      const page = await open(context);
      await active(page,mode==='write-error'?'compare':'merge');
      await select(page,'compare');
      await page.locator('#themeToggle').click();
      assert.equal(await page.locator('html').getAttribute('data-theme'),'dark');
      await select(page,'merge');
      await context.close();
    }
    assert.deepEqual(errors,[]);
    console.log('PASS homepage state: first visit defaults, last tool on refresh/new tab, local and embedded state retained while switching, uploads not persisted by homepage, invalid/legacy values ignored, disabled storage safe');
  } finally { await browser.close(); }
})().catch(error=>{console.error(error);process.exit(1);});
