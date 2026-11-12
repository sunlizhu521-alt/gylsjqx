const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '..');
const vendorRoot = process.env.QA_VENDOR_ROOT || root;
const X = require(path.join(vendorRoot, 'vendor/xlsx.full.min.js'));
const configured = new URL(process.env.QA_URL || 'https://quarter-budget.test/');
if (!configured.pathname.endsWith('/') && !configured.pathname.endsWith('.html')) configured.pathname += '/';
const appURL = new URL('quarter-budget.html', configured).href;
const origin = new URL(appURL).origin;
const headers = ['预估付款月份','供应商名称','付款主体','对公付款0A流程号','摘要','申请付款金额','是否已支付','备注','分类','跟单员','是否剔除'];
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const row = (oa, amount, supplier='供应商甲', month=46327, status='未支付') => [month,supplier,'迈德斯特',oa,'货款',amount,status,'','谈判付款','测试跟单员','否'];
function workbook(rows, extraSheet=false) {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'付款汇总明细');
  if (extraSheet) X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['旧备份不读取']]),'备份');
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
const files = {
  negotiated:workbook([row('FP001',100,'<img src=x onerror=alert(1)>'),row('FP002',50,'供应商甲',46357)],true),
  ledger:workbook([row('FP001',100,'<img src=x onerror=alert(1)>'),row('FP003',25,'供应商乙')]),
  budget:workbook([row('FP002',60),row('',10,'供应商丙')]),
  existing:workbook([row('EXT001',12.34,'既有甲',46360,'已支付'),row('EXT002',45.67,'既有乙',46407,'待核对')])
};
async function upload(page,role,bytes=files[role],name=role+'.xlsx') {
  await page.locator('#'+role+'File:not([disabled])').setInputFiles({name,mimeType:mime,buffer:bytes});
}
async function ready(page) { await page.locator('#download:not([disabled])').waitFor(); }
async function download(page) {
  const pending = page.waitForEvent('download');
  await page.locator('#download').click();
  const file = await pending, stream = await file.createReadStream(), chunks = [];
  for await (const chunk of stream) chunks.push(chunk);
  return {name:file.suggestedFilename(),book:X.read(Buffer.concat(chunks),{type:'buffer',cellNF:true})};
}
async function context(browser,storageDisabled=false) {
  const context = await browser.newContext({acceptDownloads:true});
  // Online mode uses a fresh browser context and synthetic files only. Unexpected uploads are blocked.
  await context.route('**/*',request => {
    const incoming = request.request(), url = new URL(incoming.url());
    if (url.origin !== origin || !['GET','HEAD'].includes(incoming.method())) return request.abort();
    if (process.env.QA_URL) return request.continue();
    const relative = decodeURIComponent(url.pathname).slice(1);
    if (relative.includes('..')) return request.abort();
    const local = path.join(root,relative), fallback = path.join(vendorRoot,relative);
    const file = fs.existsSync(local) ? local : fallback;
    if (!fs.existsSync(file) || !fs.statSync(file).isFile()) return request.fulfill({status:404,body:''});
    const contentType = file.endsWith('.html') ? 'text/html; charset=utf-8' : file.endsWith('.css') ? 'text/css' : 'application/javascript';
    return request.fulfill({contentType,body:fs.readFileSync(file)});
  });
  if (storageDisabled) await context.addInitScript(() => Object.defineProperty(window,'indexedDB',{get(){throw new Error('合成缓存禁用');}}));
  return context;
}

(async () => {
  const browser = await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH || (process.platform === 'darwin' ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' : undefined)});
  try {
    const ctx = await context(browser), page = await ctx.newPage(), errors = [];
    page.on('pageerror',error => errors.push(error.message));
    await page.goto(appURL,{waitUntil:'networkidle'});
    await page.waitForFunction(() => window.QuarterBudgetToolReady);
    assert.equal(await page.locator('h1').innerText(),'季度预算整理');
    assert.equal(await page.locator('input[type=file]').count(),5);
    assert(await page.locator('#download').isDisabled());
    await page.evaluate(() => localStorage.setItem('funds-export-versions','{"2026-11-10":37}'));
    for (const role of ['negotiated','ledger','budget']) await upload(page,role);
    await ready(page);
    assert.equal(await page.locator('#record-count').innerText(),'5');
    assert.equal(await page.locator('#deleted-count').innerText(),'1');
    assert.equal(await page.locator('#preview img').count(),0,'外部文本必须转义');
    assert.match(await page.locator('#reference-negotiated').innerText(),/已读取/);
    assert.doesNotMatch(await page.locator('#reference-negotiated').innerText(),/备份/);
    const summary = await page.locator('#monthly-summary').innerText();
    await page.locator('#filter-month').selectOption('2026-12');
    assert.equal(await page.locator('#visible-count').innerText(),'1');
    assert.equal(await page.locator('#monthly-summary').innerText(),summary,'网页筛选不能改变汇总');
    await page.locator('#reset-filters').click();

    const id = await page.locator('#preview [data-action=decision]').first().getAttribute('data-record-id');
    await page.locator('#preview [data-action=decision]').first().selectOption('是');
    await page.waitForFunction(async () => {
      const cache = await QuarterBudgetCache.load();
      return Object.keys(cache.files).length === 3 && Object.keys(cache.settings.decisions).length === 1;
    });
    const importedAt = await page.evaluate(async () => (await QuarterBudgetCache.load()).files.negotiated.importedAt);
    await page.reload(); await page.waitForFunction(() => window.QuarterBudgetToolReady); await ready(page);
    assert.equal(await page.locator('#preview select[data-record-id="'+id+'"]').inputValue(),'是');
    assert.equal(await page.evaluate(async () => (await QuarterBudgetCache.load()).files.negotiated.importedAt),importedAt,'恢复不能重置引用时间');
    await page.evaluate(() => { window.__quarterExport = QuarterBudgetExport; QuarterBudgetExport = async () => { throw new Error('合成导出失败'); }; });
    await page.locator('#download').click(); await page.locator('#status').filter({hasText:'合成导出失败'}).waitFor();
    assert.equal(await page.evaluate(() => localStorage.getItem('quarter-budget-export-versions')),null,'失败不能递增版本');
    await page.evaluate(() => { QuarterBudgetExport = __quarterExport; });
    await page.locator('#filter-month').selectOption('2026-12');
    const exported = await download(page);
    assert.match(exported.name,/^季度预算整理-\d{4}-\d{2}-\d{2}_V1\.xlsx$/);
    assert.deepEqual(exported.book.SheetNames,['付款汇总明细','付款预算汇总']);
    assert.equal(X.utils.decode_range(exported.book.Sheets['付款汇总明细']['!ref']).e.r,5,'网页只显示一行时仍导出全部五行');
    for (const sheet of Object.values(exported.book.Sheets)) for (const [key,cell] of Object.entries(sheet)) if (!key.startsWith('!')) assert.equal(cell.f,undefined,'导出不得含公式');
    const sheet = exported.book.Sheets['付款预算汇总'];
    assert.equal(Object.values(sheet).filter(cell => cell?.v === '付款月份').length,2,'一个汇总sheet中应有两表');
    for (const s of Object.values(exported.book.Sheets)) for (const [address,cell] of Object.entries(s)) if (/^A\d+$/.test(address)) assert.equal(cell.t,'s','导出月份应为展示文本，不参与数值求和');
    assert.equal(Object.values(sheet).filter(cell => cell?.v === '备注').length,2);
    assert.equal(await page.evaluate(() => localStorage.getItem('funds-export-versions')),'{"2026-11-10":37}','不能修改其他工具版本号');
    assert(Object.values(JSON.parse(await page.evaluate(() => localStorage.getItem('quarter-budget-export-versions')))).includes(1));
    await page.locator('#reset-filters').click();

    await upload(page,'budget',Buffer.from('not an Excel workbook'),'broken.xlsx');
    await page.locator('#status').filter({hasText:'读取失败'}).waitFor();
    assert(await page.locator('#download').isDisabled()); assert(await page.locator('#result').isHidden());
    await page.waitForFunction(async () => { const cache = await QuarterBudgetCache.load(); return !cache.files.budget && Object.keys(cache.settings.decisions).length === 0; });
    await page.locator('#clear-budget').click();
    await page.locator('#mode').selectOption('existing'); await upload(page,'existing'); await ready(page);
    assert.equal(await page.locator('#record-count').innerText(),'2');
    const pair = workbook([row('PAY001',100,'证据供应商'),row('PAY001',100,'证据供应商')]);
    await upload(page,'existing',pair,'人工核对.xlsx');
    await upload(page,'evidence',workbook([row('PAY001',100,'证据供应商',46327,'已支付')]),'付款依据.xlsx'); await ready(page);
    const choice = page.locator('#review-list [data-action=evidence]').first();
    const choiceId = await choice.getAttribute('data-record-id');
    const candidateId = await choice.locator('option').nth(1).getAttribute('value');
    await choice.selectOption(candidateId);
    await page.waitForFunction(async id => (await QuarterBudgetCache.load()).settings.matches[id],choiceId);
    await page.reload(); await page.waitForFunction(() => window.QuarterBudgetToolReady); await ready(page);
    assert.equal(await page.evaluate(async id => (await QuarterBudgetCache.load()).settings.matches[id],choiceId),candidateId);
    assert.equal(await page.locator('#preview tr[data-record-id="'+choiceId+'"] td').nth(6).innerText(),'已支付');
    await page.locator('#clear-evidence').click(); await ready(page);
    await page.waitForFunction(async () => Object.keys((await QuarterBudgetCache.load()).settings.matches).length === 0);

    // A slow read from an old selection must never replace a newer file or its cached bytes.
    await page.evaluate(() => {
      const original = File.prototype.arrayBuffer;
      File.prototype.arrayBuffer = function () {
        if (this.name === 'delayed-old.xlsx') return new Promise(resolve => { window.__releaseQuarterFile = async () => resolve(await original.call(this)); });
        return original.call(this);
      };
    });
    await upload(page,'existing',files.existing,'delayed-old.xlsx');
    await page.waitForFunction(() => typeof window.__releaseQuarterFile === 'function');
    await upload(page,'existing',workbook([row('LATEST001',67.89,'新文件')]),'latest.xlsx'); await ready(page);
    await page.evaluate(() => window.__releaseQuarterFile());
    await page.waitForFunction(async () => (await QuarterBudgetCache.load()).files.existing.filename === 'latest.xlsx');
    assert.equal(await page.locator('#record-count').innerText(),'1');
    assert.match(await page.locator('#preview').innerText(),/新文件/);
    await page.setViewportSize({width:390,height:844});
    assert(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1));
    assert(await page.locator('.uploads-scroll').evaluate(element => element.scrollWidth > element.clientWidth));
    if (process.env.QA_OUT) {
      const output = path.resolve(process.env.QA_OUT);
      assert(!output.startsWith(root + path.sep) && output !== root,'验收输出应位于仓库外');
      fs.mkdirSync(output,{recursive:true});
      await page.screenshot({path:path.join(output,'quarter-budget-mobile.png'),fullPage:true});
      await page.setViewportSize({width:1440,height:1000});
      await page.screenshot({path:path.join(output,'quarter-budget-desktop.png'),fullPage:true});
    }
    await page.locator('#clear-cache').click(); await page.locator('#cache-status').filter({hasText:'已清空'}).waitFor();
    assert(await page.locator('#download').isDisabled());
    assert.deepEqual(await page.evaluate(async () => Object.keys((await QuarterBudgetCache.load()).files)),[]);
    assert.deepEqual(errors,[]); await ctx.close();

    const blocked = await context(browser,true), fallback = await blocked.newPage();
    await fallback.goto(appURL,{waitUntil:'networkidle'}); await fallback.waitForFunction(() => window.QuarterBudgetToolReady);
    await fallback.locator('#mode').selectOption('existing'); await upload(fallback,'existing'); await ready(fallback);
    assert.match(await fallback.locator('#cache-status').innerText(),/缓存未能更新/);
    await blocked.close();
    const homeContext = await context(browser), home = await homeContext.newPage(), homeErrors = [];
    home.on('pageerror',error => homeErrors.push(error.message));
    await home.goto(new URL('index.html',configured).href,{waitUntil:'networkidle'});
    await home.evaluate(() => PaymentCache.put('quarter-qa-sentinel',{preserve:true}));
    await home.locator('[data-tool="quarter-budget"].tool-tab').click();
    let frame = await (await home.locator('iframe[data-tool="quarter-budget"]').elementHandle()).contentFrame();
    await frame.waitForFunction(() => window.QuarterBudgetToolReady);
    await frame.locator('#mode').selectOption('existing'); await upload(frame,'existing'); await ready(frame);
    assert.equal(await home.locator('#toolTitle').innerText(),'季度预算整理');
    await frame.waitForFunction(async () => Boolean((await QuarterBudgetCache.load()).files.existing));
    await home.locator('[data-tool="funds-ledger"].tool-tab').click();
    await home.locator('[data-tool="quarter-budget"].tool-tab').click();
    assert.equal(await frame.locator('#record-count').innerText(),'2','切换工具保留明细');
    await home.reload();
    await home.locator('iframe[data-tool="quarter-budget"]').waitFor();
    frame = await (await home.locator('iframe[data-tool="quarter-budget"]').elementHandle()).contentFrame();
    await frame.waitForFunction(() => window.QuarterBudgetToolReady); await ready(frame);
    assert.equal(await frame.locator('#record-count').innerText(),'2','首页刷新恢复文件');
    const homeDownload = home.waitForEvent('download'); await frame.locator('#download').click();
    assert.match((await homeDownload).suggestedFilename(),/^季度预算整理-.*_V1\.xlsx$/);
    await frame.locator('#clear-cache').click(); await frame.locator('#cache-status').filter({hasText:'已清空'}).waitFor();
    assert.deepEqual(await home.evaluate(() => PaymentCache.get('quarter-qa-sentinel')),{preserve:true},'独立清空不能影响付款工具文件');
    assert.deepEqual(homeErrors,[]); await homeContext.close();
    console.log('PASS: quarter budget upload, cache restore, persisted decisions/evidence, export/no formulas/two summaries, stale-read protection, clear/failure invalidation, mobile, unavailable storage.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
