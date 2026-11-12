const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const root = path.resolve(__dirname,'..'), base = new URL(process.env.QA_URL || 'https://budget-summary.test/');
const headers = ['预估付款月份','供应商名称','付款主体','申请付款金额','是否已支付','分类'];
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function fixture(rows) {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'付款汇总明细');
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['旧汇总'],[999999]]),'付款预算汇总');
  const other = X.utils.aoa_to_sheet([['原始表'],[12],[3]]); other.B2 = {t:'n',f:'A2+A3',v:15,c:[{a:'测试',t:'保留人工批注'}]}; other['!ref']='A1:B3';
  X.utils.book_append_sheet(book,other,'原始记录');
  book.Workbook={Sheets:[{Hidden:0},{Hidden:0},{Hidden:1}]};
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
const rows = [[46327,'供应商甲','主体甲',100,'已支付','正常付款'],[46348,'供应商甲','主体乙',25,'未支付','正常付款'],[46357,'供应商乙','主体甲',50,'待核对','谈判付款表-暂停付款'],['时间待定','','',0,'','正常付款']];
const original = fixture(rows);
async function upload(scope,bytes=original,name='合成预算.xlsx') { await scope.locator('#sourceFile:not([disabled])').setInputFiles({name,mimeType:mime,buffer:bytes}); }
async function ready(scope) { await scope.locator('#download:not([disabled])').waitFor(); }
async function download(page,scope,version,date='2026-11-11') {
  const pending = page.waitForEvent('download'); await scope.locator('#download').click();
  const file = await pending, chunks = [];
  for await (const chunk of await file.createReadStream()) chunks.push(chunk);
  assert.match(file.suggestedFilename(),/^付款预算汇总-\d{4}-\d{2}-\d{2}-V[1-9]\d*\.xlsx$/);
  if(version)assert.equal(file.suggestedFilename(),`付款预算汇总-${date}-V${version}.xlsx`);
  return Buffer.concat(chunks);
}
async function preserved(before,after) {
  const a = await Zip.loadAsync(before), b = await Zip.loadAsync(after);
  for (const name of Object.keys(a.files).filter(n => /^xl\/(worksheets\/sheet[13]\.xml|comments|drawings)/.test(n) && !a.files[n].dir)) {
    assert(b.file(name),name); assert.equal(await b.file(name).async('string'),await a.file(name).async('string'),name+' must stay unchanged');
  }
}
(async () => {
  const browser = await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH || (process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
  try {
    const context = await browser.newContext({acceptDownloads:true,viewport:{width:1440,height:1000}}), page = await context.newPage(), errors=[];
    await context.addInitScript(() => {
      const RealDate = Date;
      window.Date = class extends RealDate {
        constructor(...args) { super(...(args.length ? args : [window.__budgetTestNow || '2026-11-11T02:00:00Z'])); }
        static now() { return new RealDate(window.__budgetTestNow || '2026-11-11T02:00:00Z').getTime(); }
      };
    });
    page.on('pageerror',e=>errors.push(e.message));
    await context.route('**/*',route => {
      const req=route.request(),url=new URL(req.url());
      if(url.origin!==base.origin || !['GET','HEAD'].includes(req.method()))return route.abort();
      if(process.env.QA_URL)return route.continue();
      const relative=decodeURIComponent(url.pathname).slice(1)||'index.html';
      if(relative.split('/').includes('..'))return route.abort();
      const file=path.join(root,relative);
      if(!fs.existsSync(file)||!fs.statSync(file).isFile())return route.fulfill({status:404,body:''});
      return route.fulfill({contentType:file.endsWith('.html')?'text/html; charset=utf-8':file.endsWith('.css')?'text/css':'application/javascript',body:fs.readFileSync(file)});
    });
    await page.goto(new URL('index.html',base).href,{waitUntil:'networkidle'});
    await page.locator('.tool-tab[data-tool="funds-budget-summary"]').click();
    const frame = page.frameLocator('iframe[data-tool="funds-budget-summary"]');
    await upload(frame); await ready(frame);
    assert.match(await frame.locator('#counts').innerText(),/4 条明细，3 个分组/);
    assert.equal(await frame.locator('#preview tr').count(),5);
    assert.match(await frame.locator('#reference').innerText(),/合成预算.xlsx\n引用时间：/);
    const preview=await frame.locator('#preview').innerText();
    assert.match(preview,/175\.00\s+100\.00\s+75\.00/);
    await page.locator('.tool-tab[data-tool="quarter-budget"]').click();
    await page.locator('.tool-tab[data-tool="funds-budget-summary"]').click();
    assert.equal(await frame.locator('#preview').innerText(),preview,'切换工具保留结果');
    const bytes=await download(page,frame,1),book=X.read(bytes,{type:'buffer',cellNF:true});
    assert.deepEqual(book.SheetNames,['付款预算汇总','付款汇总明细','原始记录']);
    assert.equal(book.Sheets.付款预算汇总.A2.t,'s');assert.equal(book.Sheets.付款预算汇总.A2.v,'2026年11月');assert.equal(book.Sheets.付款预算汇总.A2.z,'@');
    assert.equal(book.Sheets.付款预算汇总.D5.v,175);
    assert.equal(book.Sheets.付款预算汇总.E5.v,100);
    assert.equal(book.Sheets.付款预算汇总.F2.f,'D2-E2');
    assert.equal(book.Workbook.Sheets[2].Hidden,1);
    await preserved(original,bytes);
    await upload(frame,bytes,'再次上传.xlsx');await ready(frame);
    assert.equal(await frame.locator('#preview').innerText(),preview,'重新导入不重复累计');
    const again=X.read(await download(page,frame,2),{type:'buffer'});assert.deepEqual(again.SheetNames,book.SheetNames);
    const invalid=fixture([[46327,'供应商','','错误金额','已支付','<img src=x onerror=alert(1)>']]);
    await upload(frame,invalid,'异常.xlsx'); await frame.locator('#status').filter({hasText:'整理未完成'}).waitFor();
    assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#result').isHidden());
    assert.equal(await frame.locator('#issues img').count(),0);assert.match(await frame.locator('#issues').innerText(),/异常.xlsx/);
    await upload(frame,Buffer.from('broken'),'损坏.xlsx');await frame.locator('#status').filter({hasText:'读取失败'}).waitFor();
    assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#result').isHidden());
    await upload(frame);await ready(frame);
    if(process.env.QA_SCREENSHOT)await page.screenshot({path:process.env.QA_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:390,height:844});
    const dimensions=await frame.locator('body').evaluate(body=>({width:body.clientWidth,scroll:body.scrollWidth}));
    assert(dimensions.scroll<=dimensions.width+2,'窄屏只有预览表格横向滚动');
    if(process.env.QA_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.QA_MOBILE_SCREENSHOT,fullPage:true});
    await frame.locator('#clear').click();assert(await frame.locator('#download').isDisabled());
    await page.goto(new URL('funds-budget-summary.html',base).href,{waitUntil:'networkidle'});
    await upload(page);await ready(page);assert.equal(await page.locator('h1').innerText(),'资金预算汇总');
    await download(page,page,3);
    await page.evaluate(() => { window.__realExport=window.FundsBudgetSummaryExport; window.FundsBudgetSummaryExport=async()=>{throw Error('测试生成失败');}; });
    await page.locator('#download').click();await page.locator('#status').filter({hasText:'测试生成失败'}).waitFor();
    assert.equal(await page.evaluate(()=>JSON.parse(localStorage.getItem('funds-budget-summary-export-versions'))['2026-11-11']),3,'生成失败不推进版本');
    await page.evaluate(()=>window.FundsBudgetSummaryExport=window.__realExport);
    await download(page,page,4);
    await page.reload({waitUntil:'networkidle'});await upload(page);await ready(page);await download(page,page,5);
    await page.evaluate(()=>window.__budgetTestNow='2026-11-11T16:00:00Z');
    await download(page,page,1,'2026-11-12');
    await page.evaluate(()=>Object.defineProperty(window,'localStorage',{get(){throw new DOMException('测试缓存不可用','SecurityError');}}));
    await download(page,page,2,'2026-11-12');await download(page,page,3,'2026-11-12');
    assert.match(await page.locator('#status').innerText(),/本页内仍会递增/);
    if(process.env.QA_SAMPLE_FILE){
      await upload(page,fs.readFileSync(process.env.QA_SAMPLE_FILE),path.basename(process.env.QA_SAMPLE_FILE));await ready(page);
      assert.match(await page.locator('#counts').innerText(),/234 条明细，19 个分组/);
      const sample=await download(page,page),s=X.read(sample,{type:'buffer',cellNF:true}).Sheets.付款预算汇总;
      assert.deepEqual([s.D21.v,s.E21.v,s.F21.v],[56608023.24,15233313.66,41374709.58]);
      if(process.env.QA_SAMPLE_OUTPUT)fs.writeFileSync(process.env.QA_SAMPLE_OUTPUT,sample);
    }
    assert.deepEqual(errors,[]);console.log('PASS: 首页/独立页、预览、下载、源表保留、切换、重复上传、错误清理、手机布局、版本递增/刷新/失败不递增/北京时间跨日/缓存不可用'+(process.env.QA_SAMPLE_FILE?'、真实样表金额':''));
  } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
