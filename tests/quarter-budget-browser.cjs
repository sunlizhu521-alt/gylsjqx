const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '..'), vendorRoot = process.env.QA_VENDOR_ROOT || root;
const X = require(path.join(vendorRoot, 'vendor/xlsx.full.min.js'));
const configured = new URL(process.env.QA_URL || 'https://quarter-budget.test/');
if (!configured.pathname.endsWith('/') && !configured.pathname.endsWith('.html')) configured.pathname += '/';
const appURL = new URL('quarter-budget.html', configured).href, origin = new URL(appURL).origin;
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const date = value => (Date.parse(value+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const people = ['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const scheduleHeaders = ['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','付款日期','申请纸质单日期','是否申请纸质单','是否已付款'];
const ledgerHeaders = ['供应商名称','付款主体','对公付款0A流程号','摘要','申请付款金额','应付时间','是否谈判付款','已付款金额','是否结清','预估付款时间'];
const budgetHeaders = ['预估付款月份','供应商名称','付款主体','摘要','申请付款金额','已支付未支付','分类','跟单员'];
const manualHeaders = ['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','付款时间','备注'];
const unsafeSupplier = '<img src=x onerror=alert(1)>';
const write = book => Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
function book(sheets) { const b=X.utils.book_new(); for(const [name,rows] of sheets) X.utils.book_append_sheet(b,X.utils.aoa_to_sheet(rows),name);return b; }
function ledger(missing=false) {
 return write(book(people.map(person=>['对账明细-'+person,[ledgerHeaders,...(person==='李奇'?[
  [unsafeSupplier,'迈德斯特','FP001','货款甲',100,date('2026-12-10'),'否',0,'否'],
  ['供应商乙','迈德斯特','FP003','货款乙',missing?null:25,date('2026-12-20'),'否',0,'否']
 ]:[])]])));
}
function manual(year=2026,duplicate=true) {
 return write(book([['Sheet1',[manualHeaders,
  ['谈判付款-谈判付款表','李奇',unsafeSupplier,'迈德斯特','FP001','货款甲',100,date(year+'-12-10'),'普通手工付款仅对比'],
  ['正常付款-正常对账表','李奇','供应商乙','迈德斯特','FP003','货款乙',25,date(year+'-12-20'),''],
  ['正常付款-预估付款','李奇','多家供应商','迈德斯特','','合并预估',2000000,date(year+'-12-31'),'手工补充预估'],
  ...(duplicate?[['正常付款-正常对账表','邢芳芳','供应商丙','迈德斯特','','未来货款',10,date(year+'-12-15'),'此款为预估']]:[])
 ]]]));
}
const files = {
 negotiated:write(book([['谈判付款付款表付款节奏',[scheduleHeaders,
  ['谈判付款-谈判付款表','李奇',unsafeSupplier,'迈德斯特','FP001','货款甲',100,date('2026-12-10'),date('2026-11-30'),'是','否'],
  ['正常付款-谈判付款表','李奇','供应商甲','迈德斯特','FP002','下一期',50,date('2027-01-10'),date('2026-12-28'),'是','是']
 ]],['谈判付款',[['供应商','申请付款金额','谈判明细'],['旧原表不重复',999,'2026年12月10日-999']]],['付款进度整理',[['辅助表不导入']]]])),
 ledger:ledger(),
 budget:write(book(people.slice(0,4).map(person=>[person,[budgetHeaders,...(person==='邢芳芳'?[
  [date('2026-12-15'),'供应商丙','迈德斯特','未来货款',10,'未支付','',''],
  ['1月','供应商丁','迈德斯特','跨年预估',5,'未支付','','']
 ]:[])]]))),
 manualBudget:manual()
};
async function upload(page,role,bytes=files[role],name=role+'.xlsx') {
 await page.locator('#'+role+'File:not([disabled])').setInputFiles({name,mimeType:mime,buffer:bytes});
 await page.waitForFunction(role=>!document.getElementById('reference-'+role).textContent.includes('正在读取'),role);
}
const ready = page=>page.locator('#download:not([disabled])').waitFor();
async function sources(page) { for(const role of ['negotiated','ledger','budget'])await upload(page,role); }
async function download(page) {
 const pending=page.waitForEvent('download');await page.locator('#download').click();const file=await pending,stream=await file.createReadStream(),chunks=[];
 for await(const chunk of stream)chunks.push(chunk);return {name:file.suggestedFilename(),book:X.read(Buffer.concat(chunks),{type:'buffer',cellNF:true})};
}
const reviewHeaders = ['供应商','付款主体','付款月份','分类','金额','OA流程号','付款状态','来源','核对提示','核对处理','付款凭据'];
async function reviewRows(page) {
 const table=page.locator('#review-list table');assert.equal(await table.count(),1,'核对区域使用一张表对齐展示');
 assert.deepEqual(await table.locator('thead th').allTextContents(),reviewHeaders);
 const rows=await table.locator('tbody tr[data-record-id]').evaluateAll(elements=>elements.map(el=>({id:el.dataset.recordId,cells:[...el.cells].map(cell=>cell.innerText)})));
 for(const row of rows){
  assert.equal(row.cells.length,reviewHeaders.length);
  const preview=page.locator('#preview tr[data-record-id="'+row.id+'"]');assert.equal(await preview.count(),1);
  const values=await preview.locator(':scope > td').allTextContents();
  for(const [review,detail]of [[0,1],[1,2],[2,0],[3,8],[5,3],[6,6]])assert.equal(row.cells[review],values[detail],'核对表字段必须与同ID明细对齐');
  if(values[5].trim())assert.equal(row.cells[4],values[5]);
  assert.match(row.cells[7],/\.xlsx/);assert(row.cells[8].trim(),'核对提示不能丢失');
  assert.equal(await table.locator('tr[data-record-id="'+row.id+'"] [data-action=decision]').count(),1);
  assert.equal(await table.locator('tr[data-record-id="'+row.id+'"] [data-action=reason]').count(),1);
 }
 return rows;
}
async function resolveDuplicate(page) {
 const all=await reviewRows(page),rows=all.filter(row=>row.cells[0]==='供应商丙');assert.equal(rows.length,2);
 const positions=rows.map(row=>all.findIndex(item=>item.id===row.id));assert.equal(positions[1],positions[0]+1,'关联记录连续排列便于对比');
 for(let index=0;index<rows.length;index++){
  const row=rows[index],value=row.cells[7].includes('manualBudget.xlsx')?'是':'否';
  await page.locator('#review-list tr[data-record-id="'+row.id+'"] [data-action=decision]').selectOption(value);
  assert.equal(await page.locator('#preview tr[data-record-id="'+row.id+'"] [data-action=decision]').inputValue(),value,'核对表处理同步付款明细');
  assert.equal(await page.locator('#preview tr[data-record-id="'+row.id+'"] > td').nth(10).innerText(),value);
  if(index===0)assert(await page.locator('#download').isDisabled(),'一组未全部确认前不能下载');
 }
 await ready(page);
}
async function reviewScreenshots(page) {
 const scroll=page.locator('#review-list .review-scroll');
 await page.setViewportSize({width:390,height:844});
 assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'手机核对表不能撑宽整页');
 assert(await scroll.evaluate(el=>el.scrollWidth>el.clientWidth),'手机核对表可横向滚动');
 const id=await page.locator('#review-list tbody tr[data-record-id]').first().getAttribute('data-record-id');
 const decision=page.locator('#review-list tr[data-record-id="'+id+'"] [data-action=decision]');
 await decision.scrollIntoViewIfNeeded();const before=await scroll.evaluate(el=>el.scrollLeft);assert(before>0);
 await decision.selectOption('否');
 await page.waitForFunction(({before})=>Math.abs(document.querySelector('#review-list .review-scroll').scrollLeft-before)<2,{before});
 assert.equal(await page.locator('#preview tr[data-record-id="'+id+'"] [data-action=decision]').inputValue(),'否');
 assert(await page.locator('#download').isDisabled(),'手机只确认一条时继续阻止下载');
 await decision.selectOption('');
 const note=page.locator('#review-list tr[data-record-id="'+id+'"] details').first();await note.locator('summary').click();
 assert(await note.evaluate(el=>el.open));assert.match(await note.locator('p').innerText(),/重复/);
 if(process.env.QA_OUT){
  const dir=path.resolve(process.env.QA_OUT);assert(!dir.startsWith(root+path.sep)&&dir!==root);fs.mkdirSync(dir,{recursive:true});
  await scroll.evaluate(el=>{el.scrollLeft=0;el.scrollTop=0;});await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-review-mobile.png')});
  await decision.scrollIntoViewIfNeeded();await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-review-mobile-controls.png')});
 }
 await page.setViewportSize({width:1440,height:1000});await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
 if(process.env.QA_OUT){
  const dir=path.resolve(process.env.QA_OUT);await scroll.evaluate(el=>{el.scrollLeft=0;el.scrollTop=0;});
  await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-review-desktop.png')});
  await note.scrollIntoViewIfNeeded();await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-review-desktop-details.png')});
 }
}
async function context(browser,storageDisabled=false) {
 const ctx=await browser.newContext({acceptDownloads:true});
 await ctx.route('**/*',route=>{
  const incoming=route.request(),url=new URL(incoming.url());
  if(url.origin!==origin||!['GET','HEAD'].includes(incoming.method()))return route.abort();
  if(process.env.QA_URL)return route.continue();
  const relative=decodeURIComponent(url.pathname).slice(1);if(relative.includes('..'))return route.abort();
  const local=path.join(root,relative),fallback=path.join(vendorRoot,relative),file=fs.existsSync(local)?local:fallback;
  if(!fs.existsSync(file)||!fs.statSync(file).isFile())return route.fulfill({status:404,body:''});
  return route.fulfill({contentType:file.endsWith('.html')?'text/html; charset=utf-8':file.endsWith('.css')?'text/css':'application/javascript',body:fs.readFileSync(file)});
 });
 if(storageDisabled)await ctx.addInitScript(()=>Object.defineProperty(window,'indexedDB',{get(){throw new Error('合成缓存禁用');}}));
 return ctx;
}
async function opened(ctx) { const page=await ctx.newPage();await page.goto(appURL,{waitUntil:'networkidle'});await page.waitForFunction(()=>window.QuarterBudgetToolReady);return page; }
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
 try {
  const ctx=await context(browser),page=await opened(ctx),errors=[];page.on('pageerror',error=>errors.push(error.message));
  assert.equal(await page.locator('input[type=file]').count(),4);assert.equal(await page.locator('#mode,#evidenceFile,#existingFile').count(),0);
  assert(await page.locator('#download').isDisabled());
  await page.evaluate(()=>localStorage.setItem('funds-export-versions','{"2026-11-10":37}'));
  await sources(page);await ready(page);assert.equal(await page.locator('#record-count').innerText(),'5');
  assert.equal(await page.locator('#deleted-count').innerText(),'1');assert.equal(await page.locator('#preview img').count(),0);
  assert(await page.locator('#comparison-panel').isHidden());
  await upload(page,'manualBudget');await page.locator('#comparison-panel:not([hidden])').waitFor();
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2026-12','手工文件唯一月份自动成为预算起月');
  assert.equal(await page.locator('#record-count').innerText(),'7','手工普通付款仅对比，两个明确预估加入预算');
  assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#review-list').innerText(),/供应商丙/);
  assert.match(await page.locator('#preview').innerText(),/多家供应商/);assert.doesNotMatch(await page.locator('#preview').innerText(),/普通手工付款仅对比/);
  assert.equal(await page.locator('#comparison-preview img').count(),0);
  await reviewRows(page);await reviewScreenshots(page);
  await resolveDuplicate(page);
  const totals=await page.locator('#comparison-totals tbody tr').innerText();assert.match(totals,/2,000,135\.00/);assert.match(totals,/0\.00/);
  assert.match(await page.locator('#comparison-preview').innerText(),/manualBudget.xlsx/);assert.match(await page.locator('#comparison-preview').innerText(),/2026年12月/);
  const overview=await page.locator('#monthly-summary').innerText();await page.locator('#filter-month').selectOption('2027-01');
  assert.equal(await page.locator('#visible-count').innerText(),'2');assert.equal(await page.locator('#monthly-summary').innerText(),overview);
  await page.locator('#reset-filters').click();
  await page.waitForFunction(async()=>{const c=await QuarterBudgetCache.load();return Object.keys(c.files).length===4&&Object.keys(c.settings.decisions).length===2;});
  const importedAt=await page.evaluate(async()=>(await QuarterBudgetCache.load()).files.manualBudget.importedAt);
  await page.reload();await page.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(page);
  assert.equal(await page.evaluate(async()=>(await QuarterBudgetCache.load()).files.manualBudget.importedAt),importedAt);
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2026-12');assert.equal(await page.locator('#record-count').innerText(),'7');
  const first=await download(page);assert.match(first.name,/_V1\.xlsx$/);
  assert.deepEqual(first.book.SheetNames,['付款汇总明细','付款预算汇总','下月预算对比']);
  for(const name of first.book.SheetNames)for(const [address,cell]of Object.entries(first.book.Sheets[name]))if(/^[A-Z]+\d+$/.test(address))assert.equal(cell.f,undefined,'导出不保留公式');
  const detail=X.utils.sheet_to_json(first.book.Sheets.付款汇总明细,{header:1,defval:''});
  const added=detail.find(r=>r[1]==='多家供应商');assert.equal(added[5],2000000);assert.equal(added[8],'正常付款-预估付款');assert.match(added[7],/manualBudget.xlsx.*Sheet1.*第4行/);
  assert.equal(first.book.Sheets.付款汇总明细['!autofilter'].ref,first.book.Sheets.付款汇总明细['!ref']);
  assert.equal(JSON.parse(await page.evaluate(()=>localStorage.getItem('funds-export-versions')))['2026-11-10'],37);
  assert.match((await download(page)).name,/_V2\.xlsx$/);

  await page.locator('#budgetStartMonth').fill('');await page.locator('#budgetStartMonth').dispatchEvent('change');
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'');assert(await page.locator('#download').isDisabled());assert(await page.locator('#result').isHidden());assert.match(await page.locator('#status').innerText(),/请选择有效的预算起月/);
  await page.locator('#budgetStartMonth').fill('2026-12');await page.locator('#budgetStartMonth').dispatchEvent('change');assert(await page.locator('#download').isDisabled(),'重新填月份后不沿用旧重复确认');await resolveDuplicate(page);
  await page.locator('#budgetStartMonth').fill('2030-12');await page.locator('#budgetStartMonth').dispatchEvent('change');
  assert(await page.locator('#download').isDisabled(),'修改起月清除重复确认');assert.match(await page.locator('#preview').innerText(),/2031年1月/);
  await page.waitForFunction(async()=>Object.keys((await QuarterBudgetCache.load()).settings.decisions).length===0);
  await upload(page,'manualBudget',manual(2027,false),'changed-month.xlsx');
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2030-12','替换手工文件不覆盖用户明确选月');
  await page.locator('#clear-manualBudget').click();await ready(page);assert(await page.locator('#comparison-panel').isHidden());
  await page.locator('#budgetStartMonth').fill('2026-12');await page.locator('#budgetStartMonth').dispatchEvent('change');await ready(page);

  await upload(page,'ledger',ledger(true),'missing-amount.xlsx');await page.locator('#result:not([hidden])').waitFor();
  assert.match(await page.locator('#reference-ledger').innerText(),/已读取/);assert(await page.locator('#download').isDisabled());
  const missingRows=await reviewRows(page),missingRow=missingRows.find(row=>row.cells[0]==='供应商乙');assert(missingRow);assert.notEqual(missingRow.cells[4],'0.00','空金额不能显示为零');
  const missingId=missingRow.id,missing=page.locator('#review-list tr[data-record-id="'+missingId+'"]');
  await missing.locator('[data-action=decision]').selectOption('是');assert(await page.locator('#download').isDisabled(),'缺金额剔除必须说明原因');
  assert.equal(await page.locator('#preview tr[data-record-id="'+missingId+'"] [data-action=decision]').inputValue(),'是');
  await missing.locator('[data-action=reason]').fill('源表金额待补，本次明确不计');await missing.locator('[data-action=reason]').dispatchEvent('change');await ready(page);
  assert.equal(await page.locator('#preview tr[data-record-id="'+missingId+'"] [data-action=reason]').inputValue(),'源表金额待补，本次明确不计');
  await upload(page,'budget',Buffer.from('invalid excel content'),'broken.xlsx');
  await page.locator('#status').filter({hasText:'读取失败'}).waitFor();assert(await page.locator('#download').isDisabled());assert(await page.locator('#result').isHidden());
  await page.waitForFunction(async()=>{const c=await QuarterBudgetCache.load();return !c.files.budget&&Object.keys(c.settings.decisions).length===0;});
  await upload(page,'budget');await upload(page,'ledger');await ready(page);

  // A late old file must not overwrite the latest file or a month selected while reading.
  await page.evaluate(()=>{const original=File.prototype.arrayBuffer;File.prototype.arrayBuffer=function(){if(this.name==='delayed-old.xlsx')return new Promise(resolve=>{window.__releaseQuarterFile=async()=>resolve(await original.call(this));});return original.call(this);};});
  await page.locator('#manualBudgetFile').setInputFiles({name:'delayed-old.xlsx',mimeType:mime,buffer:manual(2029,false)});
  await page.waitForFunction(()=>typeof window.__releaseQuarterFile==='function');
  await page.locator('#budgetStartMonth').fill('2035-06');await page.locator('#budgetStartMonth').dispatchEvent('change');
  await upload(page,'manualBudget',manual(2026,false),'latest.xlsx');await ready(page);await page.evaluate(()=>window.__releaseQuarterFile());
  await page.waitForFunction(async()=>(await QuarterBudgetCache.load()).files.manualBudget.filename==='latest.xlsx');
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2035-06');assert.match(await page.locator('#reference-manualBudget').innerText(),/latest.xlsx/);
  await page.setViewportSize({width:390,height:844});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
  assert(await page.locator('.uploads-scroll').evaluate(el=>el.scrollWidth>el.clientWidth));
  if(process.env.QA_OUT){const dir=path.resolve(process.env.QA_OUT);assert(!dir.startsWith(root+path.sep)&&dir!==root);fs.mkdirSync(dir,{recursive:true});await page.evaluate(()=>{document.activeElement?.blur();document.querySelectorAll('.preview,.uploads-scroll').forEach(el=>{el.scrollLeft=0;el.scrollTop=0;});scrollTo(0,0);});await page.screenshot({path:path.join(dir,'quarter-budget-mobile.png'),fullPage:true});await page.setViewportSize({width:1440,height:1000});await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));await page.screenshot({path:path.join(dir,'quarter-budget-desktop.png')});await page.locator('#comparison-panel').screenshot({path:path.join(dir,'quarter-budget-comparison.png')});}
  await page.locator('#clear-cache').click();await page.locator('#cache-status').filter({hasText:'已清空'}).waitFor();assert(await page.locator('#download').isDisabled());
  assert.deepEqual(await page.evaluate(async()=>Object.keys((await QuarterBudgetCache.load()).files)),[]);assert.deepEqual(errors,[]);await ctx.close();

  const migration=await context(browser),old=await opened(migration);
  await old.evaluate(async payload=>{
   const entries=[];for(const [role,bytes]of Object.entries(payload)){const buffer=new Uint8Array(bytes).buffer,digest=await crypto.subtle.digest('SHA-256',buffer);entries.push({kind:'file',role,value:{filename:role+'.xlsx',bytes:buffer,fingerprint:[...new Uint8Array(digest)].map(v=>v.toString(16).padStart(2,'0')).join(''),importedAt:'2026-11-01T00:00:00Z'}});}
   await new Promise((resolve,reject)=>{const request=indexedDB.open('gylsjqx-quarter-budget-files',1);request.onerror=()=>reject(request.error);request.onsuccess=()=>{const db=request.result,tx=db.transaction('files','readwrite'),s=tx.objectStore('files');s.clear();for(const e of entries)s.put(e,e.role);s.put({kind:'settings',value:{mode:'existing',decisions:{stale:{excluded:'是',reason:'旧选择'}},matches:{stale:'old'}}},'@settings');tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>reject(tx.error);};});
  },Object.fromEntries(['negotiated','ledger','budget','existing','evidence'].map(role=>[role,[...(files[role]||files.manualBudget)]])));
  await old.reload();await old.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(old);
  const cache=await old.evaluate(async()=>QuarterBudgetCache.load());assert.deepEqual(Object.keys(cache.files).sort(),['budget','ledger','negotiated']);assert.equal(cache.settings.schema,2);assert.deepEqual(cache.settings.decisions,{});assert.deepEqual(cache.settings.matches,{});
  assert(await old.locator('#comparison-panel').isHidden());assert.equal(await old.locator('#record-count').innerText(),'5');
  const keys=await old.evaluate(()=>new Promise((resolve,reject)=>{const r=indexedDB.open('gylsjqx-quarter-budget-files',1);r.onsuccess=()=>{const db=r.result,tx=db.transaction('files'),q=tx.objectStore('files').getAllKeys();q.onsuccess=()=>resolve(q.result);tx.oncomplete=()=>db.close();};r.onerror=()=>reject(r.error);}));
  assert(!keys.includes('existing')&&!keys.includes('evidence'));
  await old.evaluate(async()=>{const cached=await QuarterBudgetCache.load(),id=document.querySelector('#preview tbody tr').dataset.recordId;await QuarterBudgetCache.saveState({...cached.settings,budgetStartMonth:'2000-01',monthEdited:false,decisions:{[id]:{excluded:'是',reason:'旧月份选择'}}});});
  await old.reload();await old.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(old);
  assert.notEqual(await old.locator('#budgetStartMonth').inputValue(),'2000-01');assert.deepEqual(await old.evaluate(async()=>(await QuarterBudgetCache.load()).settings.decisions),{},'自动起月变化后不恢复旧月份的决定');await migration.close();

  const blocked=await context(browser,true),fallback=await opened(blocked);await sources(fallback);await ready(fallback);assert.match(await fallback.locator('#cache-status').innerText(),/缓存未能更新/);await blocked.close();
  const homeContext=await context(browser),home=await homeContext.newPage(),homeErrors=[];home.on('pageerror',error=>homeErrors.push(error.message));
  await home.goto(new URL('index.html',configured).href,{waitUntil:'networkidle'});await home.evaluate(()=>PaymentCache.put('quarter-qa-sentinel',{preserve:true}));
  await home.locator('[data-tool="quarter-budget"].tool-tab').click();let frame=await(await home.locator('iframe[data-tool="quarter-budget"]').elementHandle()).contentFrame();await frame.waitForFunction(()=>window.QuarterBudgetToolReady);
  await sources(frame);await ready(frame);assert.equal(await home.locator('#toolTitle').innerText(),'季度预算整理');
  await frame.waitForFunction(async()=>Object.keys((await QuarterBudgetCache.load()).files).length===3);
  await home.locator('[data-tool="funds-ledger"].tool-tab').click();await home.locator('[data-tool="quarter-budget"].tool-tab').click();assert.equal(await frame.locator('#record-count').innerText(),'5');
  await home.reload();await home.locator('iframe[data-tool="quarter-budget"]').waitFor();frame=await(await home.locator('iframe[data-tool="quarter-budget"]').elementHandle()).contentFrame();await frame.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(frame);assert.equal(await frame.locator('#record-count').innerText(),'5');
  const homeDownload=home.waitForEvent('download');await frame.locator('#download').click();assert.match((await homeDownload).suggestedFilename(),/_V1\.xlsx$/);
  await frame.locator('#clear-cache').click();await frame.locator('#cache-status').filter({hasText:'已清空'}).waitFor();assert.deepEqual(await home.evaluate(()=>PaymentCache.get('quarter-qa-sentinel')),{preserve:true});assert.deepEqual(homeErrors,[]);await homeContext.close();
  console.log(JSON.stringify({status:'PASS',base:appURL,fourSlots:true,manualForecast:true,manualComparison:true,reviewTable:true,reviewControlsSync:true,reviewMobile:true,duplicateGate:true,missingAmountGate:true,crossYear:true,manualMonthPriority:true,cacheMigration:true,cacheRestore:true,replacement:true,lateRead:true,export:true,version:true,mobile:true,home:true,storageFallback:true,errors}));
 } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
