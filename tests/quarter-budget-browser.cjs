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
function manual(year=2026,duplicate=true,extraRows=[]) {
 return write(book([['Sheet1',[manualHeaders,
  ['谈判付款-谈判付款表','李奇',unsafeSupplier,'迈德斯特','FP001','货款甲',100,date(year+'-12-10'),'普通手工付款仅对比'],
  ['正常付款-正常对账表','李奇','供应商乙','迈德斯特','FP003','货款乙',25,date(year+'-12-20'),''],
  ['正常付款-预估付款','李奇','多家供应商','迈德斯特','','合并预估',2000000,date(year+'-12-31'),'手工补充预估'],
  ...(duplicate?[['正常付款-正常对账表','邢芳芳','供应商丙','迈德斯特','','未来货款',10,date(year+'-12-15'),'此款为预估']]:[]),...extraRows
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
const reviewHeaders = ['供应商','跟单员','付款主体','付款月份','分类','金额','OA流程号','付款状态','来源','核对提示','核对处理','付款凭据'];
const comparisonHeaders = ['系统预算金额','手工预算金额','差异金额','核对结果','系统来源','手工来源'];
const comparisonRows = page=>page.locator('#preview tr[data-comparison-key]');
async function combinedDetails(page,expectedTotal) {
 assert.equal(await page.locator('#detail-title').innerText(),'付款明细与下月预算对比');
 assert.equal(await page.locator('#detail-panel #comparison-panel').count(),1);
 assert.equal(await page.locator('#comparison-preview,#comparison-totals table').count(),0,'对比与付款明细共用一张表');
 assert.equal(await page.locator('#preview table').count(),1);
 const headers=await page.locator('#preview thead th').allTextContents();assert.equal(headers.length,18);assert.deepEqual(headers.slice(12),comparisonHeaders);
 const rows=await comparisonRows(page).evaluateAll(elements=>elements.map(el=>[...el.cells].map(cell=>cell.innerText)));
 for(const row of rows){assert.equal(row.length,18);assert.equal(row[5],'','分组不重复展示付款申请金额');assert.match(row[4],/分组对比/);}
 for(const values of await page.locator('#preview tr[data-record-id]').evaluateAll(elements=>elements.map(el=>[...el.cells].map(cell=>cell.innerText)))){
  assert.equal(values.length,18);assert.deepEqual(values.slice(12),['','','','','',''],'明细不重复展示分组对比金额');
 }
 if(expectedTotal){
  const sums=[12,13,14].map(col=>rows.reduce((sum,row)=>sum+Math.round(Number(row[col].replaceAll(',',''))*100),0));
  assert.deepEqual(sums,expectedTotal,'各组金额只计算一次');
 }
 return rows;
}
async function detailScreenshots(page) {
 if(!process.env.QA_OUT)return;
 const dir=path.resolve(process.env.QA_OUT),preview=page.locator('#preview');
 await preview.evaluate(el=>{el.scrollTop=0;el.scrollLeft=0;});
 await page.locator('#detail-panel').screenshot({path:path.join(dir,'quarter-budget-detail-desktop.png')});
 await preview.evaluate(el=>{el.scrollLeft=el.scrollWidth;});
 await page.locator('#detail-panel').screenshot({path:path.join(dir,'quarter-budget-detail-desktop-comparison.png')});
 await page.setViewportSize({width:390,height:844});
 assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'合并表格不能撑宽手机页面');
 assert(await preview.evaluate(el=>el.scrollWidth>el.clientWidth));
 await preview.evaluate(el=>{el.scrollLeft=el.scrollWidth;});
 await page.locator('#detail-panel').screenshot({path:path.join(dir,'quarter-budget-detail-mobile.png')});
 await page.setViewportSize({width:1440,height:1000});await preview.evaluate(el=>{el.scrollLeft=0;});
}
async function reviewRows(page) {
 const table=page.locator('#review-list table');assert.equal(await table.count(),1,'核对区域使用一张表对齐展示');
 assert.deepEqual(await table.locator('thead th').allTextContents(),reviewHeaders);
 const rows=await table.locator('tbody tr[data-record-id]').evaluateAll(elements=>elements.map(el=>({id:el.dataset.recordId,cells:[...el.cells].map(cell=>cell.innerText)})));
 for(const row of rows){
  assert.equal(row.cells.length,reviewHeaders.length);
  const preview=page.locator('#preview tr[data-record-id="'+row.id+'"]');assert.equal(await preview.count(),1);
  const values=await preview.locator(':scope > td').allTextContents();
  for(const [review,detail]of [[0,1],[2,2],[3,0],[4,8],[6,3],[7,6]])assert.equal(row.cells[review],values[detail],'核对表字段必须与同ID明细对齐');
  assert.equal(row.cells[1],values[9].trim() || '未填写');
  if(values[5].trim())assert.equal(row.cells[5],values[5]);
  assert.match(row.cells[8],/\.xlsx/);assert(row.cells[9].trim(),'核对提示不能丢失');
  assert.equal(await table.locator('tr[data-record-id="'+row.id+'"] [data-action=decision]').count(),1);
  assert.equal(await table.locator('tr[data-record-id="'+row.id+'"] [data-action=reason]').count(),1);
 }
 return rows;
}
const duplicateHeaders = ['供应商','跟单员','付款主体','付款月份','金额','来源对照','是否合并','核对结果'];
async function duplicateRows(page) {
 const table=page.locator('#duplicate-review table');assert.equal(await table.count(),1);
 assert.deepEqual(await table.locator('thead th').allTextContents(),duplicateHeaders);
 const rows=await table.locator('tbody tr[data-group-id]').evaluateAll(elements=>elements.map(el=>({id:el.dataset.groupId,cells:[...el.cells].map(cell=>cell.innerText)})));
 for(const [index,row] of rows.entries()){
  assert.equal(row.cells.length,8);
  const group=table.locator('tr[data-group-id]').nth(index);assert.match(await group.locator('td').nth(5).textContent(),/\.xlsx/,'折叠来源保留文件定位');
  assert.equal(await group.locator('[data-action=merge-group]').count(),1);
 }
 assert.equal(await table.locator('[data-action=reason]').count(),0,'分组选择无需逐笔填写原因');
 return rows;
}
async function duplicateDetails(page) {
 return page.locator('#preview tr[data-record-id]').evaluateAll(elements=>elements.filter(el=>el.cells[1].innerText==='供应商丙').map(el=>({id:el.dataset.recordId,source:el.cells[7].innerText,excluded:el.cells[10].innerText})));
}
async function resolveDuplicate(page) {
 const rows=await duplicateRows(page);assert.equal(rows.length,1,'两笔同额疑似重复只展示一行');assert.equal(rows[0].cells[1],'邢芳芳','同组相同跟单员只显示一次');
 const details=await duplicateDetails(page);assert.equal(details.length,2);
 const kept=details.find(row=>!row.source.includes('manualBudget.xlsx'));assert(kept);
 await page.locator('#duplicate-review [data-action=merge-group]').selectOption('keep:'+kept.id);await ready(page);
 for(const row of details){
  const value=row.id===kept.id?'否':'是';
  assert.equal(await page.locator('#preview tr[data-record-id="'+row.id+'"] [data-action=decision]').inputValue(),value);
  assert.equal(await page.locator('#preview tr[data-record-id="'+row.id+'"] > td').nth(10).innerText(),value);
 }
}
async function duplicateChoices(page) {
 const select=page.locator('#duplicate-review [data-action=merge-group]'),rows=await duplicateDetails(page);
 assert.equal(await page.locator('#review-list tr[data-record-id]').filter({hasText:'供应商丙'}).count(),0,'重复问题不再逐笔展示');
 await select.selectOption('separate');await ready(page);
 assert((await duplicateDetails(page)).every(row=>row.excluded==='否'));
 for(const row of rows)assert((await page.locator('#preview tr[data-record-id="'+row.id+'"] [data-action=reason]').inputValue()).trim(),'分别保留选择会记录明确确认原因');
 await combinedDetails(page,[200014500,200013500,1000]);
 await select.selectOption('');assert(await page.locator('#download').isDisabled(),'重置分组选择后重新拦截下载');
 const manual=rows.find(row=>row.source.includes('manualBudget.xlsx'));assert(manual);
 await select.selectOption('keep:'+manual.id);await ready(page);
 for(const row of await duplicateDetails(page))assert.equal(row.excluded,row.id===manual.id?'否':'是','可选择保留任意来源');
 await combinedDetails(page,[200013500,200013500,0]);
 await select.selectOption('');assert(await page.locator('#download').isDisabled());
}
async function reviewScreenshots(page) {
 const scroll=page.locator('#duplicate-review .review-scroll'),select=page.locator('#duplicate-review [data-action=merge-group]');
 await page.setViewportSize({width:390,height:844});
 assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'手机核对表不能撑宽整页');
 assert(await scroll.evaluate(el=>el.scrollWidth>el.clientWidth),'手机核对表可横向滚动');
 await select.scrollIntoViewIfNeeded();const before=await scroll.evaluate(el=>el.scrollLeft);assert(before>0);
 await select.selectOption('separate');await ready(page);
 await page.waitForFunction(({before})=>Math.abs(document.querySelector('#duplicate-review .review-scroll').scrollLeft-before)<2,{before});
 assert((await duplicateDetails(page)).every(row=>row.excluded==='否'));
 await select.selectOption('');assert(await page.locator('#download').isDisabled());
 if(process.env.QA_OUT){
  const dir=path.resolve(process.env.QA_OUT);assert(!dir.startsWith(root+path.sep)&&dir!==root);fs.mkdirSync(dir,{recursive:true});
  await scroll.evaluate(el=>{el.scrollLeft=0;el.scrollTop=0;});await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-duplicate-mobile.png')});
  await select.scrollIntoViewIfNeeded();await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-duplicate-mobile-controls.png')});
 }
 await page.setViewportSize({width:1440,height:1000});await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
 if(process.env.QA_OUT){
  const dir=path.resolve(process.env.QA_OUT);await scroll.evaluate(el=>{el.scrollLeft=0;el.scrollTop=0;});
  await page.locator('#review-panel').screenshot({path:path.join(dir,'quarter-budget-duplicate-desktop.png')});
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
  assert.equal(await page.locator('#detail-title').innerText(),'付款明细');assert.equal(await page.locator('#preview thead th').count(),12);
  await upload(page,'manualBudget');await page.locator('#comparison-panel:not([hidden])').waitFor();
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2026-12','手工文件唯一月份自动成为预算起月');
  assert.equal(await page.locator('#record-count').innerText(),'7','手工普通付款仅对比，两个明确预估加入预算');
  assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#duplicate-review').innerText(),/供应商丙/);
  assert.match(await page.locator('#preview').innerText(),/多家供应商/);assert.doesNotMatch(await page.locator('#preview').innerText(),/普通手工付款仅对比/);
  assert.equal(await page.locator('#preview img').count(),0);
  const unresolvedComparison=await combinedDetails(page,[200014500,200013500,1000]);
  assert.equal(unresolvedComparison.filter(row=>row[1]==='供应商丙').length,1,'同组两笔只有一条对比汇总');
  assert.equal(unresolvedComparison.find(row=>row[1]==='供应商丙')[12],'20.00');
  await duplicateRows(page);await duplicateChoices(page);await reviewScreenshots(page);
  await resolveDuplicate(page);
  const totals=await page.locator('#comparison-totals').innerText();assert.match(totals,/2,000,135\.00/);assert.match(totals,/0\.00/);
  await combinedDetails(page,[200013500,200013500,0]);await detailScreenshots(page);
  assert.match(await comparisonRows(page).first().innerText(),/manualBudget.xlsx/);assert.match(await comparisonRows(page).first().innerText(),/2026年12月/);
  const overview=await page.locator('#monthly-summary').innerText();await page.locator('#filter-month').selectOption('2027-01');
  assert.equal(await page.locator('#visible-count').innerText(),'2');assert.equal(await page.locator('#monthly-summary').innerText(),overview);
  assert.equal(await comparisonRows(page).count(),0,'其他月份只显示付款明细');assert.equal(await page.locator('#comparison-totals').innerText(),totals);
  await page.locator('#reset-filters').click();await page.locator('#filter-excluded').selectOption('是');
  assert.equal(await page.locator('#visible-count').innerText(),'1');assert.equal(await comparisonRows(page).count(),1);
  assert.equal((await comparisonRows(page).locator('td').allTextContents())[12],'10.00','剔除筛选不重算分组金额');assert.equal(await page.locator('#comparison-totals').innerText(),totals);
  await page.locator('#reset-filters').click();
  await page.waitForFunction(async()=>{const c=await QuarterBudgetCache.load();return Object.keys(c.files).length===4&&Object.keys(c.settings.decisions).length===2;});
  const importedAt=await page.evaluate(async()=>(await QuarterBudgetCache.load()).files.manualBudget.importedAt);
  await page.reload();await page.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(page);
  assert.equal(await page.evaluate(async()=>(await QuarterBudgetCache.load()).files.manualBudget.importedAt),importedAt);
  assert.equal(await page.locator('#budgetStartMonth').inputValue(),'2026-12');assert.equal(await page.locator('#record-count').innerText(),'7');
  assert.match(await page.locator('#duplicate-review [data-action=merge-group]').inputValue(),/^keep:/,'刷新恢复分组选择');
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
  assert.equal(await page.locator('#detail-title').innerText(),'付款明细');assert.equal(await page.locator('#preview thead th').count(),12);assert.equal(await comparisonRows(page).count(),0);
  await page.locator('#budgetStartMonth').fill('2026-12');await page.locator('#budgetStartMonth').dispatchEvent('change');await ready(page);

  await upload(page,'ledger',ledger(true),'missing-amount.xlsx');await page.locator('#result:not([hidden])').waitFor();
  assert.match(await page.locator('#reference-ledger').innerText(),/已读取/);assert(await page.locator('#download').isDisabled());
  const missingRows=await reviewRows(page),missingRow=missingRows.find(row=>row.cells[0]==='供应商乙');assert(missingRow);assert.notEqual(missingRow.cells[5],'0.00','空金额不能显示为零');
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

  const groupedContext=await context(browser),grouped=await opened(groupedContext);grouped.on('pageerror',error=>errors.push(error.message));
  await grouped.locator('#budgetStartMonth').fill('2026-12');await grouped.locator('#budgetStartMonth').dispatchEvent('change');
  const multiple=write(book([['谈判付款付款表付款节奏',[scheduleHeaders,
   ['谈判付款-谈判付款表','李奇',unsafeSupplier,'迈德斯特','FP001','货款甲',100,date('2026-12-10'),date('2026-11-30'),'是','否'],
   ['谈判付款-谈判付款表','李奇',unsafeSupplier,'迈德斯特','FP004','货款甲第二笔',150,date('2026-12-20'),date('2026-12-14'),'是','否'],
   ['正常付款-谈判付款表','李奇','供应商甲','迈德斯特','FP002','下一期',50,date('2027-01-10'),date('2026-12-28'),'是','是']
  ]]]));
  await upload(grouped,'negotiated',multiple);await upload(grouped,'ledger');await upload(grouped,'budget');
  await upload(grouped,'manualBudget',manual(2026,false,[
   ['正常付款-正常对账表','李奇','仅手工供应商','迈德斯特','ONLY01','仅在手工基准中',99,date('2026-12-20'),''],
   ['正常付款-正常对账表','李奇','一月手工供应商','迈德斯特','ONLY02','其他月份基准',19,date('2027-01-20'),'']
  ]));await ready(grouped);
  let groupedRows=await combinedDetails(grouped,[200028500,200022400,6100]);
  const multipleRow=groupedRows.find(row=>row[1]===unsafeSupplier);assert.equal(multipleRow[12],'250.00');assert.match(multipleRow[4],/2条明细/);
  assert.equal(groupedRows.filter(row=>row[1]===unsafeSupplier).length,1);
  const order=await grouped.locator('#preview tbody tr').evaluateAll(rows=>rows.map(row=>({summary:row.hasAttribute('data-comparison-key'),supplier:row.cells[1].innerText})));
  const groupedStart=order.findIndex(row=>row.summary&&row.supplier===unsafeSupplier);assert(groupedStart>=0);
  assert.deepEqual(order.slice(groupedStart+1,groupedStart+3),[{summary:false,supplier:unsafeSupplier},{summary:false,supplier:unsafeSupplier}],'同组明细紧随分组行');
  const onlyManual=groupedRows.find(row=>row[1]==='仅手工供应商');assert(onlyManual);assert.deepEqual(onlyManual.slice(12,15),['0.00','99.00','-99.00']);
  assert(!groupedRows.some(row=>row[1]==='一月手工供应商'));assert.equal(await grouped.locator('#visible-count').innerText(),'7','分组行不计入明细条数');
  const groupedTotals=await grouped.locator('#comparison-totals').innerText();
  await grouped.locator('#filter-category').selectOption('正常付款-正常对账表');
  assert.match(await comparisonRows(grouped).allTextContents().then(text=>text.join('\n')),/仅手工供应商/,'分类筛选保留同分类手工独有分组');
  assert.equal(await grouped.locator('#visible-count').innerText(),'1');assert.equal(await grouped.locator('#comparison-totals').innerText(),groupedTotals);
  await grouped.locator('#reset-filters').click();await grouped.locator('#filter-status').selectOption('未支付');
  assert(!(await comparisonRows(grouped).allTextContents()).some(text=>text.includes('仅手工供应商')),'付款状态筛选不显示没有明细的组');
  assert.equal(await grouped.locator('#comparison-totals').innerText(),groupedTotals);
  await grouped.locator('#reset-filters').click();await grouped.locator('#filter-excluded').selectOption('否');
  assert(!(await comparisonRows(grouped).allTextContents()).some(text=>text.includes('仅手工供应商')),'是否剔除筛选不显示没有明细的组');
  await grouped.locator('#reset-filters').click();
  await grouped.locator('#budgetStartMonth').fill('2027-01');await grouped.locator('#budgetStartMonth').dispatchEvent('change');await ready(grouped);
  groupedRows=await combinedDetails(grouped,[5500,1900,3600]);assert(groupedRows.every(row=>row[0]==='2027年1月'));assert(groupedRows.some(row=>row[1]==='一月手工供应商'));
  assert.equal(await grouped.locator('#preview tr[data-record-id]').count(),7,'换比较月份仍保留其他月份的付款明细');
  const groupedDownload=await download(grouped);assert.deepEqual(groupedDownload.book.SheetNames,['付款汇总明细','付款预算汇总','下月预算对比']);
  const exportedDetails=X.utils.sheet_to_json(groupedDownload.book.Sheets.付款汇总明细,{header:1,defval:''});assert.equal(exportedDetails.length,8,'页面分组行不写进付款明细工作表');assert.equal(exportedDetails[0].length,11,'导出保留原有11列，不加入页面操作或分组列');
  await groupedContext.close();assert.deepEqual(errors,[]);

  const duplicateContext=await context(browser),duplicates=await opened(duplicateContext);duplicates.on('pageerror',error=>errors.push(error.message));
  const scheduleCases=[['多来源分期',10],['多来源分期',20],['同来源两笔',5],['同来源两笔',5],['零金额',0]];
  const ledgerCases=[['多来源分期',10],['多来源分期',20],['同来源两笔',5],['零金额',0]];
  await upload(duplicates,'negotiated',write(book([['谈判付款付款表付款节奏',[scheduleHeaders,...scheduleCases.map(([supplier,amount])=>['谈判付款-谈判付款表',supplier==='零金额'?'':'李奇',supplier,'迈德斯特','','同业务分期',amount,date('2026-12-10'),date('2026-11-30'),'是','否'])]]])));
  await upload(duplicates,'ledger',write(book(people.map(person=>['对账明细-'+person,[ledgerHeaders,...(person==='李奇'?ledgerCases.map(([supplier,amount])=>[supplier,'迈德斯特','','同业务分期',amount,date('2026-12-10'),'否',0,'否']):[])]]))));
  await upload(duplicates,'budget',write(book(people.slice(0,4).map(person=>[person,[budgetHeaders,...(person==='邢芳芳'?[[date('2026-12-10'),'多来源分期','迈德斯特','同业务分期',10,'未支付','','']]:[])]]))));
  const multiRows=await duplicateRows(duplicates);assert.equal(multiRows.length,4,'同业务不同金额分成独立核对行');
  const groups=duplicates.locator('#duplicate-review tr[data-group-id]'),supplier=multiRows.filter(row=>row.cells[0]==='多来源分期');
  assert.deepEqual(supplier.map(row=>row.cells[4]).sort(),['10.00','20.00']);assert.equal(multiRows.find(row=>row.cells[0]==='零金额').cells[1],'未填写、李奇','空跟单员明确显示未填写');
  let confirmed=0;
  for(let index=0;index<multiRows.length;index++){
   const row=multiRows[index],choice=groups.nth(index).locator('[data-action=merge-group]');
   const keep=await choice.locator('option').evaluateAll(options=>options.map(option=>option.value).filter(value=>value.startsWith('keep:')));
   if(row.cells[0]==='同来源两笔'){
    assert.equal(keep.length,0,'同来源多笔不能快捷合并');await choice.selectOption('separate');
   }else{
    assert.equal(keep.length,row.cells[0]==='多来源分期'&&row.cells[4]==='10.00'?3:2,'每个可合并来源均可选择保留');
    await choice.selectOption(keep[keep.length-1]);
   }
   confirmed++;if(confirmed<multiRows.length)assert(await duplicates.locator('#download').isDisabled(),'不同金额的分组须分别确认');
  }
  await ready(duplicates);
  const active=await duplicates.locator('#preview tr[data-record-id]').evaluateAll(rows=>rows.filter(row=>row.cells[10].innerText==='否').map(row=>Number(row.cells[5].innerText.replaceAll(',',''))));
  assert.equal(active.reduce((sum,value)=>sum+value,0),45,'按组保留后金额正确，零金额仍保留一条');assert.equal(active.filter(value=>value===0).length,1);
  await groups.first().locator('[data-action=merge-group]').selectOption('');assert(await duplicates.locator('#download').isDisabled());
  await duplicateContext.close();assert.deepEqual(errors,[]);

  const inferredContext=await context(browser),inferred=await opened(inferredContext);inferred.on('pageerror',error=>errors.push(error.message));
  const inferredSupplier='测试三来源同款',inferredSummary='2026.10月货款';
  await upload(inferred,'negotiated');
  await upload(inferred,'ledger',write(book(people.map(person=>['对账明细-'+person,[ledgerHeaders,...(person==='李奇'?[[inferredSupplier,'迈德斯特','',inferredSummary,14892,'','否',0,'否','']]:[])]]))),'missing-month-ledger.xlsx');
  await upload(inferred,'budget',write(book(people.slice(0,4).map(person=>[person,[budgetHeaders,...(person==='邢芳芳'?[[date('2026-12-10'),inferredSupplier+'\r','迈德斯特',inferredSummary,14892,'未支付','','']]:[])]]))),'missing-month-budget.xlsx');
  await upload(inferred,'manualBudget',write(book([['Sheet1',[manualHeaders,['正常付款-预估付款','李奇',inferredSupplier,'迈德斯特','',inferredSummary,14892,date('2026-12-10'),'手工补充预估']]]])),'missing-month-manual.xlsx');
  const inferredGroups=await duplicateRows(inferred);assert.equal(inferredGroups.length,1,'同款唯一明确月份包含缺月份来源，三条只显示一组');
  assert.equal(inferredGroups[0].cells[0],inferredSupplier);assert.equal(inferredGroups[0].cells[1],'李奇、邢芳芳','同组不同跟单员全部显示');assert.match(inferredGroups[0].cells[3],/2026年12月/);
  assert.match(inferredGroups[0].cells.join('\n'),/含月份未填来源/);assert.match(inferredGroups[0].cells[5],/3条来源/);
  const inferredRows=()=>inferred.locator('#preview tr[data-record-id]').evaluateAll((rows,supplier)=>rows.filter(row=>row.cells[1].innerText.trim()===supplier).map(row=>({id:row.dataset.recordId,month:row.cells[0].innerText,amount:row.cells[5].innerText,source:row.cells[7].innerText,excluded:row.cells[10].innerText})),inferredSupplier);
  let triplet=await inferredRows();assert.equal(triplet.length,3);
  const ledgerRow=triplet.find(row=>row.source.includes('missing-month-ledger.xlsx')),budgetRow=triplet.find(row=>row.source.includes('missing-month-budget.xlsx'));assert(ledgerRow&&budgetRow);
  assert.equal(ledgerRow.month,'月份待核对','分组推定月份不能改写源记录月份');
  const inferredSelect=inferred.locator('#duplicate-review [data-action=merge-group]');assert.equal(await inferredSelect.locator('option[value^="keep:"]').count(),3);
  assert.match(await inferredSelect.locator('option').evaluateAll((options,id)=>options.find(option=>option.value==='keep:'+id)?.textContent,ledgerRow.id),/月份待核对/,'选择空月份来源时仍提示原月份未确认');
  if(process.env.QA_OUT){await inferred.setViewportSize({width:1440,height:1000});await inferred.locator('#review-panel').screenshot({path:path.join(process.env.QA_OUT,'quarter-budget-duplicate-missing-month.png')});}
  await inferredSelect.selectOption('keep:'+budgetRow.id);await ready(inferred);triplet=await inferredRows();
  assert.equal(triplet.filter(row=>row.excluded==='否').reduce((sum,row)=>sum+Number(row.amount.replaceAll(',','')),0),14892,'合并三来源只计入指定预算一笔');
  assert(triplet.every(row=>row.excluded===(row.id===budgetRow.id?'否':'是')));assert.equal(triplet.find(row=>row.id===ledgerRow.id).month,'月份待核对');
  await inferred.waitForFunction(async()=>Object.keys((await QuarterBudgetCache.load()).settings.decisions).length===3);
  await inferred.reload();await inferred.waitForFunction(()=>window.QuarterBudgetToolReady);await ready(inferred);
  assert.equal((await duplicateRows(inferred)).length,1);assert.equal(await inferredSelect.inputValue(),'keep:'+budgetRow.id,'刷新恢复同组三来源的合并选择');
  triplet=await inferredRows();assert.equal(triplet.length,3);assert(triplet.every(row=>row.excluded===(row.id===budgetRow.id?'否':'是')));assert.equal(triplet.find(row=>row.id===ledgerRow.id).month,'月份待核对');
  await inferredSelect.selectOption('');assert(await inferred.locator('#download').isDisabled());
  triplet=await inferredRows();assert(triplet.every(row=>row.excluded==='否'),'重置恢复全部三条原记录');
  for(const row of triplet)assert.equal(await inferred.locator('#preview tr[data-record-id="'+row.id+'"] [data-action=decision]').inputValue(),'','重置清除三条逐笔确认');
  await inferred.waitForFunction(async()=>Object.keys((await QuarterBudgetCache.load()).settings.decisions).length===0);
  await inferredContext.close();assert.deepEqual(errors,[]);

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
  await old.evaluate(async()=>{const cached=await QuarterBudgetCache.load(),id=document.querySelector('#preview tbody tr[data-record-id]').dataset.recordId;await QuarterBudgetCache.saveState({...cached.settings,budgetStartMonth:'2000-01',monthEdited:false,decisions:{[id]:{excluded:'是',reason:'旧月份选择'}}});});
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
  console.log(JSON.stringify({status:'PASS',base:appURL,fourSlots:true,manualForecast:true,manualComparison:true,combinedDetails:true,comparisonFilters:true,reviewTable:true,reviewControlsSync:true,reviewMobile:true,duplicateGroups:true,inferredDuplicateMonth:true,duplicateGate:true,missingAmountGate:true,crossYear:true,manualMonthPriority:true,cacheMigration:true,cacheRestore:true,replacement:true,lateRead:true,export:true,version:true,mobile:true,home:true,storageFallback:true,errors}));
 } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
