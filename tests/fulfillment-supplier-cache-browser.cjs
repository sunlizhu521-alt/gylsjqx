const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const mimeType='application/octet-stream',url=process.env.QA_URL||'http://127.0.0.1:8940/';
function division(name,type='xlsx'){
 const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['供应商','供应商简称','正确物料编码'],[name,'简称','000123']]),'产品线明细');
 return Buffer.from(X.write(book,{type:'buffer',bookType:type==='xls'?'biff8':'xlsx'}));
}
function fulfillment(){
 const book=X.utils.book_new();for(const name of C.SHEETS){const data=name==='徐英'?[C.COLUMNS.map(key=>({'采购下单人':name,'事业部':'国内事业部','销售分渠道':'国内事业部','供应商简称':'简称','物料编码':'000123','未交付数量':3,'下单月份':'2026年11月'})[key]??'')]:[];X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...data]),name);}
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const context=await browser.newContext(),page=await context.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));let frame;
  async function open(reload=false){
   if(reload)await page.reload();else await page.goto(url);
   await page.locator('button[data-tool="fulfillment"]').click();
   frame=await(await page.locator('iframe[data-tool="fulfillment"]').elementHandle()).contentFrame();
   await frame.waitForFunction(()=>window.FulfillmentToolReady&&!document.getElementById('file').disabled&&!document.getElementById('supplier-division-file').disabled);
  }
  const input=()=>frame.locator('#supplier-division-file'),reference=()=>frame.locator('#supplier-division-reference'),clear=()=>frame.locator('#supplier-division-clear');
  const snapshot=()=>frame.evaluate(async()=>JSON.stringify(await FulfillmentCache.get(),(_,v)=>v instanceof ArrayBuffer?Array.from(new Uint8Array(v)):v));
  const saved=()=>frame.evaluate(async()=>{const f=await FulfillmentCache.getSupplierDivision();return f?{...f,bytes:Array.from(new Uint8Array(f.bytes))}:null;});
  async function upload(name,buffer){await input().setInputFiles({name,mimeType,buffer});await frame.waitForFunction(()=>!document.getElementById('supplier-division-file').disabled);}
  async function exportRows(){const pending=page.waitForEvent('download');await frame.locator('#download').click();const download=await pending,book=X.read(await fs.readFile(await download.path()),{type:'buffer'});return book.SheetNames.map(name=>[name,X.utils.sheet_to_json(book.Sheets[name],{header:1,defval:''})]);}
  await open();assert.equal(await saved(),null);assert.match(await reference().innerText(),/尚未选择/);assert.equal(await clear().isVisible(),false);
  await frame.locator('#file').setInputFiles({name:'履约更新版.xlsx',mimeType,buffer:fulfillment()});await frame.waitForFunction(()=>!document.getElementById('download').disabled);
  const originalHistory=await snapshot(),originalExport=await exportRows(),first=division('供应商甲'),second=division('供应商乙','xls');
  await upload('分工<明细>.xlsx',first);const firstCache=await saved();assert.deepEqual(Buffer.from(firstCache.bytes),first);assert.match(firstCache.savedAt,/^\d{4}-\d{2}-\d{2}T/);assert.match(await reference().innerText(),/当前引用：分工<明细>\.xlsx/);assert.match(await reference().innerText(),/已自动缓存/);assert.equal(await reference().locator('*').count(),0);assert(await clear().isVisible());
  await page.locator('button[data-tool="payment"]').click();await page.locator('button[data-tool="fulfillment"]').click();assert.equal(await input().evaluate(e=>e.files[0].name),'分工<明细>.xlsx');
  await open(true);assert.match(await reference().innerText(),/已从浏览器缓存恢复/);assert.deepEqual(await saved(),firstCache);assert.equal(await input().evaluate(e=>e.files[0].name),'分工<明细>.xlsx');
  await upload('最新分工.xls',second);const secondCache=await saved();assert.deepEqual(Buffer.from(secondCache.bytes),second);assert.equal(secondCache.name,'最新分工.xls');await open(true);assert.deepEqual(await saved(),secondCache);assert.equal(await input().evaluate(e=>e.files[0].name),'最新分工.xls');
  for(const [name,buffer]of[['空白.xlsx',Buffer.alloc(0)],['错误.txt',first],['损坏.xlsx',Buffer.from([80,75,3,4,0,0,0,0])]]){
   await upload(name,buffer);assert.match(await reference().innerText(),/新文件未保存/);assert.match(await reference().innerText(),/继续保留上次成功上传/);assert.deepEqual(await saved(),secondCache);assert.equal(await input().evaluate(e=>e.files[0].name),'最新分工.xls');
  }
  await frame.evaluate(()=>{window.originalSaveDivision=FulfillmentCache.saveSupplierDivision;FulfillmentCache.saveSupplierDivision=async()=>{throw Error('QuotaExceededError');};});
  await upload('缓存失败.xlsx',first);assert.match(await reference().innerText(),/QuotaExceededError/);assert.deepEqual(await saved(),secondCache);assert.equal(await input().evaluate(e=>e.files[0].name),'最新分工.xls');await frame.evaluate(()=>{FulfillmentCache.saveSupplierDivision=window.originalSaveDivision;});
  await frame.evaluate(()=>{window.originalRemoveDivision=FulfillmentCache.removeSupplierDivision;FulfillmentCache.removeSupplierDivision=async()=>{throw Error('清除测试失败');};});
  await clear().click();await frame.waitForFunction(()=>!document.getElementById('supplier-division-file').disabled);assert.match(await reference().innerText(),/缓存清除失败/);assert.deepEqual(await saved(),secondCache);assert(await clear().isVisible());await frame.evaluate(()=>{FulfillmentCache.removeSupplierDivision=window.originalRemoveDivision;});
  assert.deepEqual(await snapshot(),originalHistory);assert.deepEqual(await exportRows(),originalExport,'分工文件不修改履约数量、历史和导出');
  await page.setViewportSize({width:390,height:844});const slots=await frame.locator('.upload-slot').evaluateAll(es=>es.map(e=>({x:e.getBoundingClientRect().x,y:e.getBoundingClientRect().y})));assert.equal(slots.length,3);assert(slots.every(s=>s.y===slots[0].y));assert(slots[2].x>slots[1].x);
  await clear().click();await frame.waitForFunction(()=>!document.getElementById('supplier-division-file').disabled);assert.equal(await saved(),null);assert.equal(await input().evaluate(e=>e.files.length),0);await open(true);assert.match(await reference().innerText(),/尚未选择/);assert.deepEqual(await snapshot(),originalHistory);
  await frame.evaluate(async()=>{await FulfillmentCache.saveSupplierDivision({name:'损坏缓存.xlsx',bytes:new Uint8Array([80,75,3,4,0,0]).buffer,savedAt:new Date().toISOString()});});
  await open(true);assert.match(await reference().innerText(),/缓存恢复失败/);assert(await clear().isVisible());assert.equal(await input().evaluate(e=>e.files.length),0);await clear().click();await frame.waitForFunction(()=>!document.getElementById('supplier-division-file').disabled);assert.equal(await saved(),null);
  await upload('重试成功.xlsx',first);assert.equal((await saved()).name,'重试成功.xlsx');assert.deepEqual(await snapshot(),originalHistory);assert.deepEqual(errors,[]);
  console.log('PASS supplier division upload/cache: exact bytes, refresh, replacement, validation/quota/delete failures, corrupt cache recovery, history isolation, homepage switch, horizontal layout, unchanged export');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
