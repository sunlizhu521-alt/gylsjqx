const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const url=process.env.QA_URL||'http://127.0.0.1:8964/',mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function fixture(quantity){
 const book=X.utils.book_new();
 for(const [index,name]of C.SHEETS.entries()){
  const record={'采购下单人':name,'事业部':'国内事业部','销售分渠道':'国内事业部','供应商简称':'状态合成供应商','物料编码':'STATUS-'+index,'物料名称':'合成物料','下单月份':'2026年11月','未交付数量':index?1:quantity,'已下单未备料未生产':0,'已备料未生产':0,'生产中产品':0,'完工未发产品':0,'已发货数量':0};
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,C.COLUMNS.map(key=>record[key]??'')]),name);
 }
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const context=await browser.newContext(),page=await context.newPage(),errors=[];page.on('pageerror',error=>errors.push(error.message));let frame;
  async function open(reload=false){
   if(reload)await page.reload();else await page.goto(url);
   await page.locator('button[data-tool="fulfillment"]').click();
   frame=await(await page.locator('iframe[data-tool="fulfillment"]').elementHandle()).contentFrame();
   await ready();
  }
  const ready=()=>frame.waitForFunction(()=>window.FulfillmentToolReady&&!document.getElementById('file').disabled&&!document.getElementById('file').value);
  const reference=()=>frame.locator('#file-reference');
  const history=()=>frame.evaluate(async()=>JSON.parse(JSON.stringify(await FulfillmentCache.get(),(_,value)=>value instanceof ArrayBuffer?Array.from(new Uint8Array(value)):value)));
  async function upload(name,buffer){await frame.locator('#file').setInputFiles({name,mimeType,buffer});await ready();}
  async function complete(name){
   assert((await reference().innerText()).includes(name),'上传槽应显示完整文件名');
   assert.match(await reference().innerText(),/数量.*校验.*通过/);
   assert.equal(await frame.locator('#download').isDisabled(),false);
   assert.equal(await reference().getAttribute('role'),'status');
   assert.equal(await reference().locator('*').count(),0,'文件名按原文显示，不解析为HTML');
  }
  await open();assert.equal(await history(),null);assert.match(await reference().innerText(),/尚未|等待.*上传|请选择/);
  assert.equal(await frame.locator('.upload-slot').first().locator('#file-reference').count(),1,'状态在上传槽内');
  assert.equal(await frame.locator('#download').isDisabled(),true);
  const firstName='履约<状态&验收>更新版.xlsx',first=fixture(3),secondName='订单履约明细跟进表截止2026.11.13-11.19更新版.xlsx',second=fixture(5);
  const chooser=page.waitForEvent('filechooser');await frame.locator('#choose-file').click();await(await chooser).setFiles({name:firstName,mimeType,buffer:first});
  await ready();await complete(firstName);const initial=await history();assert.equal(initial.previous,null);
  await frame.evaluate(()=>{const original=FulfillmentValues.prepare;window.statusPrepareCalls=0;window.statusOriginalPrepare=original;FulfillmentValues.prepare=(...args)=>{window.statusPrepareCalls++;return original(...args);};});
  await upload(firstName,first);await complete(firstName);assert((await frame.evaluate(()=>window.statusPrepareCalls))>0,'同一文件再次选择会重新处理');assert.equal((await history()).previous,null,'重复上传不推进基准');
  await frame.evaluate(()=>{FulfillmentValues.prepare=window.statusOriginalPrepare;delete window.statusPrepareCalls;delete window.statusOriginalPrepare;});
  await frame.evaluate(()=>{
   const original=FulfillmentValues.prepare;
   FulfillmentValues.prepare=(...args)=>{FulfillmentValues.prepare=original;return new Promise(resolve=>{window.releaseStatusPrepare=()=>resolve(original(...args));});};
  });
  await frame.locator('#file').setInputFiles({name:secondName,mimeType,buffer:second});
  await frame.waitForFunction(()=>typeof window.releaseStatusPrepare==='function');
  try{
   assert((await reference().innerText()).includes(secondName),'处理中展示本次选中的文件');
   assert.match(await reference().innerText(),/正在|处理中/);
   assert.equal(await frame.locator('#file').isDisabled(),true);assert.equal(await frame.locator('#choose-file').isDisabled(),true);assert.equal(await frame.locator('#download').isDisabled(),true);
  }finally{await frame.evaluate(()=>{window.releaseStatusPrepare();delete window.releaseStatusPrepare;});}
  await ready();await complete(secondName);const stable=await history();assert.equal(stable.previous.filename,firstName);
  for(const [name,buffer,expected]of[
   ['数量异常.xlsx',fixture('100或200'),/校验未通过|校验失败/],
   ['损坏文件.xlsx',Buffer.from([80,75,3,4,0,0,0,0]),/处理失败|读取失败/]
  ]){
   await upload(name,buffer);assert((await reference().innerText()).includes(name));assert.match(await reference().innerText(),expected);
   assert.equal(await frame.locator('#download').isDisabled(),true,'异常文件不得沿用成功下载状态');assert.deepEqual(await history(),stable,'异常文件不覆盖已保存的两期');
  }
  await upload(secondName,second);await complete(secondName);const beforeSwitch=await history(),shown=await reference().innerText();
  await page.locator('button[data-tool="payment"]').click();await page.locator('button[data-tool="fulfillment"]').click();
  assert.equal(await reference().innerText(),shown,'切换工具保留当前文件状态');assert.deepEqual(await history(),beforeSwitch);
  await open(true);await complete(secondName);assert.match(await reference().innerText(),/缓存.*恢复|恢复.*缓存/);assert.deepEqual(await history(),beforeSwitch,'刷新不推进固定基准');
  const pending=page.waitForEvent('download');await frame.locator('#download').click();const download=await pending;
  assert.match(download.suggestedFilename(),/_V\d+\.xlsx$/);
  const saved=X.read(await fs.readFile(await download.path()),{type:'buffer'}),summary=X.utils.sheet_to_json(saved.Sheets.汇总表);
  assert.equal(summary.length,C.SHEETS.length);assert.equal(summary.find(row=>row.物料编码==='STATUS-0').未交付数量,5);
  assert.deepEqual(await history(),beforeSwitch,'下载不修改两期数据');assert((await reference().innerText()).includes(secondName),'下载后仍显示当前源文件');
  await frame.locator('.upload-slot').first().screenshot({path:'/tmp/fulfillment-upload-status-desktop.png'});
  await page.setViewportSize({width:390,height:844});await reference().scrollIntoViewIfNeeded();assert(await reference().isVisible());
  const layout=await reference().evaluate(element=>({height:element.clientHeight,scroll:element.scrollHeight,whiteSpace:getComputedStyle(element).whiteSpace,overflowWrap:getComputedStyle(element).overflowWrap}));
  assert(layout.height>0);assert(layout.scroll<=layout.height+1,'手机状态完整展示，不裁切说明');assert.notEqual(layout.whiteSpace,'nowrap');assert.notEqual(layout.overflowWrap,'normal');
  await frame.locator('.upload-slot').first().screenshot({path:'/tmp/fulfillment-upload-status-mobile.png'});
  assert.deepEqual(errors,[]);
  console.log(JSON.stringify({status:'PASS',url,uploadState:'empty, processing, success and failure',sameFileRetry:true,invalidFilesPreserveHistory:true,toolSwitch:true,refreshRestoresCurrent:true,downloadVerified:true,mobileStatusVisible:true}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
