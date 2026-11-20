const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const base=process.env.QA_URL||'http://127.0.0.1:8965/',mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const currentName='订单履约明细跟进表截止2026.11.13-11.19更新版.xlsx',nextName='订单履约明细跟进表截止2026.11.20-11.26更新版.xlsx';
const fullSupplier='嘉兴迈锐医疗器械有限公司',channel='海外事业一部*欧洲部*亚马逊欧盟组';
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
function fulfillment(qty){
 const book=X.utils.book_new(),record={'采购下单人':'徐英','事业部':'海外事业一部','销售分渠道':channel,'供应商简称':'迈锐','系列':'AX25','物料编码':'001','SKU':'AX25-EU','采购订单号':'CGDD000001','下单日期':'2026年11月1日','下单月份':'2026年11月','未交付数量':qty};
 for(const name of C.SHEETS)X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...(name==='徐英'?[C.COLUMNS.map(key=>record[key]??'')]:[])]),name);
 return write(book);
}
function directory(buyer='徐英'){
 const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['采购下单人','正确物料编码','供应商','供应商简称'],[buyer,'001',fullSupplier,'迈锐']]),'产品线明细');return write(book);
}
function purchase(){
 const book=X.utils.book_new(),headers=['采购日期','创建人','事业部','单据编号','供应商','物料编码','采购数量','剩余入库数量','关闭状态'];
 const row=(date,qty,supplier=fullSupplier,order='CGDD015712',remaining=1,closed='未关闭')=>[date,'孙立柱',channel,order,supplier,'001',qty,remaining,closed];
 const relevant=row('2026-11-15',123,fullSupplier,'CGDD099998');relevant[5]='999';
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,row('2026-11-13',150),row('2026-11-19 23:59:59',200,fullSupplier,'CGDD015712',0,'已关闭'),row('2026-11-12',1000),row('2026-11-20',55),row('2026-11-15',999,'其他供应商有限公司','CGDD015799'),relevant]),'采购明细');return write(book);
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const errors=[];
  async function session(){
   const context=await browser.newContext(),page=await context.newPage();page.on('pageerror',error=>errors.push(error.message));let frame;
   async function open(reload=false){
    if(reload)await page.reload();else await page.goto(base);
    await page.locator('button[data-tool="fulfillment"]').click();frame=await(await page.locator('iframe[data-tool="fulfillment"]').elementHandle()).contentFrame();await ready();
   }
   const ready=()=>frame.waitForFunction(()=>window.FulfillmentToolReady&&!document.getElementById('file').disabled&&!document.getElementById('purchase-file').disabled&&!document.getElementById('supplier-division-file').disabled);
   const upload=async(id,name,buffer)=>{await frame.locator('#'+id).setInputFiles({name,mimeType,buffer});await ready();};
   const snapshot=()=>frame.evaluate(async()=>JSON.stringify(await FulfillmentCache.get(),(_,value)=>value instanceof ArrayBuffer?Array.from(new Uint8Array(value)):value));
   const preview=()=>frame.locator('#series-preview').evaluate(el=>({headers:[...el.querySelectorAll('thead th')].map(c=>c.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(c=>c.textContent))}));
   const quantity=async()=>{const view=await preview();assert.equal(view.rows.length,1);return view.rows[0][8];};
   async function download(expected,label){
    const view=await preview();assert.equal(view.headers[7],'差异数量');assert.equal(view.headers[8],label);assert.deepEqual(view.headers.slice(9),[label.replace('新增采购订单','采购入库单数量'),'备注','判断']);assert.equal(view.rows[0][8],String(expected));
    const pending=page.waitForEvent('download');await frame.locator('#download').click();const saved=await pending;assert.match(saved.suggestedFilename(),/_V\d+\.xlsx$/);
    const book=X.read(await fs.readFile(await saved.path()),{type:'buffer',cellNF:true}),sheet=book.Sheets.系列对比,grid=X.utils.sheet_to_json(sheet,{header:1,defval:''});
    assert.deepEqual(grid[0],view.headers);assert.equal(grid.length-1,view.rows.length);
    for(let r=0;r<view.rows.length;r++)for(let c=0;c<view.headers.length;c++)assert.equal(X.utils.format_cell(sheet[X.utils.encode_cell({r:r+1,c})]),view.rows[r][c],'预览与实际下载一致');
    assert.equal(sheet.I2.v,expected);if(typeof expected==='number')assert.equal(sheet.I2.t,'n');assert.equal(sheet.I2.f,undefined);assert.equal(sheet['!autofilter'].ref,'A1:L2');assert.equal(sheet.J2.v,'未上传采购入库单');assert.equal(sheet.K2?.v??'','');
    const expectedBase=await frame.evaluate(async()=>{const h=await FulfillmentCache.get();return FulfillmentCore.seriesComparison(h.previous?.result,h.current.result).map(row=>row.map(c=>c.v));});
    assert.deepEqual(grid[1].slice(0,8),expectedBase[0].slice(0,8));assert.equal(grid[1][10],expectedBase[0][8]);assert.equal(grid[1][11],expectedBase[0][9],'备注与判断后移后原值保留');
    return book;
   }
   await open();return {context,page,get frame(){return frame;},open,ready,upload,snapshot,preview,quantity,download};
  }
  const s=await session(),po=purchase(),catalog=directory();
  await s.upload('file','订单履约明细跟进表截止2026.11.6-11.12更新版.xlsx',fulfillment(8));
  await s.upload('file',currentName,fulfillment(10));assert.equal(await s.quantity(),'未上传采购订单');const stable=await s.snapshot();
  await s.upload('supplier-division-file','供应商分工.xlsx',catalog);
  // Hold the actual file read so a download cannot use the preceding purchase input while replacement is in progress.
  await s.frame.evaluate(()=>{const original=File.prototype.arrayBuffer;File.prototype.arrayBuffer=function(){File.prototype.arrayBuffer=original;return new Promise(resolve=>{window.releasePurchaseRead=()=>resolve(original.call(this));});};});
  await s.frame.locator('#purchase-file').setInputFiles({name:'采购订单.xlsx',mimeType,buffer:po});await s.frame.waitForFunction(()=>typeof window.releasePurchaseRead==='function');
  try{assert(await s.frame.locator('#download').isDisabled());assert.match(await s.frame.locator('#purchase-reference').innerText(),/正在读取/);}finally{await s.frame.evaluate(()=>{window.releasePurchaseRead();delete window.releasePurchaseRead;});}
  await s.ready();assert.equal(await s.quantity(),'350','原采购数量按首尾日期合计，重复订单物料两行均计入');await s.frame.locator('#purchase-issues-details > summary').click();assert.match(await s.frame.locator('#purchase-issues').innerText(),/CGDD099998/);assert.doesNotMatch(await s.frame.locator('#purchase-issues').innerText(),/其他供应商有限公司/);assert.deepEqual(await s.snapshot(),stable,'采购和分工上传不推进履约历史');
  const first=await s.download(350,'11.13-11.19新增采购订单'),reviews=X.utils.sheet_to_json(first.Sheets.合并核对).filter(row=>row.核对类型==='新增采购数量');assert.equal(reviews.length,1);assert(String(reviews[0].原值).includes('CGDD099998'),'相关供应商的未匹配物料保留核对');assert(!reviews.some(row=>String(row.原值).includes('其他供应商有限公司')),'不涉及的供应商不进入合并核对');
  await s.frame.locator('#supplier-division-clear').click();await s.ready();assert.notEqual(await s.quantity(),'350','清除目录后不能保留旧归属结果');assert.deepEqual(await s.snapshot(),stable);
  await s.upload('supplier-division-file','归属改为李奇.xlsx',directory('李奇'));assert.notEqual(await s.quantity(),'350','更换分工必须重新匹配，创建人不替代采购下单人');
  await s.upload('supplier-division-file','供应商分工.xlsx',catalog);assert.equal(await s.quantity(),'350');
  await s.page.locator('button[data-tool="payment"]').click();await s.page.locator('button[data-tool="fulfillment"]').click();assert.equal(await s.quantity(),'350');assert.deepEqual(await s.snapshot(),stable);
  await s.frame.locator('#purchase-clear').click();await s.ready();assert.equal(await s.quantity(),'未上传采购订单');
  await s.upload('purchase-file','采购订单.xlsx',po);assert.equal(await s.quantity(),'350');
  await s.upload('purchase-file','损坏采购订单.xlsx',Buffer.from([80,75,3,4,0,0,0,0]));assert.notEqual(await s.quantity(),'350');assert.match(await s.frame.locator('#purchase-reference').innerText(),/失败/);assert.deepEqual(await s.snapshot(),stable);
  await s.upload('purchase-file','采购订单.xlsx',po);assert.equal(await s.quantity(),'350');
  await s.open(true);assert.equal(await s.quantity(),'未上传采购订单','刷新不把未恢复的采购订单当成已上传');assert.match(await s.frame.locator('#supplier-division-reference').innerText(),/缓存.*恢复/);assert.deepEqual(await s.snapshot(),stable);
  await s.upload('purchase-file','采购订单.xlsx',po);assert.equal(await s.quantity(),'350');
  await s.upload('file',nextName,fulfillment(10));assert.equal(await s.quantity(),'55','仅日期变动也更新列名和期间，不把新周期当旧副本');await s.download(55,'11.20-11.26新增采购订单');
  const s2=await session();await s2.upload('purchase-file','采购订单.xlsx',po);await s2.upload('supplier-division-file','供应商分工.xlsx',catalog);await s2.upload('file',currentName,fulfillment(10));assert.equal(await s2.quantity(),'350','先上传采购订单后上传履约也可统计');await s2.download(350,'11.13-11.19新增采购订单');
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,periodQuantity:350,nextPeriodQuantity:55,duplicateLinesCounted:true,originalQuantityNotRemaining:true,directoryRecalculation:true,uploadOrderIndependent:true,previewDownloadMatch:true,historyUnchangedByReferences:true,refreshRequiresPurchaseReupload:true}));
  await s2.context.close();await s.context.close();
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
