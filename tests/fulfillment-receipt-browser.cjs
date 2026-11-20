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
 const book=X.utils.book_new(),headers=['采购日期','创建人','事业部','单据编号','供应商','物料编码','采购数量','剩余入库数量'];
 const row=(date,qty)=>[date,'孙立柱',channel,'CGDD000001',fullSupplier,'001',qty,999];
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,row('2026-11-13',150),row('2026-11-19',200),row('2026-11-20',55)]),'采购明细');return write(book);
}
function receipt(onlyZero=false){
 const book=X.utils.book_new(),headers=['入库日期','单据类型','单据编号','订单单号','供应商','物料编码','应收数量','实收数量','计价数量','创建人'];
 const row=(date,qty,order='CGDD000001',document='CGRK000001',supplier=fullSupplier)=>[date,'标准采购入库单',document,order,supplier,'001',999,qty,888,'孙立柱'];
 const records=onlyZero?[row('2026-11-13',0)]:[
  row('2026-11-12',1000),row('2026-11-13',40),row('2026-11-19 23:59:59',40),
  row('2026-11-15',0.1),row('2026-11-15',0.2),row('2026-11-15',0.1),row('2026-11-15',0),row('2026-11-20',30.5),
  row('2026-11-15',777,'CGDD099999','CGDD000001','未匹配供应商有限公司')
 ];
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records]),'入库明细');return write(book);
}
function uncertainReceipt(){
 const book=X.utils.book_new(),headers=['入库日期','单据编号','订单单号','供应商','物料编码','实收数量'];
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,['2026-11-13','CGRK-CONFIRMED','CGDD000001',fullSupplier,'001',20],['2026-11-14','CGRK-REVIEW','CGDD000001','供应商冲突','001',30]]),'入库核对');return write(book);
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
   const ready=()=>frame.waitForFunction(()=>window.FulfillmentToolReady&&['file','purchase-file','receipt-file','supplier-division-file'].every(id=>!document.getElementById(id).disabled));
   const upload=async(id,name,buffer)=>{await frame.locator('#'+id).setInputFiles({name,mimeType,buffer});await ready();};
   const snapshot=()=>frame.evaluate(async()=>JSON.stringify(await FulfillmentCache.get(),(_,value)=>value instanceof ArrayBuffer?Array.from(new Uint8Array(value)):value));
   const preview=()=>frame.locator('#series-preview').evaluate(el=>({headers:[...el.querySelectorAll('thead th')].map(c=>c.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(c=>c.textContent))}));
   const quantities=async()=>{const view=await preview();assert.equal(view.rows.length,1);return view.rows[0].slice(8,10);};
   async function download(expected,period='11.13-11.19',receiptNote=''){
    const view=await preview();assert.equal(view.headers[7],'差异数量');assert.deepEqual(view.headers.slice(8),[period+'新增采购订单',period+'采购入库单数量','备注','判断']);assert.deepEqual(view.rows[0].slice(8,10),expected.map(String));
    const pending=page.waitForEvent('download');await frame.locator('#download').click();const saved=await pending;await ready();assert.match(saved.suggestedFilename(),/_V\d+\.xlsx$/);
    const book=X.read(await fs.readFile(await saved.path()),{type:'buffer',cellNF:true}),sheet=book.Sheets.系列对比,grid=X.utils.sheet_to_json(sheet,{header:1,defval:''});
    assert.deepEqual(grid[0],view.headers);assert.equal(grid.length-1,view.rows.length);
    for(let r=0;r<view.rows.length;r++)for(let c=0;c<view.headers.length;c++)assert.equal(X.utils.format_cell(sheet[X.utils.encode_cell({r:r+1,c})]),view.rows[r][c],'预览与实际下载一致');
    for(const [index,col]of ['I','J'].entries()){assert.equal(sheet[col+'2'].v,expected[index]);if(typeof expected[index]==='number')assert.equal(sheet[col+'2'].t,'n');}
    assert.equal(sheet['!autofilter'].ref,'A1:L2');assert.equal(sheet.K2?.v??'','');
    for(const data of Object.values(book.Sheets))for(const [address,cell]of Object.entries(data))if(/^[A-Z]+\d+$/.test(address))assert.equal(cell.f,undefined,'下载工作簿不保留公式');
    const original=await frame.evaluate(async()=>{const h=await FulfillmentCache.get();return FulfillmentCore.seriesComparison(h.previous?.result,h.current.result).map(row=>row.map(c=>c.v));});
    assert.deepEqual(grid[1].slice(0,8),original[0].slice(0,8));assert.equal(grid[1][10],original[0][8],'备注原样保留');
    if(receiptNote){assert(String(grid[1][11]).startsWith(original[0][9]));assert(String(grid[1][11]).includes(receiptNote));assert.equal((String(grid[1][11]).match(/入库核对：/g)||[]).length,1);}
    else assert.equal(grid[1][11],original[0][9],'没有入库待核对时保持原判断');
    return {book,name:saved.suggestedFilename()};
   }
   await open();return {context,page,get frame(){return frame;},open,ready,upload,snapshot,preview,quantities,download};
  }
  const s=await session(),po=purchase(),received=receipt(),catalog=directory();
  assert.equal(await s.frame.locator('.upload-slot').count(),4);
  await s.upload('file','订单履约明细跟进表截止2026.11.6-11.12更新版.xlsx',fulfillment(8));
  await s.upload('file',currentName,fulfillment(10));assert.deepEqual(await s.quantities(),['未上传采购订单','未上传采购入库单']);const stable=await s.snapshot();
  await s.upload('supplier-division-file','供应商分工.xlsx',catalog);await s.upload('purchase-file','采购订单.xlsx',po);
  await s.frame.evaluate(()=>{const original=File.prototype.arrayBuffer;File.prototype.arrayBuffer=function(){File.prototype.arrayBuffer=original;return new Promise(resolve=>{window.releaseReceiptRead=()=>resolve(original.call(this));});};});
  await s.frame.locator('#receipt-file').setInputFiles({name:'采购入库单.xlsx',mimeType,buffer:received});await s.frame.waitForFunction(()=>typeof window.releaseReceiptRead==='function');
  try{assert(await s.frame.locator('#download').isDisabled());assert(await s.frame.locator('#purchase-file').isDisabled());assert.match(await s.frame.locator('#receipt-reference').innerText(),/正在读取/);}finally{await s.frame.evaluate(()=>{window.releaseReceiptRead();delete window.releaseReceiptRead;});}
  await s.ready();assert.deepEqual(await s.quantities(),['350','80.4'],'采购与实收入库各自合计；首尾包含、同单多行和小数不丢失');assert.deepEqual(await s.snapshot(),stable);
  await s.frame.locator('#receipt-issues-details > summary').click();assert.match(await s.frame.locator('#receipt-issues').innerText(),/CGDD099999/,'入库单据号不冒充采购订单号');
  const first=await s.download([350,80.4]);assert(X.utils.sheet_to_json(first.book.Sheets.合并核对).some(row=>String(row.来源工作表).includes('入库')&&String(row.原值).includes('CGDD099999')),'未归属入库来源进入合并核对');
  const second=await s.download([350,80.4]);assert.equal(Number(/_V(\d+)\.xlsx$/.exec(second.name)[1]),Number(/_V(\d+)\.xlsx$/.exec(first.name)[1])+1);
  await s.upload('receipt-file','待核对入库.xlsx',uncertainReceipt());assert.deepEqual(await s.quantities(),['350','待核对']);
  const pendingJudgment=(await s.preview()).rows[0][11];
  for(const text of ['已确认入库数量：20','待核对候选数量：30','实收数量：30','CGRK-REVIEW','CGDD000001','入库核对!F3','冲突','不能计作本组入库'])assert(pendingJudgment.includes(text),text);
  await s.download([350,'待核对'],'11.13-11.19','待核对候选数量：30');
  await s.upload('supplier-division-file','供应商分工.xlsx',catalog);assert.equal((await s.preview()).rows[0][11],pendingJudgment,'重新渲染不重复追加判断');assert.deepEqual(await s.snapshot(),stable);
  await s.upload('receipt-file','采购入库单.xlsx',received);assert.deepEqual(await s.quantities(),['350','80.4']);assert(!(await s.preview()).rows[0][11].includes('入库核对：'),'替换来源后清除旧核对说明');
  await s.page.locator('button[data-tool="payment"]').click();await s.page.locator('button[data-tool="fulfillment"]').click();assert.deepEqual(await s.quantities(),['350','80.4'],'切换工具保留入库和采购来源');
  await s.frame.locator('#receipt-clear').click();await s.ready();assert.deepEqual(await s.quantities(),['350','未上传采购入库单']);
  await s.upload('receipt-file','采购入库单.xlsx',received);assert.deepEqual(await s.quantities(),['350','80.4']);
  await s.upload('receipt-file','损坏入库单.xlsx',Buffer.from([80,75,3,4,0,0,0,0]));assert.notEqual((await s.quantities())[1],'80.4');assert.match(await s.frame.locator('#receipt-reference').innerText(),/失败/);assert.equal((await s.quantities())[0],'350');assert.deepEqual(await s.snapshot(),stable);
  await s.upload('receipt-file','零数量入库单.xlsx',receipt(true));assert.deepEqual(await s.quantities(),['350','0']);await s.download([350,0]);
  await s.upload('receipt-file','采购入库单.xlsx',received);assert.deepEqual(await s.quantities(),['350','80.4']);
  await s.upload('supplier-division-file','归属改为李奇.xlsx',directory('李奇'));assert.notEqual((await s.quantities())[1],'80.4','分工替换后不能沿用旧入库归属');
  await s.upload('supplier-division-file','供应商分工.xlsx',catalog);assert.deepEqual(await s.quantities(),['350','80.4']);assert.deepEqual(await s.snapshot(),stable);
  await s.open(true);assert.deepEqual(await s.quantities(),['未上传采购订单','未上传采购入库单'],'刷新后两项临时来源均需重新选择');assert.deepEqual(await s.snapshot(),stable);
  await s.upload('receipt-file','采购入库单.xlsx',received);assert.deepEqual(await s.quantities(),['未上传采购订单','80.4'],'入库汇总不依赖是否上传采购订单');
  await s.upload('purchase-file','采购订单.xlsx',po);await s.upload('file',nextName,fulfillment(10));assert.deepEqual(await s.quantities(),['55','30.5']);await s.download([55,30.5],'11.20-11.26');
  const after=JSON.parse(await s.snapshot()),before=JSON.parse(stable);assert.deepEqual(after.previous,before.previous,'更换统计期间不变更固定上期');
  const s2=await session();await s2.upload('receipt-file','采购入库单.xlsx',received);await s2.upload('supplier-division-file','供应商分工.xlsx',catalog);await s2.upload('file',currentName,fulfillment(10));assert.deepEqual(await s2.quantities(),['未上传采购订单','80.4'],'先上传入库单后上传履约也可统计');await s2.download(['未上传采购订单',80.4]);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,periodPurchase:350,periodReceipt:80.4,nextPeriodReceipt:30.5,actualReceiptQuantity:true,sourceOrderNotReceiptNumber:true,independentInputs:true,previewDownloadMatch:true,columns:12,historyPreserved:true,refreshRequiresReferenceReupload:true}));
  await s2.context.close();await s.context.close();
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
