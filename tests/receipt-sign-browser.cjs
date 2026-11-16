const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),fs=require('node:fs/promises'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js');
(async()=>{const browser=await chromium.launch({headless:true,executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});try{
 const page=await browser.newPage(),base=process.env.QA_URL||'http://127.0.0.1:8940/',out=process.env.QA_OUTPUT||'/private/tmp/receipt-sign-qa';await fs.mkdir(out,{recursive:true});const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
 await page.goto(new URL('receipt-sign.html',base).href);assert(await page.locator('#download').isDisabled());
 await page.locator('#file:not([disabled])').waitFor();
 if(process.env.REAL_RECEIPT){assert.match(base,/^http:\/\/127\.0\.0\.1:/);await page.locator('#file').setInputFiles(process.env.REAL_RECEIPT);}
 else{const book=X.utils.book_new();for(const name of ['入库一','入库二']){const s=X.utils.aoa_to_sheet([['采购入库单'],['单据编号','TEST001'],['供 应 商 :','','测试甲公司'],['物料编码','物料名称','实收数量'],['001','模拟测试产品',0],['002','较长名称'.repeat(8),5000],['合计','',5000],['创建人','原始姓名']]);s['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}}];s['!cols']=[{wch:18},{wch:40},{wch:15}];X.utils.book_append_sheet(book,s,name);}const second=X.utils.book_new();X.utils.book_append_sheet(second,X.utils.aoa_to_sheet([['采购入库单'],['供应商：测试乙公司'],['物料编码','物料名称','实收数量'],['003','测试产品',7],['合计','',7]]),'入库一');await page.locator('#file').setInputFiles([{name:'模拟入库单.XLS',mimeType:'application/vnd.ms-excel',buffer:Buffer.from(X.write(book,{type:'buffer',bookType:'biff8'}))},{name:'第二份.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(X.write(second,{type:'buffer',bookType:'xlsx'}))}]);}
 await page.waitForFunction(()=>!document.getElementById('download').disabled);const pages=Number((await page.locator('#pageInfo').innerText()).match(/\/ (\d+) 页/)[1]);assert.equal(await page.locator('#sheet').count(),0);assert(await page.locator('#previousPage').isDisabled());
 assert.equal(await page.locator('#signature,#clearSignature,#cleanSignature').count(),0);await page.locator('#preview').screenshot({path:out+'/preview.png'});
 const before=await page.locator('#status').innerText();assert.match(before,/签名已添加/);
 await page.locator('#signatureWidth').fill('0');assert(await page.locator('#download').isDisabled());await page.locator('#signatureWidth').fill('35');assert(!(await page.locator('#download').isDisabled()));
 if(!process.env.REAL_RECEIPT){const box=await page.locator('#preview').boundingBox();await page.locator('#preview').click({position:{x:box.width/2,y:85/1684*box.height}});assert.match(await page.locator('#status').innerText(),/遮挡/);await page.locator('#preview').click({position:{x:box.width/2,y:1450/1684*box.height}});assert.match(await page.locator('#status').innerText(),/签名已添加/);}
 const firstPreview=await page.locator('#preview').evaluate(c=>c.toDataURL());
 if(pages>1){
  await page.locator('#nextPage').click();assert.equal(await page.locator('#pageInfo').innerText(),`第 2 / ${pages} 页`);
  if(!process.env.REAL_RECEIPT){assert.match(await page.locator('#pageName').innerText(),/模拟入库单.XLS \/ 入库二/);assert.notEqual(await page.locator('#preview').evaluate(c=>c.toDataURL()),firstPreview);}
  await page.locator('#previousPage').click();assert.equal(await page.locator('#preview').evaluate(c=>c.toDataURL()),firstPreview,'翻页保留每页签名位置');
  for(let i=1;i<pages;i++)await page.locator('#nextPage').click();assert(await page.locator('#nextPage').isDisabled());
  if(!process.env.REAL_RECEIPT)assert.match(await page.locator('#pageName').innerText(),/第二份.xlsx/);
 }
 let downloadCount=0;page.on('download',()=>downloadCount++);
 const event=page.waitForEvent('download');await page.locator('#download').click();const download=await event;await download.saveAs(out+'/签字.pdf');
 const bytes=await fs.readFile(out+'/签字.pdf');const info=await page.evaluate(async bytes=>{const pdf=await PDFLib.PDFDocument.load(new Uint8Array(bytes));return {count:pdf.getPageCount(),sizes:pdf.getPages().map(p=>p.getSize())};},[...bytes]);assert.equal(info.count,pages);assert.equal(downloadCount,1,'只下载一个合并PDF');await page.waitForFunction(()=>!document.getElementById('file').disabled);assert(await page.locator('#nextPage').isDisabled());if(!process.env.REAL_RECEIPT){assert.equal(pages,3);assert.equal(download.suggestedFilename(),'入库单-测试甲公司、测试乙公司-10007.pdf');}assert(info.sizes.every(s=>Math.abs(s.width-595.28)<.1&&Math.abs(s.height-841.89)<.1));
 await page.evaluate(async bytes=>{const pdfjs=await import('./vendor/pdf.min.mjs');pdfjs.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.min.mjs',location.href).href;const pdf=await pdfjs.getDocument({data:new Uint8Array(bytes)}).promise;for(let i=1;i<=pdf.numPages;i++){const p=await pdf.getPage(i),view=p.getViewport({scale:2});const canvas=document.createElement('canvas');canvas.id='pdf-check-'+i;canvas.width=view.width;canvas.height=view.height;document.body.append(canvas);await p.render({canvasContext:canvas.getContext('2d'),viewport:view}).promise;}},[...bytes]);for(let i=1;i<=pages;i++)await page.locator('#pdf-check-'+i).screenshot({path:out+'/pdf-page'+i+'.png'});
 await page.reload();assert(await page.locator('#download').isDisabled());
 const repeatBook=X.utils.book_new();X.utils.book_append_sheet(repeatBook,X.utils.aoa_to_sheet([['采购入库单'],['供应商：','测试甲公司'],['物料编码','物料名称','实收数量'],['004','小数数量',12.5],['合计','',12.5]]),'Sheet1');
 await page.locator('#file:not([disabled])').waitFor();
 await page.locator('#file').setInputFiles({name:'复用签名.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(X.write(repeatBook,{type:'buffer',bookType:'xlsx'}))});
 await page.waitForFunction(()=>!document.getElementById('download').disabled);
 assert.equal(await page.locator('#pageInfo').innerText(),'第 1 / 1 页');assert(await page.locator('#previousPage').isDisabled());assert(await page.locator('#nextPage').isDisabled());
 const restoredDownload=page.waitForEvent('download');await page.locator('#download').click();assert.equal((await restoredDownload).suggestedFilename(),'入库单-测试甲公司-12.5.pdf');

 // The homepage must also handle multiple receipts stacked in a single sheet.
 const host=await browser.newPage();host.on('pageerror',e=>errors.push(e.message));
 await host.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
 await host.goto(base);await host.locator('.tool-tab[data-tool="receipt-sign"]').click();const frame=host.frameLocator('iframe[data-tool="receipt-sign"]');
 const stacked=X.utils.aoa_to_sheet([]),numbers=['MULTI001','MULTI002','MULTI001'],totals=[.1,.2,2.5];stacked['!merges']=[];stacked['!cols']=[{wch:18},{wch:40},{wch:15}];let offset=0,lastTotal;
 for(let i=0;i<numbers.length;i++){
  const rows=[['采购入库单'],['单据编号：',numbers[i]],['供应商：'+(i===1?'乙公司':'甲公司')],['物料编码','物料名称','实收数量'],...Array.from({length:i+1},(_,j)=>['000'+j,'第'+(i+1)+'单产品'+j,j===0?totals[i]:0]),['合计','',totals[i]],['创建人','原创建人'],['审核日期','2026/11/16']];
  X.utils.sheet_add_aoa(stacked,rows,{origin:{r:offset,c:0}});stacked['!merges'].push({s:{r:offset,c:0},e:{r:offset,c:2}});lastTotal=X.utils.encode_cell({r:offset+4+i+1,c:2});offset+=rows.length+i+1;
 }
 const multi=X.utils.book_new();X.utils.book_append_sheet(multi,stacked,'连续入库单');
 const multiBuffer=()=>Buffer.from(X.write(multi,{type:'buffer',bookType:'biff8'})),originalBytes=multiBuffer();
 const uploadMulti=bytes=>frame.locator('#file:not([disabled])').setInputFiles({name:'连续多单.XLS',mimeType:'application/vnd.ms-excel',buffer:bytes});
 await uploadMulti(originalBytes);await frame.locator('#download:not([disabled])').waitFor();
 assert.match(await frame.locator('#status').innerText(),/1 个文件、3 张入库单/);const pictures=[];
 for(let i=0;i<3;i++){
  assert.equal(await frame.locator('#pageInfo').innerText(),`第 ${i+1} / 3 页`);assert.match(await frame.locator('#pageName').innerText(),new RegExp('第'+(i+1)+'张 / 单据编号：'+numbers[i]));
  pictures.push(await frame.locator('#preview').evaluate(c=>c.toDataURL()));
  if(i<2)await frame.locator('#nextPage').click();
 }
 assert.equal(new Set(pictures).size,3,'单据内容独立，重复单号也不去重');assert(await frame.locator('#nextPage').isDisabled());
 await host.locator('.tool-tab[data-tool="funds"]').click();await host.locator('.tool-tab[data-tool="receipt-sign"]').click();assert.equal(await frame.locator('#pageInfo').innerText(),'第 3 / 3 页');
 await frame.locator('#previousPage').click();assert.equal(await frame.locator('#preview').evaluate(c=>c.toDataURL()),pictures[1]);
 const multiEvent=host.waitForEvent('download');await frame.locator('#download').click();const multiDownload=await multiEvent;assert.equal(multiDownload.suggestedFilename(),'入库单-甲公司、乙公司-2.8.pdf');
 await multiDownload.saveAs(out+'/多单签字.pdf');const multiBytes=await fs.readFile(out+'/多单签字.pdf');
 const multiInfo=await frame.locator('body').evaluate(async(_,bytes)=>{const doc=await PDFLib.PDFDocument.load(new Uint8Array(bytes));return {count:doc.getPageCount(),sizes:doc.getPages().map(p=>p.getSize())};},[...multiBytes]);assert.equal(multiInfo.count,3);assert(multiInfo.sizes.every(s=>Math.abs(s.width-595.28)<.1&&Math.abs(s.height-841.89)<.1));
 await frame.locator('body').evaluate(async(_,bytes)=>{const pdfjs=await import('./vendor/pdf.min.mjs');pdfjs.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.min.mjs',location.href).href;const pdf=await pdfjs.getDocument({data:new Uint8Array(bytes)}).promise;for(let i=1;i<=pdf.numPages;i++){const p=await pdf.getPage(i),view=p.getViewport({scale:1.5}),c=document.createElement('canvas');c.id='multi-pdf-'+i;c.width=view.width;c.height=view.height;document.body.append(c);await p.render({canvasContext:c.getContext('2d'),viewport:view}).promise;}},[...multiBytes]);
 for(let i=1;i<=3;i++)await fs.writeFile(out+'/multi-pdf-page'+i+'.png',Buffer.from(await frame.locator('#multi-pdf-'+i).evaluate(c=>c.toDataURL('image/png').split(',')[1]),'base64'));
 // A later invalid receipt cannot leave an earlier signed page downloadable.
 stacked[lastTotal]={t:'n',f:'SUM(C1:C2)'};await uploadMulti(multiBuffer());await frame.locator('#status').filter({hasText:'读取失败'}).waitFor();assert(await frame.locator('#download').isDisabled());assert.match(await frame.locator('#status').innerText(),new RegExp('第3张.*'+lastTotal));assert.equal(await frame.locator('#pageInfo').innerText(),'第 0 / 0 页');
 await uploadMulti(originalBytes);await frame.locator('#download:not([disabled])').waitFor();assert.equal(await frame.locator('#pageInfo').innerText(),'第 1 / 3 页');
 assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',pages,out,filename:download.suggestedFilename()}));
}finally{await browser.close();}})().catch(e=>{console.error(e);process.exitCode=1;});
