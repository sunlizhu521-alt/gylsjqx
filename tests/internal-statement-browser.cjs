const fs=require('node:fs/promises'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../internal-statement-core.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const headers=['供应商','采购组织','单据编号','入库日期','物料编码','物料名称','实收数量','含税单价','价税合计','金额','备注'];
function synthetic(){const b=X.utils.book_new(),rows=[];for(let i=0;i<70;i++)rows.push(['合成供应商甲','合成采购组织甲','RK'+i,'2026/6/30','001'+String(i).padStart(4,'0'),'内部交易合成物料'+i,2,0.1695,0.34,0.3,'采购备注 '+i+'；完整保留规格及中文说明']);rows.push([...rows[0]],[...rows[1]]);rows.push(['合成供应商乙','合成采购组织甲','RK72','2026/6/29','00001','另一个供应商',-1,1.23,-1.23,-1,'退货']);rows.push(['合成供应商甲','合成采购组织乙','RK73','2026/6/29','00001','另一个采购组织',10,1.23,12.3,10,'']);X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows]),'Sheet1');return Buffer.from(X.write(b,{type:'buffer',bookType:'xlsx'}));}
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/';if(process.env.QA_SOURCE&&!/^http:\/\/(127\.0\.0\.1|localhost):/.test(base))throw Error('真实文件仅限本地验收');
 const bytes=process.env.QA_SOURCE?await fs.readFile(process.env.QA_SOURCE):synthetic(),expected=C.parse(X.read(bytes),'入库明细.xlsx');
 const browser=await chromium.launch({headless:true,executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1500,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));await page.goto(base,{waitUntil:'networkidle'});
  await page.locator('.tool-tab[data-tool="internal-statement"]').click();const frame=page.frameLocator('iframe[data-tool="internal-statement"]');await frame.locator('#sourceFile:not([disabled])').setInputFiles({name:'入库明细.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
  await frame.locator('#buyers button:not([disabled])').first().waitFor();assert.equal(await frame.locator('#statement option').count(),expected.groups.length);const status=await frame.locator('#status').innerText();assert(status.includes(expected.count+' 条明细'));assert(status.includes(C.amountText(expected.amount)));
  const buyers=[...new Set(expected.groups.map(g=>g.buyer))];assert.equal(await frame.locator('#buyers .buyer').count(),buyers.length);
  const out=process.env.QA_OUT;if(out)await fs.mkdir(out,{recursive:true});
  const result=[];
  for(const [b,buyer]of buyers.entries()){
   const groups=expected.groups.filter(g=>g.buyer===buyer);
   for(const format of ['xlsx','pdf']){
    const event=page.waitForEvent('download',{timeout:120000});await frame.locator('#buyers .buyer').nth(b).locator('button[data-format="'+format+'"]').click();const downloaded=await event;assert.equal(downloaded.suggestedFilename(),C.fileName(buyer,groups)+'.'+format);
    const output=await fs.readFile(await downloaded.path());if(out)await fs.writeFile(out+'/buyer-'+b+'.'+format,output);
    if(format==='xlsx'){
     const saved=X.read(output,{cellStyles:true,cellNF:true});assert.equal(saved.SheetNames.length,groups.length);
     for(const [index,g]of groups.entries()){
      const s=saved.Sheets[saved.SheetNames[index]];assert.equal(s.A1.v,buyer);assert(s.A4.v.includes(g.supplier));assert.equal(s.A5.v,'物料编码');
      for(const [n,item]of g.items.entries()){const r=n+6;assert.equal(s['A'+r].v,item.code);assert.equal(s['A'+r].t,'s');assert.equal(s['A'+r].z,'@');assert.equal(s['B'+r].v,item.name);assert.equal(s['C'+r].v,Number(item.price));assert.equal(s['C'+r].z,'0.00');assert.equal(s['C'+r].w,X.SSF.format('0.00',Number(item.price)));assert.equal(s['D'+r].v,Number(item.quantity));assert.equal(s['D'+r].w,item.quantity);assert.equal(s['E'+r].v,Number(C.amountText(item.amount)));assert.equal(s['F'+r].v,item.note);}
      const t=g.items.length+6;assert.equal(s['E'+t].v,Number(C.amountText(g.amount)));assert.equal(s['E'+t].f,'SUM(E6:E'+(t-1)+')');assert.equal(s['D'+t].v,Number(g.quantity));assert.equal(s['C'+(t+1)].v,'');assert.equal(s['C'+(t+6)].v,'');assert(s['C'+(t+6)].f.includes('COUNT('));assert(s['!merges'].some(m=>m.s.r===0&&m.e.c===5));
     }
    }else{
     const meta=await frame.locator('body').evaluate(async(_,data)=>{const pdf=await PDFLib.PDFDocument.load(Uint8Array.from(data));return {pages:pdf.getPageCount(),sizes:pdf.getPages().map(p=>[p.getWidth(),p.getHeight()])};},[...output]);assert(meta.pages>=groups.length);assert(meta.sizes.every(([w,h])=>Math.abs(w-595.28)<.01&&Math.abs(h-841.89)<.01));result.push({buyer,pages:meta.pages});
    }
   }
  }
  await frame.locator('#statement').selectOption('0');await frame.locator('#balance').fill('0');await frame.locator('#receipts').fill('20');await frame.locator('#afterSales').fill('0');await frame.locator('#returns').fill('1.23');await frame.locator('#note').fill('本次手动备注');await frame.locator('#update').click();assert(!((await frame.locator('#status').innerText()).startsWith('请核实')));
  await page.locator('.tool-tab[data-tool="funds"]').click();await page.locator('.tool-tab[data-tool="internal-statement"]').click();assert.equal(await frame.locator('#note').inputValue(),'本次手动备注');
  const event=page.waitForEvent('download');await frame.locator('#buyers .buyer').first().locator('button[data-format="xlsx"]').click();const written=X.read(await fs.readFile(await(await event).path()));const s=written.Sheets[written.SheetNames[0]],t=expected.groups[0].items.length+6;assert.equal(s['C'+(t+1)].v,0);assert.equal(s['C'+(t+6)].v,Number(C.amountText(BigInt(expected.groups[0].amount)-2123n)));assert.equal(s['A'+(t+11)].v,'备注：本次手动备注');
  if(out){await frame.locator('#preview').screenshot({path:out+'/preview.png'});await frame.locator('#previewSection').scrollIntoViewIfNeeded();await page.screenshot({path:out+'/ui.png'});}
  await frame.locator('#sourceFile').setInputFiles({name:'最新入库.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:synthetic()});await frame.locator('#buyers button:not([disabled])').first().waitFor();assert.equal(await frame.locator('#note').inputValue(),'');assert.equal(await frame.locator('#statement option').count(),3);
  await page.setViewportSize({width:390,height:844});await frame.locator('#sourceFile').scrollIntoViewIfNeeded();
  assert(await frame.locator('body').evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'手机宽度不挤出页面');
  const invalid=X.utils.book_new();X.utils.book_append_sheet(invalid,X.utils.aoa_to_sheet([['无效表头']]),'Sheet1');
  await frame.locator('#sourceFile').setInputFiles({name:'错误文件.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(X.write(invalid,{type:'buffer',bookType:'xlsx'}))});
  await frame.locator('#sourceFile:not([disabled])').waitFor();assert((await frame.locator('#status').innerText()).startsWith('读取失败：'));assert(await frame.locator('#downloadSection').isHidden());
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,count:expected.count,groups:expected.groups.length,result,exports:true,manualFields:true,switching:true,latestUpload:true,mobile:true,invalidUpload:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
