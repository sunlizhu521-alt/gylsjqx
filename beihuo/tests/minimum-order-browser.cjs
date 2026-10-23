const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs/promises'),XLSX=require('../vendor/xlsx.full.min.js');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage();
  await page.goto((process.env.QA_URL||'http://127.0.0.1:8940/')+'beihuo/index.html');
  await page.waitForFunction(()=>document.querySelectorAll('[data-upload]').length===4);
  await page.evaluate(async()=>{
   const record=(rows,sheetName)=>{
    const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet(rows),sheetName);
    const bytes=XLSX.write(w,{type:'array',bookType:'xlsx'});
    return {name:'synthetic.xlsx',file:{name:'synthetic.xlsx',arrayBuffer:async()=>bytes}};
   };
   state.demandRows=await buildDemandAllocationRows({
    demand:record([['创建人','数量','识别码','流程号'],['甲',40,'A','OA1'],['甲',60,'A','OA2'],['甲',99,'B','OA3'],['甲',0,'C','OA4'],['甲',10,'D','OA5']],'需求'),
    purchaseDivision:record([['起订量','物料编码','采购订单下单人','供应商简称'],[100,'A','采购甲','供应商甲'],[100,'B','采购甲','供应商甲'],[0,'C','采购甲','供应商甲'],['','D','采购甲','供应商甲']],'产品线明细')
   });
   rebuildFilterOptions();renderDemandDashboard();
  });
  assert.deepEqual(await page.evaluate(()=>state.filteredRows.map(r=>[r.materialCode,r.minimumOrderQuantity,r.minimumOrderStatus])),[['A','100','满足'],['B','100','不满足'],['C','0','满足'],['D','','']]);
  const headers=await page.locator('.detail-table th').allTextContents();
  assert.equal(await page.locator('#detailTableBody tr').nth(1).locator('td').nth(headers.indexOf('起订量是否满足')).innerText(),'不满足');
  const event=page.waitForEvent('download');await page.locator('#downloadDetailButton').click();const dl=await event;
  const w=XLSX.read(await fs.readFile(await dl.path()),{type:'buffer'}),rows=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1,defval:''});
  const col=rows[0].indexOf('起订量是否满足');
  assert.deepEqual(rows.slice(1).map(r=>r[col]),['满足','不满足','满足','']);
  console.log('PASS moved MOQ column, merged quantity, preview and exported status, zero and blank');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
