const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
function file(items){const w=X.utils.book_new();for(const [i,name] of C.SHEETS.entries()){const rows=i?[]:items.map(([code,qty,owner='采购员'])=>C.COLUMNS.map(k=>({'物料编码':code,'未交付数量':qty,'采购下单人':owner,'供应商简称':'供应商','SKU':'SKU-'+code,'物料名称':'物料'+code})[k]??''));X.utils.book_append_sheet(w,X.utils.aoa_to_sheet([C.COLUMNS,...rows]),name);}return Buffer.from(X.write(w,{type:'buffer',bookType:'xlsx'}));}
(async()=>{const browser=await chromium.launch({executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:true});try{
 const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));const url=new URL('fulfillment.html',process.env.QA_URL||'http://127.0.0.1:8939/').href;
 if(process.env.QA_FAIL_SCRIPT){let failed=0;await page.route('**/fulfillment-tool.js*',route=>{if(failed++===0)return route.abort();return route.continue();});}
 await page.goto(url);await page.waitForFunction(()=>!document.getElementById('file').disabled);
 async function upload(buffer,name){await page.locator('#file').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});await page.waitForFunction(()=>!document.getElementById('file').disabled);}
 async function history(){return page.evaluate(()=>FulfillmentCache.get());}
 await upload(file([['001',0.1],['001',0.2],['002',5],['004',10]]),'上一版.xlsx');assert.equal((await history()).previous,null);
 const second=file([['001',0.3],['004',2,'甲'],['004',3,'乙'],['005',4]]);
 await upload(second,'本次更新版.xlsx');let h=await history();assert.equal(h.previous.filename,'上一版.xlsx');assert.match(await page.locator('#history').innerText(),/3 个物料/);
 await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);assert.match(await page.locator('#history').innerText(),/上一版.xlsx/);
 await upload(second,'重复副本.xlsx');assert.deepEqual(await history(),h);
 await upload(file([['001','错误数量']]),'异常.xlsx');assert.match(await page.locator('#status').innerText(),/数量校验未通过/);assert.deepEqual(await history(),h);assert(await page.locator('#download').isDisabled());
 await upload(second,'本次更新版.xlsx');assert.deepEqual(await history(),h);assert(!(await page.locator('#download').isDisabled()));
 const saved=page.waitForEvent('download');await page.locator('#download').click();const download=await saved;assert.equal(download.suggestedFilename(),'本次汇总版.xlsx');
 const output=process.env.QA_OUTPUT||'/private/tmp/fulfillment-history-qa';await fs.mkdir(output,{recursive:true});const out=path.join(output,download.suggestedFilename());await download.saveAs(out);
 const book=X.read(await fs.readFile(out),{type:'buffer',cellStyles:true});assert.deepEqual(book.SheetNames,['汇总表','差异',...C.SHEETS]);const rows=X.utils.sheet_to_json(book.Sheets.差异,{header:1});assert.deepEqual(rows[0],C.DIFF_COLUMNS);assert.deepEqual(rows.slice(1).map(r=>[r[2],r[5]]),[['004',-5],['005',4],['002',-5]]);assert.equal(rows[1][0],'甲 / 乙');assert.equal(book.Sheets.差异.A1.s.fgColor.rgb,'315B9D');
 await page.screenshot({path:path.join(output,'desktop.png'),fullPage:true});await page.setViewportSize({width:390,height:844});await page.screenshot({path:path.join(output,'mobile.png'),fullPage:true});assert.deepEqual(errors,[]);
 if(process.env.QA_FAIL_SCRIPT){
 const blocked=await browser.newPage();await blocked.route('**/fulfillment-core.js*',route=>route.abort());await blocked.goto(url);await blocked.locator('#retry-load').waitFor({state:'visible'});assert.equal(await blocked.locator('#file').isDisabled(),true);assert.match(await blocked.locator('#status').innerText(),/加载失败：fulfillment-core.js/);await blocked.unroute('**/fulfillment-core.js*');await blocked.locator('#retry-load').click();await blocked.waitForFunction(()=>!document.getElementById('file').disabled);await blocked.close();
 }
 // A subsequent successful upload advances exactly one snapshot.
 await upload(file([['005',1]]),'第三版.xlsx');assert.equal((await history()).previous.filename,'本次更新版.xlsx');
 console.log(JSON.stringify({url,sheets:book.SheetNames,differences:rows.slice(1).length,persistence:'passed',duplicate:'passed',failureRollback:'passed',output}));
}finally{await browser.close();}})().catch(e=>{console.error(e);process.exitCode=1;});
