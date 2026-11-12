const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs/promises');
(async()=>{const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});try{
 const page=await browser.newPage();await page.goto((process.env.QA_URL||'http://127.0.0.1:8940/')+'beihuo/index.html');await page.waitForFunction(()=>typeof renderDemandDashboard==='function');
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.evaluate(()=>{state.demandRows=[{buyer:'甲/乙',materialCode:'A',quantityNumber:0,stockReason:'原因A',deliveryDays:2},{buyer:'甲/乙',materialCode:'B',quantityNumber:3,stockReason:'原因B'},{buyer:'甲:乙',materialCode:'C',quantityNumber:4,stockReason:'原因C'},{buyer:'',materialCode:'D',quantityNumber:5,stockReason:'原因D'}];state.demandRows.forEach((row,i)=>row.materialCodeValid=i<2?'是':'否');rebuildFilterOptions();renderDemandDashboard();});
 assert.equal(await page.locator('.detail-table th').count(),19);assert(!(await page.locator('.detail-table').innerText()).includes('原因A'));
 const event=page.waitForEvent('download');await page.locator('#downloadDetailButton').click();const download=await event;assert(download.suggestedFilename().endsWith('.zip'));
 const bytes=await fs.readFile(await download.path());
 const files=await page.evaluate(async bytes=>{const z=await JSZip.loadAsync(new Uint8Array(bytes));return Promise.all(Object.values(z.files).filter(f=>!f.dir).map(async f=>{const b=await f.async('uint8array'),w=XLSX.read(b,{type:'array',cellNF:true}),s=w.Sheets[w.SheetNames[0]];const inner=await JSZip.loadAsync(b);return {name:f.name,rows:XLSX.utils.sheet_to_json(s,{header:1,defval:''}),highlight:w.Custprops.KSOReadingLayout,formula:s.N2.f,dateFormat:s.N2.z,dateType:s.N2.t,todayType:s.L2.t,styles:await inner.file('xl/styles.xml').async('string'),xml:await inner.file('xl/worksheets/sheet1.xml').async('string')};}));},[...bytes]);
 assert.equal(files.length,3);assert.equal(new Set(files.map(f=>f.name)).size,3);
 assert.deepEqual(files.map(f=>f.rows.slice(1).map(r=>r[6])),[['A','B'],['C'],['D']]);
 assert.deepEqual(files.map(f=>f.rows.slice(1).map(r=>r[18])),[['原因A','原因B'],['原因C'],['原因D']]);
 for(const f of files){assert.equal(f.rows[0][18],'备货原因');assert.equal(f.rows[0][1],'是否新加坡公司下单');assert(f.rows.slice(1).every(row=>row[1]===''));assert(!f.rows[0].includes('采购分工明细是否存在'));assert.equal(f.highlight,true);assert.equal(f.formula,'IF(M2="","",TEXT(DATE(VALUE(LEFT(L2,4)),VALUE(MID(L2,6,2)),VALUE(RIGHT(L2,2)))+M2,"yyyy-mm-dd"))');assert.equal(f.dateFormat,'@');assert.equal(f.dateType,'s');assert.equal(f.todayType,'s');assert.match(f.styles,/FF1F4E78/);assert.equal((f.styles.match(/style="thin"><color rgb="FF404040"/g)||[]).length,4);assert.equal((f.styles.match(/style="hair"><color rgb="FF404040"/g)||[]).length,4);}
 for(const [index,f] of files.entries()) {
  assert.match(f.xml,/<c[^>]*r="B2"[^>]*s="6"/,'blank Singapore result has no border');
  assert.match(f.xml,/<c[^>]*r="J2"[^>]*s="4"/,'numeric quantity including zero retains hair border');
  assert.match(f.xml,new RegExp('<c[^>]*r="N2"[^>]*s="'+(index===0?'5':'9')+'"'),'blank delivery formula keeps text style without border');
 }
 assert((await page.locator('.detail-table th').allTextContents()).includes('采购分工明细是否存在'));
 await page.locator('[data-filter-button="materialCodeValid"]').click();
 await page.locator('[data-filter-key="materialCodeValid"][value="是"]').check();
 assert.equal(await page.locator('#detailTableBody tr').count(),2);
 const filteredEvent=page.waitForEvent('download');await page.locator('#downloadDetailButton').click();const filtered=await filteredEvent;
 assert(filtered.suggestedFilename().endsWith('.xlsx'));
 const filteredBytes=await fs.readFile(await filtered.path());
 const filteredCodes=await page.evaluate(bytes=>{const w=XLSX.read(new Uint8Array(bytes),{type:'array'});return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1}).slice(1).map(r=>r[6]);},[...filteredBytes]);
 assert.deepEqual(filteredCodes,['A','B']);
 await page.locator('[data-filter-button="materialCodeValid"]').click();
 await page.locator('[data-filter-key="materialCodeValid"][value="是"]').uncheck();
 await page.locator('[data-filter-key="materialCodeValid"][value="否"]').check();
 assert.equal(await page.locator('#detailTableBody tr').count(),2);
 assert.deepEqual(await page.evaluate(()=>state.filteredRows.map(r=>r.materialCode)),['C','D']);
 await page.locator('[data-filter-key="materialCodeValid"][value="__all__"]').check();
 assert.equal(await page.locator('#detailTableBody tr').count(),4);
 assert(files[2].name.includes('未填写采购下单人'));assert.equal(files[0].rows[1][9],0);assert.deepEqual(errors,[]);console.log('PASS buyer ZIP, no cross-buyer rows, blank buyer, filename collisions, hidden preview reason, exported reasons, formulas and highlight');
}finally{await browser.close();}})().catch(e=>{console.error(e);process.exitCode=1;});
