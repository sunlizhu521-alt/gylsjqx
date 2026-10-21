const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs/promises');
(async()=>{const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});try{
 const page=await browser.newPage();await page.goto((process.env.QA_URL||'http://127.0.0.1:8940/')+'beihuo/index.html');await page.waitForFunction(()=>typeof renderDemandDashboard==='function');
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.evaluate(()=>{state.demandRows=[{buyer:'甲/乙',materialCode:'A',quantityNumber:0,stockReason:'原因A',deliveryDays:2},{buyer:'甲/乙',materialCode:'B',quantityNumber:3,stockReason:'原因B'},{buyer:'甲:乙',materialCode:'C',quantityNumber:4,stockReason:'原因C'},{buyer:'',materialCode:'D',quantityNumber:5,stockReason:'原因D'}];renderDemandDashboard();});
 assert.equal(await page.locator('.detail-table th').count(),18);assert(!(await page.locator('.detail-table').innerText()).includes('原因A'));
 const event=page.waitForEvent('download');await page.locator('#downloadDetailButton').click();const download=await event;assert(download.suggestedFilename().endsWith('.zip'));
 const bytes=await fs.readFile(await download.path());
 const files=await page.evaluate(async bytes=>{const z=await JSZip.loadAsync(new Uint8Array(bytes));return Promise.all(Object.values(z.files).filter(f=>!f.dir).map(async f=>{const b=await f.async('uint8array'),w=XLSX.read(b,{type:'array',cellNF:true}),s=w.Sheets[w.SheetNames[0]];const inner=await JSZip.loadAsync(b);return {name:f.name,rows:XLSX.utils.sheet_to_json(s,{header:1,defval:''}),highlight:w.Custprops.KSOReadingLayout,formula:s.N2.f,dateFormat:s.N2.z,styles:await inner.file('xl/styles.xml').async('string')};}));},[...bytes]);
 assert.equal(files.length,3);assert.equal(new Set(files.map(f=>f.name)).size,3);
 assert.deepEqual(files.map(f=>f.rows.slice(1).map(r=>r[4])),[['A','B'],['C'],['D']]);
 assert.deepEqual(files.map(f=>f.rows.slice(1).map(r=>r[18])),[['原因A','原因B'],['原因C'],['原因D']]);
 for(const f of files){assert.equal(f.rows[0][18],'备货原因');assert.equal(f.highlight,true);assert.equal(f.formula,'IF(M2="","",L2+M2)');assert.equal(f.dateFormat,'yyyy-mm-dd');assert.match(f.styles,/FF1F4E78/);}
 assert(files[2].name.includes('未填写采购下单人'));assert.equal(files[0].rows[1][7],0);assert.deepEqual(errors,[]);console.log('PASS buyer ZIP, no cross-buyer rows, blank buyer, filename collisions, hidden preview reason, exported reasons, formulas and highlight');
}finally{await browser.close();}})().catch(e=>{console.error(e);process.exitCode=1;});
