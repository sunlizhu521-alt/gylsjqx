const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),assert=require('node:assert/strict'),fs=require('node:fs/promises');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const divisions=['海外事业一部','海外事业二部','国内事业部'];
function fixture(departments){
 const book=X.utils.book_new();
 for(const name of [...C.SHEETS,'孙立柱']){
  const rows=name==='徐英'?departments.map((d,i)=>C.COLUMNS.map(k=>({'事业部':d,'物料编码':'M'+i,'未交付数量':i,'下单日期':'2024年1月12日','下单月份':'2026年10月23日','合同约定交期':'10月陆续交货','OA备货单号':'OA'+i})[k]??'')):name==='孙立柱'?[C.COLUMNS.map(k=>({'事业部':divisions[0],'物料编码':'EXCLUDED','未交付数量':99})[k]??'')]:[];
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...rows]),name);
 }
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.goto(new URL('fulfillment.html',process.env.QA_URL||'http://127.0.0.1:8940/').href);
  await page.waitForFunction(()=>!document.getElementById('file').disabled);
  for(const d of divisions)assert(await page.locator('[data-download-division="'+d+'"]').isDisabled());
  const departments=[...divisions,divisions[0],'海外事业一部附属',''];
  await page.locator('#file').setInputFiles({name:'订单2026.10.23更新版.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:fixture(departments)});
  await page.waitForFunction(()=>!document.getElementById('download').disabled);
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);
  for(const d of divisions){
   const event=page.waitForEvent('download');await page.locator('[data-download-division="'+d+'"]').click();const dl=await event;
   assert.equal(dl.suggestedFilename(),'订单2026.10.23汇总版-'+d+'.xlsx');
   const w=X.read(await fs.readFile(await dl.path()),{type:'buffer',cellNF:true,cellStyles:true});
   assert.deepEqual(w.SheetNames,['汇总表']);
   const rows=X.utils.sheet_to_json(w.Sheets.汇总表,{header:1,defval:''});
   assert.deepEqual(rows[0],C.SUMMARY_COLUMNS);assert.equal(rows[0][3],'下单月份');assert.equal(rows[0].length,21);
   assert.deepEqual(rows.slice(1).map(r=>r[8]),departments.flatMap((name,i)=>name===d?['M'+i]:[]));
   assert(rows.slice(1).every(r=>r[0]===d));assert.equal(w.Sheets.汇总表.D2.t,'n');assert.equal(w.Sheets.汇总表.D2.z,'yyyy"年"m"月"');assert.equal(w.Sheets.汇总表.D2.w,'2026年10月');
   assert.equal(w.Sheets.汇总表.A1.s.fgColor.rgb,'315B9D');
   if(d===divisions[0])assert.equal(rows[1][11],0);
  }
  const full=page.waitForEvent('download');await page.locator('#download').click();const w=X.read(await fs.readFile(await(await full).path()),{type:'buffer'});
  assert.deepEqual(w.SheetNames,['汇总表','差异',...C.SHEETS,'孙立柱']);
  assert.equal(X.utils.sheet_to_json(w.Sheets.汇总表).length,departments.length);
  await page.locator('#file').setInputFiles({name:'仅国内.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:fixture(['国内事业部'])});
  await page.waitForFunction(()=>!document.getElementById('file').disabled);
  assert(await page.locator('[data-download-division="海外事业一部"]').isDisabled());assert(await page.locator('[data-download-division="海外事业二部"]').isDisabled());
  assert(!(await page.locator('[data-download-division="国内事业部"]').isDisabled()));assert.deepEqual(errors,[]);
  console.log('PASS division-only downloads, exact matching, row order, zero/date/style, cache restore, full export and empty groups');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
