const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8963/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  const frame=page.frameLocator('iframe[data-tool="funds"]'),input=fixture();
  const source=input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]],existingOA=source.C2.v;
  source.C1.v='对公付款OA流程号';
  X.utils.sheet_add_aoa(source,[
   ['不支付合成','迈德斯特','不支付',900,0,900,'否',46332,'保留人工备注','11月2日－900',900,'一致',900,0],
   ['零额合成','迈德斯特','未提交OA',0,0,0,'是',46332,'不支付待确认','11月2日－0',0,'一致',0,0]
  ],{origin:-1});
  X.utils.sheet_add_aoa(source,[['2026年11月支付','2026年12月支付'],[80,20],[900],[0]],{origin:'O1'});
  source.I3.c=[{a:'人工',t:'不支付记录仍保留批注'}];
  async function upload(role){await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:'合成-'+role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:write(input[role].book)});}
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  async function verify(expectedRows,excluded,monthMismatch=false){
   await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();
   assert.equal(await frame.locator('#pendingCount').innerText(),'0');
   assert.match(await frame.locator('#totalValidationStats').innerText(),new RegExp(`核对 ${expectedRows.length} 条；已排除“不支付” ${excluded} 条`));
   const snapshot=()=>frame.locator('#totalValidationPreview').evaluate(el=>({headers:[...el.querySelectorAll('th')].map(e=>e.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(td=>td.textContent))}));
   const preview=await snapshot(),event=page.waitForEvent('download');await frame.locator('#download').click();const download=await event;
   const output=X.read(await fs.readFile(await download.path()),{cellNF:true});
   const sheet=output.Sheets.总额校验核对,grid=X.utils.sheet_to_json(sheet,{header:1,defval:''});
   assert.deepEqual(grid[0],C.TOTAL_VALIDATION_HEADERS);assert.deepEqual(grid.slice(1).map(r=>r[2]),expectedRows);
   assert.equal(sheet['!ref'],'A1:S'+(expectedRows.length+1));assert.equal(sheet['!autofilter'].ref,sheet['!ref']);
   if(expectedRows.includes(2)){
    const first=grid[1];assert.equal(first[5],100);assert.equal(first[9],100,'总额相等也要按月核对');
    assert.equal(first[17],monthMismatch?'不一致':'一致');
    if(monthMismatch){for(const part of ['2026年11月','2026年12月','谈判明细J2合计100.00','月度支付O2金额80.00','差额20.00','月度支付P2金额20.00','差额-20.00'])assert(first[18].includes(part),part);}
    else assert.doesNotMatch(first[18],/差额/,'改正月份后旧提示清除');
   }
   assert.deepEqual(preview.headers,grid[0]);assert.equal(preview.rows.length,expectedRows.length);
   for(let r=0;r<expectedRows.length;r++)for(let c=0;c<grid[0].length;c++)assert.equal(preview.rows[r][c],X.utils.format_cell(sheet[X.utils.encode_cell({r:r+1,c})]));
   assert.deepEqual(await snapshot(),preview,'下载后预览与下载前一致');
   const original=output.Sheets.谈判付款,gridWithHeader=X.utils.sheet_to_json(original,{header:1,defval:''}),rows=[gridWithHeader[0],...gridWithHeader.slice(2)],headers=rows[0];
   assert.equal(original['!autofilter'].ref,original['!ref'].replace(/^A1:/,'A2:'),'筛选在下方日期表头行');
   const col=names=>headers.findIndex(h=>names.includes(h));
   assert.equal(rows.length,4,'原谈判付款三条记录均保留');
   for(let r=1;r<=3;r++){
    // Existing OA matching is unchanged; validation always uses the uploaded marker.
    assert.equal(rows[r][col(['对公付款0A流程号','对公付款OA流程号'])],r===1?existingOA:source['C'+(r+1)].v);
    assert.equal(rows[r][col(['申请付款金额','OA-申请付款金额'])],source['D'+(r+1)].v);
   }
   assert.equal(rows[2][col(['备注'])],'保留人工备注');
   assert.equal(original[X.utils.encode_cell({r:3,c:col(['备注'])})].c[0].t,'不支付记录仍保留批注');
   assert.equal(original.F3.f,'D3-E3','原金额公式随新增表头下移');
  }
  await verify([2,4],1,true);
  source.O2.v=100;source.P2.v=0;await upload('negotiated');await verify([2,4],1);
  source.C3.v='未提交OA';await upload('negotiated');await verify([2,3,4],0);
  for(const r of [2,3,4])source['C'+r].v='不支付';await upload('negotiated');await verify([],3);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,monthlyMismatch:true,correctedMonths:true,excluded:true,sourcePreserved:true,replacement:true,allExcluded:true,previewDownloadMatch:true}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
