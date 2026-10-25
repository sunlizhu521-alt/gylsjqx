const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),{fixtures,REVIEW_HEADERS}=require('./funds-fixtures.cjs');
(async()=>{
 const data=await fixtures(),app=X.read(data.applications),s=app.Sheets.付款申请汇总,rows=X.utils.sheet_to_json(s,{header:1});
 rows[0].push('付款分类');for(const row of rows.slice(1))row[10]=row[1].startsWith('谈判')?'谈判付款':'正常付款';
 const normal=rows.find(r=>r[1]==='正常供应商1'),negotiated=rows.find(r=>r[1]==='谈判甲公司'&&r[9]===100);
 rows.push(normal.slice(),['2026-10-25',...normal.slice(1)],['2026-10-25',...negotiated.slice(1)]);
 const external=normal.slice();external[1]='未关联测试供应商';external[3]='MDST-DGFK20991231099';external[9]=12.34;
 rows.push(external.slice(),['2026-10-26',...external.slice(1)],Object.assign(external.slice(),{10:'谈判付款'}));
 app.Sheets.付款申请汇总=X.utils.aoa_to_sheet(rows);data.applications=Buffer.from(X.write(app,{type:'buffer',bookType:'xlsx'}));
 const input={};for(const key of ['ledger','applications','negotiated','oa'])input[key]={name:key+'.xlsx',book:X.read(data[key],{cellNF:true,cellStyles:true})};input.negotiated.sheetName=data.sheetName;
 const result=C.build(input);assert.equal(result.stats.pending,0);assert.equal(result.reviewRows.length,3);assert.equal(result.normalRows.length,5);assert.equal(result.stats.negotiatedRows,2);
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.goto(process.env.QA_URL||'http://127.0.0.1:8940/',{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const key of ['ledger','applications','oa','negotiated'])await frame.locator('#'+key+'File:not([disabled])').setInputFiles({name:key+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:data[key]});
  await frame.locator('#negotiatedSheet option').filter({hasText:data.sheetName}).first().waitFor({state:'attached'});await frame.locator('#negotiatedSheet').selectOption({label:data.sheetName});await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:60000});
  assert.equal(await frame.locator('#pendingCount').innerText(),'0');assert.deepEqual(await frame.locator('#reviewPreview th').allTextContents(),REVIEW_HEADERS);assert.equal(await frame.locator('#reviewPreview tbody tr').count(),3);
  assert.match(await frame.locator('#reviewPreview').innerText(),/未关联测试供应商/);assert.match(await frame.locator('#reviewPreview').innerText(),/共 3 次，重复 2 次/);
  const event=page.waitForEvent('download');await frame.locator('#download').click();const downloaded=await event,bytes=await fs.readFile(await downloaded.path()),saved=X.read(bytes,{cellStyles:true,cellNF:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);const review=saved.Sheets.付款申请核对;
  for(let r=0;r<result.reviewRows.length;r++)for(let c=0;c<REVIEW_HEADERS.length;c++)assert.equal(review[X.utils.encode_cell({r:r+1,c})]?.v??'',result.reviewRows[r][c].v);
  assert.equal(review['!autofilter'].ref,'A1:J4');assert.equal(review.D2.t,'n');assert.equal(review.J2.t,'n');for(const row of review['!rows'])assert.equal(row.hpt,30);
  const n=saved.Sheets.谈判付款,original=input.negotiated.book.Sheets[data.sheetName];assert.deepEqual(n['!merges'],original['!merges']);
  for(const address of ['E2','N2','Q2','K2','O2']){assert.equal(n[address].v,original[address].v);assert.equal(n[address].f,original[address].f);}
  const headers=Object.fromEntries(result.negotiated.headers.map((h,i)=>[h,i]));assert.equal(result.negotiated.rows[0][headers['是否提交付款申请单']].v,'已经提交部分');
  assert.equal(X.utils.sheet_to_json(saved.Sheets.正常付款).length,5);assert.deepEqual(errors,[]);
  const name=process.env.QA_URL?'live':'local';await fs.writeFile('/private/tmp/funds-duplicates-'+name+'.xlsx',bytes);
  console.log('PASS '+name+': duplicate groups in both categories and unmatched requests, one review per group, complete evidence, no blocking, three-sheet export and financial preservation');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
