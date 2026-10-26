const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js'),{fixtures}=require('./funds-fixtures.cjs');
function fields(sheet){return Object.fromEntries(X.utils.sheet_to_json(sheet,{header:1})[0].map((name,i)=>[name,i]));}
function assign(sheet,row,values){const headers=fields(sheet);for(const [key,value] of Object.entries(values)){let col=headers[key];if(col===undefined){col=X.utils.decode_range(sheet['!ref']).e.c+1;headers[key]=col;X.utils.sheet_add_aoa(sheet,[[key]],{origin:{r:0,c:col}});}X.utils.sheet_add_aoa(sheet,[[value]],{origin:{r:row-1,c:col}});}}
(async()=>{
 const data=await fixtures(),ledger=X.read(data.ledger,{cellNF:true,cellStyles:true}),negotiated=X.read(data.negotiated,{cellNF:true,cellStyles:true});
 const oldNormal='MDST-DGFK20261024002',candidateNormal='MDST-DGFK20261024001',oldNegotiated='MDST-DGFK20261024003',candidateNegotiated='MDST-DGFK20261024010';
 const originalFields={'流程编号':'原流程信息保留','当前节点':'原节点保留','未操作者':'原操作者保留','流程状态':'原流程状态保留','申请人':'原申请人保留','提交时间汇总':'2026-10-01-20.00','提交付款申请单日期':'2026-10-01'};
 assign(ledger.Sheets['对账明细-李奇'],2,{...originalFields,'对公付款0A流程号':oldNormal,'是否提交付款申请单':'已提交'});
 assign(negotiated.Sheets[data.sheetName],2,{...originalFields,'对公付款0A流程号':oldNegotiated,'是否提交付款申请单':'已经提交部分'});
 data.ledger=Buffer.from(X.write(ledger,{type:'buffer',bookType:'xlsx'}));data.negotiated=Buffer.from(X.write(negotiated,{type:'buffer',bookType:'xlsx'}));
 const input={};for(const kind of ['ledger','applications','negotiated','oa'])input[kind]={name:kind+'.xlsx',book:X.read(data[kind],{cellNF:true,cellStyles:true})};input.negotiated.sheetName=data.sheetName;
 const result=C.build(input);assert.equal(result.stats.pending,0);assert.equal(result.stats.negotiatedRows,2);assert.equal(result.normalRows.length,5);
 const expectedReviews=result.reviewRows.filter(row=>row[5].v.startsWith('OA供应商冲突'));assert.equal(expectedReviews.length,2);
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.goto(process.env.QA_URL||'http://127.0.0.1:8940/',{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const kind of ['ledger','applications','oa','negotiated'])await frame.locator('#'+kind+'File:not([disabled])').setInputFiles({name:kind+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:data[kind]});
  await frame.locator('#negotiatedSheet option').filter({hasText:data.sheetName}).first().waitFor({state:'attached'});await frame.locator('#negotiatedSheet').selectOption(data.sheetName);await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:60000});
  assert.equal(await frame.locator('#pendingCount').textContent(),'0');
  const event=page.waitForEvent('download');await frame.locator('#download').click();const download=await event,saved=X.read(await fs.readFile(await download.path()),{cellNF:true,cellStyles:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);
  const normal=saved.Sheets.正常付款,neg=saved.Sheets.谈判付款,review=saved.Sheets.付款申请核对,normalFields=fields(normal),negFields=fields(neg);
  for(const [sheetName,sheet,headers,expectedRows,expectedHeaders] of [['正常付款',normal,normalFields,result.normalRows,result.normalHeaders],['付款申请核对',review,fields(review),result.reviewRows,result.reviewHeaders]]){
   assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],expectedHeaders);
   for(let r=0;r<expectedRows.length;r++)for(let c=0;c<expectedHeaders.length;c++)assert.equal(sheet[X.utils.encode_cell({r:r+1,c})]?.v??'',sheetName==='付款申请核对'&&c===5?C.formatReviewText(expectedRows[r][c].v):expectedRows[r][c].v);
  }
  for(const [sheet,headers,preview,oldId,state] of [[normal,normalFields,'normalPreview',oldNormal,'已提交'],[neg,negFields,'negotiatedPreview',oldNegotiated,'已经提交部分']]){
   const cell=key=>sheet[X.utils.encode_cell({r:1,c:headers[key]})]?.v??'';
   assert.equal(cell('对公付款0A流程号'),oldId);assert.equal(cell('是否提交付款申请单'),state);
   for(const [key,value] of Object.entries(originalFields)){if(headers[key]===undefined)continue;assert.equal(cell(key),value,'原信息应保留：'+key);}
   const previewHeaders=await frame.locator('#'+preview+' th').allTextContents();
   for(const key of ['对公付款0A流程号','流程编号','当前节点','未操作者','流程状态','申请人','是否提交付款申请单','提交时间汇总'])assert.equal(await frame.locator('#'+preview+' tbody tr').first().locator('td').nth(previewHeaders.indexOf(key)).textContent(),String(cell(key)));
  }
  const reviews=X.utils.sheet_to_json(review),conflicts=reviews.filter(row=>row['核对说明'].includes('OA供应商冲突'));assert.equal(conflicts.length,2);
  for(const [supplier,other,oldId,candidate,amount,otherAmount] of [['正常供应商1','正常供应商2',oldNormal,candidateNormal,'100.00','101'],['谈判甲公司','正常供应商3',oldNegotiated,candidateNegotiated,'300.00','102']]){
   const row=conflicts.find(row=>row['供应商名称']===supplier);assert(row);const text=row['核对说明'];
   for(const term of [supplier,other,oldId,candidate,amount,'付款金额：'+otherAmount+'｜','oa.xlsx/Sheet0!'])assert(text.includes(term),'核对说明缺少 '+term);
   assert(!text.includes('已更新为'));assert(text.includes('未自动更新'));assert.equal(await frame.locator('#reviewPreview td.review-detail').filter({hasText:supplier}).filter({hasText:'OA供应商冲突'}).textContent(),text);
  }
  const before=input.negotiated.book.Sheets[data.sheetName];assert.deepEqual(neg['!merges'],before['!merges']);assert.equal(X.utils.decode_range(neg['!ref']).e.r,X.utils.decode_range(before['!ref']).e.r);
  for(const address of ['D2','E2','G2','K2','N2','Q2','O2']){assert.equal(neg[address].v,before[address].v,'原谈判内容保留 '+address);assert.equal(neg[address].f,before[address].f,'原谈判公式保留 '+address);}
  assert.equal(normal.P2.f,'I2-N2');assert.equal(normal.P2.v,100);assert.deepEqual(errors,[]);
  console.log('PASS: negotiated and normal cross-supplier OA conflicts preserve original OA, flow and application fields; both identities, amounts, flow IDs and source rows appear in review; preview equals three-sheet download; negotiated records, merges and financial formulas preserved.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
