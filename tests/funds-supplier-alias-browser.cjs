const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js'),{fixtures}=require('./funds-fixtures.cjs');
const names=['中山荣杰医疗器械工业有限公司','中山荣杰医疗器器材工业有限公司'];
const ids=['MDST-DGFK20261021013','MDST-DGFK20261021014'];
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
function replace(book,name,records,extra=[]){
 const headers=[...X.utils.sheet_to_json(book.Sheets[name],{header:1})[0],...extra];
 book.Sheets[name]=X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(h=>record[h]??''))]);
 return book.Sheets[name];
}
function value(sheet,header,row=2){const headers=X.utils.sheet_to_json(sheet,{header:1})[0];return sheet[X.utils.encode_cell({r:row-1,c:headers.indexOf(header)})];}
(async()=>{
 const data=await fixtures(),input={};
 for(const kind of ['ledger','applications','negotiated','oa'])input[kind]={name:kind+'.xlsx',book:X.read(data[kind],{cellNF:true,cellStyles:true})};
 input.negotiated.sheetName=data.sheetName;
 const base={'月份':'2026-11-01','供应商名称':names[1],'付款主体':'国源','开票时间':'2026-11-01','摘要':'合成核验货款','应付款时间':'2026-11-16','OA合同编号':'HT-RONGJIE','是否提交付款申请单':''};
 const ledgerRows=[
  {...base,'对公付款0A流程号':ids[0],'开票金额':79500,'申请付款金额':79500,'合同总金额':79500,'已付款金额':79500,'剩余付款金额':0,'是否结清':'是','是否谈判付款':'是','单据号【财务应付单】':'AP000001'},
  {...base,'对公付款0A流程号':ids[1],'开票金额':5000,'申请付款金额':5000,'合同总金额':5000,'已付款金额':0,'剩余付款金额':5000,'是否结清':'否','是否谈判付款':'否','单据号【财务应付单】':'AP000002'}
 ];
 for(const name of input.ledger.book.SheetNames)replace(input.ledger.book,name,name==='对账明细-李奇'?ledgerRows:[]);
 const source=replace(input.negotiated.book,data.sheetName,[{'供应商':names[0],'付款主体':'国源','对公付款0A流程号':ids[0],'合同总金额':79500,'已付金额':0,'摘要':'合成核验货款','申请付款金额':79500,'单据号【财务应付单】':'AP000001','创建人':'李奇','是否结清':'否','应付款时间':'2026-11-16','剩余付款金额':79500,'备注':'人工备注保持',46325:79500}]);
 source.N2={t:'n',v:79500,f:'G2-E2'};
 replace(input.applications.book,'付款申请汇总',ledgerRows.map((row,i)=>({'申请日期':'2026-11-13','供应商全称':names[1-i],'付款主体':'国源','流程编码':ids[i],'合同金额':row['申请付款金额'],'已付金额':0,'付款事由':'合成核验货款','备注':row['单据号【财务应付单】'],'对接人':'李奇','本次申请付款金额':row['申请付款金额'],'付款分类':i?'正常付款':'谈判付款'})),['付款分类']);
 replace(input.oa.book,'Sheet0',ledgerRows.map((row,i)=>({'流程编号':ids[i],'当前节点':'财务审核','未操作者':'测试审核人','流程状态':'未归档','申请人':'李奇','收款单位':names[i],'付款金额':row['申请付款金额']})));
 const before=JSON.stringify(input),expected=C.build(input);
 assert.equal(expected.stats.pending,0);assert.equal(expected.normalRows.length,1);assert.equal(expected.reviewRows.length,0);assert.equal(JSON.stringify(input),before);
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.goto(process.env.QA_URL||'http://127.0.0.1:8943/',{waitUntil:'networkidle'});
  await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const kind of ['ledger','applications','oa','negotiated'])await frame.locator('#'+kind+'File:not([disabled])').setInputFiles({name:input[kind].name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:write(input[kind].book)});
  await frame.locator('#negotiatedSheet option').filter({hasText:data.sheetName}).first().waitFor({state:'attached'});
  await frame.locator('#negotiatedSheet').selectOption(data.sheetName);await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:60000});
  assert.equal(await frame.locator('#pendingCount').textContent(),'0');
  assert.doesNotMatch(await frame.locator('#reviewPreview').innerText(),/OA供应商冲突|已付金额同步核对/);
  const event=page.waitForEvent('download');await frame.locator('#download').click();const download=await event,saved=X.read(await fs.readFile(await download.path()),{cellNF:true,cellStyles:true});
  assert.equal(X.utils.sheet_to_json(saved.Sheets.付款申请核对).length,0);
  const negotiated=saved.Sheets.谈判付款,normal=saved.Sheets.正常付款;
  const checks=[
   [negotiated,'negotiatedPreview',{'供应商':names[0],'对公付款0A流程号':ids[0],'已付金额':79500,'剩余付款金额':0,'是否结清':'是','是否提交付款申请单':'已经提交全部','当前节点':'财务审核'}],
   [normal,'normalPreview',{'供应商名称':names[1],'对公付款0A流程号':ids[1],'已付金额':0,'剩余金额':5000,'是否提交付款申请单':'已提交','当前节点':'财务审核','未操作者':'测试审核人','流程状态':'未归档','申请人':'李奇'}]
  ];
  for(const [sheet,preview,fields] of checks){
   const headers=await frame.locator('#'+preview+' th').allTextContents();
   for(const [header,wanted] of Object.entries(fields)){
    const cell=value(sheet,header);assert.equal(cell?.v,wanted,header);
    const shown=await frame.locator('#'+preview+' tbody tr').first().locator('td').nth(headers.indexOf(header)).textContent();
    assert.equal(typeof wanted==='number'?Number(shown.replace(/,/g,'')):shown,wanted,'预览 '+header);
   }
  }
  assert.equal(value(negotiated,'剩余付款金额').f,'G2-E2');assert.equal(negotiated.P2.v,79500);assert.equal(value(negotiated,'备注').v,'人工备注保持');
  assert.equal(value(saved.Sheets.付款进度整理,'供应商名称',2).v,names[1]);assert.equal(value(saved.Sheets.付款进度整理,'供应商名称',3).v,names[0]);
  assert.deepEqual(errors,[]);assert.equal(JSON.stringify(input),before);
  console.log('PASS: confirmed supplier aliases match across all four sources; preview/download preserve source names and OA, synchronize 79500 paid / 0 remaining / settled with original formula, normal workflow/submission matches, no alias-related reviews.');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exit(1);});
