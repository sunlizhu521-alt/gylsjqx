const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
 const page=await browser.newPage();await page.goto(new URL('fulfillment.html',process.env.QA_URL||'http://127.0.0.1:8940/').href);await page.waitForFunction(()=>window.FulfillmentToolReady);
 const result=await page.evaluate(async()=>{
  const w=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet([['原文','合并'],[2,3],[5]]);
  sheet.A3.f='SUM(A2:B2)';sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}}];
  XLSX.utils.book_append_sheet(w,sheet,'汇总表');XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([['额外原表']]),'差异');
  w.Workbook={WBProps:{date1904:true},Sheets:[{name:'汇总表',Hidden:0},{name:'差异',Hidden:1}],Names:[{Name:'范围',Ref:"'汇总表'!$A$2",Sheet:0}]};
  const source=XLSX.write(w,{type:'array',bookType:'xlsx'});
  const generated=FulfillmentCore.workbook({rows:[FulfillmentCore.COLUMNS.map(k=>k==='合同约定交期'?FulfillmentCore.dateCell({v:'2026-10-31'},false):{t:'s',v:''})],differences:[]});
  const blob=new Blob([XLSX.write(generated,{type:'array',bookType:'xlsx'})]);
  let rejectedFormula=false;try{await FulfillmentExport(blob,source);}catch(error){rejectedFormula=/公式/.test(error.message);}
  const prepared=await FulfillmentValues.prepare(source);if(prepared.issues.length)throw Error(JSON.stringify(prepared.issues));
  const output=await FulfillmentExport(blob,prepared.bytes),bytes=await output.arrayBuffer(),book=XLSX.read(bytes,{type:'array',cellNF:true});
  const a=await JSZip.loadAsync(source),b=await JSZip.loadAsync(bytes);
  const originalXml=await a.file('xl/worksheets/sheet2.xml').async('string'),outputXml=await b.file('xl/worksheets/sheet2.xml').async('string');
  const date=book.Sheets[book.SheetNames[0]].R2;
  return {names:book.SheetNames,unchanged:originalXml===outputXml,rejectedFormula,formula:book.Sheets.汇总表.A3.f,value:book.Sheets.汇总表.A3.v,merges:book.Sheets.汇总表['!merges'],hidden:book.Workbook.Sheets[3].Hidden,defined:book.Workbook.Names[0].Sheet,date:XLSX.SSF.format(date.z,date.v,{date1904:true})};
 });
 assert.deepEqual(result.names,['汇总表（本次）','差异（本次）','汇总表','差异']);
 assert.equal(result.rejectedFormula,true,'禁止把未转值的源表直接导出');
 assert.equal(result.unchanged,true);assert.equal(result.formula,undefined);assert.equal(result.value,5);assert.equal(result.merges.length,1);assert.equal(result.hidden,1);assert.equal(result.defined,2);assert.equal(result.date,'2026年10月31日');
 const legacy=await page.evaluate(async()=>{
  const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([['原文'],[12]]),'原表');
  const g=FulfillmentCore.workbook({rows:[],differences:[]});
  const blob=await FulfillmentExport(new Blob([XLSX.write(g,{type:'array',bookType:'xlsx'})]),XLSX.write(w,{type:'array',bookType:'biff8'}));
  const b=XLSX.read(await blob.arrayBuffer(),{type:'array'});return {names:b.SheetNames,value:b.Sheets.原表.A2.v};
 });
 assert.deepEqual(legacy.names,['汇总表','差异','原表']);assert.equal(legacy.value,12);
 console.log('PASS original XML/merges, formulas converted to values, colliding names, hidden sheet, local names, 1904 date system');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
