const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage();await page.goto(process.env.QA_URL||'http://127.0.0.1:8940/');
  await page.addScriptTag({path:require('path').resolve('funds-export.js')});
  const book=X.utils.book_new();const source=X.utils.aoa_to_sheet([['供应商','申请付款金额','已付金额','剩余付款金额','付款计划'],['测试甲',100,0,100,'保持备注'],['测试乙',200,20,180,'合并备注'],['测试丙',300,30,270],['合计',600,50,550]]);
  source.D2={t:'n',v:100,f:'B2-C2'};source.E2={t:'n',v:100,f:'D2'};source.D5={t:'n',v:550,f:'SUM(D2:D4)'};source['!merges']=[{s:{r:2,c:4},e:{r:3,c:4}}];source['!cols']=[{wch:30},{wch:20},{wch:20},{wch:20},{wch:35}];source['!rows']=[{hpt:33},{hpt:40},{hpt:27}];
  X.utils.book_append_sheet(book,source,'未来10月谈判付款');X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['不应保留'],['secret']]),'其他');
  const input=Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
  const result={normalHeaders:['登记月份','供应商名称','申请付款金额','已付金额','剩余金额'],normalRows:[[{t:'n',v:46319,z:'yyyy年m月'},{t:'s',v:'测试正常公司'},{t:'n',v:12.34,z:'#,##0.00'},{t:'n',v:2.34,z:'#,##0.00'},{t:'n',v:10,f:'C2-D2',z:'#,##0.00'}]],negotiated:{sheetName:'未来10月谈判付款',headerRow:0,updates:[{address:'C2',cell:{t:'n',v:40}},{address:'D2',cell:{t:'n',v:60,f:'B2-C2'}},{address:'F1',cell:{t:'s',v:'提交付款申请单日期'}},{address:'F2',cell:{t:'s',v:'2026-10-01\n2026-10-24'}}]}};
  async function exportBytes(bytes,data){return Buffer.from(await page.evaluate(async({bytes,result})=>[...new Uint8Array(await (await FundsExport(new Uint8Array(bytes),result)).arrayBuffer())],{bytes:[...bytes],result:data}));}
  const output=await exportBytes(input,result),saved=X.read(output,{cellStyles:true,cellNF:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款']);const s=saved.Sheets.谈判付款;
  assert.equal(s.C2.v,40);assert.equal(s.D2.v,60);assert.equal(s.E2.v,60);assert.equal(s.D5.v,510);assert.equal(s.D5.f,'SUM(D2:D4)');assert.deepEqual(s['!merges'],source['!merges']);assert.equal(s['!rows'][1].hpt,40);assert.equal(s.E3.v,'合并备注');assert.equal(saved.Sheets.正常付款.A2.v,46319);assert.equal(saved.Sheets.正常付款.E2.v,10);
  assert.equal(saved.Sheets.正常付款['!autofilter'].ref,'A1:E2');assert.equal(saved.Sheets.正常付款.A1.s.fgColor.rgb,'315B9D');
  const parts=await page.evaluate(async bytes=>{const z=await JSZip.loadAsync(new Uint8Array(bytes));return {files:Object.keys(z.files),xml:await z.file('xl/worksheets/funds-normal.xml').async('string')};},[...output]);assert(!parts.files.includes('xl/worksheets/sheet2.xml'));assert.match(parts.xml,/state="frozen"/);
  const again=await exportBytes(output,{...result,negotiated:{...result.negotiated,sheetName:'谈判付款'}});assert.deepEqual(X.read(again).SheetNames,['谈判付款','正常付款']);
  const b1904=X.utils.book_new();b1904.Workbook={WBProps:{date1904:true}};X.utils.book_append_sheet(b1904,source,'未来10月谈判付款');const out1904=await exportBytes(Buffer.from(X.write(b1904,{type:'buffer',bookType:'xlsx'})),result);assert.equal(X.read(out1904).Sheets.正常付款.A2.v,46319-1462);
  source.E4={f:"'其他'!A1",t:'n',v:0};const invalid=X.utils.book_new();X.utils.book_append_sheet(invalid,source,'未来10月谈判付款');X.utils.book_append_sheet(invalid,X.utils.aoa_to_sheet([[1]]),'其他');await assert.rejects(exportBytes(Buffer.from(X.write(invalid,{type:'buffer',bookType:'xlsx'})),result),/引用了其他工作表/);
  await fs.writeFile('/private/tmp/funds-export-synthetic.xlsx',output);
  console.log('PASS: native negotiated data/merges/formulas, two tabs, normal styles/dates, repeated export, dependent caches, cross-sheet guard');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
