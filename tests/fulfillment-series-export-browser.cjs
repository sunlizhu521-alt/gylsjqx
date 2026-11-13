const assert=require('node:assert/strict');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');

(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage();await page.goto(new URL('fulfillment.html',process.env.QA_URL||'http://127.0.0.1:8940/').href);await page.waitForFunction(()=>window.FulfillmentToolReady);
  const result=await page.evaluate(async()=>{
   const X=XLSX,source=X.utils.book_new(),raw=X.utils.aoa_to_sheet([['原始采购内容'],['跳转到历史系列表']]),series=X.utils.aoa_to_sheet([['人工系列校验','保留内容'],['旧系列',99]]);
   raw.A2.l={Target:"#'系列对比'!A2"};series.A2.c=[{a:'人工',t:'不能覆盖的人工批注'}];series['!cols']=[{wch:32},{wch:18}];series['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}}];
   X.utils.book_append_sheet(source,raw,'原始采购表');X.utils.book_append_sheet(source,series,'系列对比');X.utils.book_append_sheet(source,X.utils.aoa_to_sheet([['更早的校验'],[88]]),'系列对比（原表）');
   source.Workbook={Sheets:[{name:'原始采购表',Hidden:0},{name:'系列对比',Hidden:1},{name:'系列对比（原表）',Hidden:0}],Names:[
    {Name:'SeriesRange',Ref:"'系列对比'!$A$2",Sheet:1},{Name:'Unquoted',Ref:'系列对比!$B$2'},
    {Name:'Literal',Ref:'"系列对比!A2"'},{Name:'External',Ref:"'[外部.xlsx]系列对比'!A2"},
    {Name:'ExternalUnquoted',Ref:'[外部.xlsx]系列对比!A2'},{Name:'Span',Ref:"'原始采购表:系列对比'!A1"}
   ]};
   const zip=await JSZip.loadAsync(X.write(source,{type:'array',bookType:'xlsx'}));
   zip.file('xl/charts/chart1.xml','<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:f>系列对比!$B$2</c:f><c:f>"系列对比!B2"</c:f><c:f>\'[外部.xlsx]系列对比\'!B2</c:f></c:chart></c:chartSpace>');
   zip.file('xl/pivotCache/pivotCacheDefinition1.xml','<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="系列对比" ref="A1:B2"/></cacheSource></pivotCacheDefinition>');
   const external='<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><definedName>系列对比!A2</definedName></externalLink>';zip.file('xl/externalLinks/externalLink1.xml',external);
   const bytes=await zip.generateAsync({type:'uint8array'}),beforeXml=await zip.file('xl/worksheets/sheet2.xml').async('string');
   const generated=X.utils.book_new();X.utils.book_append_sheet(generated,X.utils.aoa_to_sheet([['本次系列'],[123]]),'系列对比');
   const output=await (await FulfillmentExport(new Blob([X.write(generated,{type:'array',bookType:'xlsx'})]),bytes)).arrayBuffer();
   const saved=X.read(output,{cellStyles:true}),after=await JSZip.loadAsync(output);
   const next=await (await FulfillmentExport(new Blob([X.write(generated,{type:'array',bookType:'xlsx'})]),output)).arrayBuffer(),repeated=X.read(next,{cellStyles:true});
   const plain=X.utils.book_new();X.utils.book_append_sheet(plain,raw,'原始采购表');const plainBytes=X.write(plain,{type:'array',bookType:'xlsx'}),plainBefore=await JSZip.loadAsync(plainBytes),plainOut=await (await FulfillmentExport(new Blob([X.write(generated,{type:'array',bookType:'xlsx'})]),plainBytes)).arrayBuffer(),plainAfter=await JSZip.loadAsync(plainOut);
   return {names:saved.SheetNames,value:saved.Sheets.系列对比.A2.v,oldValue:saved.Sheets['系列对比（原表2）'].B2.v,olderValue:saved.Sheets['系列对比（原表）'].A2.v,hidden:saved.Workbook.Sheets.find(s=>s.name==='系列对比（原表2）').Hidden,
    nativeUnchanged:beforeXml===await after.file('xl/worksheets/sheet2.xml').async('string'),comments:saved.Sheets['系列对比（原表2）'].A2.c,merges:saved.Sheets['系列对比（原表2）']['!merges'],
    link:saved.Sheets.原始采购表.A2.l.Target,defined:saved.Workbook.Names,chart:await after.file('xl/charts/chart1.xml').async('string'),pivot:await after.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string'),externalUnchanged:external===await after.file('xl/externalLinks/externalLink1.xml').async('string'),
    repeatedNames:repeated.SheetNames,repeatedValue:repeated.Sheets['系列对比（原表3）'].A2.v,repeatedOldValue:repeated.Sheets['系列对比（原表2）'].B2.v,
    originalNames:X.read(bytes).SheetNames,plainNames:X.read(plainOut).SheetNames,plainXmlUnchanged:await plainBefore.file('xl/worksheets/sheet1.xml').async('string')===await plainAfter.file('xl/worksheets/sheet1.xml').async('string')};
  });
  assert.deepEqual(result.names,['系列对比','原始采购表','系列对比（原表2）','系列对比（原表）']);assert.equal(result.value,123);assert.equal(result.oldValue,99);assert.equal(result.olderValue,88);assert.equal(result.hidden,1);assert(result.nativeUnchanged);
  assert.equal(result.comments[0].t,'不能覆盖的人工批注');assert.equal(result.merges.length,1);assert.equal(result.link,"#'系列对比（原表2）'!A2");
  const names=Object.fromEntries(result.defined.map(n=>[n.Name,n]));assert.equal(names.SeriesRange.Ref,"'系列对比（原表2）'!$A$2");assert.equal(names.SeriesRange.Sheet,2);assert.equal(names.Unquoted.Ref,"'系列对比（原表2）'!$B$2");assert.equal(names.Literal.Ref,'"系列对比!A2"');assert.equal(names.External.Ref,"'[外部.xlsx]系列对比'!A2");assert.equal(names.ExternalUnquoted.Ref,'[外部.xlsx]系列对比!A2');assert.equal(names.Span.Ref,"'原始采购表:系列对比（原表2）'!A1");
  assert.match(result.chart,/'系列对比（原表2）'!\$B\$2/);assert.match(result.chart,/"系列对比!B2"/);assert.match(result.chart,/\[外部.xlsx\]系列对比/);assert.match(result.pivot,/sheet="系列对比（原表2）"/);assert(result.externalUnchanged);
  assert.deepEqual(result.repeatedNames,['系列对比','系列对比（原表3）','原始采购表','系列对比（原表2）','系列对比（原表）']);assert.equal(result.repeatedValue,123);assert.equal(result.repeatedOldValue,99);
  assert.deepEqual(result.originalNames,['原始采购表','系列对比','系列对比（原表）']);assert.deepEqual(result.plainNames,['系列对比','原始采购表']);assert(result.plainXmlUnchanged);
  console.log('PASS: fixed current series sheet name, preserved original native sheet/notes/merges/hidden state, collision backups, internal names/links/chart/pivot references, external/string references, repeated and non-colliding exports');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
