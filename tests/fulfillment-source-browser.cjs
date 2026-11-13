const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
 const page=await browser.newPage();await page.goto(new URL('fulfillment.html',process.env.QA_URL||'http://127.0.0.1:8940/').href);await page.waitForFunction(()=>window.FulfillmentToolReady);
 const result=await page.evaluate(async()=>{
  const w=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet([['原文','合并'],[2,3],[5]]);
  sheet.A3.f='SUM(A2:B2)';sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}}];
  XLSX.utils.book_append_sheet(w,sheet,'汇总表');
  for(const name of ['差异','差异（本次）','差异（本次2）','差异说明','辅助保留','差异（本次1）'])XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([[name],[12]]),name);
  w.Sheets.差异.A2.c=[{a:'测试',t:'随旧差异工作表移除'}];w.Sheets.差异说明.A2.c=[{a:'人工',t:'说明表内容保留'}];
  w.Workbook={WBProps:{date1904:true},Sheets:w.SheetNames.map(name=>({name,Hidden:['差异','差异说明'].includes(name)?1:0})),Names:[
   {Name:'范围',Ref:"'汇总表'!$A$2",Sheet:0},{Name:'旧差异局部名称',Ref:"'差异'!$A$2",Sheet:1},{Name:'辅助范围',Ref:"'辅助保留'!$A$2",Sheet:5},
   {Name:'文字常量',Ref:'"差异!A2"'},{Name:'外部范围',Ref:"'[外部.xlsx]差异'!$A$2"}
  ]};
  const source=XLSX.write(w,{type:'array',bookType:'xlsx'});
  const generated=FulfillmentCore.workbook({rows:[FulfillmentCore.COLUMNS.map(k=>k==='合同约定交期'?FulfillmentCore.dateCell({v:'2026-10-31'},false):{t:'s',v:''})],differences:[]});
  const blob=new Blob([XLSX.write(generated,{type:'array',bookType:'xlsx'})]);
  let rejectedFormula=false;try{await FulfillmentExport(blob,source);}catch(error){rejectedFormula=/公式/.test(error.message);}
  const prepared=await FulfillmentValues.prepare(source);if(prepared.issues.length)throw Error(JSON.stringify(prepared.issues));
  const input=await JSZip.loadAsync(prepared.bytes),parse=text=>new DOMParser().parseFromString(text,'application/xml'),all=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)],serializer=new XMLSerializer();
  const rels=parse(await input.file('xl/_rels/workbook.xml.rels').async('string')),chain=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');chain.setAttribute('Id','rIdTestCalcChain');chain.setAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain');chain.setAttribute('Target','calcChain.xml');rels.documentElement.append(chain);input.file('xl/_rels/workbook.xml.rels',serializer.serializeToString(rels));
  const types=parse(await input.file('[Content_Types].xml').async('string')),type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/xl/calcChain.xml');type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml');types.documentElement.append(type);input.file('[Content_Types].xml',serializer.serializeToString(types));input.file('xl/calcChain.xml','<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A2" i="2"/></calcChain>');
  const preparedBytes=await input.generateAsync({type:'uint8array'}),output=await FulfillmentExport(blob,preparedBytes),bytes=await output.arrayBuffer(),book=XLSX.read(bytes,{type:'array',cellNF:true}),zip=await JSZip.loadAsync(bytes);
  const repeated=XLSX.read(await (await FulfillmentExport(blob,bytes)).arrayBuffer(),{type:'array'});
  const paths=['xl/worksheets/sheet2.xml','xl/worksheets/sheet3.xml','xl/worksheets/sheet4.xml','xl/worksheets/_rels/sheet2.xml.rels','xl/calcChain.xml'];
  const finalRels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string')),finalTypes=parse(await zip.file('[Content_Types].xml').async('string')),app=parse(await zip.file('docProps/app.xml').async('string'));
  const remainingNames=book.Workbook.Names,date=book.Sheets[book.SheetNames[0]].S2;
  const referenceInput=XLSX.read(preparedBytes,{type:'array'});referenceInput.Workbook.Names.push({Name:'仍在使用的范围',Ref:"'差异'!A2"});let dependencyRejected=false;
  try{await FulfillmentExport(blob,XLSX.write(referenceInput,{type:'array',bookType:'xlsx'}));}catch(error){dependencyRejected=/仍引用已移除的差异表/.test(error.message);}
  return {names:book.SheetNames,originalNames:XLSX.read(source).SheetNames,repeatedNames:repeated.SheetNames,rejectedFormula,dependencyRejected,
   formula:book.Sheets.汇总表.A3.f,value:book.Sheets.汇总表.A3.v,merges:book.Sheets.汇总表['!merges'],hidden:book.Workbook.Sheets.find(s=>s.name==='差异说明').Hidden,remainingNames,
   unchanged:await input.file('xl/worksheets/sheet5.xml').async('string')===await zip.file('xl/worksheets/sheet5.xml').async('string'),comment:book.Sheets.差异说明.A2.c[0].t,
   deletedParts:paths.every(path=>!zip.file(path)),staleRelations:all(finalRels,'Relationship').some(rel=>['worksheets/sheet2.xml','worksheets/sheet3.xml','worksheets/sheet4.xml','calcChain.xml'].includes(rel.getAttribute('Target'))),staleTypes:all(finalTypes,'Override').some(t=>paths.includes(t.getAttribute('PartName').replace(/^\//,''))),
   appTitles:all(all(app,'TitlesOfParts')[0],'lpstr').map(t=>t.textContent),date:XLSX.SSF.format(date.z,date.v,{date1904:true})};
 });
 assert.deepEqual(result.names,['汇总表（本次）','汇总表','差异说明','辅助保留','差异（本次1）']);
 assert.deepEqual(result.originalNames,['汇总表','差异','差异（本次）','差异（本次2）','差异说明','辅助保留','差异（本次1）']);
 assert(!result.repeatedNames.some(name=>/^差异(?:（本次(?:[2-9]|[1-9]\d+)?）)?$/.test(name)),'重导出不恢复旧差异表');
 assert(result.rejectedFormula);assert(result.dependencyRejected);assert(result.unchanged);assert(result.deletedParts);assert(!result.staleRelations);assert(!result.staleTypes);assert.deepEqual(result.appTitles,result.names);
 assert.equal(result.formula,undefined);assert.equal(result.value,5);assert.equal(result.merges.length,1);assert.equal(result.hidden,1);assert.equal(result.comment,'说明表内容保留');assert.equal(result.date,'2026年10月31日');
 const names=Object.fromEntries(result.remainingNames.map(n=>[n.Name,n]));assert.equal(names.范围.Sheet,1);assert.equal(names.辅助范围.Sheet,3);assert.equal(names.旧差异局部名称,undefined);assert.equal(names.文字常量.Ref,'"差异!A2"');assert.equal(names.外部范围.Ref,"'[外部.xlsx]差异'!$A$2");
 const legacy=await page.evaluate(async()=>{
  const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([['原文'],[12]]),'原表');XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([['旧差异']]),'差异');
  const g=FulfillmentCore.workbook({rows:[],differences:[]}),blob=await FulfillmentExport(new Blob([XLSX.write(g,{type:'array',bookType:'xlsx'})]),XLSX.write(w,{type:'array',bookType:'biff8'}));
  const b=XLSX.read(await blob.arrayBuffer(),{type:'array'});return {names:b.SheetNames,value:b.Sheets.原表.A2.v};
 });
 assert.deepEqual(legacy.names,['汇总表','原表']);assert.equal(legacy.value,12);
 console.log('PASS original XML/merges/notes/hidden states and 1904 dates preserved; old difference sheets/relations/local names/calculation chain removed, metadata reindexed, dependencies guarded, legacy/repeated exports verified');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
