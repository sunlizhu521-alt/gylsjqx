/* Add generated sheets to the original XLSX package without rewriting its worksheets. */
window.FulfillmentExport = async function(generatedBlob, sourceBytes) {
 const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main', R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const parser=new DOMParser(),serializer=new XMLSerializer();
 const parse=async(zip,path)=>{const file=zip.file(path);if(!file)throw Error('原文件缺少 '+path);const doc=parser.parseFromString(await file.async('string'),'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('原文件 XML 无法解析');return doc;};
 const elements=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)];
 const save=(zip,path,doc)=>zip.file(path,serializer.serializeToString(doc));
 let original;
 try {original=await JSZip.loadAsync(sourceBytes);}
 catch (_) {
  // Legacy XLS has no OOXML package; convert it before adding the generated sheets.
  const book=XLSX.read(sourceBytes,{type:'array',cellStyles:true,cellNF:true});
  original=await JSZip.loadAsync(XLSX.write(book,{type:'array',bookType:'xlsx',cellStyles:true}));
 }
 const generated=await JSZip.loadAsync(await generatedBlob.arrayBuffer());
 const workbook=await parse(original,'xl/workbook.xml'),rels=await parse(original,'xl/_rels/workbook.xml.rels');
 const types=await parse(original,'[Content_Types].xml');
 const styleRel=elements(rels,'Relationship').find(e=>e.getAttribute('Type')===R+'/styles');
 const resolve=target=>target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
 const stylePath=styleRel?resolve(styleRel.getAttribute('Target')):'xl/styles.xml';
 const styles=await parse(original,stylePath),addedStyles=await parse(generated,'xl/styles.xml');
 const formats=new Map(),targetFormats=elements(styles,'numFmts')[0]||styles.documentElement.insertBefore(styles.createElementNS(S,'numFmts'),styles.documentElement.firstChild);
 let nextFormat=Math.max(163,...elements(styles,'numFmt').map(e=>Number(e.getAttribute('numFmtId'))))+1;
 for(const f of elements(addedStyles,'numFmt')){const copy=styles.importNode(f,true),id=nextFormat++;formats.set(Number(f.getAttribute('numFmtId')),id);copy.setAttribute('numFmtId',id);targetFormats.append(copy);}
 targetFormats.setAttribute('count',targetFormats.children.length);
 const offsets={};
 for(const name of ['fonts','fills','borders','cellXfs']){
  const target=elements(styles,name)[0],source=elements(addedStyles,name)[0];
  if(!target||!source)throw Error('无法保留原工作簿样式：'+name);
  offsets[name]=target.children.length;
  for(const item of source.children){
   const copy=styles.importNode(item,true);
   if(name==='cellXfs'){
    for(const [attr,list] of [['fontId','fonts'],['fillId','fills'],['borderId','borders']])copy.setAttribute(attr,Number(copy.getAttribute(attr)||0)+offsets[list]);
    const format=Number(copy.getAttribute('numFmtId')||0);if(formats.has(format))copy.setAttribute('numFmtId',formats.get(format));
    copy.setAttribute('xfId','0');
   }
   target.append(copy);
  }
  target.setAttribute('count',target.children.length);
 }
 const generatedBook=await parse(generated,'xl/workbook.xml'),generatedRels=await parse(generated,'xl/_rels/workbook.xml.rels');
 const sheets=elements(workbook,'sheets')[0],first=sheets.firstChild;
 const names=new Set(elements(workbook,'sheet').map(e=>e.getAttribute('name').toLowerCase()));
 const ids=new Set(elements(rels,'Relationship').map(e=>e.getAttribute('Id')));
 let nextSheet=Math.max(0,...elements(workbook,'sheet').map(e=>Number(e.getAttribute('sheetId'))))+1;
 const addedSheets=elements(generatedBook,'sheet');
 const date1904=['1','true'].includes(elements(workbook,'workbookPr')[0]?.getAttribute('date1904'));
 for(const [i,sheet] of addedSheets.entries()){
  let name=sheet.getAttribute('name'),suffix=1;const base=name;
  while(names.has(name.toLowerCase()))name=base+'（本次'+(suffix++===1?'':suffix-1)+'）';
  names.add(name.toLowerCase());
  let rid='rIdFulfillment'+(i+1);while(ids.has(rid))rid+='x';ids.add(rid);
  let path='xl/worksheets/fulfillment'+(i+1)+'.xml';while(original.file(path))path=path.replace('.xml','x.xml');
  const relation=elements(generatedRels,'Relationship').find(e=>e.getAttribute('Id')===sheet.getAttributeNS(R,'id'));
  const doc=await parse(generated,resolve(relation.getAttribute('Target')));
  for(const c of elements(doc,'c')){
   const style=Number(c.getAttribute('s')||0);
   if(date1904){
    const xf=elements(addedStyles,'cellXfs')[0].children[style],numFmt=Number(xf?.getAttribute('numFmtId')||0);
    const format=elements(addedStyles,'numFmt').find(e=>Number(e.getAttribute('numFmtId'))===numFmt)?.getAttribute('formatCode')||XLSX.SSF.get_table()[numFmt];
    const v=elements(c,'v')[0];if(v&&(!c.getAttribute('t')||c.getAttribute('t')==='n')&&format&&XLSX.SSF.is_date(format))v.textContent=Number(v.textContent)-1462;
   }
   c.setAttribute('s',style+offsets.cellXfs);
  }
  save(original,path,doc);
  const entry=workbook.createElementNS(S,'sheet');entry.setAttribute('name',name);entry.setAttribute('sheetId',nextSheet++);entry.setAttributeNS(R,'r:id',rid);sheets.insertBefore(entry,first);
  const rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/worksheet');rel.setAttribute('Target',path.slice(3));rels.documentElement.append(rel);
  const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+path);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.append(type);
 }
 for(const name of elements(workbook,'definedName'))if(name.hasAttribute('localSheetId'))name.setAttribute('localSheetId',Number(name.getAttribute('localSheetId'))+addedSheets.length);
 for(const view of elements(workbook,'workbookView')){view.setAttribute('activeTab','0');view.setAttribute('firstSheet','0');}
 save(original,'xl/workbook.xml',workbook);save(original,'xl/_rels/workbook.xml.rels',rels);save(original,'[Content_Types].xml',types);save(original,stylePath,styles);
 return original.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
};
