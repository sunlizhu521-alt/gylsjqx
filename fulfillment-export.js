/* Add generated sheets while preserving the values-only source package and native parts. */
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
 const sourceSheets=elements(workbook,'sheet'),removedSheets=sourceSheets.filter(sheet=>/^差异(?:（本次(?:[2-9]|[1-9]\d+)?）)?$/.test(sheet.getAttribute('name')));
 const removedNames=new Set(removedSheets.map(sheet=>sheet.getAttribute('name')));
 if(removedSheets.length){
  const retainedSheets=sourceSheets.filter(sheet=>!removedSheets.includes(sheet));
  for(const name of elements(workbook,'definedName'))if(name.hasAttribute('localSheetId')){
   const sheet=sourceSheets[Number(name.getAttribute('localSheetId'))],index=retainedSheets.indexOf(sheet);
   if(index<0)name.remove();else name.setAttribute('localSheetId',index);
  }
  const removePart=path=>{original.remove(path);original.remove(path.replace(/([^/]+)$/,'_rels/$1.rels'));for(const type of elements(types,'Override'))if(type.getAttribute('PartName').replace(/^\//,'')===path)type.remove();};
  for(const sheet of removedSheets){
   const relation=elements(rels,'Relationship').find(e=>e.getAttribute('Id')===sheet.getAttributeNS(R,'id'));
   if(!relation)throw Error('差异工作表关系缺失，无法安全移除');
   removePart(new URL(relation.getAttribute('Target'),'https://workbook.local/xl/workbook.xml').pathname.slice(1));relation.remove();sheet.remove();
  }
  // All retained cells are values-only, so a saved calculation chain is obsolete.
  for(const relation of elements(rels,'Relationship').filter(e=>e.getAttribute('Type')===R+'/calcChain')){removePart(new URL(relation.getAttribute('Target'),'https://workbook.local/xl/workbook.xml').pathname.slice(1));relation.remove();}
  const pointsToRemoved=text=>text.split(/("(?:[^"]|"")*")/g).some((part,index)=>index%2===0&&[...part.matchAll(/(?<![\w\u0080-\uFFFF.\]'])('(?:[^']|'')+'|[A-Z0-9_.\u0080-\uFFFF]+(?::[A-Z0-9_.\u0080-\uFFFF]+)?)!/gi)].some(match=>{const name=match[1].replace(/^'|'$/g,'').replace(/''/g,"'");return !name.includes('[')&&name.split(':').some(end=>removedNames.has(end));}));
  const inspectReferences=(doc,path)=>{
   const reference=[...['definedName','f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula'].flatMap(tag=>elements(doc,tag).map(node=>node.textContent)),...elements(doc,'hyperlink').map(node=>node.getAttribute('location')||'')].find(pointsToRemoved);
   if(reference||elements(doc,'worksheetSource').some(node=>removedNames.has(node.getAttribute('sheet'))&&!node.hasAttributeNS(R,'id')))throw Error('保留内容 '+path+' 仍引用已移除的差异表，请先调整引用后重新上传');
  };
  inspectReferences(workbook,'工作簿命名范围');
  const visited=new Set();
  const inspectPart=async path=>{
   if(visited.has(path)||!original.file(path)||path.startsWith('xl/externalLinks/'))return;visited.add(path);
   if(path.endsWith('.xml'))inspectReferences(await parse(original,path),path);
   const relations=path.replace(/([^/]+)$/,'_rels/$1.rels');if(!original.file(relations))return;
   for(const relation of elements(await parse(original,relations),'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await inspectPart(new URL(relation.getAttribute('Target'),'https://workbook.local/'+path).pathname.slice(1));
  };
  for(const relation of elements(rels,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await inspectPart(new URL(relation.getAttribute('Target'),'https://workbook.local/xl/workbook.xml').pathname.slice(1));
 }
 for(const relation of elements(rels,'Relationship').filter(e=>e.getAttribute('Type')===R+'/worksheet')){
  const path=new URL(relation.getAttribute('Target'),'https://workbook.local/xl/workbook.xml').pathname.slice(1);
  if(elements(await parse(original,path),'f').length)throw Error('原工作表仍含公式，请重新上传并完成公式转数值后下载');
 }
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
 const previousSeries=elements(workbook,'sheet').find(sheet=>sheet.getAttribute('name')==='系列对比');
 if(previousSeries&&addedSheets.some(sheet=>sheet.getAttribute('name')==='系列对比')){
  let backup='系列对比（原表）',suffix=2;while(names.has(backup.toLowerCase()))backup='系列对比（原表'+suffix+++'）';
  previousSeries.setAttribute('name',backup);names.delete('系列对比');names.add(backup.toLowerCase());
  const rename=text=>text.split(/("(?:[^"]|"")*")/g).map((part,index)=>index%2?part:part.replace(/(?<![\w\u0080-\uFFFF.\]'])('(?:[^']|'')+'|[A-Z0-9_.\u0080-\uFFFF]+(?::[A-Z0-9_.\u0080-\uFFFF]+)?)!/gi,(match,token)=>{
   const name=token.replace(/^'|'$/g,'').replace(/''/g,"'");if(name.includes('['))return match;
   const ends=name.split(':');if(!ends.includes('系列对比'))return match;
   return "'"+ends.map(end=>end==='系列对比'?backup:end).join(':').replace(/'/g,"''")+"'!";
  })).join('');
  const updateReferences=doc=>{
   let changed=false;
   for(const tag of ['definedName','f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula'])for(const node of elements(doc,tag)){
    const value=rename(node.textContent);if(value!==node.textContent){node.textContent=value;changed=true;}
   }
   for(const link of elements(doc,'hyperlink'))if(link.hasAttribute('location')){
    const value=rename(link.getAttribute('location'));if(value!==link.getAttribute('location')){link.setAttribute('location',value);changed=true;}
   }
   for(const source of elements(doc,'worksheetSource'))if(source.getAttribute('sheet')==='系列对比'&&!source.hasAttributeNS(R,'id')){source.setAttribute('sheet',backup);changed=true;}
   return changed;
  };
  updateReferences(workbook);
  // Keep old references attached to the preserved sheet, not the newly generated comparison.
  for(const path of Object.keys(original.files))if(/^xl\/.*\.xml$/.test(path)&&!path.startsWith('xl/externalLinks/')&&path!=='xl/workbook.xml'&&path!==stylePath){
   const text=await original.file(path).async('string');if(!text.includes('系列对比'))continue;
   const doc=await parse(original,path);if(updateReferences(doc))save(original,path,doc);
  }
 }
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
 if(removedSheets.length&&original.file('docProps/app.xml')){
  const app=await parse(original,'docProps/app.xml'),pairs=elements(elements(app,'HeadingPairs')[0]||app,'vector')[0],titles=elements(elements(app,'TitlesOfParts')[0]||app,'vector')[0];
  if(pairs&&titles){let offset=0;for(let i=0;i+1<pairs.children.length;i+=2){const count=elements(pairs.children[i+1],'i4')[0],size=Number(count?.textContent||0);if(elements(pairs.children[i],'lpstr')[0]?.textContent==='Worksheets'){
   const remaining=[...titles.children],next=elements(workbook,'sheet').map(sheet=>{const title=app.createElementNS(titles.namespaceURI,'vt:lpstr');title.textContent=sheet.getAttribute('name');return title;});remaining.splice(offset,size,...next);titles.replaceChildren(...remaining);titles.setAttribute('size',remaining.length);count.textContent=next.length;break;
  }offset+=size;}save(original,'docProps/app.xml',app);}
 }
 save(original,'xl/workbook.xml',workbook);save(original,'xl/_rels/workbook.xml.rels',rels);save(original,'[Content_Types].xml',types);save(original,stylePath,styles);
 return original.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
};
