/* Recalculate cell formulas, then patch only their OOXML cells. Original files are never written. */
(()=>{'use strict';
const NS={s:'http://schemas.openxmlformats.org/spreadsheetml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',p:'http://schemas.openxmlformats.org/package/2006/relationships',ct:'http://schemas.openxmlformats.org/package/2006/content-types',xdr:'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',a:'http://schemas.openxmlformats.org/drawingml/2006/main'};
const all=(node,name)=>[...node.getElementsByTagNameNS('*',name)],children=(node,name)=>[...node.children].filter(c=>c.localName===name);
function parse(text){const doc=new DOMParser().parseFromString(text,'application/xml');if(all(doc,'parsererror').length)throw Error('工作簿 XML 无法读取');return doc;}
const xml=doc=>new XMLSerializer().serializeToString(doc);
const partPath=(target,base)=>new URL(target,'https://workbook.local/'+base).pathname.slice(1);
const relPath=path=>path.replace(/([^/]+)$/,'_rels/$1.rels');
function relative(target,base){const a=base.split('/');a.pop();const b=target.split('/');while(a.length&&b.length&&a[0]===b[0]){a.shift();b.shift();}return '../'.repeat(a.length)+b.join('/');}
function el(doc,namespace,name,attrs={}){const node=doc.createElementNS(namespace,name);for(const [key,value] of Object.entries(attrs))node.setAttribute(key,String(value));return node;}
function appendText(doc,parent,name,value){const node=el(doc,NS.xdr,'xdr:'+name);node.textContent=String(value);parent.append(node);}
function emptyRels(){return parse('<Relationships xmlns="'+NS.p+'"/>');}
function nextId(doc){const used=new Set(all(doc,'Relationship').map(n=>n.getAttribute('Id')));let id=1;while(used.has('rId'+id))id++;return 'rId'+id;}
const issue=(sheet,cell,message)=>({sheet,cell,field:'公式转数值',kind:'formula',message});
function replaceValue(doc,cell,value,type){
 for(const node of [...cell.children])if(['f','v','is'].includes(node.localName))node.remove();
 cell.removeAttribute('t');
 if(type==='z'||value==null)return;
 if(type==='s'||typeof value==='string'){cell.setAttribute('t','inlineStr');const is=el(doc,NS.s,'is'),t=el(doc,NS.s,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=value;is.append(t);cell.append(is);}
 else{const v=el(doc,NS.s,'v');if(type==='b'||typeof value==='boolean'){cell.setAttribute('t','b');v.textContent=value?'1':'0';}else{if(typeof value!=='number'||!Number.isFinite(value))throw Error('公式结果不是可保存的数值');v.textContent=String(value);}cell.append(v);}
}
function cellBox(doc,address){
 let point=XLSX.utils.decode_cell(address),end={...point};
 for(const merge of all(doc,'mergeCell')){const range=XLSX.utils.decode_range(merge.getAttribute('ref'));if(point.r>=range.s.r&&point.r<=range.e.r&&point.c>=range.s.c&&point.c<=range.e.c){point=range.s;end=range.e;break;}}
 const format=all(doc,'sheetFormatPr')[0],defaultWidth=Number(format?.getAttribute('defaultColWidth'))||8.43,defaultHeight=Number(format?.getAttribute('defaultRowHeight'))||15;
 const cols=all(doc,'col'),rows=new Map(all(doc,'row').map(n=>[Number(n.getAttribute('r'))-1,n]));
 let width=0,height=0;
 for(let c=point.c;c<=end.c;c++){const col=cols.find(n=>Number(n.getAttribute('min'))<=c+1&&Number(n.getAttribute('max'))>=c+1);const w=Number(col?.getAttribute('width'))||defaultWidth;width+=Math.max(1,Math.floor(w*7+5));}
 for(let r=point.r;r<=end.r;r++){const h=Number(rows.get(r)?.getAttribute('ht'))||defaultHeight;height+=h*96/72;}
 return {col:point.c,row:point.r,width:Math.max(1,width-4),height:Math.max(1,height-4)};
}
async function imageDefinitions(zip){
 const source=zip.file('xl/cellimages.xml'),relationships=zip.file('xl/_rels/cellimages.xml.rels');
 if(!source||!relationships)return new Map();
 const doc=parse(await source.async('string')),rels=parse(await relationships.async('string')),targets=new Map(all(rels,'Relationship').filter(n=>n.getAttribute('TargetMode')!=='External').map(n=>[n.getAttribute('Id'),partPath(n.getAttribute('Target'),'xl/cellimages.xml')])),definitions=new Map();
 for(const pic of all(doc,'pic')){
  const name=all(pic,'cNvPr')[0]?.getAttribute('name'),blips=all(pic,'blip'),ext=all(pic,'xfrm')[0]&&all(all(pic,'xfrm')[0],'ext')[0];
  if(!name||blips.length!==1||definitions.has(name)){if(name)definitions.set(name,null);continue;}
  const media=targets.get(blips[0].getAttributeNS(NS.r,'embed')),cx=Number(ext?.getAttribute('cx')),cy=Number(ext?.getAttribute('cy'));
  definitions.set(name,media&&zip.file(media)&&cx>0&&cy>0?{pic,media,cx,cy}:null);
 }
 return definitions;
}
async function addImages(zip,sheet,plans,types){
 let sheetRels=zip.file(relPath(sheet.path))?parse(await zip.file(relPath(sheet.path)).async('string')):emptyRels();
 const drawingNodes=children(sheet.doc.documentElement,'drawing');
 if(drawingNodes.length>1)throw Error(sheet.name+'：发现多个绘图关系，无法安全保存图片');
 let path,doc,rels;
 if(drawingNodes.length){
  const id=drawingNodes[0].getAttributeNS(NS.r,'id'),relationship=all(sheetRels,'Relationship').find(n=>n.getAttribute('Id')===id);
  if(!relationship||relationship.getAttribute('TargetMode')==='External')throw Error(sheet.name+'：图片绘图关系不完整');
  path=partPath(relationship.getAttribute('Target'),sheet.path);
  if(!zip.file(path))throw Error(sheet.name+'：图片绘图文件不存在');
  doc=parse(await zip.file(path).async('string'));rels=zip.file(relPath(path))?parse(await zip.file(relPath(path)).async('string')):emptyRels();
 }else{
  let index=1;while(zip.file('xl/drawings/drawing'+index+'.xml'))index++;
  path='xl/drawings/drawing'+index+'.xml';doc=parse('<xdr:wsDr xmlns:xdr="'+NS.xdr+'" xmlns:a="'+NS.a+'" xmlns:r="'+NS.r+'"/>');rels=emptyRels();
  const id=nextId(sheetRels);sheetRels.documentElement.append(el(sheetRels,NS.p,'Relationship',{Id:id,Type:NS.r+'/drawing',Target:relative(path,sheet.path)}));
  const drawing=el(sheet.doc,NS.s,'drawing');drawing.setAttributeNS(NS.r,'r:id',id);
  const following=[...sheet.doc.documentElement.children].find(n=>['legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'].includes(n.localName));
  sheet.doc.documentElement.insertBefore(drawing,following||null);
  types.documentElement.append(el(types,NS.ct,'Override',{PartName:'/'+path,ContentType:'application/vnd.openxmlformats-officedocument.drawing+xml'}));
 }
 let shapeId=Math.max(0,...all(doc,'cNvPr').map(n=>Number(n.getAttribute('id'))||0))+1;
 for(const plan of plans){
  const id=nextId(rels);rels.documentElement.append(el(rels,NS.p,'Relationship',{Id:id,Type:NS.r+'/image',Target:relative(plan.definition.media,path)}));
  const box=cellBox(sheet.doc,plan.cell),ratio=Math.min(box.width*9525/plan.definition.cx,box.height*9525/plan.definition.cy),cx=Math.max(1,Math.round(plan.definition.cx*ratio)),cy=Math.max(1,Math.round(plan.definition.cy*ratio));
  const anchor=el(doc,NS.xdr,'xdr:oneCellAnchor'),from=el(doc,NS.xdr,'xdr:from');
  appendText(doc,from,'col',box.col);appendText(doc,from,'colOff',19050+Math.round((box.width*9525-cx)/2));appendText(doc,from,'row',box.row);appendText(doc,from,'rowOff',19050+Math.round((box.height*9525-cy)/2));
  anchor.append(from,el(doc,NS.xdr,'xdr:ext',{cx,cy}));
  const pic=doc.importNode(plan.definition.pic,true),properties=all(pic,'cNvPr')[0];properties.setAttribute('id',String(shapeId++));properties.setAttribute('name',plan.cell+' '+plan.name);
  all(pic,'blip')[0].setAttributeNS(NS.r,'r:embed',id);
  const transform=all(pic,'xfrm')[0];for(const off of all(transform,'off')){off.setAttribute('x','0');off.setAttribute('y','0');}for(const ext of all(transform,'ext')){ext.setAttribute('cx',String(cx));ext.setAttribute('cy',String(cy));}
  anchor.append(pic,el(doc,NS.xdr,'xdr:clientData'));doc.documentElement.append(anchor);
 }
 zip.file(path,xml(doc));zip.file(relPath(path),xml(rels));zip.file(relPath(sheet.path),xml(sheetRels));
}
async function prepare(input){
 const original=input instanceof ArrayBuffer?input:input.buffer.slice(input.byteOffset,input.byteOffset+input.byteLength);
 let bytes=original,book=XLSX.read(bytes,{type:'array',cellDates:false,cellNF:true,cellStyles:true});
 const signature=new Uint8Array(bytes),converted=signature[0]!==80||signature[1]!==75;
 if(converted)bytes=XLSX.write(book,{type:'array',bookType:'xlsx'});
 const zip=await JSZip.loadAsync(bytes),workbookPart=zip.file('xl/workbook.xml'),relsPart=zip.file('xl/_rels/workbook.xml.rels');
 if(!workbookPart||!relsPart)throw Error('文件缺少 XLSX 工作簿结构');
 const workbook=parse(await workbookPart.async('string')),workbookRels=parse(await relsPart.async('string')),targets=new Map(all(workbookRels,'Relationship').map(n=>[n.getAttribute('Id'),n])),sheets=[],issues=[],images=[];
 let formulaCount=0;
 for(const node of all(workbook,'sheet')){
  const name=node.getAttribute('name'),relation=targets.get(node.getAttributeNS(NS.r,'id'));
  if(!relation||relation.getAttribute('TargetMode')==='External')throw Error(name+'：工作表关系无法读取');
  const path=partPath(relation.getAttribute('Target'),'xl/workbook.xml'),part=zip.file(path);if(!part)throw Error(name+'：工作表文件不存在');
  const doc=parse(await part.async('string')),sheet=book.Sheets[name];if(!sheet)throw Error(name+'：工作表数据无法读取');
  sheet['!fulfillmentFilterActive']=all(doc,'filterColumn').length>0||all(doc,'sheetPr').some(n=>['1','true'].includes(n.getAttribute('filterMode')));
  if(converted&&sheet['!autofilter']&&!sheet['!fulfillmentFilterActive'])delete sheet['!fulfillmentFilterActive'];
  const cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c]));sheets.push({name,path,doc,cells});
  const shared=new Map();for(const [address,cell] of cells)for(const f of children(cell,'f').filter(n=>n.getAttribute('t')==='shared')){const id=f.getAttribute('si');if(!shared.has(id))shared.set(id,[]);shared.get(id).push({address,f});}
  for(const [id,group] of shared){
   const masters=group.filter(entry=>entry.f.textContent.trim()),master=masters[0];let range;
   try{if(id==null||masters.length!==1||!master.f.getAttribute('ref'))throw Error();range=XLSX.utils.decode_range(master.f.getAttribute('ref'));}catch{issues.push(issue(name,group[0].address,'共享公式定义不完整，无法安全展开。'));continue;}
   for(const entry of group){const point=XLSX.utils.decode_cell(entry.address);if(point.r<range.s.r||point.r>range.e.r||point.c<range.s.c||point.c>range.e.c)issues.push(issue(name,entry.address,'共享公式超出定义范围，无法安全展开。'));}
  }
  for(const [address,cell] of cells){
   const formulas=children(cell,'f');if(!formulas.length)continue;formulaCount++;
   const f=formulas[0],kind=f.getAttribute('t'),value=sheet[address],text=String(value?.f||f.textContent||'');
   if(formulas.length!==1||['array','dataTable'].includes(kind)){issues.push(issue(name,address,'数组或数据表公式无法安全转为数值，请先在 WPS/Excel 计算并粘贴为值。'));continue;}
   if(!value||typeof value.f!=='string'||!value.f.trim()){issues.push(issue(name,address,'公式未完整读取（可能为共享公式），不能使用保存的旧结果。'));continue;}
   if(/DISPIMG\s*\(/i.test(text)){
    const match=/^(?:_xlfn\.)?DISPIMG\(\s*"([^"]+)"\s*,\s*1\s*\)$/i.exec(text.trim());
    if(!match){issues.push(issue(name,address,'图片公式格式无法识别，无法保留图片。'));continue;}
    images.push({sheet:name,cell:address,name:match[1]});
   }
  }
 }
 const failed=()=>({book,bytes:original,issues,formulaCount,correctedCount:0,imageCount:0,converted});
 if(issues.length)return failed();
 const tableParts=[];for(const path of Object.keys(zip.files).filter(p=>/^xl\/tables\/[^/]+\.xml$/.test(p))){const doc=parse(await zip.file(path).async('string')),formulas=[...all(doc,'calculatedColumnFormula'),...all(doc,'totalsRowFormula')];if(formulas.length)tableParts.push({path,doc,formulas});}
 const hasChain=!!zip.file('xl/calcChain.xml')||all(workbookRels,'Relationship').some(n=>/\/calcChain$/.test(n.getAttribute('Type')||''));
 if(!formulaCount&&!tableParts.length&&!hasChain)return {book,bytes,issues:[],formulaCount:0,correctedCount:0,imageCount:0,converted};
 if(images.length){
  const definitions=await imageDefinitions(zip);
  for(const image of images){image.definition=definitions.get(image.name);if(!image.definition)issues.push(issue(image.sheet,image.cell,'图片资源或尺寸无法识别，已停止处理以保留原图片。'));}
  if(issues.length)return failed();
  for(const image of images){const cell=book.Sheets[image.sheet][image.cell];delete cell.f;delete cell.F;delete cell.D;delete cell.w;cell.t='z';delete cell.v;}
 }
 const materialized=FulfillmentValidation.materialize(book);book=materialized.book;issues.push(...materialized.issues);
 if(issues.length)return failed();
 const updates=new Map();for(const update of materialized.updates){const key=update.sheet+'\0'+update.cell;if(updates.has(key))throw Error('公式转换结果出现重复单元格');updates.set(key,update);}
 for(const image of images)updates.set(image.sheet+'\0'+image.cell,{...image,type:'z',value:null});
 if(updates.size!==formulaCount)throw Error('公式数量与转换结果不一致，未生成文件');
 const types=parse(await zip.file('[Content_Types].xml').async('string'));
 for(const sheet of sheets){
  let changed=false;
  for(const [address,cell] of sheet.cells){if(!children(cell,'f').length)continue;const update=updates.get(sheet.name+'\0'+address);if(!update)throw Error(sheet.name+'!'+address+' 公式没有核算结果');replaceValue(sheet.doc,cell,update.value,update.type);changed=true;}
  const plans=images.filter(image=>image.sheet===sheet.name);if(plans.length){await addImages(zip,sheet,plans,types);changed=true;}
  if(changed)zip.file(sheet.path,xml(sheet.doc));
 }
 // Table formulas can repopulate values on opening. Their already-calculated cells above remain intact.
 for(const table of tableParts){table.formulas.forEach(n=>n.remove());zip.file(table.path,xml(table.doc));}
 for(const node of all(workbookRels,'Relationship').filter(n=>/\/calcChain$/.test(n.getAttribute('Type')||''))){const path=partPath(node.getAttribute('Target'),'xl/workbook.xml');zip.remove(path);for(const type of all(types,'Override').filter(n=>n.getAttribute('PartName')==='/'+path))type.remove();node.remove();}
 if(zip.file('xl/calcChain.xml'))zip.remove('xl/calcChain.xml');
 for(const type of all(types,'Override').filter(n=>/\/calcChain\.xml$/.test(n.getAttribute('PartName')||'')))type.remove();
 zip.file('xl/_rels/workbook.xml.rels',xml(workbookRels));zip.file('[Content_Types].xml',xml(types));
 bytes=await zip.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
 const check=XLSX.read(bytes,{type:'array',cellDates:false,cellNF:true});
 for(const [name,sheet] of Object.entries(check.Sheets))for(const [address,cell] of Object.entries(sheet))if(/^[A-Z]+[1-9]\d*$/.test(address)&&cell&&(cell.f!=null||cell.F!=null||cell.D))throw Error(name+'!'+address+' 仍有公式，文件未通过完整性检查');
 for(const update of updates.values()){const cell=check.Sheets[update.sheet]?.[update.cell];if(update.type==='z'){if(cell&&cell.v!=null&&cell.v!=='')throw Error(update.sheet+'!'+update.cell+' 图片公式未清除');}else if(!cell||cell.t!==update.type||cell.v!==update.value)throw Error(update.sheet+'!'+update.cell+' 转换结果重读不一致');}
 // Drawing names and internal relationships are checked before returning the downloadable file.
 for(const sheet of sheets){const plans=images.filter(image=>image.sheet===sheet.name);if(!plans.length)continue;
  const sheetRels=parse(await zip.file(relPath(sheet.path)).async('string')),drawing=children(sheet.doc.documentElement,'drawing')[0],relation=all(sheetRels,'Relationship').find(n=>n.getAttribute('Id')===drawing.getAttributeNS(NS.r,'id')),path=partPath(relation.getAttribute('Target'),sheet.path),doc=parse(await zip.file(path).async('string')),rels=parse(await zip.file(relPath(path)).async('string'));
  for(const plan of plans){const pic=all(doc,'pic').find(p=>all(p,'cNvPr')[0]?.getAttribute('name')===plan.cell+' '+plan.name),id=pic&&all(pic,'blip')[0]?.getAttributeNS(NS.r,'embed'),imageRel=all(rels,'Relationship').find(n=>n.getAttribute('Id')===id);if(!imageRel||!zip.file(partPath(imageRel.getAttribute('Target'),path)))throw Error(sheet.name+'!'+plan.cell+' 图片未完整保留');}
 }
 return {book,bytes,issues:[],formulaCount,correctedCount:materialized.correctedCount,imageCount:images.length,converted};
}
window.FulfillmentValues={prepare};
})();
