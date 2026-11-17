(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'),require('./vendor/jszip.min.js'),require('./fulfillment-check-core.js'));else root.FulfillmentCheckExport=factory(root.XLSX,root.JSZip,root.FulfillmentCheckCore);})(typeof globalThis!=='undefined'?globalThis:this,function(X,Zip,C){
'use strict';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
const dec=v=>v.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-f]+|#\d+);/gi,(_,k)=>k[0]==='#'?String.fromCodePoint(parseInt(k.slice(k[1].toLowerCase()==='x'?2:1),k[1].toLowerCase()==='x'?16:10)):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[k]));
const re=n=>new RegExp('<((?:[\\w.-]+:)?'+n+')\\b[^>]*?(?:/>|>[\\s\\S]*?</\\1\\s*>)','g');
const tags=(x,n)=>x.match(re(n))||[],first=(x,n)=>tags(x,n)[0]||'',body=x=>x.replace(/^<[^>]+>/,'').replace(/<\/[^>]+>$/,''),head=x=>x.match(/^<[^>]+>/)[0],prefix=x=>x.match(/^<([\w.-]+:)?/)?.[1]||'';
const attr=(x,n)=>{const m=head(x).match(new RegExp('\\s'+n+'=(["\'])([\\s\\S]*?)\\1'));return m?dec(m[2]):null;};
function set(x,n,v){const h=head(x),rx=new RegExp('\\s'+n+'=(["\'])[\\s\\S]*?\\1');return (v===null?h.replace(rx,''):rx.test(h)?h.replace(rx,' '+n+'="'+esc(v)+'"'):h.replace(/\/?>(?=$)/,end=>' '+n+'="'+esc(v)+'"'+end))+x.slice(h.length);}
const qualify=(x,p)=>x.replace(/<(\/?)([A-Za-z][\w]*)(?=[\s/>])/g,'<$1'+p+'$2');
const append=(x,n,items)=>x.replace(new RegExp('</((?:[\\w.-]+:)?'+n+')\\s*>'),(_,tag)=>qualify(items,tag.includes(':')?tag.split(':')[0]+':':'')+'</'+tag+'>');
const relpath=path=>path.replace(/([^/]+)$/,'_rels/$1.rels');
function resolve(path,target){const out=[];for(const p of (target.startsWith('/')?target.slice(1):path.slice(0,path.lastIndexOf('/')+1)+target).split('/')){if(p==='..'){if(!out.length)throw Error('来源文件包含无效路径');out.pop();}else if(p&&p!=='.')out.push(p);}return out.join('/');}
async function read(zip,path){if(!zip.file(path))throw Error('来源文件缺少 '+path+'，无法保真导出');return zip.file(path).async('string');}
function appendStyle(xml,name,items){
 const child={fonts:'font',fills:'fill',borders:'border',numFmts:'numFmt',cellStyleXfs:'xf',cellXfs:'xf',dxfs:'dxf',cellStyles:'cellStyle'}[name];
 const old=first(xml,name),offset=tags(body(old),child).length;
 let next=old&&!/\/>$/.test(old)?append(old,name,items):`<${name}>${items}</${name}>`;
 next=set(next,'count',offset+tags(items,child).length);
 if(old)return {xml:xml.replace(old,next),offset};
 const order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst'];
 const after=order.slice(order.indexOf(name)+1).map(n=>first(xml,n)).find(Boolean);
 return {xml:after?xml.replace(after,next+after):append(xml,'styleSheet',next),offset};
}
function explicitTheme(xml,theme){
 const colors=['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].map(n=>{const node=first(theme,n),c=first(node,'srgbClr')||first(node,'sysClr');return c&&(attr(c,'lastClr')||attr(c,'val'));});
 return xml.replace(re('color'),node=>{const index=attr(node,'theme');if(index===null)return node;if(!colors[+index])throw Error('采购来源主题颜色无法解析');return set(set(node,'theme',null),'rgb','FF'+colors[+index]);})
 .replace(/<(?:[\w.-]+:)?(?:fgColor|bgColor)\b[^>]*\/>/g,node=>{const index=attr(node,'theme');if(index===null)return node;if(!colors[+index])throw Error('采购来源主题颜色无法解析');return set(set(node,'theme',null),'rgb','FF'+colors[+index]);})
 // Keep the source's explicit font name, rather than adopting the other workbook's theme font.
 .replace(re('scheme'),'');
}
function mergeStyles(base,donor,theme){
 const counts={},maps={},formats=new Map();
 let nextId=Math.max(163,...tags(first(base,'numFmts'),'numFmt').map(n=>+attr(n,'numFmtId')))+1;
 const numItems=tags(first(donor,'numFmts'),'numFmt').map(n=>{const id=nextId++;formats.set(+attr(n,'numFmtId'),id);return set(n,'numFmtId',id);}).join('');
 if(numItems)base=appendStyle(base,'numFmts',numItems).xml;
 for(const name of ['fonts','fills','borders','cellStyleXfs','cellXfs','dxfs']){
  const child={fonts:'font',fills:'fill',borders:'border',cellStyleXfs:'xf',cellXfs:'xf',dxfs:'dxf'}[name];
  const items=tags(body(first(donor,name)),child).map(n=>{
   n=explicitTheme(n,theme);
   if(name==='cellStyleXfs'||name==='cellXfs'){
    for(const [attribute,section] of [['fontId','fonts'],['fillId','fills'],['borderId','borders'],['xfId','cellStyleXfs']])if(attr(n,attribute)!==null&&counts[section]!==undefined)n=set(n,attribute,+attr(n,attribute)+counts[section]);
    if(formats.has(+attr(n,'numFmtId')))n=set(n,'numFmtId',formats.get(+attr(n,'numFmtId')));
   }
   return n;
  });
  const merged=appendStyle(base,name,items.join(''));base=merged.xml;counts[name]=merged.offset;maps[name]=items;
 }
 return {xml:base,offset:counts.cellXfs,dxfOffset:counts.dxfs};
}
const numberFormat=value=>{
 const precision=(value.toLocaleString('en-US',{useGrouping:false,maximumSignificantDigits:21}).split('.')[1]||'').length;
 return precision>250?'0.0000000000000000E+00':precision?'0.'+'0'.repeat(precision):'0';
};
function addResultStyles(xml,rows){
 let m=appendStyle(xml,'fonts','<font><sz val="11"/><name val="等线"/><color rgb="FF172033"/></font><font><b/><sz val="11"/><name val="等线"/><color rgb="FFFFFFFF"/></font>');xml=m.xml;const font=m.offset;
 m=appendStyle(xml,'fills','<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315B9D"/><bgColor indexed="64"/></patternFill></fill>');xml=m.xml;const fill=m.offset;
 m=appendStyle(xml,'borders',['thin','hair'].map(s=>'<border>'+['left','right','top','bottom'].map(side=>`<${side} style="${s}"><color rgb="FF404040"/></${side}>`).join('')+'<diagonal/></border>').join(''));xml=m.xml;const border=m.offset;
 let numId=Math.max(163,...tags(first(xml,'numFmts'),'numFmt').map(n=>+attr(n,'numFmtId')))+1;
 const formats=new Map([['0',1]]);
 for(const row of rows)for(const value of row)if(typeof value==='number'){
  if(!Number.isFinite(value))throw Error('核对数量超出有效范围');
  const format=numberFormat(value);if(!formats.has(format)){formats.set(format,numId);xml=appendStyle(xml,'numFmts',`<numFmt numFmtId="${numId++}" formatCode="${format}"/>`).xml;}
 }
 const xf=(header,wrap,format=49)=>`<xf numFmtId="${format}" fontId="${font+(header?1:0)}" fillId="${fill+(header?1:0)}" borderId="${border+(header?0:1)}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="${wrap?'left':'center'}" vertical="center" wrapText="${wrap?1:0}"/></xf>`;
 m=appendStyle(xml,'cellXfs',xf(true,false)+xf(false,false)+xf(false,true)+[...formats.values()].map(id=>xf(false,false,id)).join(''));
 return {xml:m.xml,header:m.offset,text:m.offset+1,wrap:m.offset+2,numbers:new Map([...formats.keys()].map((format,i)=>[format,m.offset+3+i]))};
}
const width=v=>Math.max(...String(v??'').split('\n').map(line=>[...line].reduce((n,c)=>n+(c.codePointAt(0)>255?2:1),0)));
function resultSheet(headers,rows,styles,detail){
 const data=[headers,...rows],notes=new Set(detail?[8,13,14,16]:[]);
 const widths=headers.map((h,c)=>Math.min(255,Math.max(width(h)+5,...rows.map(row=>Math.min(notes.has(c)?60:100,width(row[c])+3)))));
 const cellText=v=>esc(String(v).replace(/_x[0-9a-f]{4}_/gi,m=>'_x005F_'+m.slice(1)).replace(/[\u0000-\u0008\u000B-\u001F]/g,''));
 const xml=data.map((row,r)=>{
  const lines=r?Math.max(1,...row.map((v,c)=>notes.has(c)?String(v??'').split('\n').reduce((n,l)=>n+Math.max(1,Math.ceil(width(l)/(widths[c]-2))),0):1)):1;
  // Excel caps a row at 409 points; full text also remains in the cell and edit bar.
  return `<row r="${r+1}" ht="${Math.min(409,Math.max(28,lines*16+8))}" customHeight="1">`+headers.map((_,c)=>{
   const v=row[c],s=r===0?styles.header:typeof v==='number'?styles.numbers.get(numberFormat(v)):notes.has(c)?styles.wrap:styles.text,a=X.utils.encode_cell({r,c});
   if(v===null||v===undefined)return `<c r="${a}" s="${s}"/>`;
   if(typeof v==='number'){if(!Number.isFinite(v))throw Error('核对数量超出有效范围');return `<c r="${a}" s="${s}"><v>${v}</v></c>`;}
   if(String(v).length>32767)throw Error('核对说明超过 Excel 单元格上限，请缩小文件范围后重试');
   return `<c r="${a}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${cellText(v)}</t></is></c>`;
  }).join('')+'</row>';
 }).join('');
 const ref='A1:'+X.utils.encode_cell({r:rows.length,c:headers.length-1});
 return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${S}"><dimension ref="${ref}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="28"/><cols>${widths.map((w,c)=>`<col min="${c+1}" max="${c+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${xml}</sheetData><autoFilter ref="${ref}"/></worksheet>`;
}
async function packageInfo(bytes){
 const zip=await Zip.loadAsync(bytes),wb=await read(zip,'xl/workbook.xml'),rels=await read(zip,'xl/_rels/workbook.xml.rels'),types=await read(zip,'[Content_Types].xml');
 if(!types.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'))throw Error('仅支持原始 .xlsx 文件，不能保真合并此格式');
 const entries=tags(first(wb,'sheets'),'sheet'),relationships=tags(rels,'Relationship');
 const relation=type=>relationships.find(r=>attr(r,'Type').endsWith('/'+type));
 const path=type=>relation(type)?resolve('xl/workbook.xml',attr(relation(type),'Target')):null;
 const sheetPaths=entries.map(s=>{const rel=relationships.find(r=>attr(r,'Id')===attr(s,'(?:[\\w.-]+:)?id'));if(!rel||!attr(rel,'Type').endsWith('/worksheet'))throw Error('存在非普通工作表，暂不能保真合并');return resolve('xl/workbook.xml',attr(rel,'Target'));});
 return {zip,wb,rels,types,entries,names:entries.map(s=>attr(s,'name')),sheetPaths,relationships,path};
}
return async function(purchaseBytes,fulfillmentBytes,result){
 if(!result?.canExport)throw Error('尚未完成核对');
 const base=await packageInfo(fulfillmentBytes),donor=await packageInfo(purchaseBytes),zip=base.zip;
 if(base.names.some(n=>['汇总','差异明细'].includes(n)))throw Error('履约原文件已有“汇总”或“差异明细”，请先重命名原工作表，避免覆盖原内容');
 if((attr(first(base.wb,'workbookPr')||'<workbookPr/>','date1904')||'0')!==(attr(first(donor.wb,'workbookPr')||'<workbookPr/>','date1904')||'0'))throw Error('两份文件日期系统不同，无法保持原日期，请统一后上传');
 if(donor.relationships.some(r=>!/\/(worksheet|styles|sharedStrings|theme|calcChain|sheetMetadata)$/.test(attr(r,'Type'))))throw Error('采购文件包含外部连接或其他工作簿级功能，暂不能保真合并，请使用原始采购订单导出');
 const donorNames=tags(first(donor.wb,'definedNames'),'definedName');
 if(donorNames.some(n=>attr(n,'localSheetId')===null)||first(donor.wb,'externalReferences'))throw Error('采购文件含全局定义名称或外部引用，暂不能保真合并，请使用原始采购订单导出');
 const usedNames=new Set(['汇总','差异明细',...base.names].map(n=>n.toLowerCase())),names=donor.names.map(name=>{let next=name,i=1;while(usedNames.has(next.toLowerCase()))next=('采购-'+name).slice(0,27)+'-'+i++;usedNames.add(next.toLowerCase());return next;});
 // References that depend on worksheet indices or named ranges cannot be blindly moved between workbooks.
 const donorXml=await Promise.all(donor.sheetPaths.map(path=>read(donor.zip,path)));
 if(donorXml.some(x=>/<(?:\w+:)?c\b[^>]*\s(?:cm|vm)=/.test(x)))throw Error('采购表含动态数组或单元格元数据，暂不能保真合并');
 if(donorXml.some(x=>/<(?:\w+:)?f\b/.test(x))&&names.some((n,i)=>n!==donor.names[i]))throw Error('采购表重名且含公式，请先重命名来源表后上传，避免公式引用改变');
 if(names.some((n,i)=>n!==donor.names[i])&&donorXml.some(x=>/<(?:\w+:)?(?:formula1|formula2)\b|<hyperlink\b[^>]*\blocation=/.test(x)))throw Error('采购表重名且含内部链接或验证公式，请先重命名来源表后上传');
 if(donorXml.some(x=>/<(?:\w+:)?(?:extLst|tableParts)\b/.test(x)))throw Error('采购表含扩展或结构化表，暂不能保真合并，请使用原始采购订单导出');
 const stylePath=base.path('styles'),donorStyle=await read(donor.zip,donor.path('styles'));
 if(first(donorStyle,'colors')&&first(donorStyle,'colors')!==first(await read(zip,stylePath),'colors'))throw Error('采购表使用不同的自定义调色板，无法保持原颜色');
 const merged=mergeStyles(await read(zip,stylePath),donorStyle,donor.path('theme')?await read(donor.zip,donor.path('theme')):''),styles=addResultStyles(merged.xml,[...result.summary,...result.detail]);
 const strings=donor.path('sharedStrings')?tags(await read(donor.zip,donor.path('sharedStrings')),'si'):[];
 let folder='xl/fulfillment-check-import/';while(Object.keys(zip.files).some(p=>p.startsWith(folder)))folder=folder.replace(/\/$/,'x/');
 let wb=base.wb,rels=base.rels,types=base.types;
 const newEntries=[],idStart=Math.max(...base.entries.map(e=>+attr(e,'sheetId')))+1;
 const register=(name,path,index,state)=>{
  let id='rIdFulfillmentCheck'+index;while(tags(rels,'Relationship').some(r=>attr(r,'Id')===id))id+='x';
  rels=append(rels,'Relationships',`<Relationship Id="${id}" Type="${R}/worksheet" Target="/${path}"/>`);
  if(!tags(types,'Override').some(n=>attr(n,'PartName')==='/'+path))types=append(types,'Types',`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  const source=base.entries[0],rp=head(source).match(/\s((?:[\w.-]+:)?id)=/)?.[1]||'r:id';
  newEntries.push(`<${prefix(source)}sheet name="${esc(name)}" sheetId="${idStart+index}" ${rp}="${id}"${state?' state="'+esc(state)+'"':''}/>`);
 };
 for(const [index,name] of ['汇总','差异明细'].entries()){
  const path=folder+'result'+index+'.xml';zip.file(path,resultSheet(index?C.DETAIL:C.SUMMARY,index?result.detail:result.summary,styles,Boolean(index)));register(name,path,index);
 }
 const copied=new Set(),allowed=new Set(['comments','vmlDrawing','drawing','image','printerSettings','hyperlink']);
 async function copyRelated(path){
  if(copied.has(path))return;copied.add(path);
  zip.file(folder+path,await donor.zip.file(path)?.async('uint8array')||await read(donor.zip,path));
  const type=tags(donor.types,'Override').find(n=>attr(n,'PartName')==='/'+path);
  if(type)types=append(types,'Types',set(type,'PartName','/'+folder+path));
  else{const extension=path.split('.').pop(),d=tags(donor.types,'Default').find(n=>attr(n,'Extension')===extension);if(d&&!tags(types,'Default').some(n=>attr(n,'Extension')===extension))types=append(types,'Types',d);else if(d&&tags(types,'Default').some(n=>attr(n,'Extension')===extension&&attr(n,'ContentType')!==attr(d,'ContentType')))throw Error('附件类型冲突，无法保持原文件');}
  const rp=relpath(path);if(!donor.zip.file(rp))return;
  let xml=await read(donor.zip,rp);
  for(const rel of tags(xml,'Relationship')){
   const kind=attr(rel,'Type').split('/').pop();if(!allowed.has(kind))throw Error('采购表含不支持的附件类型：'+kind+'，已停止导出以保留原内容');
   if(attr(rel,'TargetMode')==='External')continue;
   const target=resolve(path,attr(rel,'Target'));await copyRelated(target);
   xml=xml.replace(rel,set(rel,'Target','/'+folder+target));
  }
  zip.file(folder+rp,xml);
 }
 for(let i=0;i<donor.entries.length;i++){
  const path=donor.sheetPaths[i];await copyRelated(path);
  let xml=donorXml[i];
  xml=xml.replace(re('c'),node=>{
   node=set(node,'s',+(attr(node,'s')||0)+merged.offset);
   if(attr(node,'t')==='s'){const value=first(node,'v'),item=strings[+body(value)];if(!item)throw Error('采购表共享文本索引无效');node=set(node,'t','inlineStr').replace(value,'<is>'+body(item)+'</is>');}
   return node;
  });
  for(const [tag,attribute] of [['row','s'],['col','style']])xml=xml.replace(re(tag),n=>attr(n,attribute)===null?n:set(n,attribute,+attr(n,attribute)+merged.offset));
  xml=xml.replace(re('cfRule'),n=>attr(n,'dxfId')===null?n:set(n,'dxfId',+attr(n,'dxfId')+merged.dxfOffset));
  // Theme references in drawings require a drawing-level theme override; reject rather than change appearance.
  for(const part of copied)if(/drawings\/.*\.xml$/.test(part)&&/<(?:\w+:)?(?:schemeClr|fontRef)\b/.test(await read(donor.zip,part)))throw Error('采购图片或图形引用主题，暂不能保真迁移，请使用无主题图形的导出');
  zip.file(folder+path,xml);register(names[i],folder+path,i+2,attr(donor.entries[i],'state'));
 }
 const originalSheets=first(wb,'sheets');wb=wb.replace(originalSheets,head(originalSheets)+newEntries.join('')+body(originalSheets)+`</${prefix(originalSheets)}sheets>`);
 wb=wb.replace(re('definedName'),n=>attr(n,'localSheetId')===null?n:set(n,'localSheetId',+attr(n,'localSheetId')+newEntries.length));
 const localNames=donorNames.map(n=>{const i=+attr(n,'localSheetId');if(!names[i])throw Error('采购局部名称的工作表无效');if(names[i]!==donor.names[i])throw Error('采购表重名且含局部名称，请先重命名来源表，避免引用改变');return set(n,'localSheetId',i+2);}).join('');
 if(localNames){if(first(wb,'definedNames'))wb=append(wb,'definedNames',localNames);else wb=wb.replace(/<\/((?:[\w.-]+:)?sheets)>/,m=>m+'<definedNames>'+localNames+'</definedNames>');}
 wb=wb.replace(re('workbookView'),n=>set(set(n,'activeTab',0),'firstSheet',0));
 zip.file('xl/workbook.xml',wb);zip.file('xl/_rels/workbook.xml.rels',rels);zip.file(stylePath,styles.xml);zip.file('[Content_Types].xml',types);
 if(zip.file('docProps/app.xml')){
  let app=await read(zip,'docProps/app.xml'),titles=first(app,'TitlesOfParts'),vector=first(titles,'vector');
  if(vector){const p=prefix(first(vector,'lpstr')||'<vt:lpstr>'),order=['汇总','差异明细',...names,...base.names];app=app.replace(vector,set(head(vector),'size',order.length)+order.map(n=>`<${p}lpstr>${esc(n)}</${p}lpstr>`).join('')+`</${prefix(vector)}vector>`);}
  const pairs=first(app,'HeadingPairs'),variants=tags(pairs,'variant');for(let i=0;i<variants.length-1;i++)if(dec(body(first(variants[i],'lpstr')))==='Worksheets'){const n=first(variants[i+1],'i4');app=app.replace(n,head(n)+(base.names.length+newEntries.length)+`</${prefix(n)}i4>`);}
  zip.file('docProps/app.xml',app);
 }
 const bytes=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),check=X.read(bytes,{type:'array'});
 if(check.SheetNames.length!==base.names.length+donor.names.length+2||check.SheetNames[0]!=='汇总'||check.SheetNames[1]!=='差异明细')throw Error('导出工作表校验失败');
 return bytes;
};
});
