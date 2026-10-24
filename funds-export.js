/* Keep the negotiated worksheet's native XML and patch only approved fields. */
window.FundsExport = async function(sourceBytes, result) {
  'use strict';
  window.FundsExport.lastWarnings=[];
  const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const P='http://schemas.openxmlformats.org/package/2006/relationships';
  const parser=new DOMParser(), serializer=new XMLSerializer();
  const elements=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)];
  const first=(doc,name)=>elements(doc,name)[0];
  const parseText=text=>{const doc=parser.parseFromString(text,'application/xml');if(first(doc,'parsererror'))throw Error('工作簿 XML 无法解析');return doc;};
  const parse=async(zip,path)=>{if(!zip.file(path))throw Error('工作簿缺少 '+path);return parseText(await zip.file(path).async('string'));};
  const save=(path,doc)=>zip.file(path,serializer.serializeToString(doc));
  const create=(doc,name,attrs={})=>{const e=doc.createElementNS(S,name);for(const [key,value] of Object.entries(attrs))e.setAttribute(key,String(value));return e;};
  const resolve=(base,target)=>new URL(target,'https://xlsx.invalid/'+base).pathname.slice(1);
  const relPath=path=>path.replace(/([^/]+)$/, '_rels/$1.rels');
  let zip;
  try {zip=await JSZip.loadAsync(sourceBytes);} catch (_) {throw Error('谈判付款汇总请上传 .xlsx 文件，以保留原表格式及公式');}
  const workbook=await parse(zip,'xl/workbook.xml'),rels=await parse(zip,'xl/_rels/workbook.xml.rels');
  const types=await parse(zip,'[Content_Types].xml');
  const sheets=first(workbook,'sheets'),entries=elements(sheets,'sheet');
  const selected=entries.find(e=>e.getAttribute('name')===result.negotiated.sheetName);
  if(!selected)throw Error('未找到选中的谈判付款工作表');
  const selectedIndex=entries.indexOf(selected),oldName=selected.getAttribute('name');
  const selectedRel=elements(rels,'Relationship').find(e=>e.getAttribute('Id')===selected.getAttributeNS(R,'id'));
  if(!selectedRel)throw Error('谈判付款工作表关系缺失');
  const sheetPath=resolve('xl/workbook.xml',selectedRel.getAttribute('Target'));
  const sheet=await parse(zip,sheetPath),sheetData=first(sheet,'sheetData');
  if(!sheetData)throw Error('谈判付款表缺少单元格数据');
  const removedNames=entries.filter(e=>e!==selected).map(e=>e.getAttribute('name'));
  const quote=name=>"'"+name.replace(/'/g,"''")+"'";
  const escapeRx=value=>value.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  const refers=(formula,name)=>new RegExp('(?:'+escapeRx(quote(name))+'|(?<![\\w\\u4e00-\\u9fff])'+escapeRx(name)+')!','i').test(formula);
  const renameSelf=formula=>formula.replace(new RegExp('(?:'+escapeRx(quote(oldName))+'|(?<![\\w\\u4e00-\\u9fff])'+escapeRx(oldName)+')!','g'),quote('谈判付款')+'!');
  const forbiddenNames=[];
  for(const name of elements(workbook,'definedName')) {
    const local=name.getAttribute('localSheetId');
    if((local!==null&&Number(local)!==selectedIndex)||removedNames.some(n=>refers(name.textContent,n))) {
      forbiddenNames.push(name.getAttribute('name'));name.remove();
    } else {if(local!==null)name.setAttribute('localSheetId','0');name.textContent=renameSelf(name.textContent);}
  }
  for(const formula of [...elements(sheet,'f'),...elements(sheet,'formula'),...elements(sheet,'formula1'),...elements(sheet,'formula2')]) {
    if(removedNames.some(n=>refers(formula.textContent,n))||forbiddenNames.some(n=>new RegExp('\\b'+escapeRx(n)+'\\b').test(formula.textContent)))throw Error('谈判付款表公式引用了其他工作表，无法在只输出两张表时完整保留，请先将相关公式调整为本表引用');
    formula.textContent=renameSelf(formula.textContent);
  }
  for(const link of elements(sheet,'hyperlink'))if(link.hasAttribute('location')) {
    const location=link.getAttribute('location');
    if(removedNames.some(n=>refers(location,n)))throw Error('谈判付款表含指向其他工作表的链接，请先调整该链接');
    link.setAttribute('location',renameSelf(location));
  }
  const styleRel=elements(rels,'Relationship').find(e=>e.getAttribute('Type')===R+'/styles');
  if(!styleRel)throw Error('工作簿样式关系缺失');
  const stylePath=resolve('xl/workbook.xml',styleRel.getAttribute('Target')),styles=await parse(zip,stylePath);
  const list=name=>{let item=first(styles,name);if(!item){item=create(styles,name,{count:0});styles.documentElement.append(item);}return item;};
  const append=(name,item)=>{const target=list(name),id=target.children.length;target.append(item);target.setAttribute('count',target.children.length);return id;};
  const fonts=list('fonts'),fills=list('fills'),borders=list('borders'),xfs=list('cellXfs');
  const font=create(styles,'font');font.append(create(styles,'sz',{val:12}),create(styles,'name',{val:'Microsoft YaHei'}),create(styles,'color',{rgb:'FF182333'}));
  const fontId=append('fonts',font),headerFont=font.cloneNode(true);first(headerFont,'color').setAttribute('rgb','FFFFFFFF');headerFont.append(create(styles,'b'));
  const headerFontId=append('fonts',headerFont);
  const fill=create(styles,'fill'),pattern=create(styles,'patternFill',{patternType:'solid'});pattern.append(create(styles,'fgColor',{rgb:'FF315B9D'}),create(styles,'bgColor',{indexed:64}));fill.append(pattern);const fillId=append('fills',fill);
  const border=create(styles,'border');for(const edge of ['left','right','top','bottom']){const side=create(styles,edge,{style:'thin'});side.append(create(styles,'color',{rgb:'FFD9DEE7'}));border.append(side);}const borderId=append('borders',border);
  let numFmts=first(styles,'numFmts');if(!numFmts){numFmts=create(styles,'numFmts',{count:0});styles.documentElement.insertBefore(numFmts,styles.documentElement.firstChild);}
  const formats=new Map(elements(numFmts,'numFmt').map(e=>[e.getAttribute('formatCode'),Number(e.getAttribute('numFmtId'))]));
  let nextFormat=Math.max(163,...formats.values())+1;
  const formatId=code=>{if(!code||code==='General')return 0;if(formats.has(code))return formats.get(code);const id=nextFormat++;numFmts.append(create(styles,'numFmt',{numFmtId:id,formatCode:code}));numFmts.setAttribute('count',numFmts.children.length);formats.set(code,id);return id;};
  const styleIds=new Map();
  const styleFor=(code='',header=false)=>{const key=header?'header':code;if(styleIds.has(key))return styleIds.get(key);const xf=create(styles,'xf',{numFmtId:formatId(code),fontId:header?headerFontId:fontId,fillId:header?fillId:0,borderId,xfId:0,applyFont:1,applyFill:1,applyBorder:1,applyAlignment:1,applyNumberFormat:1});xf.append(create(styles,'alignment',{horizontal:'center',vertical:'center',wrapText:1}));const id=append('cellXfs',xf);styleIds.set(key,id);return id;};
  const date1904=['1','true'].includes(first(workbook,'workbookPr')?.getAttribute('date1904'));
  const writeCell=(doc,address,cell,styleId,convertDates=false)=>{
    const c=create(doc,'c',{r:address});if(styleId!==undefined)c.setAttribute('s',styleId);
    if(cell?.f){const f=create(doc,'f');f.textContent=String(cell.f).replace(/^=/,'');c.append(f);}
    if(cell?.v===null||cell?.v===undefined)return c;
    let value=cell.v;
    const numeric=typeof value==='number'&&cell.t!=='s';
    if(numeric&&!Number.isFinite(value))throw Error('单元格 '+address+' 的金额或日期不是有效数值');
    if(numeric&&convertDates&&date1904&&cell.z&&XLSX.SSF.is_date(cell.z))value-=1462;
    if(numeric||cell.t==='b'||typeof value==='boolean'){
      if(cell.t==='b'||typeof value==='boolean'){c.setAttribute('t','b');value=value?1:0;}
      const v=create(doc,'v');v.textContent=String(value);c.append(v);
    } else {
      if(cell.t==='e')throw Error('单元格 '+address+' 含错误结果，请先核对');
      c.setAttribute('t','inlineStr');const is=create(doc,'is'),t=create(doc,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=String(value);is.append(t);c.append(is);
    }
    return c;
  };
  const rows=new Map(elements(sheetData,'row').map(r=>[Number(r.getAttribute('r')),r]));
  const cells=new Map(elements(sheetData,'c').map(c=>[c.getAttribute('r'),c]));
  const originalRange=XLSX.utils.decode_range(first(sheet,'dimension')?.getAttribute('ref')||'A1');
  let maxRow=originalRange.e.r,maxColumn=originalRange.e.c;
  const changed=new Set();
  for(const {address,cell} of result.negotiated.updates){
    if(!/^[A-Z]+[1-9]\d*$/.test(address))throw Error('谈判付款更新地址无效');
    const point=XLSX.utils.decode_cell(address),old=cells.get(address);let row=rows.get(point.r+1);
    if(!row){row=create(sheet,'row',{r:point.r+1});const after=[...rows].filter(([r])=>r>point.r+1).sort((a,b)=>a[0]-b[0])[0]?.[1];sheetData.insertBefore(row,after||null);rows.set(point.r+1,row);}
    const isHeader=point.r===result.negotiated.headerRow;
    const styleId=old?.getAttribute('s')??styleFor(cell?.z||'',isHeader);
    const next=writeCell(sheet,address,cell,styleId);
    // Retain unrelated cell metadata while replacing only value/formula payloads.
    if(old){for(const attr of old.attributes)if(!['r','s','t'].includes(attr.name))next.setAttribute(attr.name,attr.value);old.replaceWith(next);}
    else {const after=[...row.children].find(c=>c.localName==='c'&&XLSX.utils.decode_cell(c.getAttribute('r')).c>point.c);row.insertBefore(next,after||null);}
    cells.set(address,next);changed.add(address);maxRow=Math.max(maxRow,point.r);maxColumn=Math.max(maxColumn,point.c);
    // New multi-line fields must remain visible without reducing the source row height.
    if(point.c>originalRange.e.c){const lines=String(cell?.v??'').split('\n').length;const needed=lines*17+6;if(needed>Number(row.getAttribute('ht')||15)){row.setAttribute('ht',needed);row.setAttribute('customHeight','1');}}
  }
  let dimension=first(sheet,'dimension');if(!dimension){dimension=create(sheet,'dimension');sheet.documentElement.insertBefore(dimension,sheet.documentElement.firstChild);}dimension.setAttribute('ref',XLSX.utils.encode_range({s:originalRange.s,e:{r:maxRow,c:maxColumn}}));
  if(maxColumn>originalRange.e.c){let cols=first(sheet,'cols');if(!cols){cols=create(sheet,'cols');sheet.documentElement.insertBefore(cols,sheetData);}cols.append(create(sheet,'col',{min:originalRange.e.c+2,max:maxColumn+1,width:28,customWidth:1}));}
  // Update simple dependent caches; unsupported affected formulas are recalculated by Excel/WPS.
  const formulaCells=[...cells].filter(([,c])=>first(c,'f'));
  const affected=new Set(changed);let grew=true;
  while(grew){grew=false;for(const [address,c] of formulaCells){const f=first(c,'f'),formula=f.textContent.toUpperCase().replace(/\$/g,'');const refs=formula.match(/\b[A-Z]{1,3}[1-9]\d*\b/g)||[];
    for(const match of formula.matchAll(/\b([A-Z]{1,3}[1-9]\d*):([A-Z]{1,3}[1-9]\d*)\b/g)){const a=XLSX.utils.decode_cell(match[1]),b=XLSX.utils.decode_cell(match[2]);if((b.r-a.r+1)*(b.c-a.c+1)<=20000)for(let r=a.r;r<=b.r;r++)for(let col=a.c;col<=b.c;col++)refs.push(XLSX.utils.encode_cell({r,c:col}));}if(!affected.has(address)&&(f.getAttribute('t')==='shared'||refs.some(r=>affected.has(r)))){affected.add(address);grew=true;}}}
  const evaluating=new Set(),values=new Map();
  function evaluateCell(address){
    if(values.has(address))return values.get(address);if(evaluating.has(address))throw Error('循环公式');evaluating.add(address);
    try{const c=cells.get(address);if(!c)return 0;const f=first(c,'f');let value;
      if(f&&affected.has(address)){value=evaluateFormula(f.textContent);}
      else {const raw=first(c,'v')?.textContent;if(c.getAttribute('t')&&c.getAttribute('t')!=='n')throw Error('非数字引用');value=raw===undefined||raw===''?0:Number(raw);}
      if(!Number.isFinite(value))throw Error('无效结果');values.set(address,value);return value;
    }finally{evaluating.delete(address);}
  }
  function evaluateFormula(formula){
    let source=formula.replace(/^=/,'').replace(/\$/g,'').replace(/'谈判付款'!/g,'').replace(/谈判付款!/g,'').toUpperCase().replace(/\s+/g,'');
    source=source.replace(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/g,(_,start,end)=>{const a=XLSX.utils.decode_cell(start),b=XLSX.utils.decode_cell(end);if((b.r-a.r+1)*(b.c-a.c+1)>20000)throw Error('范围过大');let total=0;for(let r=a.r;r<=b.r;r++)for(let c=a.c;c<=b.c;c++)total+=evaluateCell(XLSX.utils.encode_cell({r,c}));return String(total);});
    const tokens=source.match(/(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?|[A-Z]+[1-9]\d*|[()+\-*/]/g)||[];
    if(tokens.join('')!==source||!tokens.length)throw Error('需要工作簿重算');let at=0;
    const primary=()=>{const token=tokens[at++];if(token==='+'||token==='-')return(token==='-'?-1:1)*primary();if(token==='('){const v=add();if(tokens[at++]!==')')throw Error('括号');return v;}if(/^[A-Z]+\d+$/.test(token||''))return evaluateCell(token);if(!/^(?:\d|\.)/.test(token||''))throw Error('公式');return Number(token);};
    const multiply=()=>{let v=primary();while(['*','/'].includes(tokens[at])){const op=tokens[at++],n=primary();v=op==='*'?v*n:v/n;}return v;};
    const add=()=>{let v=multiply();while(['+','-'].includes(tokens[at])){const op=tokens[at++],n=multiply();v=op==='+'?v+n:v-n;}return v;};
    const value=add();if(at!==tokens.length)throw Error('公式');return value;
  }
  for(const [address,c] of formulaCells)if(affected.has(address)){
    for(const v of elements(c,'v'))v.remove();
    try{const value=evaluateCell(address);c.removeAttribute('t');const v=create(sheet,'v');v.textContent=String(Math.round((value+Number.EPSILON)*1e8)/1e8);c.append(v);}catch(_){window.FundsExport.lastWarnings.push('谈判付款!'+address+' 保留公式，需在 Excel/WPS 打开时重新计算。');}
  }
  const normal=parseText(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><cols/><sheetData/></worksheet>`);
  const normalData=first(normal,'sheetData'),normalCols=first(normal,'cols');
  const widths=result.normalHeaders.map((h,i)=>Math.min(52,Math.max(16,...[h,...result.normalRows.map(row=>row[i]?.v??'')].map(value=>Math.max(...String(value).split('\n').map(line=>[...line].reduce((n,c)=>n+(/[\u0000-\u00ff]/.test(c)?1:2),0)))+2))));
  widths.forEach((width,i)=>normalCols.append(create(normal,'col',{min:i+1,max:i+1,width,customWidth:1})));
  const allRows=[result.normalHeaders.map(v=>({t:'s',v})),...result.normalRows];
  for(const [r,values] of allRows.entries()){
    let lineCount=1;values.forEach((cell,c)=>{const lines=String(cell?.v??'').split('\n');lineCount=Math.max(lineCount,lines.reduce((n,line)=>n+Math.max(1,Math.ceil([...line].reduce((sum,ch)=>sum+(/[\u0000-\u00ff]/.test(ch)?1:2),0)/(widths[c]-2))),0));});
    const row=create(normal,'row',{r:r+1,ht:Math.max(r===0?32:24,lineCount*17+6),customHeight:1});
    for(let c=0;c<result.normalHeaders.length;c++){const cell=values[c]||{t:'s',v:''};row.append(writeCell(normal,XLSX.utils.encode_cell({r,c}),cell,styleFor(cell.z||'',r===0),true));}
    normalData.append(row);
  }
  normal.documentElement.append(create(normal,'autoFilter',{ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:allRows.length-1,c:result.normalHeaders.length-1}})}));
  normal.documentElement.insertBefore(create(normal,'dimension',{ref:`A1:${XLSX.utils.encode_cell({r:allRows.length-1,c:result.normalHeaders.length-1})}`}),normal.documentElement.firstChild);
  for(const entry of entries)if(entry!==selected){const id=entry.getAttributeNS(R,'id');elements(rels,'Relationship').find(e=>e.getAttribute('Id')===id)?.remove();entry.remove();}
  selected.setAttribute('name','谈判付款');selected.removeAttribute('state');
  let newId='rIdFundsNormal',normalPath='xl/worksheets/funds-normal.xml';
  while(elements(rels,'Relationship').some(e=>e.getAttribute('Id')===newId))newId+='x';
  while(zip.file(normalPath))normalPath=normalPath.replace('.xml','x.xml');
  const nextId=Math.max(...entries.map(e=>Number(e.getAttribute('sheetId'))),0)+1;
  const entry=create(workbook,'sheet',{name:'正常付款',sheetId:nextId});entry.setAttributeNS(R,'r:id',newId);sheets.append(entry);
  const relationship=rels.createElementNS(P,'Relationship');relationship.setAttribute('Id',newId);relationship.setAttribute('Type',R+'/worksheet');relationship.setAttribute('Target',normalPath.slice(3));rels.documentElement.append(relationship);
  const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+normalPath);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.append(type);
  for(const relationship of elements(rels,'Relationship'))if(relationship.getAttribute('Type')===R+'/calcChain')relationship.remove();
  for(const view of elements(workbook,'workbookView')){view.setAttribute('activeTab','0');view.setAttribute('firstSheet','0');}
  let calc=first(workbook,'calcPr');if(!calc){calc=create(workbook,'calcPr');workbook.documentElement.append(calc);}calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');
  for(const view of elements(sheet,'sheetView')){view.setAttribute('tabSelected','1');view.setAttribute('workbookViewId','0');}
  save(sheetPath,sheet);save(normalPath,normal);save(stylePath,styles);save('xl/workbook.xml',workbook);save('xl/_rels/workbook.xml.rels',rels);
  // Keep only reachable OOXML parts, including the selected sheet's drawings and comments.
  const reachable=new Set(['[Content_Types].xml','_rels/.rels']);
  const visit=async(path)=>{if(reachable.has(path))return;reachable.add(path);const relationPath=relPath(path);if(!zip.file(relationPath))return;reachable.add(relationPath);const doc=await parse(zip,relationPath);for(const relation of elements(doc,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve(path,relation.getAttribute('Target')));};
  const root=await parse(zip,'_rels/.rels');for(const relation of elements(root,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve('',relation.getAttribute('Target')));
  for(const path of Object.keys(zip.files))if(!zip.files[path].dir&&!reachable.has(path))zip.remove(path);
  for(const override of elements(types,'Override'))if(!zip.file(override.getAttribute('PartName').replace(/^\//,'')))override.remove();
  save('[Content_Types].xml',types);
  if(zip.file('docProps/app.xml')){const app=await parse(zip,'docProps/app.xml');for(const title of elements(app,'TitlesOfParts')){const vector=first(title,'vector');if(vector){vector.setAttribute('size','2');vector.replaceChildren();for(const value of ['谈判付款','正常付款']){const text=app.createElementNS(vector.namespaceURI,'vt:lpstr');text.textContent=value;vector.append(text);}}}const pairs=first(app,'HeadingPairs');if(pairs){const count=first(pairs,'i4');if(count)count.textContent='2';}save('docProps/app.xml',app);}
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
};
