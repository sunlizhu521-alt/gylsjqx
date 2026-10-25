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
    if(removedNames.some(n=>refers(formula.textContent,n))||forbiddenNames.some(n=>new RegExp('\\b'+escapeRx(n)+'\\b').test(formula.textContent)))throw Error('谈判付款表公式引用了其他工作表，无法在仅保留所选谈判原表时完整保留，请先将相关公式调整为本表引用');
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
  const wrappedStyles=new Map();
  const wrapStyle=id=>{
    const original=xfs.children[Number(id)];if(!original)throw Error('谈判付款单元格样式缺失');
    if(first(original,'alignment')?.getAttribute('wrapText')==='1'&&original.getAttribute('applyAlignment')!=='0')return id;
    if(wrappedStyles.has(String(id)))return wrappedStyles.get(String(id));
    const xf=original.cloneNode(true);let alignment=first(xf,'alignment');
    if(!alignment){alignment=create(styles,'alignment');xf.insertBefore(alignment,first(xf,'protection')||null);}
    alignment.setAttribute('wrapText','1');xf.setAttribute('applyAlignment','1');
    const next=append('cellXfs',xf);wrappedStyles.set(String(id),next);return next;
  };
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
  const sharedRel=elements(rels,'Relationship').find(e=>e.getAttribute('Type')===R+'/sharedStrings');
  const sharedTexts=sharedRel?elements(await parse(zip,resolve('xl/workbook.xml',sharedRel.getAttribute('Target'))),'si').map(si=>elements(si,'t').map(t=>t.textContent).join('')):[];
  const cellText=c=>c?.getAttribute('t')==='s'?sharedTexts[Number(first(c,'v')?.textContent)]||'':c?.getAttribute('t')==='inlineStr'?elements(c,'t').map(t=>t.textContent).join(''):c?first(c,'v')?.textContent??'':'';
  const layoutStyles=new Map(),headerFonts=new Map(),borderStyles=new Map();
  const layoutStyle=(id,header,flow,edges=['left','right','top','bottom'])=>{
    const key=[id,header,flow,edges.join(',')].join('|');if(layoutStyles.has(key))return layoutStyles.get(key);
    const original=xfs.children[Number(id)],xf=original.cloneNode(true),oldBorder=Number(xf.getAttribute('borderId')||0),borderKey=oldBorder+':'+edges.join(',');
    if(!borderStyles.has(borderKey)){
      const originalBorder=borders.children[oldBorder],next=originalBorder.cloneNode(true);
      for(const edge of edges){let side=first(next,edge);if(!side||!side.getAttribute('style')||side.getAttribute('style')==='none'){
        const replacement=first(border,edge).cloneNode(true);if(side)side.replaceWith(replacement);else{const order=['start','end','left','right','top','bottom','diagonal','vertical','horizontal'];next.insertBefore(replacement,[...next.children].find(e=>order.indexOf(e.localName)>order.indexOf(edge))||null);}
      }}
      borderStyles.set(borderKey,next.isEqualNode(originalBorder)?oldBorder:append('borders',next));
    }
    xf.setAttribute('borderId',borderStyles.get(borderKey));xf.setAttribute('applyBorder','1');
    if(header){
      const oldFont=Number(xf.getAttribute('fontId')||0);
      if(!headerFonts.has(oldFont)){const originalFont=fonts.children[oldFont],next=originalFont.cloneNode(true),color=create(styles,'color',{rgb:'FFFFFFFF'}),oldColor=first(next,'color');if(oldColor)oldColor.replaceWith(color);else next.append(color);headerFonts.set(oldFont,next.isEqualNode(originalFont)?oldFont:append('fonts',next));}
      xf.setAttribute('fontId',headerFonts.get(oldFont));xf.setAttribute('applyFont','1');xf.setAttribute('fillId',fillId);xf.setAttribute('applyFill','1');
    }
    if(flow){let alignment=first(xf,'alignment');if(!alignment){alignment=create(styles,'alignment');xf.insertBefore(alignment,first(xf,'protection')||null);}alignment.setAttribute('wrapText','0');alignment.setAttribute('shrinkToFit','1');xf.setAttribute('applyAlignment','1');}
    const next=xf.isEqualNode(original)?Number(id):append('cellXfs',xf);layoutStyles.set(key,next);return next;
  };
  const fitSheet=(doc,end,headerRow=0)=>{
    const data=first(doc,'sheetData');let cols=first(doc,'cols');
    if(!cols){cols=create(doc,'cols');doc.documentElement.insertBefore(cols,data);}
    const oldCols=[...cols.children],widths=Array(end.c+1).fill(12),merges=elements(doc,'mergeCell').map(m=>XLSX.utils.decode_range(m.getAttribute('ref'))),merged=new Map(merges.map(range=>[XLSX.utils.encode_cell(range.s),range.e.c-range.s.c+1]));
    const rows=new Map(elements(data,'row').map(row=>[Number(row.getAttribute('r'))-1,row])),cells=new Map(elements(data,'c').map(c=>[c.getAttribute('r'),c]));
    const ensureCell=(r,c)=>{const address=XLSX.utils.encode_cell({r,c});if(cells.has(address))return cells.get(address);let row=rows.get(r);if(!row){row=create(doc,'row',{r:r+1});data.insertBefore(row,[...data.children].find(e=>Number(e.getAttribute('r'))>r+1)||null);rows.set(r,row);}const cell=create(doc,'c',{r:address});row.insertBefore(cell,[...row.children].find(e=>XLSX.utils.decode_cell(e.getAttribute('r')).c>c)||null);cells.set(address,cell);return cell;};
    const styleId=c=>{const point=XLSX.utils.decode_cell(c.getAttribute('r')),column=oldCols.find(e=>Number(e.getAttribute('min'))<=point.c+1&&Number(e.getAttribute('max'))>=point.c+1);return Number(c.getAttribute('s')??c.parentElement.getAttribute('s')??column?.getAttribute('style')??0);};
    const flowColumns=new Set();for(let c=0;c<=end.c;c++)if(cellText(ensureCell(headerRow,c)).normalize('NFKC').replace(/\s+/g,'')==='流程编号')flowColumns.add(c);
    const populatedColumns=new Set([...cells.values()].filter(c=>cellText(c)!==''||first(c,'f')).map(c=>XLSX.utils.decode_cell(c.getAttribute('r')).c));
    // Keep the grid continuous through blank cells within every populated column.
    for(const c of populatedColumns)for(let r=headerRow;r<=end.r;r++)ensureCell(r,c);
    for(const c of elements(data,'c')){
      const address=c.getAttribute('r'),point=XLSX.utils.decode_cell(address),header=point.r===headerRow,flow=flowColumns.has(point.c);
      if(flow&&!header&&!first(c,'f')&&/[\r\n]/.test(cellText(c))){const text=cellText(c).replace(/\r\n|[\r\n]/g,'；');for(const child of [...c.children])if(['v','is'].includes(child.localName))child.remove();const is=create(doc,'is'),t=create(doc,'t');t.textContent=text;is.append(t);c.setAttribute('t','inlineStr');c.insertBefore(is,c.firstChild);}
      if(header||flow||populatedColumns.has(point.c))c.setAttribute('s',layoutStyle(styleId(c),header,flow));
      const xf=xfs.children[styleId(c)];
      const size=Number(first(fonts.children[Number(xf?.getAttribute('fontId')||0)],'sz')?.getAttribute('val')||11),type=c.getAttribute('t'),raw=first(c,'v')?.textContent??'';
      let text=type==='s'?sharedTexts[Number(raw)]||'':type==='inlineStr'?elements(c,'t').map(t=>t.textContent).join(''):raw;
      if((!type||type==='n')&&raw!=='')try{const id=Number(xf?.getAttribute('numFmtId')||0),format=[...formats].find(([,n])=>n===id)?.[0]??id;text=XLSX.SSF.format(format,Number(raw),{date1904});}catch(_){}
      const width=Math.max(...String(text).split(/\r?\n/).map(line=>[...line].reduce((n,ch)=>n+(/[\u0000-\u00ff]/.test(ch)?1:2),0)))*size/11/(merged.get(address)||1)+3;
      widths[point.c]=Math.min(flow?255:80,Math.max(widths[point.c]||12,width));
    }
    // A populated merged cell also needs borders on its perimeter, without changing the merge.
    for(const range of merges){const anchor=cells.get(XLSX.utils.encode_cell(range.s));if(cellText(anchor)===''&&!anchor?.getElementsByTagNameNS('*','f').length)continue;for(let r=range.s.r;r<=range.e.r;r++)for(let c=range.s.c;c<=range.e.c;c++){const edges=[c===range.s.c&&'left',c===range.e.c&&'right',r===range.s.r&&'top',r===range.e.r&&'bottom'].filter(Boolean);if(edges.length){const cell=ensureCell(r,c);cell.setAttribute('s',layoutStyle(styleId(cell),r===headerRow,flowColumns.has(c),edges));}}}
    cols.replaceChildren(...widths.map((width,c)=>{
      const old=oldCols.find(e=>Number(e.getAttribute('min'))<=c+1&&Number(e.getAttribute('max'))>=c+1),col=old?old.cloneNode(true):create(doc,'col');
      for(const [key,value] of Object.entries({min:c+1,max:c+1,width:Math.ceil(width*100)/100,bestFit:1,customWidth:1}))col.setAttribute(key,value);return col;
    }),...oldCols.filter(c=>Number(c.getAttribute('max'))>end.c+1).map(c=>{const col=c.cloneNode(true);col.setAttribute('min',Math.max(Number(col.getAttribute('min')),end.c+2));return col;}));
    let format=first(doc,'sheetFormatPr');if(!format){format=create(doc,'sheetFormatPr');doc.documentElement.insertBefore(format,cols);}format.setAttribute('defaultRowHeight','30');format.setAttribute('customHeight','1');
    for(const row of elements(data,'row')){row.setAttribute('ht','30');row.setAttribute('customHeight','1');row.removeAttribute('hidden');}
    const range=XLSX.utils.encode_range({s:{r:headerRow,c:0},e:end});let filter=first(doc,'autoFilter');
    if(!filter){filter=create(doc,'autoFilter');let after=data.nextElementSibling;while(after&&['sheetCalcPr','sheetProtection','protectedRanges','scenarios'].includes(after.localName))after=after.nextElementSibling;doc.documentElement.insertBefore(filter,after);}
    filter.setAttribute('ref',range);filter.replaceChildren();
    if(doc===sheet)for(const name of elements(workbook,'definedName'))if(name.getAttribute('name')==='_xlnm._FilterDatabase'&&name.getAttribute('localSheetId')==='0')name.textContent="'谈判付款'!$A$"+(headerRow+1)+':$'+XLSX.utils.encode_col(end.c)+'$'+(end.r+1);
  };
  const rows=new Map(elements(sheetData,'row').map(r=>[Number(r.getAttribute('r')),r]));
  const cells=new Map(elements(sheetData,'c').map(c=>[c.getAttribute('r'),c]));
  const originalRange=XLSX.utils.decode_range(first(sheet,'dimension')?.getAttribute('ref')||'A1');
  const submissionColumns=new Set((result.negotiated.headers||[]).flatMap((h,c)=>['提交时间汇总','提交付款申请单日期'].includes(String(h).normalize('NFKC').replace(/\s+/g,''))?[c]:[]));
  const sourceColumns=elements(sheet,'col');
  let maxRow=originalRange.e.r,maxColumn=originalRange.e.c;
  const changed=new Set();
  for(const {address,cell} of result.negotiated.updates){
    if(!/^[A-Z]+[1-9]\d*$/.test(address))throw Error('谈判付款更新地址无效');
    const point=XLSX.utils.decode_cell(address),old=cells.get(address);let row=rows.get(point.r+1);
    if(!row){row=create(sheet,'row',{r:point.r+1});const after=[...rows].filter(([r])=>r>point.r+1).sort((a,b)=>a[0]-b[0])[0]?.[1];sheetData.insertBefore(row,after||null);rows.set(point.r+1,row);}
    const isHeader=point.r===result.negotiated.headerRow;
    const submission=!isHeader&&submissionColumns.has(point.c),column=sourceColumns.find(c=>Number(c.getAttribute('min'))<=point.c+1&&Number(c.getAttribute('max'))>=point.c+1);
    let styleId=old?.getAttribute('s')??(old&&submission?(row.getAttribute('s')??column?.getAttribute('style')??0):styleFor(cell?.z||'',isHeader));
    if(submission)styleId=wrapStyle(styleId);
    const next=writeCell(sheet,address,cell,styleId);
    // Retain unrelated cell metadata while replacing only value/formula payloads.
    if(old){for(const attr of old.attributes)if(!['r','s','t'].includes(attr.name))next.setAttribute(attr.name,attr.value);old.replaceWith(next);}
    else {const after=[...row.children].find(c=>c.localName==='c'&&XLSX.utils.decode_cell(c.getAttribute('r')).c>point.c);row.insertBefore(next,after||null);}
    cells.set(address,next);changed.add(address);maxRow=Math.max(maxRow,point.r);maxColumn=Math.max(maxColumn,point.c);
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
  // Move the existing column XML, including styles and references, instead of rebuilding the source sheet.
  const order=result.negotiated.columnOrder;
  if(order?.some((old,index)=>old!==index)){
    if(new Set(order).size!==order.length||order.some(c=>!Number.isInteger(c)||c<0||c>=order.length))throw Error('谈判付款列顺序无效');
    const positions=new Map(order.map((old,index)=>[old,index])),column=c=>positions.get(c)??c;
    const reference=ref=>{const m=ref.match(/^(\$?)([A-Z]{1,3})(\$?\d+)?$/i);if(!m)throw Error('无法保留的列引用：'+ref);return {c:XLSX.utils.decode_col(m[2].toUpperCase()),absolute:m[1],row:m[3]||''};};
    const groups=(start,end)=>{
      const cols=[];for(let c=start;c<=end;c++)cols.push(column(c));cols.sort((a,b)=>a-b);
      const out=[];for(const c of cols){const last=out.at(-1);if(last&&last[1]+1===c)last[1]=c;else out.push([c,c]);}return out;
    };
    const ranges=ref=>{
      const [left,right]=ref.split(':'),a=reference(left),b=reference(right||left),format=(c,r)=>r.absolute+XLSX.utils.encode_col(c)+r.row;
      if(!right)return [format(column(a.c),a)];
      return groups(a.c,b.c).map(([start,end])=>format(start,a)+':'+format(end,b));
    };
    const singleRange=ref=>{const parts=ranges(ref);if(parts.length!==1)throw Error('日期列调整会拆分原合并或表格范围 '+ref+'，已停止导出以保留原结构');return parts[0];};
    const formula=text=>text.split(/("(?:[^"]|"")*")/g).map((part,index)=>index%2?part:part.replace(/(^|[^A-Z0-9_.])((?:'(?:[^']|'')+'|[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![A-Z0-9_(])/gi,(match,prefix,token,ref)=>{
      const name=token?.slice(0,-1).replace(/^'|'$/g,'').replace(/''/g,"'");if(name&&name!=='谈判付款'&&name!==oldName)return match;
      const parts=ranges(ref).map(range=>(token||'')+range);return prefix+(parts.length>1?'('+parts.join(',')+')':parts[0]);
    })).join('');
    const relocateDocument=doc=>{
      for(const el of doc.getElementsByTagName('*')){
        for(const attr of ['ref','sqref','activeCell','topLeftCell'])if(el.hasAttribute(attr)){
          const value=el.getAttribute(attr);if(!/^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?(?:\s+\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)*$/.test(value))continue;
          el.setAttribute(attr,attr==='sqref'?value.split(/\s+/).flatMap(ranges).join(' '):singleRange(value));
        }
        if(['f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula'].includes(el.localName))el.textContent=formula(el.textContent);
        if(el.localName==='hyperlink'&&el.hasAttribute('location'))el.setAttribute('location',formula(el.getAttribute('location')));
      }
    };
    for(const c of cells.values()){
      const address=c.getAttribute('r'),f=first(c,'f');
      if(f?.getAttribute('t')==='shared'){
        const p=XLSX.utils.decode_cell(address),expanded=result.negotiated.rows[p.r-result.negotiated.headerRow-1]?.[p.c]?.f;
        if(!expanded)throw Error('无法展开共享公式 '+address+'，已保留原文件');
        f.textContent=expanded;for(const attr of ['t','si','ref'])f.removeAttribute(attr);
      }
      c.setAttribute('r',singleRange(address));
    }
    const remapFilters=doc=>{for(const filter of elements(doc,'autoFilter')){
      const start=XLSX.utils.decode_range(filter.getAttribute('ref')).s.c,nextStart=XLSX.utils.decode_range(singleRange(filter.getAttribute('ref'))).s.c;
      for(const f of elements(filter,'filterColumn'))f.setAttribute('colId',column(start+Number(f.getAttribute('colId')))-nextStart);
    }};
    remapFilters(sheet);relocateDocument(sheet);
    for(const row of rows.values()){
      const children=[...row.children].filter(c=>c.localName==='c').sort((a,b)=>XLSX.utils.decode_cell(a.getAttribute('r')).c-XLSX.utils.decode_cell(b.getAttribute('r')).c);
      for(const c of children)row.insertBefore(c,first(row,'extLst')||null);row.removeAttribute('spans');
    }
    const cols=first(sheet,'cols');if(cols){const moved=[];for(const col of [...cols.children])for(const [start,end] of groups(Number(col.getAttribute('min'))-1,Number(col.getAttribute('max'))-1)){const next=col.cloneNode(true);next.setAttribute('min',start+1);next.setAttribute('max',end+1);moved.push(next);}cols.replaceChildren(...moved.sort((a,b)=>Number(a.getAttribute('min'))-Number(b.getAttribute('min'))));}
    for(const name of elements(workbook,'definedName'))name.textContent=formula(name.textContent);
    const visited=new Set();
    const relocateRelated=async path=>{
      const relationPath=relPath(path);if(!zip.file(relationPath))return;
      for(const rel of elements(await parse(zip,relationPath),'Relationship')){
        if(rel.getAttribute('TargetMode')==='External')continue;
        const target=resolve(path,rel.getAttribute('Target'));if(visited.has(target)||!zip.file(target)||!(/\.(?:xml|vml)$/.test(target)))continue;visited.add(target);
        const doc=await parse(zip,target),table=first(doc,'table');
        if(table){const start=XLSX.utils.decode_range(table.getAttribute('ref')).s.c,list=first(table,'tableColumns');if(list)list.replaceChildren(...[...list.children].map((el,i)=>({el,c:column(start+i)})).sort((a,b)=>a.c-b.c).map(x=>x.el));}
        remapFilters(doc);relocateDocument(doc);
        for(const tag of ['from','to'])for(const anchor of elements(doc,tag)){const col=first(anchor,'col');if(col)col.textContent=column(Number(col.textContent));}
        for(const el of elements(doc,'Column'))el.textContent=column(Number(el.textContent));
        for(const el of elements(doc,'Anchor')){const values=el.textContent.split(',').map(v=>v.trim());if(values.length===8){values[0]=column(Number(values[0]));values[4]=column(Number(values[4]));el.textContent=values.join(', ');}}
        save(target,doc);await relocateRelated(target);
      }
    };
    await relocateRelated(sheetPath);
  }
  fitSheet(sheet,{r:maxRow,c:maxColumn},result.negotiated.headerRow||0);
  for(const entry of entries)if(entry!==selected){const id=entry.getAttributeNS(R,'id');elements(rels,'Relationship').find(e=>e.getAttribute('Id')===id)?.remove();entry.remove();}
  selected.setAttribute('name','谈判付款');selected.removeAttribute('state');
  let nextSheetId=Math.max(...entries.map(e=>Number(e.getAttribute('sheetId'))),0)+1;
  const generatedSheets=[['正常付款','Normal',result.normalHeaders,result.normalRows],['付款申请核对','Review',result.reviewHeaders,result.reviewRows]];
  for(const [name,key,headers,data] of generatedSheets){
    const generated=parseText(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><cols/><sheetData/></worksheet>`);
    const generatedData=first(generated,'sheetData');
    const allRows=[headers.map(v=>({t:'s',v})),...data];
    for(const [r,values] of allRows.entries()){
      const row=create(generated,'row',{r:r+1});
      for(let c=0;c<headers.length;c++){const cell=values[c]||{t:'s',v:''};row.append(writeCell(generated,XLSX.utils.encode_cell({r,c}),cell,styleFor(cell.z||'',r===0),true));}
      generatedData.append(row);
    }
    const range=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:allRows.length-1,c:headers.length-1}});
    generated.documentElement.insertBefore(create(generated,'dimension',{ref:range}),generated.documentElement.firstChild);
    fitSheet(generated,{r:allRows.length-1,c:headers.length-1});
    let newId='rIdFunds'+key,generatedPath='xl/worksheets/funds-'+key.toLowerCase()+'.xml';
    while(elements(rels,'Relationship').some(e=>e.getAttribute('Id')===newId))newId+='x';
    while(zip.file(generatedPath))generatedPath=generatedPath.replace('.xml','x.xml');
    const entry=create(workbook,'sheet',{name,sheetId:nextSheetId++});entry.setAttributeNS(R,'r:id',newId);sheets.append(entry);
    const relationship=rels.createElementNS(P,'Relationship');relationship.setAttribute('Id',newId);relationship.setAttribute('Type',R+'/worksheet');relationship.setAttribute('Target',generatedPath.slice(3));rels.documentElement.append(relationship);
    const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+generatedPath);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.append(type);
    save(generatedPath,generated);
  }
  for(const relationship of elements(rels,'Relationship'))if(relationship.getAttribute('Type')===R+'/calcChain')relationship.remove();
  for(const view of elements(workbook,'workbookView')){view.setAttribute('activeTab','0');view.setAttribute('firstSheet','0');}
  let calc=first(workbook,'calcPr');if(!calc){calc=create(workbook,'calcPr');workbook.documentElement.append(calc);}calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');
  for(const view of elements(sheet,'sheetView')){view.setAttribute('tabSelected','1');view.setAttribute('workbookViewId','0');}
  save(sheetPath,sheet);save(stylePath,styles);save('xl/workbook.xml',workbook);save('xl/_rels/workbook.xml.rels',rels);
  // Keep only reachable OOXML parts, including the selected sheet's drawings and comments.
  const reachable=new Set(['[Content_Types].xml','_rels/.rels']);
  const visit=async(path)=>{if(reachable.has(path))return;reachable.add(path);const relationPath=relPath(path);if(!zip.file(relationPath))return;reachable.add(relationPath);const doc=await parse(zip,relationPath);for(const relation of elements(doc,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve(path,relation.getAttribute('Target')));};
  const root=await parse(zip,'_rels/.rels');for(const relation of elements(root,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve('',relation.getAttribute('Target')));
  for(const path of Object.keys(zip.files))if(!zip.files[path].dir&&!reachable.has(path))zip.remove(path);
  for(const override of elements(types,'Override'))if(!zip.file(override.getAttribute('PartName').replace(/^\//,'')))override.remove();
  save('[Content_Types].xml',types);
  if(zip.file('docProps/app.xml')){const app=await parse(zip,'docProps/app.xml');for(const title of elements(app,'TitlesOfParts')){const vector=first(title,'vector');if(vector){vector.setAttribute('size','3');vector.replaceChildren();for(const value of ['谈判付款','正常付款','付款申请核对']){const text=app.createElementNS(vector.namespaceURI,'vt:lpstr');text.textContent=value;vector.append(text);}}}const pairs=first(app,'HeadingPairs');if(pairs){const count=first(pairs,'i4');if(count)count.textContent='3';}save('docProps/app.xml',app);}
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
};
