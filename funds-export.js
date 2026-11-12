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
  const oldName=selected.getAttribute('name');
  const plan=FundsCore.outputSheetPlan(entries.map(e=>e.getAttribute('name')),oldName);
  const sourceEntries=new Map(entries.map(e=>[e.getAttribute('name'),e])),sourceNames=[...sourceEntries.keys()];
  const renamed=new Map(plan.filter(item=>item.sourceName&&item.name!==item.sourceName).map(item=>[item.sourceName.toLowerCase(),item.name]));
  const selectedRel=elements(rels,'Relationship').find(e=>e.getAttribute('Id')===selected.getAttributeNS(R,'id'));
  if(!selectedRel)throw Error('谈判付款工作表关系缺失');
  const sheetPath=resolve('xl/workbook.xml',selectedRel.getAttribute('Target'));
  const sheet=await parse(zip,sheetPath),sheetData=first(sheet,'sheetData');
  if(!sheetData)throw Error('谈判付款表缺少单元格数据');
  const originalRange=result.negotiated.sourceRange||XLSX.utils.decode_range(first(sheet,'dimension')?.getAttribute('ref')||'A1');
  const managedColumns=result.negotiated.managedColumns?new Set(result.negotiated.managedColumns):null;
  if(managedColumns&&result.negotiated.columnOrder?.some((c,i)=>c!==i))throw Error('谈判付款原列顺序不能变更，未生成移动保留列的文件');
  const managedColumn=c=>!managedColumns||managedColumns.has(c);
  const managedAddress=address=>managedColumn(XLSX.utils.decode_cell(address).c);
  const preserveRowLayout=managedColumns&&Array.from({length:originalRange.e.c+1},(_,c)=>c).some(c=>!managedColumn(c));
  const quote=name=>"'"+name.replace(/'/g,"''")+"'";
  const renameSelf=formula=>formula.split(/("(?:[^"]|"")*")/g).map((part,index)=>{
    if(index%2)return part;
    for(const match of part.matchAll(/(?:'((?:[^']|'')+)'|([A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*:[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*))!/gi)){
      const names=(match[1]||match[2]).replace(/''/g,"'").split(':');if(names.length!==2||!names.every(name=>sourceEntries.has(name)))continue;
      const start=sourceNames.indexOf(names[0]),end=sourceNames.indexOf(names[1]),span=sourceNames.slice(Math.min(start,end),Math.max(start,end)+1);
      const a=plan.findIndex(item=>item.sourceName===names[0]),b=plan.findIndex(item=>item.sourceName===names[1]);
      if(names.some(name=>renamed.has(name.toLowerCase()))||JSON.stringify(span)!==JSON.stringify(plan.slice(Math.min(a,b),Math.max(a,b)+1).map(item=>item.sourceName)))throw Error('原表三维公式引用 '+match[0]+' 会受工作表顺序变化影响，请改为逐个工作表引用后再导出');
    }
    return part.replace(/(?<![\w\u4e00-\u9fff.\]'])('(?:[^']|'')+'|[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*)!/gi,(match,token)=>{
    const name=token.replace(/^'|'$/g,'').replace(/''/g,"'");
    return renamed.has(name.toLowerCase())?quote(renamed.get(name.toLowerCase()))+'!':match;
  });}).join('');
  for(const name of elements(workbook,'definedName')) {
    const local=name.getAttribute('localSheetId');
    const sourceName=local===null?null:entries[Number(local)]?.getAttribute('name');
    if(sourceName&&FundsCore.OUTPUT_SHEETS.slice(1).includes(sourceName))name.remove();
    else {if(local!==null)name.setAttribute('localSheetId',plan.findIndex(item=>item.sourceName===sourceName));name.textContent=renameSelf(name.textContent);}
  }
  for(const formula of [...elements(sheet,'f'),...elements(sheet,'formula'),...elements(sheet,'formula1'),...elements(sheet,'formula2')]) {
    formula.textContent=renameSelf(formula.textContent);
  }
  for(const link of elements(sheet,'hyperlink'))if(link.hasAttribute('location')) {
    const location=link.getAttribute('location');
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
  const border=create(styles,'border');for(const edge of ['left','right','top','bottom']){const side=create(styles,edge,{style:'hair'});side.append(create(styles,'color',{rgb:'FF404040'}));border.append(side);}const borderId=append('borders',border);
  const headerBorder=border.cloneNode(true);for(const side of headerBorder.children)side.setAttribute('style','thin');const headerBorderId=append('borders',headerBorder);
  let numFmts=first(styles,'numFmts');if(!numFmts){numFmts=create(styles,'numFmts',{count:0});styles.documentElement.insertBefore(numFmts,styles.documentElement.firstChild);}
  const formats=new Map(elements(numFmts,'numFmt').map(e=>[e.getAttribute('formatCode'),Number(e.getAttribute('numFmtId'))]));
  let nextFormat=Math.max(163,...formats.values())+1;
  const formatId=code=>{if(!code||code==='General')return 0;if(formats.has(code))return formats.get(code);const id=nextFormat++;numFmts.append(create(styles,'numFmt',{numFmtId:id,formatCode:code}));numFmts.setAttribute('count',numFmts.children.length);formats.set(code,id);return id;};
  const styleIds=new Map();let appliedFillId;
  const redFill=()=>{
    if(appliedFillId===undefined){const fill=create(styles,'fill'),pattern=create(styles,'patternFill',{patternType:'solid'});pattern.append(create(styles,'fgColor',{rgb:'FFFF0000'}),create(styles,'bgColor',{rgb:'FFFF0000'}));fill.append(pattern);appliedFillId=[...fills.children].findIndex(existing=>existing.isEqualNode(fill));if(appliedFillId<0)appliedFillId=append('fills',fill);}
    return appliedFillId;
  };
  const styleFor=(code='',header=false,applied=false)=>{
    const key=JSON.stringify([code,header,applied]);if(styleIds.has(key))return styleIds.get(key);
    // Native fills follow records when WPS sorts them and remain manually editable.
    const xf=create(styles,'xf',{numFmtId:formatId(code),fontId:header?headerFontId:fontId,fillId:header?fillId:applied?redFill():0,borderId:header?headerBorderId:borderId,xfId:0,applyFont:1,applyFill:1,applyBorder:1,applyAlignment:1,applyNumberFormat:1});xf.append(create(styles,'alignment',{horizontal:'center',vertical:'center',wrapText:1}));const id=append('cellXfs',xf);styleIds.set(key,id);return id;
  };
  const appliedFormula='N("GYLSJQX_FUNDS_APPLIED_V1")=0',appliedStylePrefix='GYLSJQX_FUNDS_APPLIED_V2_';
  // Remove only our old fixed-address rules; they override manual colors and drift after sorting.
  for(const container of elements(sheet,'conditionalFormatting')){
    // Withdraw old tool rules only from business columns, retaining custom-column rules.
    const owned=[...container.children].filter(rule=>rule.namespaceURI===S&&rule.localName==='cfRule'&&first(rule,'formula')?.textContent.trim()===appliedFormula);
    if(owned.length&&managedColumns){
      const retained=[];let touchesManaged=false;
      for(const ref of (container.getAttribute('sqref')||'').trim().split(/\s+/).filter(Boolean)){
        const range=XLSX.utils.decode_range(ref);let start=null;
        for(let c=range.s.c;c<=range.e.c+1;c++){
          if(c<=range.e.c&&!managedColumn(c)){if(start===null)start=c;}
          else {if(c<=range.e.c)touchesManaged=true;if(start!==null){retained.push(XLSX.utils.encode_range({s:{r:range.s.r,c:start},e:{r:range.e.r,c:c-1}}));start=null;}}
        }
      }
      if(!touchesManaged)continue;
      if(retained.length){const copy=container.cloneNode(false);copy.setAttribute('sqref',retained.join(' '));for(const rule of owned)copy.append(rule.cloneNode(true));container.after(copy);}
    }
    let removed=false;
    for(const rule of [...container.children])if(rule.namespaceURI===S&&rule.localName==='cfRule'&&first(rule,'formula')?.textContent.trim()===appliedFormula){rule.remove();removed=true;}
    if(removed&&!container.children.length)container.remove();
  }
  let baseXfs=first(styles,'cellStyleXfs'),namedStyles=first(styles,'cellStyles');
  if(!baseXfs){baseXfs=create(styles,'cellStyleXfs',{count:0});styles.documentElement.insertBefore(baseXfs,xfs);}
  if(!namedStyles){namedStyles=create(styles,'cellStyles',{count:0});styles.documentElement.insertBefore(namedStyles,xfs.nextSibling);}
  const ownedBases=new Set([...namedStyles.children].filter(n=>n.getAttribute('name')?.startsWith(appliedStylePrefix)).map(n=>Number(n.getAttribute('xfId'))));
  const restoredStyles=new Map(),appliedStyles=new Map();
  // Hidden native cell styles keep the previous fill with the cell, including after sorting/save.
  for(const c of elements(sheetData,'c')){
    if(!managedAddress(c.getAttribute('r')))continue;
    const id=Number(c.getAttribute('s')||0),xf=xfs.children[id],parent=Number(xf?.getAttribute('xfId')||0);if(!ownedBases.has(parent))continue;
    if(!restoredStyles.has(id)){
      const base=baseXfs.children[parent];if(!base)throw Error('付款申请标色的原样式缺失');
      const restored=xf.cloneNode(true),pattern=first(fills.children[Number(xf.getAttribute('fillId')||0)],'patternFill');
      if(pattern?.getAttribute('patternType')==='solid'&&first(pattern,'fgColor')?.getAttribute('rgb')==='FFFF0000')restored.setAttribute('fillId',base.getAttribute('fillId')||'0');
      restored.setAttribute('xfId',base.getAttribute('xfId')||'0');restoredStyles.set(id,append('cellXfs',restored));
    }
    c.setAttribute('s',restoredStyles.get(id));
  }
  const applyEditableFill=c=>{
    const id=Number(c.getAttribute('s')||0);if(!appliedStyles.has(id)){
      const original=xfs.children[id];let base=[...ownedBases].find(i=>baseXfs.children[i].isEqualNode(original));
      if(base===undefined){base=append('cellStyleXfs',original.cloneNode(true));append('cellStyles',create(styles,'cellStyle',{name:appliedStylePrefix+base,xfId:base,hidden:1}));ownedBases.add(base);}
      const xf=original.cloneNode(true);xf.setAttribute('fillId',redFill());xf.setAttribute('applyFill','1');xf.setAttribute('xfId',base);appliedStyles.set(id,append('cellXfs',xf));
    }
    c.setAttribute('s',appliedStyles.get(id));
  };
  const highlightedCells=result.highlights?.negotiatedCells??[],highlightedRows=result.highlights?.normalRows??[];
  if(!Array.isArray(highlightedCells)||!Array.isArray(highlightedRows))throw Error('付款申请标红位置无效');
  for(const row of highlightedRows)if(!Number.isInteger(row)||row<1||row>result.normalRows.length+1)throw Error('正常付款标红行号无效');
  const appliedNormalRows=new Set(highlightedRows);
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
  const textDateStyles=new Map();
  const writeCell=(doc,address,cell,styleId,convertDates=false)=>{
    // Store generated dates and schedule date labels as text, leaving monetary values numeric.
    if(typeof cell?.v==='number'&&cell.t!=='s'&&cell.z&&XLSX.SSF.is_date(cell.z)){
      if(!Number.isFinite(cell.v))throw Error('单元格 '+address+' 的日期不是有效数值');
      const format=cell.z.replace(/"[^"]*"|\\.|[年月日时分秒]/g,part=>part.length===1?'"'+part+'"':part);
      cell={t:'s',v:XLSX.SSF.format(format,cell.v,{date1904:convertDates?false:date1904}),z:'@'};
      if(styleId!==undefined){
        const key=String(styleId);
        if(!textDateStyles.has(key)){const xf=xfs.children[Number(styleId)].cloneNode(true);xf.setAttribute('numFmtId','49');xf.setAttribute('applyNumberFormat','1');textDateStyles.set(key,append('cellXfs',xf));}
        styleId=textDateStyles.get(key);
      }
    }
    const c=create(doc,'c',{r:address});if(styleId!==undefined)c.setAttribute('s',styleId);
    if(cell?.f){const f=create(doc,'f');f.textContent=String(cell.f).replace(/^=/,'');c.append(f);}
    if(cell?.v===null||cell?.v===undefined)return c;
    let value=cell.v;
    const numeric=typeof value==='number'&&cell.t!=='s';
    if(numeric&&!Number.isFinite(value))throw Error('单元格 '+address+' 的金额或日期不是有效数值');
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
  const emptyBorderId=append('borders',create(styles,'border')),emptyBorderStyles=new Map();
  const withoutBorder=id=>{if(!emptyBorderStyles.has(id)){const xf=xfs.children[id].cloneNode(true);xf.setAttribute('borderId',emptyBorderId);xf.setAttribute('applyBorder','1');emptyBorderStyles.set(id,append('cellXfs',xf));}return emptyBorderStyles.get(id);};
  const layoutStyles=new Map(),headerFonts=new Map(),borderStyles=new Map();
  const layoutStyle=(id,header,flow,edges=['left','right','top','bottom'])=>{
    const key=[id,header,flow,edges.join(',')].join('|');if(layoutStyles.has(key))return layoutStyles.get(key);
    const original=xfs.children[Number(id)],xf=original.cloneNode(true),oldBorder=Number(xf.getAttribute('borderId')||0),borderKey=[oldBorder,header,edges.join(',')].join(':');
    if(!borderStyles.has(borderKey)){
      const originalBorder=borders.children[oldBorder],next=originalBorder.cloneNode(true);
      for(const edge of edges){const side=first(next,edge),replacement=first(header?headerBorder:border,edge).cloneNode(true);if(side)side.replaceWith(replacement);else{const order=['start','end','left','right','top','bottom','diagonal','vertical','horizontal'];next.insertBefore(replacement,[...next.children].find(e=>order.indexOf(e.localName)>order.indexOf(edge))||null);}}
      borderStyles.set(borderKey,next.isEqualNode(originalBorder)?oldBorder:append('borders',next));
    }
    xf.setAttribute('borderId',borderStyles.get(borderKey));xf.setAttribute('applyBorder','1');
    if(header){
      const oldFont=Number(xf.getAttribute('fontId')||0);
      if(!headerFonts.has(oldFont)){const originalFont=fonts.children[oldFont],next=originalFont.cloneNode(true),color=create(styles,'color',{rgb:'FFFFFFFF'}),oldColor=first(next,'color');if(oldColor)oldColor.replaceWith(color);else next.append(color);headerFonts.set(oldFont,next.isEqualNode(originalFont)?oldFont:append('fonts',next));}
      xf.setAttribute('fontId',headerFonts.get(oldFont));xf.setAttribute('applyFont','1');xf.setAttribute('fillId',fillId);xf.setAttribute('applyFill','1');
    }
    if(header||flow){let alignment=first(xf,'alignment');if(!alignment){alignment=create(styles,'alignment');xf.insertBefore(alignment,first(xf,'protection')||null);}alignment.setAttribute('wrapText','0');alignment.setAttribute('shrinkToFit',flow?'1':'0');xf.setAttribute('applyAlignment','1');}
    const next=xf.isEqualNode(original)?Number(id):append('cellXfs',xf);layoutStyles.set(key,next);return next;
  };
  const measure=document.createElement('canvas').getContext('2d'),dateDisplayStyles=new Map();
  // WPS format 31 displays a Chinese date, while SSF falls back to m/d/yy.
  const dateDisplayStyle=id=>{
    if(dateDisplayStyles.has(id))return dateDisplayStyles.get(id);
    const xf=xfs.children[id].cloneNode(true);xf.setAttribute('numFmtId',formatId('yyyy"年"m"月"d"日"'));xf.setAttribute('applyNumberFormat','1');
    const next=append('cellXfs',xf);dateDisplayStyles.set(id,next);return next;
  };
  const fitSheet=(doc,end,headerRow=0,review=false,progress=false)=>{
    const scoped=c=>doc!==sheet||managedColumn(result.negotiated.columnOrder?.[c]??c);
    const batch=doc===sheet&&result.negotiated.batchLayout,detailColumn=batch?result.negotiated.headers.findIndex(h=>String(h).trim()==='谈判明细'):-1;
    const data=first(doc,'sheetData');let cols=first(doc,'cols');
    if(!cols){cols=create(doc,'cols');doc.documentElement.insertBefore(cols,data);}
    const oldCols=[...cols.children],widths=Array(end.c+1).fill(12),merges=elements(doc,'mergeCell').map(m=>XLSX.utils.decode_range(m.getAttribute('ref'))),merged=new Map(merges.map(range=>[XLSX.utils.encode_cell(range.s),range.e.c-range.s.c+1]));
    const rows=new Map(elements(data,'row').map(row=>[Number(row.getAttribute('r'))-1,row])),cells=new Map(elements(data,'c').map(c=>[c.getAttribute('r'),c]));
    const ensureCell=(r,c)=>{const address=XLSX.utils.encode_cell({r,c});if(cells.has(address))return cells.get(address);let row=rows.get(r);if(!row){row=create(doc,'row',{r:r+1});data.insertBefore(row,[...data.children].find(e=>Number(e.getAttribute('r'))>r+1)||null);rows.set(r,row);}const cell=create(doc,'c',{r:address});row.insertBefore(cell,[...row.children].find(e=>XLSX.utils.decode_cell(e.getAttribute('r')).c>c)||null);cells.set(address,cell);return cell;};
    const styleId=c=>{const point=XLSX.utils.decode_cell(c.getAttribute('r')),column=oldCols.find(e=>Number(e.getAttribute('min'))<=point.c+1&&Number(e.getAttribute('max'))>=point.c+1);return Number(c.getAttribute('s')??c.parentElement.getAttribute('s')??column?.getAttribute('style')??0);};
    const flowColumns=new Set();for(let c=0;c<=end.c;c++)if(scoped(c)&&['流程编号','付款自然周','申请付款的日期'].includes(cellText(ensureCell(headerRow,c)).normalize('NFKC').replace(/\s+/g,'')))flowColumns.add(c);
    const singleLineColumns=new Set(flowColumns),dateColumns=new Set();
    if(doc!==sheet)for(let c=0;c<=end.c;c++){const title=cellText(ensureCell(headerRow,c));if(['对公付款0A流程号','对公付款OA流程号','开票时间','OA合同编号',...(progress?['对公付款OA号','申请日期','应付时间']:[])].includes(title))singleLineColumns.add(c);if(title==='开票时间'||progress&&['申请日期','应付时间'].includes(title))dateColumns.add(c);}
    const populatedColumns=new Set(),populatedRows=new Set();
    for(const cell of cells.values())if(cellText(cell).trim()){const point=XLSX.utils.decode_cell(cell.getAttribute('r'));populatedRows.add(point.r);if(scoped(point.c))populatedColumns.add(point.c);}
    for(const range of merges)if(cellText(cells.get(XLSX.utils.encode_cell(range.s))).trim()){for(let r=range.s.r;r<=range.e.r;r++)populatedRows.add(r);for(let c=range.s.c;c<=range.e.c;c++)if(scoped(c))populatedColumns.add(c);}
    // Explicit blank cells override inherited column borders without changing their values.
    for(const c of populatedColumns)for(let r=headerRow;r<=end.r;r++)ensureCell(r,c);
    for(const c of elements(data,'c')){
      const address=c.getAttribute('r'),point=XLSX.utils.decode_cell(address),header=point.r===headerRow,flow=singleLineColumns.has(point.c);
      if(point.c>end.c||point.r>end.r||!scoped(point.c))continue;
      if(doc!==sheet&&flow&&!header&&!first(c,'f')&&/[\r\n]/.test(cellText(c))){const text=cellText(c).replace(/\r\n|[\r\n]/g,'；');for(const child of [...c.children])if(['v','is'].includes(child.localName))child.remove();const is=create(doc,'is'),t=create(doc,'t');t.textContent=text;is.append(t);c.setAttribute('t','inlineStr');c.insertBefore(is,c.firstChild);}
      if(header&&['',null,'n'].includes(c.getAttribute('t'))){
        const id=Number(xfs.children[styleId(c)].getAttribute('numFmtId')||0),raw=first(c,'v')?.textContent;
        const sourceFormat=id===31?FundsCore.DATE_FORMAT:[...formats].find(([,n])=>n===id)?.[0]??XLSX.SSF.get_table()[id];
        if((raw?.trim()||first(c,'f'))&&sourceFormat&&XLSX.SSF.is_date(sourceFormat)){
          const formula=first(c,'f');
          if(!raw?.trim())throw Error('日期表头 '+address+' 缺少公式计算结果，请在 Excel/WPS 重算并保存后重新上传');
          if(formula&&(formula.getAttribute('t')&&formula.getAttribute('t')!=='normal'||!formula.textContent.trim()))throw Error('日期表头 '+address+' 含共享或数组公式，请在源表改为普通日期公式后重新上传');
          // Keep the full date so re-uploading month-only labels does not lose the original day.
          const format=FundsCore.DATE_FORMAT;
          const next=writeCell(doc,address,{t:'n',v:Number(raw),z:format},styleId(c));
          for(const child of [...c.children])if(['v','is'].includes(child.localName))child.remove();
          c.setAttribute('s',next.getAttribute('s'));
          if(formula){
            formula.textContent='TEXT('+formula.textContent+',"'+format.replace(/"/g,'""')+'")';
            c.setAttribute('t','str');const value=create(doc,'v');value.textContent=cellText(next);formula.after(value);
          }else{c.setAttribute('t','inlineStr');c.insertBefore(first(next,'is'),c.firstChild);}
        }
      }
      if(header||flow||populatedColumns.has(point.c))c.setAttribute('s',layoutStyle(styleId(c),header,flow));
      if(!header&&(!populatedRows.has(point.r)||!populatedColumns.has(point.c)))c.setAttribute('s',withoutBorder(styleId(c)));
      const xf=xfs.children[styleId(c)];
      const size=Number(first(fonts.children[Number(xf?.getAttribute('fontId')||0)],'sz')?.getAttribute('val')||11),type=c.getAttribute('t'),raw=first(c,'v')?.textContent??'';
      let text=type==='s'?sharedTexts[Number(raw)]||'':type==='inlineStr'?elements(c,'t').map(t=>t.textContent).join(''):raw;
      if((!type||type==='n')&&raw!=='')try{const id=Number(xf?.getAttribute('numFmtId')||0),format=[...formats].find(([,n])=>n===id)?.[0]??id;text=XLSX.SSF.format(format,Number(raw),{date1904});}catch(_){}
      let width=Math.max(...String(text).split(/\r?\n/).map(line=>[...line].reduce((n,ch)=>n+(/[\u0000-\u00ff]/.test(ch)?1:2),0)))*size/11/(merged.get(address)||1)+3;
      if(header||flow){const font=fonts.children[Number(xf?.getAttribute('fontId')||0)];if(measure){measure.font=(first(font,'b')?'bold ':'')+size+'pt '+JSON.stringify(first(font,'name')?.getAttribute('val')||'Arial')+', sans-serif';width=Math.max(width,(measure.measureText(String(text)).width*(header?1.2:1.15)+(header?36:20))/7/(merged.get(address)||1));}if(flow)width=Math.max(width,dateColumns.has(point.c)?22:32);}
      widths[point.c]=Math.max(widths[point.c]||12,Math.min(header||flow?255:80,width));
    }
    if(batch||review){
      const detailStyles=new Map();
      for(const c of elements(data,'c')){
        const point=XLSX.utils.decode_cell(c.getAttribute('r'));if(!scoped(point.c)||(!review&&point.c!==detailColumn)||point.r<=headerRow)continue;
        const id=styleId(c);if(!detailStyles.has(id)){const xf=xfs.children[id].cloneNode(true);let alignment=first(xf,'alignment');if(!alignment){alignment=create(styles,'alignment');xf.insertBefore(alignment,first(xf,'protection')||null);}alignment.setAttribute('horizontal','left');alignment.setAttribute('vertical',review?'top':'center');alignment.setAttribute('wrapText','1');alignment.setAttribute('shrinkToFit','0');xf.setAttribute('applyAlignment','1');detailStyles.set(id,append('cellXfs',xf));}c.setAttribute('s',detailStyles.get(id));
      }
    }
    const reviewHeight=c=>{
      const point=XLSX.utils.decode_cell(c.getAttribute('r')),xf=xfs.children[styleId(c)],font=fonts.children[Number(xf.getAttribute('fontId')||0)],size=Number(first(font,'sz')?.getAttribute('val')||12);
      const available=Math.max(1,(widths[point.c]-3)*7);let lines=0;
      if(measure)measure.font=size+'pt '+JSON.stringify(first(font,'name')?.getAttribute('val')||'Arial')+', sans-serif';
      for(const line of cellText(c).split(/\r?\n/)){let used=0;lines++;for(const ch of line){const w=measure?measure.measureText(ch).width:((ch.charCodeAt(0)>255?2:1)*size);if(used&&used+w>available){lines++;used=0;}used+=w;}}
      return lines*(size*1.5+2)+10;
    };
    if(review)for(const c of cells.values()){
      const point=XLSX.utils.decode_cell(c.getAttribute('r'));if(point.r<=headerRow)continue;
      while(reviewHeight(c)>409.5&&widths[point.c]<255)widths[point.c]=Math.min(255,widths[point.c]*1.25);
      if(reviewHeight(c)>409.5)throw Error('付款申请核对第'+(point.r+1)+'行内容超过单行可完整显示的高度，请精简源说明后重新生成；未导出截断内容');
    }
    // A populated merged cell also needs borders on its perimeter, without changing the merge.
    for(const range of merges){if(Array.from({length:range.e.c-range.s.c+1},(_,i)=>range.s.c+i).some(c=>!scoped(c)))continue;const anchor=cells.get(XLSX.utils.encode_cell(range.s));if(!cellText(anchor).trim())continue;for(let r=range.s.r;r<=range.e.r;r++)for(let c=range.s.c;c<=range.e.c;c++){const edges=[c===range.s.c&&'left',c===range.e.c&&'right',r===range.s.r&&'top',r===range.e.r&&'bottom'].filter(Boolean);if(edges.length){const cell=ensureCell(r,c);cell.setAttribute('s',layoutStyle(styleId(cell),r===headerRow,flowColumns.has(c),edges));}}}
    cols.replaceChildren(...widths.flatMap((width,c)=>{
      const old=oldCols.find(e=>Number(e.getAttribute('min'))<=c+1&&Number(e.getAttribute('max'))>=c+1),col=old?old.cloneNode(true):create(doc,'col');
      if(!scoped(c)){if(!old)return [];col.setAttribute('min',c+1);col.setAttribute('max',c+1);return [col];}
      for(const [key,value] of Object.entries({min:c+1,max:c+1,width:Math.ceil(width*100)/100,bestFit:1,customWidth:1}))col.setAttribute(key,value);return col;
    }),...oldCols.filter(c=>Number(c.getAttribute('max'))>end.c+1).map(c=>{const col=c.cloneNode(true);col.setAttribute('min',Math.max(Number(col.getAttribute('min')),end.c+2));return col;}));
    if(doc!==sheet||!preserveRowLayout){
    let format=first(doc,'sheetFormatPr');if(!format){format=create(doc,'sheetFormatPr');doc.documentElement.insertBefore(format,cols);}format.setAttribute('defaultRowHeight','30');format.setAttribute('customHeight','1');
    for(const row of elements(data,'row')){
      let height=30;
      if(batch){
        height=Math.max(height,Number(row.getAttribute('ht'))||0);
        const detail=cells.get(XLSX.utils.encode_cell({r:Number(row.getAttribute('r'))-1,c:detailColumn}));
        if(detail){const xf=xfs.children[styleId(detail)],size=Number(first(fonts.children[Number(xf.getAttribute('fontId')||0)],'sz')?.getAttribute('val')||11),width=Math.max(1,widths[detailColumn]-3);const lines=cellText(detail).split(/\r?\n/).reduce((n,line)=>n+Math.max(1,Math.ceil([...line].reduce((w,ch)=>w+(/[\u0000-\u00ff]/.test(ch)?1:2),0)*size/11/width)),0);height=Math.max(height,lines*(size+5)+8);}
      }
      if((review||progress)&&Number(row.getAttribute('r'))>headerRow+1)for(const c of elements(row,'c'))height=Math.max(height,reviewHeight(c));
      row.setAttribute('ht',Math.min(409.5,height));row.setAttribute('customHeight','1');row.removeAttribute('hidden');
    }
    }
    const range=XLSX.utils.encode_range({s:{r:headerRow,c:0},e:end});let filter=first(doc,'autoFilter');
    if(!filter){filter=create(doc,'autoFilter');let after=data.nextElementSibling;while(after&&['sheetCalcPr','sheetProtection','protectedRanges','scenarios'].includes(after.localName))after=after.nextElementSibling;doc.documentElement.insertBefore(filter,after);}
    filter.setAttribute('ref',range);if(doc!==sheet||!preserveRowLayout)filter.replaceChildren();
    if(doc===sheet)for(const name of elements(workbook,'definedName'))if(name.getAttribute('name')==='_xlnm._FilterDatabase'&&name.getAttribute('localSheetId')==='0')name.textContent="'谈判付款'!$A$"+(headerRow+1)+':$'+XLSX.utils.encode_col(end.c)+'$'+(end.r+1);
  };
  const rows=new Map(elements(sheetData,'row').map(r=>[Number(r.getAttribute('r')),r]));
  const cells=new Map(elements(sheetData,'c').map(c=>[c.getAttribute('r'),c]));
  const submissionColumns=new Set((result.negotiated.headers||[]).flatMap((h,c)=>['提交时间汇总','提交付款申请单日期'].includes(String(h).normalize('NFKC').replace(/\s+/g,''))?[c]:[]));
  const sourceColumns=elements(sheet,'col');
  let maxRow=originalRange.e.r,maxColumn=originalRange.e.c;
  const changed=new Set();
  // A changed shared-formula master must not leave untouched followers without their formula.
  const touchedShared=new Set(result.negotiated.updates.map(u=>cells.has(u.address)?first(cells.get(u.address),'f'):null).filter(f=>f?.getAttribute('t')==='shared').map(f=>f.getAttribute('si')));
  for(const [address,c] of cells)if(!managedAddress(address)&&first(c,'f')?.getAttribute('t')==='shared'&&touchedShared.has(first(c,'f').getAttribute('si')))throw Error('谈判付款!'+address+' 的共享公式跨越业务列和保留列，无法安全更新；请先在原表将两列公式分开后重新上传');
  for(const [address,c] of cells){
    const f=first(c,'f');if(!f||f.getAttribute('t')!=='shared'||!touchedShared.has(f.getAttribute('si')))continue;
    const point=XLSX.utils.decode_cell(address),expanded=result.negotiated.rows?.[point.r-result.negotiated.headerRow-1]?.[point.c]?.f||f.textContent;
    if(!expanded)throw Error('谈判付款!'+address+' 的共享公式无法展开，请在 Excel/WPS 重算后上传');
    f.textContent=renameSelf(expanded);for(const key of ['t','si','ref'])f.removeAttribute(key);
  }
  const amountStyles=new Map();
  for(const {address,cell} of result.negotiated.updates){
    if(!/^[A-Z]+[1-9]\d*$/.test(address))throw Error('谈判付款更新地址无效');
    if(!managedAddress(address))throw Error('谈判付款更新超出业务列：'+address+'，未改动保留列');
    const point=XLSX.utils.decode_cell(address),old=cells.get(address);let row=rows.get(point.r+1);
    if(!row){row=create(sheet,'row',{r:point.r+1});const after=[...rows].filter(([r])=>r>point.r+1).sort((a,b)=>a[0]-b[0])[0]?.[1];sheetData.insertBefore(row,after||null);rows.set(point.r+1,row);}
    const isHeader=point.r===result.negotiated.headerRow;
    const submission=!isHeader&&submissionColumns.has(point.c),column=sourceColumns.find(c=>Number(c.getAttribute('min'))<=point.c+1&&Number(c.getAttribute('max'))>=point.c+1);
    let styleId=old?.getAttribute('s')??(old&&submission?(row.getAttribute('s')??column?.getAttribute('style')??0):styleFor(cell?.z||'',isHeader));
    if(typeof cell?.v==='number'&&cell.z&&XLSX.SSF.is_date(cell.z))styleId=dateDisplayStyle(Number(styleId));
    if(typeof cell?.v==='number'&&cell.z==='#,##0.00'){
      const key=String(styleId);if(!amountStyles.has(key)){const xf=xfs.children[Number(styleId)].cloneNode(true);xf.setAttribute('numFmtId',formatId(cell.z));xf.setAttribute('applyNumberFormat','1');amountStyles.set(key,append('cellXfs',xf));}styleId=amountStyles.get(key);
    }
    if(submission)styleId=wrapStyle(styleId);
    const next=writeCell(sheet,address,cell?.f?{...cell,f:renameSelf(cell.f)}:cell,styleId);
    // Retain unrelated cell metadata while replacing only value/formula payloads.
    if(old){for(const attr of old.attributes)if(!['r','s','t'].includes(attr.name))next.setAttribute(attr.name,attr.value);old.replaceWith(next);}
    else {const after=[...row.children].find(c=>c.localName==='c'&&XLSX.utils.decode_cell(c.getAttribute('r')).c>point.c);row.insertBefore(next,after||null);}
    cells.set(address,next);changed.add(address);maxRow=Math.max(maxRow,point.r);maxColumn=Math.max(maxColumn,point.c);
  }
  let dimension=first(sheet,'dimension');if(!dimension){dimension=create(sheet,'dimension');sheet.documentElement.insertBefore(dimension,sheet.documentElement.firstChild);}dimension.setAttribute('ref',XLSX.utils.encode_range({s:originalRange.s,e:{r:maxRow,c:maxColumn}}));
  if(maxColumn>originalRange.e.c){let cols=first(sheet,'cols');if(!cols){cols=create(sheet,'cols');sheet.documentElement.insertBefore(cols,sheetData);}cols.append(create(sheet,'col',{min:originalRange.e.c+2,max:maxColumn+1,width:28,customWidth:1}));}
  // Update simple dependent caches; unsupported affected formulas are recalculated by Excel/WPS.
  const formulaCells=[...cells].filter(([,c])=>first(c,'f'));
  const formulaText=(address,c)=>{const point=XLSX.utils.decode_cell(address);return first(c,'f')?.textContent||renameSelf(result.negotiated.rows?.[point.r-result.negotiated.headerRow-1]?.[point.c]?.f||'');};
  const localDependencies=formula=>formula.split(/("(?:[^"]|"")*")/g).map((part,index)=>index%2?'':part.replace(/('(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Z_\u4e00-\u9fff][A-Z0-9_.:\u4e00-\u9fff]*)!(\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?)/gi,(match,token,ref)=>token.replace(/^'|'$/g,'').replace(/''/g,"'")==='谈判付款'?ref:'')).join('').toUpperCase().replace(/\$/g,'');
  const affected=new Set(changed);let grew=true;
  while(grew){grew=false;for(const [address,c] of formulaCells){const f=first(c,'f'),formula=localDependencies(formulaText(address,c));const refs=formula.match(/\b[A-Z]{1,3}[1-9]\d*\b/g)||[];
    for(const match of formula.matchAll(/\b([A-Z]{1,3}[1-9]\d*):([A-Z]{1,3}[1-9]\d*)\b/g)){const a=XLSX.utils.decode_cell(match[1]),b=XLSX.utils.decode_cell(match[2]);if((b.r-a.r+1)*(b.c-a.c+1)<=20000)for(let r=a.r;r<=b.r;r++)for(let col=a.c;col<=b.c;col++)refs.push(XLSX.utils.encode_cell({r,c:col}));}if(!affected.has(address)&&((f.getAttribute('t')==='shared'&&!formula)||refs.some(r=>affected.has(r)))){affected.add(address);grew=true;}}}
  const evaluating=new Set(),values=new Map();
  function evaluateCell(address){
    if(values.has(address))return values.get(address);if(evaluating.has(address))throw Error('循环公式');evaluating.add(address);
    try{const c=cells.get(address);if(!c)return 0;const f=first(c,'f');let value;
      if(f&&affected.has(address)){value=evaluateFormula(formulaText(address,c));}
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
  for(const [address,c] of formulaCells)if(affected.has(address)&&managedAddress(address)){
    for(const v of elements(c,'v'))v.remove();
    try{const value=evaluateCell(address);c.removeAttribute('t');const v=create(sheet,'v');v.textContent=String(Math.round((value+Number.EPSILON)*1e8)/1e8);c.append(v);}catch(_){window.FundsExport.lastWarnings.push('谈判付款!'+address+' 保留公式，需在 Excel/WPS 打开时重新计算。');}
  }
  // Move the existing column XML, including styles and references, instead of rebuilding the source sheet.
  // Patch native notes before column relocation so notes and their VML anchors move together.
  const commentPrefix='申请日期（付款节奏跟进）：',commentAuthor='付款节奏跟进';
  const isToolComment=line=>line.startsWith(commentPrefix)||line.startsWith('申请日期（资金整理）：');
  const commentEntries=Object.entries(result.applicationComments||{}).filter(([address])=>highlightedCells.includes(address));
  const sheetRelationsPath=relPath(sheetPath),sheetRelations=zip.file(sheetRelationsPath)?await parse(zip,sheetRelationsPath):parseText(`<Relationships xmlns="${P}"/>`);
  const findRelation=type=>elements(sheetRelations,'Relationship').find(r=>r.getAttribute('Type')===R+'/'+type);
  const oldCommentRelation=findRelation('comments');
  if(oldCommentRelation||commentEntries.length){
    const comments=oldCommentRelation?await parse(zip,resolve(sheetPath,oldCommentRelation.getAttribute('Target'))):parseText(`<comments xmlns="${S}"><authors/><commentList/></comments>`);
    const authors=first(comments,'authors'),list=first(comments,'commentList');
    if(!authors||!list)throw Error('谈判付款原批注结构无法识别，未覆盖原批注');
    const textOf=node=>elements(node,'t').map(t=>t.textContent).join('');
    // Locate complete tool lines before editing runs: rich-text boundaries are not line boundaries.
    for(const comment of elements(list,'comment')){
      if(!managedAddress(comment.getAttribute('ref')))continue;
      const text=first(comment,'text');if(!text)continue;
      const original=textOf(text),removed=[];
      for(const line of original.matchAll(/[^\r\n]*(?:\r\n|\r|\n|$)/g))if(isToolComment(line[0].replace(/[\r\n]+$/,''))){
        const end=line.index+line[0].length,last=removed.at(-1);
        if(last?.[1]===line.index)last[1]=end;else removed.push([line.index,end]);
      }
      if(!removed.length)continue;
      const last=removed.at(-1);
      if(last[1]===original.length&&last[0]>0&&!/[\r\n]$/.test(original))last[0]-=original.slice(last[0]-2,last[0])==='\r\n'?2:1;
      let offset=0;
      for(const t of elements(text,'t')){
        const start=offset,end=offset+t.textContent.length;offset=end;let cursor=start,clean='';
        for(const [from,to] of removed){if(to<=start)continue;if(from>=end)break;clean+=original.slice(cursor,Math.max(cursor,from));cursor=Math.max(cursor,Math.min(end,to));}
        clean+=original.slice(cursor,end);if(clean===t.textContent)continue;
        if(clean){t.textContent=clean;t.setAttribute('xml:space','preserve');}
        else {const parent=t.parentElement;t.remove();if(parent.localName==='r'&&!elements(parent,'t').length)parent.remove();}
      }
      if(!textOf(text))comment.remove();
    }
    let authorId=elements(authors,'author').findIndex(a=>a.textContent===commentAuthor);
    if(authorId<0&&commentEntries.length){authorId=authors.children.length;const author=create(comments,'author');author.textContent=commentAuthor;authors.append(author);}
    for(const [address,message] of commentEntries){
      if(!/^[A-Z]+[1-9]\d*$/.test(address)||typeof message!=='string'||!message.startsWith(commentPrefix)||/[\r\n]/.test(message))throw Error('申请日期批注数据无效');
      if(!managedAddress(address))throw Error('谈判付款批注超出业务列：'+address+'，未改动保留列');
      let comment=elements(list,'comment').find(c=>c.getAttribute('ref')===address);
      if(!comment){comment=create(comments,'comment',{ref:address,authorId});comment.append(create(comments,'text'));list.append(comment);}
      const text=first(comment,'text');
      for(const t of [...text.children].filter(c=>c.localName==='t')){const run=create(comments,'r');t.replaceWith(run);run.append(t);}
      const run=create(comments,'r'),t=create(comments,'t');t.setAttribute('xml:space','preserve');t.textContent=(textOf(text)?'\n':'')+message;run.append(t);text.append(run);
    }
    const addRelation=(type,path)=>{
      let id='rIdFunds'+type;while(elements(sheetRelations,'Relationship').some(r=>r.getAttribute('Id')===id))id+='x';
      const rel=sheetRelations.createElementNS(P,'Relationship');for(const [k,v] of Object.entries({Id:id,Type:R+'/'+type,Target:'/'+path}))rel.setAttribute(k,v);sheetRelations.documentElement.append(rel);return rel;
    };
    const uniquePath=base=>{let path=base;while(zip.file(path))path=path.replace(/(\.[^.]+)$/,'x$1');return path;};
    const commentPath=oldCommentRelation?resolve(sheetPath,oldCommentRelation.getAttribute('Target')):uniquePath('xl/comments-funds.xml');
    if(!oldCommentRelation){addRelation('comments',commentPath);const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+commentPath);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');types.documentElement.append(type);}
    save(commentPath,comments);
    // Reuse SheetJS's native VML note shapes, keeping existing shapes (including controls) intact.
    const legacy=first(sheet,'legacyDrawing'),vmlRelation=legacy?elements(sheetRelations,'Relationship').find(r=>r.getAttribute('Id')===legacy.getAttributeNS(R,'id')):null;
    let vmlPath=vmlRelation?resolve(sheetPath,vmlRelation.getAttribute('Target')):null,vml=vmlPath?await parse(zip,vmlPath):null;
    const refs=new Set(elements(list,'comment').map(c=>c.getAttribute('ref')));
    const noteAddress=shape=>{const data=first(shape,'ClientData');return data?.getAttribute('ObjectType')==='Note'?XLSX.utils.encode_cell({r:Number(first(data,'Row')?.textContent),c:Number(first(data,'Column')?.textContent)}):null;};
    if(vml)for(const shape of elements(vml,'shape')){const address=noteAddress(shape);if(address&&managedAddress(address)&&!refs.has(address))shape.remove();}
    const existing=new Set(vml?elements(vml,'shape').map(noteAddress):[]),missing=[...refs].filter(ref=>managedAddress(ref)&&!existing.has(ref));
    if(missing.length){
      const scratch=XLSX.utils.book_new(),notes={};for(const ref of missing){notes[ref]={t:'s',v:'',c:[{a:commentAuthor,t:'申请日期'}]};notes[ref].c.hidden=true;}
      notes['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxRow,c:maxColumn}});XLSX.utils.book_append_sheet(scratch,notes,'批注');
      const parts=await JSZip.loadAsync(XLSX.write(scratch,{type:'array',bookType:'xlsx'})),path=Object.keys(parts.files).find(p=>p.endsWith('.vml'));
      if(!path)throw Error('申请日期批注图形生成失败');const generated=await parse(parts,path);
      if(!vml){vml=generated;vmlPath=uniquePath('xl/drawings/funds-comments.vml');const relation=addRelation('vmlDrawing',vmlPath),drawing=create(sheet,'legacyDrawing');drawing.setAttributeNS(R,'r:id',relation.getAttribute('Id'));sheet.documentElement.insertBefore(drawing,[...sheet.documentElement.children].find(n=>['legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'].includes(n.localName))||null);}
      else{
        for(const type of elements(generated,'shapetype'))if(!elements(vml,'shapetype').some(t=>t.getAttribute('id')===type.getAttribute('id')))vml.documentElement.append(vml.importNode(type,true));
        let id=1024;const ids=new Set(elements(vml,'shape').map(s=>s.getAttribute('id')));
        for(const shape of elements(generated,'shape')){while(ids.has('_x0000_s'+id))id++;const next=vml.importNode(shape,true);next.setAttribute('id','_x0000_s'+id);ids.add('_x0000_s'+id++);vml.documentElement.append(next);}
      }
      if(!elements(types,'Default').some(t=>t.getAttribute('Extension')==='vml')){const type=types.createElementNS(types.documentElement.namespaceURI,'Default');type.setAttribute('Extension','vml');type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.vmlDrawing');types.documentElement.append(type);}
    }
    if(vml)save(vmlPath,vml);save(sheetRelationsPath,sheetRelations);
  }
  const order=result.negotiated.columnOrder;
  let relocateExternalFormula=text=>text;
  const relocatedParts=new Set();
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
    const formula=(text,local=true)=>text.split(/("(?:[^"]|"")*")/g).map((part,index)=>index%2?part:part.replace(/(^|[^A-Z0-9_.\]'])((?:'(?:[^']|'')+'|[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![A-Z0-9_(])/gi,(match,prefix,token,ref)=>{
      const name=token?.slice(0,-1).replace(/^'|'$/g,'').replace(/''/g,"'");if(name?name!=='谈判付款':!local)return match;
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
        f.textContent=renameSelf(expanded);for(const attr of ['t','si','ref'])f.removeAttribute(attr);
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
    for(const name of elements(workbook,'definedName'))name.textContent=formula(name.textContent,name.getAttribute('localSheetId')==='0');
    relocateExternalFormula=text=>formula(text,false);
    const visited=new Set();
    const relocateRelated=async path=>{
      const relationPath=relPath(path);if(!zip.file(relationPath))return;
      for(const rel of elements(await parse(zip,relationPath),'Relationship')){
        if(rel.getAttribute('TargetMode')==='External')continue;
        const target=resolve(path,rel.getAttribute('Target'));if(visited.has(target)||!zip.file(target)||!(/\.(?:xml|vml)$/.test(target)))continue;visited.add(target);
        const doc=await parse(zip,target),table=first(doc,'table');
        for(const tag of ['f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula'])for(const f of elements(doc,tag))f.textContent=renameSelf(f.textContent);
        for(const link of elements(doc,'hyperlink'))if(link.hasAttribute('location'))link.setAttribute('location',renameSelf(link.getAttribute('location')));
        if(table){const start=XLSX.utils.decode_range(table.getAttribute('ref')).s.c,list=first(table,'tableColumns');if(list)list.replaceChildren(...[...list.children].map((el,i)=>({el,c:column(start+i)})).sort((a,b)=>a.c-b.c).map(x=>x.el));}
        remapFilters(doc);relocateDocument(doc);
        for(const tag of ['from','to'])for(const anchor of elements(doc,tag)){const col=first(anchor,'col');if(col)col.textContent=column(Number(col.textContent));}
        for(const el of elements(doc,'Column'))el.textContent=column(Number(el.textContent));
        for(const el of elements(doc,'Anchor')){const values=el.textContent.split(',').map(v=>v.trim());if(values.length===8){values[0]=column(Number(values[0]));values[4]=column(Number(values[4]));el.textContent=values.join(', ');}}
        save(target,doc);relocatedParts.add(target);await relocateRelated(target);
      }
    };
    await relocateRelated(sheetPath);
  }
  fitSheet(sheet,{r:maxRow,c:maxColumn},result.negotiated.headerRow||0);
  const appliedCells=highlightedCells.flatMap(address=>{
    if(typeof address!=='string'||!/^[A-Z]+[1-9]\d*$/.test(address))throw Error('谈判付款标红地址无效');
    if(!managedAddress(address))throw Error('谈判付款标红超出业务列：'+address+'，未改动保留列');
    const point=XLSX.utils.decode_cell(address);if(point.r<=(result.negotiated.headerRow||0))return [];
    if(point.r>maxRow||point.c>maxColumn)throw Error('谈判付款标红地址超出明细范围');
    const moved=order?.indexOf(point.c);return [XLSX.utils.encode_cell({r:point.r,c:moved>=0?moved:point.c})];
  });
  const appliedSet=new Set(appliedCells);for(const c of elements(sheetData,'c'))if(appliedSet.has(c.getAttribute('r')))applyEditableFill(c);
  const applicationNoteCells=new Set(commentEntries.map(([address])=>{const point=XLSX.utils.decode_cell(address),moved=order?.indexOf(point.c);return XLSX.utils.encode_cell({r:point.r,c:moved>=0?moved:point.c});}));
  // Only this export's application-date notes need repositioning; retain manual note placement.
  if(applicationNoteCells.size&&first(sheet,'legacyDrawing')){
    const legacy=first(sheet,'legacyDrawing'),relations=await parse(zip,sheetRelationsPath);
    const relation=legacy&&elements(relations,'Relationship').find(r=>r.getAttribute('Id')===legacy.getAttributeNS(R,'id'));
    if(relation){
      const path=resolve(sheetPath,relation.getAttribute('Target')),vml=await parse(zip,path),cols=elements(sheet,'col');
      const columnPixels=c=>{const col=cols.find(e=>Number(e.getAttribute('min'))<=c+1&&Number(e.getAttribute('max'))>=c+1);return col?.getAttribute('hidden')==='1'?0:Math.max(1,Math.floor((Number(col?.getAttribute('width')||8.43)+Math.round(128/7)/256)*7));};
      const rowPixels=r=>{const row=rows.get(r+1);return row?.getAttribute('hidden')==='1'?0:Number(row?.getAttribute('ht')||first(sheet,'sheetFormatPr')?.getAttribute('defaultRowHeight')||30)*4/3;};
      const edge=(start,offset,size,limit)=>{while(start<limit&&offset>=size(start)){offset-=size(start);start++;}return [start,Math.round(offset)];};
      for(const shape of elements(vml,'shape')){
        const data=first(shape,'ClientData');if(data?.getAttribute('ObjectType')!=='Note')continue;
        const r=Number(first(data,'Row')?.textContent),c=Number(first(data,'Column')?.textContent);
        if(!Number.isInteger(r)||!Number.isInteger(c)||r<0||c<0||r>1048575||c>16383)continue;
        if(!applicationNoteCells.has(XLSX.utils.encode_cell({r,c})))continue;
        const style=new Map((shape.getAttribute('style')||'').split(';').filter(p=>p.includes(':')).map(p=>{const at=p.indexOf(':');return [p.slice(0,at).trim(),p.slice(at+1).trim()];}));
        const dimension=(key,fallback)=>{const m=/^([\d.]+)(pt|px)$/i.exec(style.get(key)||'');return m&&Number(m[1])>0?Number(m[1])*(m[2].toLowerCase()==='pt'?4/3:1):fallback;};
        const width=dimension('width',288),height=dimension('height',128),startCol=Math.min(c+1,16383);
        const [endCol,endX]=edge(startCol,6+width,columnPixels,16383),[endRow,endY]=edge(r,2+height,rowPixels,1048575);
        let left=6,top=2;for(let i=0;i<startCol;i++)left+=columnPixels(i);for(let i=0;i<r;i++)top+=rowPixels(i);
        for(const [key,value] of Object.entries({position:'absolute',left:'0pt',top:'0pt','margin-left':left*.75+'pt','margin-top':top*.75+'pt',width:width*.75+'pt',height:height*.75+'pt'}))style.set(key,value);
        shape.setAttribute('style',[...style].map(([k,v])=>k+':'+v).join(';')+';');
        let anchor=first(data,'Anchor');if(!anchor){anchor=vml.createElementNS(data.namespaceURI,'Anchor');data.insertBefore(anchor,first(data,'Row'));}
        anchor.textContent=[startCol,6,r,2,endCol,endX,endRow,endY].join(', ');
      }
      save(path,vml);
    }
  }
  for(const item of plan)if(item.sourceName)sourceEntries.get(item.sourceName).setAttribute('name',item.name);
  selected.removeAttribute('state');
  const outputEntries=new Map([['谈判付款',selected]]),outputPaths=new Set([sheetPath]);
  let nextSheetId=Math.max(...entries.map(e=>Number(e.getAttribute('sheetId'))),0)+1;
  const generatedSheets=[['正常付款','Normal',result.normalHeaders,result.normalRows],['付款申请核对','Review',result.reviewHeaders,result.reviewRows],['付款进度整理','Progress',result.progressHeaders||FundsCore.PROGRESS_HEADERS,result.progressRows||[]]];
  for(const [name,key,headers,data] of generatedSheets){
    const generated=parseText(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><cols/><sheetData/></worksheet>`);
    const generatedData=first(generated,'sheetData');
    const allRows=[headers.map(v=>({t:'s',v})),...data];
    for(const [r,values] of allRows.entries()){
      const row=create(generated,'row',{r:r+1});
      for(let c=0;c<headers.length;c++){const original=values[c]||{t:'s',v:''},cell=key==='Review'&&r>0&&headers[c]==='核对说明'?{...original,v:FundsCore.formatReviewText(original.v)}:original;row.append(writeCell(generated,XLSX.utils.encode_cell({r,c}),cell,styleFor(cell.z||'',r===0,r>0&&(key==='Normal'&&appliedNormalRows.has(r+1)||key==='Progress'&&headers[c]==='是否已付款'&&cell.v==='已付款')),true));}
      generatedData.append(row);
    }
    const range=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:allRows.length-1,c:headers.length-1}});
    generated.documentElement.insertBefore(create(generated,'dimension',{ref:range}),generated.documentElement.firstChild);
    fitSheet(generated,{r:allRows.length-1,c:headers.length-1},0,key==='Review',key==='Progress');
    if(key==='Normal'){
      for(const col of elements(generated,'col'))if(FundsCore.NORMAL_HIDDEN_HEADERS.includes(headers[Number(col.getAttribute('min'))-1]))col.setAttribute('hidden','1');
    }
    let entry=sourceEntries.get(name),generatedPath;
    if(entry){
      const relationship=elements(rels,'Relationship').find(e=>e.getAttribute('Id')===entry.getAttributeNS(R,'id'));
      if(!relationship||relationship.getAttribute('Type')!==R+'/worksheet')throw Error(name+' 不是可更新的工作表');
      generatedPath=resolve('xl/workbook.xml',relationship.getAttribute('Target'));
      zip.remove(relPath(generatedPath));entry.removeAttribute('state');
    }else{
      let newId='rIdFunds'+key;generatedPath='xl/worksheets/funds-'+key.toLowerCase()+'.xml';
      while(elements(rels,'Relationship').some(e=>e.getAttribute('Id')===newId))newId+='x';
      while(zip.file(generatedPath))generatedPath=generatedPath.replace('.xml','x.xml');
      entry=create(workbook,'sheet',{name,sheetId:nextSheetId++});entry.setAttributeNS(R,'r:id',newId);
      const relationship=rels.createElementNS(P,'Relationship');relationship.setAttribute('Id',newId);relationship.setAttribute('Type',R+'/worksheet');relationship.setAttribute('Target',generatedPath.slice(3));rels.documentElement.append(relationship);
      const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/'+generatedPath);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.append(type);
    }
    outputEntries.set(name,entry);outputPaths.add(generatedPath);save(generatedPath,generated);
  }
  sheets.replaceChildren(...plan.map(item=>outputEntries.get(item.name)||sourceEntries.get(item.sourceName)));
  // Only adjust references into renamed/reordered sheets; other original XML stays byte-identical.
  for(const path of Object.keys(zip.files)){
    if(!/^xl\/.*\.xml$/.test(path)||outputPaths.has(path)||relocatedParts.has(path)||path==='xl/workbook.xml'||path===stylePath)continue;
    const text=await zip.file(path).async('string');
    if(!/<(?:[\w]+:)?(?:f|formula|formula1|formula2|calculatedColumnFormula|totalsRowFormula|hyperlink)\b/.test(text))continue;
    const doc=parseText(text);let changed=false;
    for(const tag of ['f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula'])for(const f of elements(doc,tag)){
      const updated=relocateExternalFormula(renameSelf(f.textContent));if(updated!==f.textContent){f.textContent=updated;changed=true;}
    }
    for(const link of elements(doc,'hyperlink'))if(link.hasAttribute('location')){
      const original=link.getAttribute('location'),updated=relocateExternalFormula(renameSelf(original));if(updated!==original){link.setAttribute('location',updated);changed=true;}
    }
    if(changed)save(path,doc);
  }
  for(const relationship of elements(rels,'Relationship'))if(relationship.getAttribute('Type')===R+'/calcChain')relationship.remove();
  for(const view of elements(workbook,'workbookView')){view.setAttribute('activeTab','0');view.setAttribute('firstSheet','0');}
  let calc=first(workbook,'calcPr');if(!calc){calc=create(workbook,'calcPr');workbook.documentElement.append(calc);}calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');
  for(const view of elements(sheet,'sheetView')){view.setAttribute('tabSelected','1');view.setAttribute('workbookViewId','0');}
  save(sheetPath,sheet);save(stylePath,styles);save('xl/workbook.xml',workbook);save('xl/_rels/workbook.xml.rels',rels);
  // All retained worksheets keep their relationships, drawings, comments and other native parts.
  const reachable=new Set(['[Content_Types].xml','_rels/.rels']);
  const visit=async(path)=>{if(reachable.has(path))return;reachable.add(path);const relationPath=relPath(path);if(!zip.file(relationPath))return;reachable.add(relationPath);const doc=await parse(zip,relationPath);for(const relation of elements(doc,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve(path,relation.getAttribute('Target')));};
  const root=await parse(zip,'_rels/.rels');for(const relation of elements(root,'Relationship'))if(relation.getAttribute('TargetMode')!=='External')await visit(resolve('',relation.getAttribute('Target')));
  for(const path of Object.keys(zip.files))if(!zip.files[path].dir&&!reachable.has(path))zip.remove(path);
  for(const override of elements(types,'Override'))if(!zip.file(override.getAttribute('PartName').replace(/^\//,'')))override.remove();
  save('[Content_Types].xml',types);
  if(zip.file('docProps/app.xml')){const app=await parse(zip,'docProps/app.xml');for(const title of elements(app,'TitlesOfParts')){const vector=first(title,'vector');if(vector){vector.setAttribute('size',plan.length);vector.replaceChildren();for(const {name} of plan){const text=app.createElementNS(vector.namespaceURI,'vt:lpstr');text.textContent=name;vector.append(text);}}}const pairs=first(app,'HeadingPairs');if(pairs){const count=first(pairs,'i4');if(count)count.textContent=plan.length;}save('docProps/app.xml',app);}
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
};
