/* Add a display-only month header without rebuilding the source workbook. */
window.FundsMonthHeader = async function(zip, negotiated) {
  'use strict';
  const bands=negotiated.headerBands||[];
  if(!bands.length)return zip;
  const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const parser=new DOMParser(),serializer=new XMLSerializer();
  const all=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)],first=(doc,name)=>all(doc,name)[0];
  const parse=text=>{const doc=parser.parseFromString(text,'application/xml');if(first(doc,'parsererror'))throw Error('月份表头：工作簿 XML 无法解析');return doc;};
  const read=async path=>{if(!zip.file(path))throw Error('月份表头：工作簿缺少 '+path);return parse(await zip.file(path).async('string'));};
  const save=(path,doc)=>zip.file(path,serializer.serializeToString(doc));
  const create=(doc,name,attrs={})=>{const el=doc.createElementNS(S,name);for(const [k,v]of Object.entries(attrs))el.setAttribute(k,String(v));return el;};
  const resolve=(base,target)=>new URL(target,'https://xlsx.invalid/'+base).pathname.slice(1),relPath=path=>path.replace(/([^/]+)$/,'_rels/$1.rels');
  const workbook=await read('xl/workbook.xml'),rels=await read('xl/_rels/workbook.xml.rels'),entries=all(workbook,'sheet');
  const entry=entries.find(e=>e.getAttribute('name')==='谈判付款');if(!entry)throw Error('月份表头：未找到谈判付款工作表');
  const relation=all(rels,'Relationship').find(e=>e.getAttribute('Id')===entry.getAttributeNS(R,'id'));
  if(!relation)throw Error('月份表头：谈判付款工作表关系缺失');
  const path=resolve('xl/workbook.xml',relation.getAttribute('Target')),doc=await read(path),data=first(doc,'sheetData');
  const oldHeader=negotiated.headerRow,insert=negotiated.headerTopRow==null;
  if(!Number.isInteger(oldHeader)||oldHeader<0||oldHeader>=1048575)throw Error('月份表头：表头行号无效');
  const top=insert?oldHeader:negotiated.headerTopRow,lower=insert?oldHeader+1:oldHeader;
  if(!Number.isInteger(top)||lower!==top+1||!data)throw Error('月份表头：双层表头结构无效');
  const occupied=new Set();
  for(const b of bands){
    if(!Number.isInteger(b.startCol)||!Number.isInteger(b.endCol)||b.startCol<0||b.endCol<b.startCol||b.endCol>16383||!FundsTotalValidation.isMonthBandLabel(b.label))throw Error('月份表头：月份分组无效');
    for(let c=b.startCol;c<=b.endCol;c++){if(occupied.has(c))throw Error('月份表头：月份分组重叠');occupied.add(c);}
  }
  const originalRows=all(data,'row'),originalCells=all(data,'c');
  const lastCol=Math.max(...originalCells.map(c=>XLSX.utils.decode_cell(c.getAttribute('r')).c),...bands.map(b=>b.endCol));
  const shiftRow=r=>{const next=insert&&r>=oldHeader?r+1:r;if(next>1048575)throw Error('月份表头：插入行将超出工作表行数上限');return next;};
  const point=(p,anchor=false)=>({...p,r:insert&&anchor&&p.r===oldHeader&&!occupied.has(p.c)?top:shiftRow(p.r)});
  const reference=(ref,anchor=false)=>{
    const parts=ref.split(':'),points=parts.map(value=>{const m=/^(\$?)([A-Z]{1,3})(\$?)(\d+)$/i.exec(value);if(!m)throw Error('月份表头：无法调整单元格引用 '+ref);const p=point({c:XLSX.utils.decode_col(m[2].toUpperCase()),r:Number(m[4])-1},anchor);return {p,prefix:m[1],rowPrefix:m[3]};});
    if(points.length===2){const min=Math.min(points[0].p.r,points[1].p.r),max=Math.max(points[0].p.r,points[1].p.r);points[0].p.r=min;points[1].p.r=max;}
    return points.map(({p,prefix,rowPrefix})=>prefix+XLSX.utils.encode_col(p.c)+rowPrefix+(p.r+1)).join(':');
  };
  const rowReference=ref=>ref.split(':').map(part=>{const m=/^(\$?)(\d+)$/.exec(part);if(!m)throw Error('月份表头：无法调整行引用 '+ref);return m[1]+(shiftRow(Number(m[2])-1)+1);}).join(':');
  // Protect literal strings and structured-reference column names from A1 matching.
  const outsideLiterals=(text,change)=>{
    let out='',start=0,i=0;
    while(i<text.length){
      if(text[i]==="'"){i++;while(i<text.length){if(text[i++]==="'"){if(text[i]==="'"){i++;continue;}break;}}continue;}
      if(text[i]!== '"'&&text[i]!=='['){i++;continue;}
      out+=change(text.slice(start,i));const begin=i;
      if(text[i]==='"'){i++;while(i<text.length){if(text[i++]==='"'){if(text[i]==='"'){i++;continue;}break;}}}
      else{
        let depth=0;do{if(text[i]==='[')depth++;else if(text[i]===']')depth--;i++;}while(i<text.length&&depth);
        const external=/^[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/i.exec(text.slice(i));if(external)i+=external[0].length;
      }
      out+=text.slice(begin,i);start=i;
    }
    return out+change(text.slice(start));
  };
  const names=entries.map(e=>e.getAttribute('name'));
  const columnRange=/^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}$/i;
  const checkHeaderRange=ref=>{
    if(!ref.includes(':'))return;
    if(columnRange.test(ref)){
      const columns=ref.split(':').map(value=>XLSX.utils.decode_col(value.replace('$','').toUpperCase()));
      for(let c=Math.min(...columns);c<=Math.max(...columns);c++)if(occupied.has(c))throw Error('月份表头：公式范围 '+ref+' 包含将拆分的表头，无法安全保持原计算结果');
      return;
    }
    if(/^\$?\d+:\$?\d+$/.test(ref)){
      const rows=ref.split(':').map(value=>Number(value.replace('$',''))-1);
      if(oldHeader>=Math.min(...rows)&&oldHeader<=Math.max(...rows))throw Error('月份表头：公式范围 '+ref+' 包含将拆分的表头，无法安全保持原计算结果');
      return;
    }
    const range=XLSX.utils.decode_range(ref.replace(/\$/g,'')),start=Math.min(range.s.r,range.e.r),end=Math.max(range.s.r,range.e.r);
    if(oldHeader<start||oldHeader>end)return;
    let dates=false,other=false;for(let c=Math.min(range.s.c,range.e.c);c<=Math.max(range.s.c,range.e.c);c++)occupied.has(c)?dates=true:other=true;
    if(start<oldHeader||dates&&other||start!==end&&other)throw Error('月份表头：公式范围 '+ref+' 包含将拆分的表头，无法安全保持原计算结果');
  };
  const formula=(text,local=false,displayRange=false)=>!insert?text:outsideLiterals(text,part=>{
    for(const match of part.matchAll(/(?:'((?:[^']|'')+)'|([A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*:[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*))!/gi)){
      const endpoints=(match[1]||match[2]).replace(/''/g,"'").split(':');if(endpoints.length!==2)continue;
      const a=names.indexOf(endpoints[0]),b=names.indexOf(endpoints[1]),target=names.indexOf('谈判付款');
      if(a>=0&&b>=0&&target>=Math.min(a,b)&&target<=Math.max(a,b))throw Error('月份表头：涉及谈判付款的三维公式引用无法安全插行，请改为逐表引用');
    }
    return part.replace(/(^|[^A-Z0-9_.\[\]'])((?:'(?:[^']|'')+'|[A-Z_\u4e00-\u9fff][A-Z0-9_.\u4e00-\u9fff]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?\d+:\$?\d+|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![A-Z0-9_(])/gi,(match,prefix,token,ref)=>{
      const name=token?.slice(0,-1).replace(/^'|'$/g,'').replace(/''/g,"'");if(name?name!=='谈判付款':!local)return match;
      if(!displayRange)checkHeaderRange(ref);
      if(columnRange.test(ref))return match;
      return prefix+(token||'')+(/^\$?\d/.test(ref)?rowReference(ref):reference(ref,true));
    });
  });
  const formulas=new Set(['f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula']);
  const coordinate=/^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?:\s+\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)*$/i;
  const relocate=(target,local,coordinates)=>{
    let changed=false;
    for(const el of target.getElementsByTagName('*')){
      if(coordinates)for(const attr of ['ref','sqref','activeCell','topLeftCell'])if(el.hasAttribute(attr)&&coordinate.test(el.getAttribute(attr))){
        const next=el.getAttribute(attr).split(/\s+/).map(value=>reference(value,['comment','threadedComment','hyperlink'].includes(el.localName))).join(' ');if(next!==el.getAttribute(attr)){el.setAttribute(attr,next);changed=true;}
      }
      if(formulas.has(el.localName)){const next=formula(el.textContent,local);if(next!==el.textContent){el.textContent=next;changed=true;}}
      if(el.localName==='hyperlink'&&el.hasAttribute('location')){const next=formula(el.getAttribute('location'),local);if(next!==el.getAttribute('location')){el.setAttribute('location',next);changed=true;}}
      if(el.localName==='worksheetSource'&&el.getAttribute('sheet')==='谈判付款'&&el.hasAttribute('ref')){el.setAttribute('ref',reference(el.getAttribute('ref')));changed=true;}
    }
    return changed;
  };
  const visited=new Set([path]),related=[];
  const collect=async base=>{
    const rp=relPath(base);if(!zip.file(rp))return;
    for(const rel of all(await read(rp),'Relationship')){
      if(rel.getAttribute('TargetMode')==='External')continue;
      const target=resolve(base,rel.getAttribute('Target'));if(visited.has(target)||!zip.file(target)||!(/\.(?:xml|vml)$/.test(target)))continue;visited.add(target);
      related.push({path:target,type:rel.getAttribute('Type'),doc:await read(target)});await collect(target);
    }
  };
  await collect(path);
  for(const item of related){const table=first(item.doc,'table');if(table){const range=XLSX.utils.decode_range(table.getAttribute('ref'));if(range.s.r<=oldHeader&&range.e.r>=oldHeader)throw Error('月份表头：原表头位于 Excel 结构化表格内，合并表头会破坏表格结构，请先转为普通区域');}}
  if(insert){
    const headerViews=all(doc,'sheetView').filter(view=>view.hasAttribute('topLeftCell')&&XLSX.utils.decode_cell(view.getAttribute('topLeftCell')).r===oldHeader);
    relocate(doc,true,true);
    for(const view of headerViews){const p=XLSX.utils.decode_cell(view.getAttribute('topLeftCell'));p.r=top;view.setAttribute('topLeftCell',XLSX.utils.encode_cell(p));}
    for(const cell of originalCells)cell.setAttribute('r',reference(cell.getAttribute('r')));
    for(const row of originalRows){row.setAttribute('r',shiftRow(Number(row.getAttribute('r'))-1)+1);row.removeAttribute('spans');}
    for(const pane of all(doc,'pane'))if(/^frozen/.test(pane.getAttribute('state')||'')){const count=Number(pane.getAttribute('ySplit')||0);if(count>oldHeader)pane.setAttribute('ySplit',count+1);}
    for(const breaks of all(doc,'rowBreaks'))for(const brk of all(breaks,'brk'))brk.setAttribute('id',shiftRow(Number(brk.getAttribute('id'))));
    const headerHeight=Number(originalRows.find(row=>Number(row.getAttribute('r'))===lower+1)?.getAttribute('ht')||30);
    for(const item of related){
      const {doc:part}=item;relocate(part,!/\/chart(?:s)?\//.test(item.path),true);
      for(const tag of ['from','to'])for(const anchor of all(part,tag)){const row=first(anchor,'row');if(row)row.textContent=shiftRow(Number(row.textContent));}
      for(const client of all(part,'ClientData')){
        const row=first(client,'Row'),col=first(client,'Column'),oldRow=row?Number(row.textContent):null;
        if(row)row.textContent=oldRow===oldHeader&&col&&!occupied.has(Number(col.textContent))?top:shiftRow(oldRow);
        const anchor=first(client,'Anchor');if(anchor){const a=anchor.textContent.split(',').map(v=>v.trim());if(a.length===8){a[2]=String(shiftRow(Number(a[2])));a[6]=String(shiftRow(Number(a[6])));anchor.textContent=a.join(', ');}}
        if(oldRow!==null&&oldRow>=oldHeader){const shape=client.parentNode,style=shape.getAttribute('style');if(style)shape.setAttribute('style',style.replace(/(margin-top\s*:\s*)([-\d.]+)(pt|px)/i,(_,prefix,value,unit)=>prefix+(Number(value)+headerHeight*(unit==='px'?4/3:1))+unit));}
      }
      save(item.path,part);
    }
    for(const name of all(workbook,'definedName')){
      const own=Number(name.getAttribute('localSheetId'))===entries.indexOf(entry)&&name.hasAttribute('localSheetId'),original=name.textContent;
      name.textContent=formula(original,own,['_xlnm.Print_Area','_xlnm.Print_Titles','_xlnm._FilterDatabase'].includes(name.getAttribute('name')));
      if(own&&name.getAttribute('name')==='_xlnm.Print_Titles')name.textContent=name.textContent.replace(/(\$?)(\d+):(\$?)(\d+)/g,(match,a,start,b,end)=>Number(start)===lower+1?a+(top+1)+':'+b+end:match);
      if(own&&name.getAttribute('name')==='_xlnm.Print_Area')name.textContent=name.textContent.replace(/(\$?[A-Z]{1,3}\$?)(\d+)(?=:)/gi,(match,prefix,row)=>Number(row)===lower+1?prefix+(top+1):match);
    }
    // Other sheets retain their local coordinates; only references into the changed sheet move.
    for(const name of Object.keys(zip.files)){
      if(!/^xl\/.*\.xml$/.test(name)||visited.has(name)||name==='xl/workbook.xml')continue;
      const text=await zip.file(name).async('string');if(!/谈判付款/.test(text))continue;
      const external=parse(text);if(relocate(external,false,false))save(name,external);
    }
  }
  const rows=new Map(all(data,'row').map(row=>[Number(row.getAttribute('r'))-1,row]));
  let topRow=rows.get(top);const lowerRow=rows.get(lower);if(!lowerRow)throw Error('月份表头：原日期表头行不存在');
  if(!topRow){topRow=create(doc,'row',{r:top+1,ht:lowerRow.getAttribute('ht')||30,customHeight:'1'});data.insertBefore(topRow,lowerRow);}
  const cells=new Map(all(data,'c').map(cell=>[cell.getAttribute('r'),cell]));
  const ensure=(r,c,template)=>{const address=XLSX.utils.encode_cell({r,c});let cell=cells.get(address);if(!cell){cell=create(doc,'c',{r:address});if(template?.hasAttribute('s'))cell.setAttribute('s',template.getAttribute('s'));const row=r===top?topRow:lowerRow;row.append(cell);cells.set(address,cell);}return cell;};
  const payload=cell=>[...cell.children].some(el=>el.localName==='f'||el.localName==='v'&&el.textContent!==''||el.localName==='is'&&el.textContent!=='');
  const clear=cell=>{for(const child of [...cell.children])if(['f','v','is'].includes(child.localName))child.remove();cell.removeAttribute('t');};
  const mergeContainer=first(doc,'mergeCells')||create(doc,'mergeCells');
  if(!mergeContainer.parentNode){const after=[...doc.documentElement.children].find(el=>['phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'].includes(el.localName));doc.documentElement.insertBefore(mergeContainer,after||null);}
  for(const merge of [...mergeContainer.children]){
    const range=XLSX.utils.decode_range(merge.getAttribute('ref'));
    if(range.e.r<top||range.s.r>lower)continue;
    if(range.s.r<top||range.e.r>lower)throw Error('月份表头：已有合并范围 '+merge.getAttribute('ref')+' 跨越表头与数据，不能覆盖');
    if(insert)throw Error('月份表头：原表头存在合并范围 '+merge.getAttribute('ref')+'，请先确认表头结构');
    merge.remove();
  }
  const addMerge=(s,e)=>{if(s.r!==e.r||s.c!==e.c)mergeContainer.append(create(doc,'mergeCell',{ref:XLSX.utils.encode_range({s,e})}));};
  for(let c=0;c<=lastCol;c++)if(!occupied.has(c)){
    const bottom=ensure(lower,c),upper=ensure(top,c,bottom);
    if(bottom.hasAttribute('s'))upper.setAttribute('s',bottom.getAttribute('s'));
    if(payload(bottom)){
      if(payload(upper)&&upper.textContent!==bottom.textContent)throw Error('月份表头：'+upper.getAttribute('r')+' 与 '+bottom.getAttribute('r')+' 内容冲突，未覆盖原表头');
      const replacement=bottom.cloneNode(true);replacement.setAttribute('r',upper.getAttribute('r'));upper.replaceWith(replacement);cells.set(replacement.getAttribute('r'),replacement);clear(bottom);
    }
    addMerge({r:top,c},{r:lower,c});
  }
  for(const band of bands){
    for(let c=band.startCol;c<=band.endCol;c++){
      const bottom=ensure(lower,c),cell=ensure(top,c,bottom);if(bottom.hasAttribute('s'))cell.setAttribute('s',bottom.getAttribute('s'));clear(cell);
      if(c===band.startCol){cell.setAttribute('t','inlineStr');const value=create(doc,'is'),text=create(doc,'t');text.textContent=band.label;value.append(text);cell.append(value);}
    }
    addMerge({r:top,c:band.startCol},{r:top,c:band.endCol});
  }
  mergeContainer.setAttribute('count',mergeContainer.children.length);
  for(const row of [topRow,lowerRow]){row.removeAttribute('spans');const cells=[...row.children].filter(el=>el.localName==='c').sort((a,b)=>XLSX.utils.decode_cell(a.getAttribute('r')).c-XLSX.utils.decode_cell(b.getAttribute('r')).c);for(const cell of cells)row.insertBefore(cell,first(row,'extLst')||null);}
  const dimension=first(doc,'dimension');if(dimension){const range=XLSX.utils.decode_range(dimension.getAttribute('ref'));range.s.r=Math.min(range.s.r,top);dimension.setAttribute('ref',XLSX.utils.encode_range(range));}
  const filter=first(doc,'autoFilter');if(filter){const range=XLSX.utils.decode_range(filter.getAttribute('ref'));range.s.r=lower;filter.setAttribute('ref',XLSX.utils.encode_range(range));for(const name of all(workbook,'definedName'))if(name.getAttribute('name')==='_xlnm._FilterDatabase'&&Number(name.getAttribute('localSheetId'))===entries.indexOf(entry))name.textContent="'谈判付款'!$"+XLSX.utils.encode_col(range.s.c)+'$'+(lower+1)+':$'+XLSX.utils.encode_col(range.e.c)+'$'+(range.e.r+1);}
  save(path,doc);save('xl/workbook.xml',workbook);
  return zip;
};
