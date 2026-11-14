(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'), require('./product-dimension-core.js'));
  else root.ProductDimensionExport = factory(root.XLSX, root.JSZip, root.ProductDimensionCore);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip, Core) {
  'use strict';
  const escape = value => String(value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  const cellText = value => escape(String(value).replace(/_x[0-9a-f]{4}_/gi,m=>'_x005F_'+m.slice(1)).replace(/[\u0000-\u0008\u000b-\u001f]/g,c=>'_x'+c.charCodeAt(0).toString(16).padStart(4,'0')+'_'));
  const decode = value => value.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-f]+|#\d+);/gi,(all,key)=>key[0]==='#'?String.fromCodePoint(parseInt(key.slice(key[1].toLowerCase()==='x'?2:1),key[1].toLowerCase()==='x'?16:10)):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[key]||all));
  const tagRE = (name,flags='g') => new RegExp(`<((?:[\\w.-]+:)?${name})\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/\\1\\s*>)`,flags);
  const tags = (xml,name) => xml.match(tagRE(name)) || [];
  const first = (xml,name) => tags(xml,name)[0];
  const opening = xml => xml.match(/^<[^>]+>/)[0];
  const content = xml => /\/>$/.test(xml)?'':xml.replace(/^<[^>]+>/,'').replace(/<\/[^>]+>$/,'');
  const prefix = xml => xml.match(/^<([\w.-]+:)?/)?.[1] || '';
  const attr = (xml,name) => {const m=opening(xml).match(new RegExp('(?:^|\\s)'+name+'=(["\'])([\\s\\S]*?)\\1'));return m?decode(m[2]):null;};
  function setAttr(xml,name,value) {
    const head=opening(xml),re=new RegExp('\\s+'+name+'=(["\'])[\\s\\S]*?\\1');
    const next=value===null?head.replace(re,''):re.test(head)?head.replace(re,' '+name+'="'+escape(value)+'"'):head.replace(/\/?>(?=$)/,end=>' '+name+'="'+escape(value)+'"'+end);
    return next+xml.slice(head.length);
  }
  const body = (xml,value) => opening(xml).replace(/\/>$/, '>')+value+`</${xml.match(/^<([^\s/>]+)/)[1]}>`;
  const relationPath = path => path.replace(/([^/]+)$/,'_rels/$1.rels');
  function resolve(base,target) {
    const result=[];
    for(const part of (target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target).split('/')){
      if(part==='..')result.pop();else if(part&&part!=='.')result.push(part);
    }
    return result.join('/');
  }
  async function read(zip,path) {
    if(!zip.file(path))throw Error('工作簿缺少 '+path+'，无法保真导出。');
    return zip.file(path).async('string');
  }
  function appendStyle(xml,name,item,child) {
    const old=first(xml,name);if(!old)throw Error('工作簿样式结构不完整：'+name);
    const p=prefix(old),qualified=item.replace(/<(\/?)([A-Za-z][\w]*)(?=[\s/>])/g,'<$1'+p+'$2');
    return xml.replace(old,setAttr(body(old,content(old)+qualified),'count',tags(content(old),child).length+tags(item,child).length));
  }
  function addStyles(xml,formats) {
    const count=(name,child)=>tags(content(first(xml,name)||'<none/>'),child).length;
    const fill=count('fills','fill'),border=count('borders','border'),start=count('cellXfs','xf');
    xml=appendStyle(xml,'fills','<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>','fill');
    xml=appendStyle(xml,'borders','<border>'+['left','right','top','bottom'].map(side=>`<${side} style="hair"><color rgb="FF404040"/></${side}>`).join('')+'<diagonal/></border>','border');
    const formatIds=new Map([['General',0],['@',49]]),existing=tags(first(xml,'numFmts')||'','numFmt');
    for(const node of existing)formatIds.set(attr(node,'formatCode'),Number(attr(node,'numFmtId')));
    let next=Math.max(163,...existing.map(node=>Number(attr(node,'numFmtId'))))+1;
    for(const format of formats)if(!formatIds.has(format)){
      if(!first(xml,'numFmts')){const fonts=first(xml,'fonts'),p=prefix(fonts);xml=xml.replace(fonts,`<${p}numFmts count="0"/>`+fonts);}
      xml=appendStyle(xml,'numFmts',`<numFmt numFmtId="${next}" formatCode="${escape(format)}"/>`,'numFmt');formatIds.set(format,next++);
    }
    const styles=new Map();
    for(const format of new Set(['General','@',...formats])){
      styles.set(format,start+styles.size);
      xml=appendStyle(xml,'cellXfs',`<xf numFmtId="${formatIds.get(format)}" fontId="0" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`,'xf');
    }
    return {xml,styles};
  }
  function cachedValue(cell,sheetName) {
    const f=first(cell,'f');if(!f)return cell;
    if(['array','dataTable'].includes(attr(f,'t')))throw Error(`${sheetName}!${attr(cell,'r')}使用数组或数据表公式，请先转为值后再维护。`);
    const v=first(cell,'v'),type=attr(cell,'t'),value=v?decode(content(v)):null;
    if(!v || type==='e' || (!type||type==='n')&&(value===''||!Number.isFinite(Number(value))))throw Error(`${sheetName}!${attr(cell,'r')}的公式缺少有效的已保存结果，请先在Excel中计算并保存。`);
    cell=cell.replace(f,'');
    if(type==='str')cell=body(setAttr(cell,'t','inlineStr'),`<${prefix(cell)}is><${prefix(cell)}t xml:space="preserve">${content(v)}</${prefix(cell)}t></${prefix(cell)}is>`);
    return cell;
  }
  function rangeMapper(rowMap) {
    const entries=[...rowMap].sort((a,b)=>a[0]-b[0]);
    return function(ref) {
      const output=[];
      for(const part of String(ref||'').trim().split(/\s+/).filter(Boolean)){
        const m=/^(\$?[A-Z]{1,3})(\$?)(\d+)(?::(\$?[A-Z]{1,3})(\$?)(\d+))?$/.exec(part);
        if(!m)throw Error('暂不支持主表范围 '+part+'，请改为普通单元格范围。');
        const start=Number(m[3]),end=Number(m[6]||m[3]);
        const retained=entries.filter(([row])=>row>=start&&row<=end);
        if(!retained.length)continue;
        const a=m[1]+m[2]+retained[0][1],b=(m[4]||m[1])+(m[5]||m[2])+retained[retained.length-1][1];
        output.push(a===b?a:a+':'+b);
      }
      return output.join(' ');
    };
  }
  async function remapNotes(zip,rels,sheetPath,rowMap,mapRange) {
    for(const rel of tags(rels,'Relationship')){
      const type=attr(rel,'Type'),path=resolve(sheetPath,attr(rel,'Target'));
      if(type.endsWith('/comments')){
        let xml=await read(zip,path);
        xml=xml.replace(tagRE('comment'),node=>{const ref=mapRange(attr(node,'ref'));return ref?setAttr(node,'ref',ref):'';});zip.file(path,xml);
      }else if(type.endsWith('/vmlDrawing')){
        let xml=await read(zip,path);
        xml=xml.replace(tagRE('shape'),node=>{
          const data=first(node,'ClientData');
          if(!data||attr(data,'ObjectType')!=='Note')throw Error('主表包含非批注的VML控件，暂不支持重排行，请先移除控件。');
          const row=first(data,'Row'),oldRow=Number(content(row||'<Row/>'))+1,next=rowMap.get(oldRow);
          if(!next)return '';
          node=node.replace(row,body(row,String(next-1)));
          const anchor=first(node,'Anchor');if(anchor){const parts=content(anchor).split(',').map(v=>Number(v.trim()));if(parts.length!==8||parts.some(v=>!Number.isFinite(v)))throw Error('主表批注锚点无效。');parts[2]=Math.max(0,parts[2]+next-oldRow);parts[6]=Math.max(parts[2],parts[6]+next-oldRow);node=node.replace(anchor,body(anchor,parts.join(',')));}
          return node;
        });zip.file(path,xml);
      }else if(!type.endsWith('/hyperlink'))throw Error('主表包含暂不支持重排的绘图、表格或扩展关系，请先移除后再维护。');
    }
  }
  return async function(sourceBytes,analysis,result) {
    if(!result?.canExport||!Array.isArray(result.rows)||result.issues?.some(item=>item.severity==='error'))throw Error('存在待处理错误，不能下载商品维度表。');
    const {dimensionSheet:name,headerRow,headers}=analysis||{};
    if(!name||!Number.isSafeInteger(headerRow)||headerRow<1||!headers?.length||headers.length>16384||result.rows.length+headerRow>1048576)throw Error('商品维度表结果或行列范围无效。');
    const codeColumn=headers.indexOf('物料编码');if(codeColumn<0)throw Error('商品维度表缺少物料编码列。');
    const errors=Core.validateResult(analysis,result);
    if(errors.length)throw Error('商品维度表校验未通过：'+errors.map(error=>error.message||error.reason).join('；'));
    const addedCodes=new Set(analysis.added.map(group=>group.code)),isNew=row=>addedCodes.has(row.code);
    let zip;try{zip=await Zip.loadAsync(sourceBytes);}catch(_){throw Error('商品维度表必须为原始.xlsx文件，请另存为.xlsx后再上传。');}
    let workbook=await read(zip,'xl/workbook.xml'),workbookRels=await read(zip,'xl/_rels/workbook.xml.rels'),types=await read(zip,'[Content_Types].xml');
    if(!tags(types,'Override').some(n=>attr(n,'PartName')==='/xl/workbook.xml'&&attr(n,'ContentType')==='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'))throw Error('商品维度表仅支持.xlsx格式。');
    const entries=tags(first(workbook,'sheets')||'','sheet'),sheetIndex=entries.findIndex(n=>attr(n,'name')===name),entry=entries[sheetIndex];
    const relationships=tags(workbookRels,'Relationship'),rel=entry&&relationships.find(n=>attr(n,'Id')===attr(entry,'(?:[\\w.-]+:)?id'));
    if(!rel)throw Error('原工作簿未找到商品维度表。');
    const path=resolve('xl/workbook.xml',attr(rel,'Target')),styleRel=relationships.find(n=>attr(n,'Type').endsWith('/styles'));
    if(!styleRel)throw Error('原工作簿缺少样式。');
    let sheet=await read(zip,path),sheetRels=zip.file(relationPath(path))?await read(zip,relationPath(path)):null;
    const data=first(sheet,'sheetData');if(!data)throw Error('商品维度表缺少sheetData。');
    for(const tag of ['mergeCells','tableParts','sheetProtection','protectedRanges','dataValidations','drawing','oleObjects','controls','scenarios','customSheetViews'])if(first(sheet,tag))throw Error('商品维度表包含暂不支持的'+tag+'，请先移除后再维护。');
    const originalRows=new Map(tags(data,'row').map(row=>[Number(attr(row,'r')),row]));
    if(!originalRows.has(headerRow))throw Error('原商品维度表表头行不存在。');
    // Inspect raw formula cells as SheetJS can omit a string formula with no cached value.
    for(const row of originalRows.values())for(const cell of tags(row,'c'))cachedValue(cell,name);
    const rowMap=new Map();for(let row=1;row<=headerRow;row++)rowMap.set(row,row);
    const codes=new Set();let last=headerRow,newSeen=false;
    result.rows.forEach((row,index)=>{
      if(typeof row.code!=='string'||!row.code||codes.has(row.code))throw Error('结果包含空白或重复物料编码，请重新整理。');codes.add(row.code);
      if(isNew(row))newSeen=true;
      else{if(newSeen||!Number.isSafeInteger(row.sourceRow)||row.sourceRow<=last||!originalRows.has(row.sourceRow))throw Error('原商品行顺序或来源无效，请重新整理。');rowMap.set(row.sourceRow,headerRow+index+1);last=row.sourceRow;}
    });
    const mapRange=rangeMapper(rowMap),moved=[...rowMap].some(([a,b])=>a!==b)||[...originalRows.keys()].some(row=>!rowMap.has(row)),end=headerRow+result.rows.length,lastCol=X.utils.encode_col(headers.length-1),range=`A${headerRow}:${lastCol}${end}`;
    const target=new RegExp("(?:'"+name.replace(/'/g,"''").replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+"'|"+name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+")!",'i');
    if(moved){
      for(const other of entries){if(other===entry)continue;const otherRel=relationships.find(n=>attr(n,'Id')===attr(other,'(?:[\\w.-]+:)?id'));if(!otherRel)continue;const otherXml=await read(zip,resolve('xl/workbook.xml',attr(otherRel,'Target')));if(tags(otherXml,'f').some(f=>target.test(decode(content(f)).replace(/"(?:[^"]|"")*"/g,''))))throw Error('其他工作表公式引用了商品维度主表，去重会改变引用位置，请先核对并转为值。');}
    }
    const dateColumns=new Set(['金蝶创建时间','当前日期'].map(header=>headers.indexOf(header))),formats=new Set();
    if(result.rows.some(isNew))formats.add('yyyy-mm-dd');
    const stylePath=resolve('xl/workbook.xml',attr(styleRel,'Target')),styles=addStyles(await read(zip,stylePath),formats),p=prefix(data);
    function valueCell(address,value,style,kind='s') {
      const attrs=`r="${address}" s="${style}"`;
      if(value==null||value==='')return `<${p}c ${attrs}/>`;
      if(kind==='n'){if(!Number.isFinite(value))throw Error('新增行包含无效数值。');return `<${p}c ${attrs}><${p}v>${value}</${p}v></${p}c>`;}
      if(kind==='b')return `<${p}c ${attrs} t="b"><${p}v>${value?1:0}</${p}v></${p}c>`;
      return `<${p}c ${attrs} t="inlineStr"><${p}is><${p}t xml:space="preserve">${cellText(value)}</${p}t></${p}is></${p}c>`;
    }
    function originalRow(row,number,code) {
      row=setAttr(setAttr(setAttr(row,'r',number),'hidden',null),'collapsed',null);
      row=row.replace(tagRE('c'),cell=>{
        const old=attr(cell,'r'),match=/^([A-Z]+)\d+$/.exec(old||'');if(!match)throw Error('主表单元格地址无效。');
        cell=setAttr(cachedValue(cell,name),'r',match[1]+number);
        if(code!==undefined&&match[1]===X.utils.encode_col(codeColumn))cell=body(setAttr(cell,'t','inlineStr'),`<${prefix(cell)}is><${prefix(cell)}t xml:space="preserve">${cellText(code)}</${prefix(cell)}t></${prefix(cell)}is>`);
        return cell;
      });return row;
    }
    const output=[];for(let row=1;row<=headerRow;row++)if(originalRows.has(row))output.push(originalRow(originalRows.get(row),row));
    result.rows.forEach((row,index)=>{
      const number=headerRow+index+1;
      if(!isNew(row)){output.push(originalRow(originalRows.get(row.sourceRow),number,row.code));return;}
      if(!Array.isArray(row.values)||row.values.length!==headers.length)throw Error('新增商品列结构与原表不一致。');
      output.push(`<${p}row r="${number}" ht="30" customHeight="1">`+headers.map((header,col)=>{
        const value=col===codeColumn?row.code:row.values[col],text=[codeColumn,headers.indexOf('SKU'),headers.indexOf('领星SKU')].includes(col),kind=text?'s':typeof value==='number'?'n':typeof value==='boolean'?'b':'s';
        const format=dateColumns.has(col)&&kind==='n'?'yyyy-mm-dd':kind==='s'?'@':'General';
        return valueCell(X.utils.encode_col(col)+number,value,styles.styles.get(format),kind);
      }).join('')+`</${p}row>`);
    });
    sheet=sheet.replace(data,body(data,output.join('')));
    const dimension=first(sheet,'dimension');if(!dimension)throw Error('商品维度表缺少dimension范围。');sheet=sheet.replace(dimension,setAttr(dimension,'ref',`A1:${lastCol}${end}`));
    sheet=sheet.replace(tagRE('sheetPr'),node=>setAttr(node,'filterMode',null));
    const filter=first(sheet,'autoFilter'),filterXml=`<${p}autoFilter ref="${range}"/>`;
    const sheetCalc=first(sheet,'sheetCalcPr');
    sheet=filter?sheet.replace(filter,filterXml):sheetCalc?sheet.replace(sheetCalc,sheetCalc+filterXml):sheet.replace(/<\/((?:[\w.-]+:)?sheetData)\s*>/,m=>m+filterXml);
    sheet=sheet.replace(tagRE('sortState'),'');
    sheet=sheet.replace(tagRE('conditionalFormatting'),node=>{if(moved&&first(node,'formula'))throw Error('主表包含公式条件格式，去重后需要人工核对公式范围。');const ref=mapRange(attr(node,'sqref'));return ref?setAttr(node,'sqref',ref):'';});
    sheet=sheet.replace(tagRE('ignoredError'),node=>{const ref=mapRange(attr(node,'sqref'));return ref?setAttr(node,'sqref',ref):'';});
    const usedLinkIds=new Set();
    sheet=sheet.replace(tagRE('hyperlink'),node=>{const ref=mapRange(attr(node,'ref'));if(!ref)return '';const location=attr(node,'location');if(moved&&location)throw Error('主表包含内部超链接，去重后需要人工核对链接位置。');const id=attr(node,'(?:[\\w.-]+:)?id');if(id)usedLinkIds.add(id);return setAttr(node,'ref',ref);});
    sheet=sheet.replace(tagRE('selection'),node=>setAttr(setAttr(node,'activeCell',`A${headerRow+1}`),'sqref',`A${headerRow+1}`));
    sheet=sheet.replace(tagRE('pane'),node=>{const cell=attr(node,'topLeftCell');return cell?setAttr(node,'topLeftCell',mapRange(cell)||(cell.match(/^[A-Z]+/)?.[0]||'A')+(headerRow+1)):node;});
    if(first(sheet,'rowBreaks'))sheet=sheet.replace(tagRE('rowBreaks'),node=>{let count=0;node=node.replace(tagRE('brk'),brk=>{const row=rowMap.get(Number(attr(brk,'id')));if(!row)return '';count++;return setAttr(brk,'id',row);});return count?setAttr(setAttr(node,'count',count),'manualBreakCount',count):'';});
    if(sheetRels){await remapNotes(zip,sheetRels,path,rowMap,mapRange);sheetRels=sheetRels.replace(tagRE('Relationship'),node=>attr(node,'Type').endsWith('/hyperlink')&&!usedLinkIds.has(attr(node,'Id'))?'':node);zip.file(relationPath(path),sheetRels);}
    const quoted="'"+name.replace(/'/g,"''")+"'!",filterRef=quoted+`$A$${headerRow}:$${lastCol}$${end}`;
    let foundFilter=false;
    workbook=workbook.replace(tagRE('definedName'),node=>{
      if(attr(node,'name')==='_xlnm._FilterDatabase'&&Number(attr(node,'localSheetId'))===sheetIndex){foundFilter=true;return body(node,escape(filterRef));}
      const value=decode(content(node));
      if(moved&&(value.startsWith(quoted)||value.startsWith(name+'!'))){const marker=value.startsWith(quoted)?quoted:name+'!',ref=value.slice(marker.length);let next;try{next=mapRange(ref);}catch(_){throw Error('命名范围 '+attr(node,'name')+'引用主表且无法安全重排，请先核对。');}return next?body(node,escape(marker+next)):'';}
      if(moved&&target.test(value.replace(/"(?:[^"]|"")*"/g,'')))throw Error('命名范围 '+attr(node,'name')+'包含引用主表的公式，请先核对。');
      return node;
    });
    if(!foundFilter){const node=`<${prefix(entry)}definedName name="_xlnm._FilterDatabase" localSheetId="${sheetIndex}" hidden="1">${escape(filterRef)}</${prefix(entry)}definedName>`,names=first(workbook,'definedNames'),insert=`<${prefix(entry)}definedNames>${node}</${prefix(entry)}definedNames>`;if(names)workbook=workbook.replace(names,body(names,content(names)+node));else{const after=first(workbook,'externalReferences')||first(workbook,'functionGroups')||first(workbook,'sheets');workbook=workbook.replace(after,after+insert);}}
    for(const chainRel of relationships.filter(n=>attr(n,'Type').endsWith('/calcChain'))){const chainPath=resolve('xl/workbook.xml',attr(chainRel,'Target'));let chain=await read(zip,chainPath),current=0,remaining=0;chain=chain.replace(tagRE('c'),node=>{if(attr(node,'i')!==null)current=Number(attr(node,'i'));if(current===Number(attr(entry,'sheetId')))return '';remaining++;return setAttr(node,'i',current);});if(remaining)zip.file(chainPath,chain);else{zip.remove(chainPath);workbookRels=workbookRels.replace(chainRel,'');types=types.replace(tagRE('Override'),node=>attr(node,'PartName')==='/'+chainPath?'':node);}}
    zip.file(path,sheet);zip.file(stylePath,styles.xml);zip.file('xl/workbook.xml',workbook);zip.file('xl/_rels/workbook.xml.rels',workbookRels);zip.file('[Content_Types].xml',types);
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  };
});
