(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'));
  else root.FundsBudgetSummaryExport = factory(root.XLSX, root.JSZip);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip) {
  'use strict';
  const NAME = '付款预算汇总';
  const HEADERS = ['付款月份','分类','状态','申请总金额','已付款金额','剩余付款金额','备注'];
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MAIN = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  const escape = value => String(value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  const cellText = value => escape(String(value).replace(/_x[0-9a-f]{4}_/gi,match=>'_x005F_'+match.slice(1)).replace(/[\u0000-\u0008\u000b-\u001f]/g,char=>'_x'+char.charCodeAt(0).toString(16).padStart(4,'0')+'_'));
  const decode = value => value.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-f]+|#\d+);/gi, (all, key) => key[0] === '#' ? String.fromCodePoint(parseInt(key.slice(key[1].toLowerCase() === 'x' ? 2 : 1), key[1].toLowerCase() === 'x' ? 16 : 10)) : ({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[key] || all));
  const tagRE = (name, flags = 'g') => new RegExp(`<((?:[\\w.-]+:)?${name})\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/\\1\\s*>)`, flags);
  const tags = (xml, name) => xml.match(tagRE(name)) || [];
  const first = (xml, name) => tags(xml, name)[0];
  const opening = xml => xml.match(/^<[^>]+>/)[0];
  const attr = (xml, name) => {const hit = opening(xml).match(new RegExp('(?:^|\\s)' + name + '=(["\'])([\\s\\S]*?)\\1')); return hit ? decode(hit[2]) : null;};
  function setAttr(xml, name, value) {
    const head = opening(xml), re = new RegExp('\\s+' + name + '=(["\'])[\\s\\S]*?\\1');
    const next = value === null ? head.replace(re,'') : re.test(head) ? head.replace(re, ' ' + name + '="' + escape(value) + '"') : head.replace(/\/?>(?=$)/, end => ' ' + name + '="' + escape(value) + '"' + end);
    return next + xml.slice(head.length);
  }
  const content = xml => xml.replace(/^<[^>]+>/,'').replace(/<\/[^>]+>$/,'');
  const prefix = xml => xml.match(/^<([\w.-]+:)?/)?.[1] || '';
  const qualified = (xml, p) => xml.replace(/<(\/?)([A-Za-z][\w]*)(?=[\s/>])/g, '<$1' + p + '$2');
  const relationPath = path => path.replace(/([^/]+)$/, '_rels/$1.rels');
  function resolve(base, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : base.slice(0,base.lastIndexOf('/') + 1) + target).split('/'), result = [];
    for (const part of parts) {if (part === '..') result.pop(); else if (part && part !== '.') result.push(part);}
    return result.join('/');
  }
  async function read(zip, path) {
    if (!zip.file(path)) throw Error('工作簿缺少 ' + path + '，无法保真导出。');
    return zip.file(path).async('string');
  }
  function append(xml, name, items) {
    const old = first(xml,name);
    if (!old || /\/>$/.test(old)) throw Error('工作簿样式结构不完整：' + name);
    const p = prefix(old), next = setAttr(old.replace(new RegExp('</' + p + name + '\\s*>$'), qualified(items,p) + '</' + p + name + '>'), 'count', tags(content(old), {fonts:'font',fills:'fill',borders:'border',cellXfs:'xf',numFmts:'numFmt'}[name]).length + tags(items, {fonts:'font',fills:'fill',borders:'border',cellXfs:'xf',numFmts:'numFmt'}[name]).length);
    return xml.replace(old,next);
  }
  function addStyles(xml) {
    const count = name => {const node=first(xml,name); if(!node)throw Error('工作簿缺少样式：'+name); return tags(content(node),{fonts:'font',fills:'fill',borders:'border',cellXfs:'xf'}[name]).length;};
    const font=count('fonts'), fill=count('fills'), border=count('borders'), style=count('cellXfs');
    const f = color => `<font><sz val="11"/><color rgb="${color}"/><name val="等线"/></font>`;
    xml=append(xml,'fonts',f('FF172033')+f('FFFFFFFF'));
    xml=append(xml,'fills',['FFFFFFFF','FF315B9D'].map(color=>`<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`).join(''));
    xml=append(xml,'borders',['thin','hair'].map(style=>'<border>'+['left','right','top','bottom'].map(side=>`<${side} style="${style}"><color rgb="FF404040"/></${side}>`).join('')+'<diagonal/></border>').join('')+'<border><left/><right/><top/><bottom/><diagonal/></border>');
    const xf=(format,header,blank=false)=>`<xf numFmtId="${format}" fontId="${font+(header?1:0)}" fillId="${fill+(header?1:0)}" borderId="${border+(header?0:blank?2:1)}" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="0" shrinkToFit="0"/></xf>`;
    xml=append(xml,'cellXfs',xf(0,true)+[false,true].map(blank=>[0,49,4].map(format=>xf(format,false,blank)).join('')).join(''));
    return {xml,header:style,text:style+1,month:style+2,money:style+3};
  }
  function validate(result) {
    if(!result || result.canExport !== true || result.errors?.length)throw Error('汇总存在待处理错误，不能导出。');
    if(!Array.isArray(result.groups)||!result.totals)throw Error('缺少付款预算汇总结果。');
    const totals={totalCents:0,paidCents:0,remainingCents:0};
    for(const group of [...result.groups,result.totals]){
      for(const key of Object.keys(totals)){
        if(!Number.isSafeInteger(group[key]))throw Error('汇总金额无效，请重新整理。');
        if(Math.round(group[key]/100*100)!==group[key])throw Error('汇总金额转换为 Excel 元数值后会丢失分位精度，不能导出。');
      }
      if(group.totalCents-group.paidCents!==group.remainingCents)throw Error('汇总剩余金额与申请金额、已付款金额不一致。');
    }
    for(const group of result.groups)for(const key of Object.keys(totals)){totals[key]+=group[key];if(!Number.isSafeInteger(totals[key]))throw Error('汇总金额超出安全范围。');}
    if(Object.keys(totals).some(key=>totals[key]!==result.totals[key]))throw Error('汇总合计与分组金额不一致。');
    if(result.groups.length>1048574)throw Error('汇总行数超过 Excel 上限。');
  }
  function monthCell(key) {
    if(!key)return {value:'时间待定'};
    const match=/^(\d{4})-(0[1-9]|1[0-2])$/.exec(key);
    if(!match||Number(match[1])<1900)throw Error('汇总付款月份无效：'+key);
    return {value:Number(match[1])+'年'+Number(match[2])+'月',month:true};
  }
  function sheetXml(result,styles,namespace) {
    const rows=[HEADERS.map(value=>({value})),...result.groups.map((g,index)=>[monthCell(g.monthKey),{value:g.category??''},{value:g.state??''},{value:g.totalCents/100,money:true},{value:g.paidCents/100,money:true},{value:g.remainingCents/100,money:true,formula:`D${index+2}-E${index+2}`},{value:''}])];
    const end=rows.length,total=result.totals;
    rows.push([{value:'合计'},{value:''},{value:''},...['totalCents','paidCents','remainingCents'].map((key,index)=>({value:total[key]/100,money:true,formula:end>1?`SUM(${String.fromCharCode(68+index)}2:${String.fromCharCode(68+index)}${end})`:'SUM(0)'})),{value:''}]);
    const contentRows=new Set(),contentColumns=new Set();
    rows.forEach((row,r)=>row.forEach((cell,c)=>{
      if(cell.value==null||(typeof cell.value==='string'&&!cell.value.trim()))return;
      contentRows.add(r);contentColumns.add(c);
    }));
    const widths=HEADERS.map((_,col)=>Math.min(255,rows.reduce((size,row)=>Math.max(size,[...String(row[col].display??(row[col].money?X.SSF.format('#,##0.00',row[col].value):row[col].value))].reduce((n,c)=>n+(c.codePointAt(0)>255?2:1),0)),12)+2));
    const data=rows.map((row,r)=>`<row r="${r+1}" ht="26" customHeight="1">`+row.map((cell,c)=>{
      const blank=!contentRows.has(r)||!contentColumns.has(c);
      const ref=String.fromCharCode(65+c)+(r+1),style=r===0?styles.header:(cell.month?styles.month:cell.money?styles.money:styles.text)+(blank?3:0);
      if(typeof cell.value==='number')return `<c r="${ref}" s="${style}">${cell.formula?'<f>'+escape(cell.formula)+'</f>':''}<v>${cell.value}</v></c>`;
      return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${cellText(cell.value)}</t></is></c>`;
    }).join('')+'</row>').join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${namespace}"><dimension ref="A1:G${rows.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="26"/><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${data}</sheetData><autoFilter ref="A1:G${end}"/></worksheet>`;
  }
  function check3D(formula,before,after) {
    const index=(names,name)=>names.findIndex(n=>n.toLowerCase()===name.toLowerCase());
    for(const part of formula.split(/("(?:[^"]|"")*")/g).filter((_,i)=>i%2===0))for(const match of part.matchAll(/(?:'((?:[^']|'')+)'|((?:\[[^\]]+\])?[\p{L}\p{N}_.]+))(?::(?:'((?:[^']|'')+)'|([\p{L}\p{N}_.]+)))?!/gu)){
      const token=(match[1]||match[2]).replace(/''/g,"'"),second=match[3]||match[4];
      if(token.includes('['))continue;
      const pair=second?[token,second.replace(/''/g,"'")]:token.split(':');if(pair.length!==2)continue;
      const [a,b]=pair.map(name=>index(before,name)),[c,d]=pair.map(name=>index(after,name));if(a<0||b<0)continue;
      if(JSON.stringify(before.slice(Math.min(a,b),Math.max(a,b)+1))!==JSON.stringify(after.slice(Math.min(c,d),Math.max(c,d)+1)))throw Error('原表三维公式引用 '+match[0]+' 会受工作表顺序变化影响，请先改为逐表引用后再导出。');
    }
  }
  async function descendants(zip,path,visited=new Set()) {
    if(visited.has(path))return visited;visited.add(path);
    const rel=relationPath(path);if(!zip.file(rel))return visited;visited.add(rel);
    for(const node of tags(await read(zip,rel),'Relationship'))if(attr(node,'TargetMode')!=='External')await descendants(zip,resolve(path,attr(node,'Target')),visited);
    return visited;
  }
  return async function(sourceBytes,result) {
    validate(result);
    let zip;try{zip=await Zip.loadAsync(sourceBytes);}catch(_){throw Error('请上传原始 .xlsx 工作簿，以保留原表公式和格式。');}
    let workbook=await read(zip,'xl/workbook.xml'),rels=await read(zip,'xl/_rels/workbook.xml.rels'),types=await read(zip,'[Content_Types].xml');
    if(!tags(types,'Override').some(node=>attr(node,'PartName')==='/xl/workbook.xml'&&attr(node,'ContentType')===MAIN))throw Error('仅支持 .xlsx 工作簿，请先另存为 .xlsx 后导出。');
    const sheets=first(workbook,'sheets'),entries=tags(sheets||'','sheet'),names=entries.map(node=>attr(node,'name'));
    if(!entries.length||new Set(names).size!==names.length)throw Error('工作簿工作表清单无效。');
    const oldIndex=names.indexOf(NAME),order=[NAME,...names.filter(name=>name!==NAME)];
    const relationships=tags(rels,'Relationship'),sheetRel=node=>relationships.find(rel=>attr(rel,'Id')===attr(node,'(?:[\\w.-]+:)?id'));
    let entry=oldIndex>=0?entries[oldIndex]:null,path;
    const removed=new Set();
    if(entry){const rel=sheetRel(entry);if(!rel||!attr(rel,'Type').endsWith('/worksheet'))throw Error('付款预算汇总不是可更新的工作表。');path=resolve('xl/workbook.xml',attr(rel,'Target'));for(const p of await descendants(zip,path))if(p!==path)removed.add(p);entry=setAttr(entry,'state',null);}
    else{
      path='xl/worksheets/funds-budget-summary.xml';while(zip.file(path))path=path.replace('.xml','x.xml');
      let id='rIdFundsBudgetSummary';while(relationships.some(node=>attr(node,'Id')===id))id+='x';
      const p=prefix(entries[0]),relationshipIdName=opening(entries[0]).match(/\s((?:[\w.-]+:)?id)=/)?.[1]||'r:id';
      entry=`<${p}sheet name="${NAME}" sheetId="${Math.max(...entries.map(node=>Number(attr(node,'sheetId'))))+1}" ${relationshipIdName}="${id}"/>`;
      const relationType=attr(relationships.find(node=>attr(node,'Type').endsWith('/worksheet'))||'<Relationship Type="'+R+'/worksheet"/>','Type');
      rels=rels.replace(/<\/([\w.-]+:)?Relationships\s*>/,match=>`<${prefix(first(rels,'Relationship')||'<Relationship>')}Relationship Id="${id}" Type="${relationType}" Target="${path.slice(3)}"/>`+match);
      types=types.replace(/<\/([\w.-]+:)?Types\s*>/,match=>`<${prefix(first(types,'Override')||'<Override>')}Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`+match);
    }
    for(const file of Object.keys(zip.files)){
      if(zip.files[file].dir||file===path||!file.endsWith('.xml'))continue;
      const xml=await read(zip,file);
      for(const name of ['f','formula','formula1','formula2','calculatedColumnFormula','totalsRowFormula','definedName'])for(const node of tags(xml,name))check3D(decode(content(node)),names,order);
    }
    workbook=workbook.replace(sheets,opening(sheets)+entry+entries.filter((_,i)=>i!==oldIndex).join('')+`</${prefix(sheets)}sheets>`);
    workbook=workbook.replace(tagRE('definedName'),node=>{
      const local=attr(node,'localSheetId');if(local===null)return node;
      const name=names[Number(local)];if(!name)throw Error('工作簿局部名称引用的工作表不存在。');
      if(name===NAME&&/^_xlnm\.(?:_FilterDatabase|Print_Area|Print_Titles)$/.test(attr(node,'name')||''))return '';
      return setAttr(node,'localSheetId',order.indexOf(name));
    });
    workbook=workbook.replace(tagRE('workbookView'),node=>setAttr(setAttr(node,'activeTab',0),'firstSheet',0));
    const calc=first(workbook,'calcPr');if(calc)workbook=workbook.replace(calc,setAttr(setAttr(calc,'fullCalcOnLoad',1),'forceFullCalc',1));
    const styleRel=relationships.find(node=>attr(node,'Type').endsWith('/styles'));if(!styleRel)throw Error('工作簿样式关系缺失。');
    const stylePath=resolve('xl/workbook.xml',attr(styleRel,'Target')),styles=addStyles(await read(zip,stylePath));
    const properties=first(workbook,'workbookPr'),date1904=properties&&/^(1|true)$/i.test(attr(properties,'date1904')||'');
    if(result.date1904!==undefined&&Boolean(result.date1904)!==Boolean(date1904))throw Error('汇总结果与原工作簿日期系统不一致，请重新分析原文件。');
    const namespace=workbook.match(/xmlns(?::[\w.-]+)?=["']([^"']*spreadsheetml[^"']*)["']/)?.[1]||S;
    zip.file(path,sheetXml(result,styles,namespace));zip.remove(relationPath(path));
    for(const rel of relationships)if(attr(rel,'Type').endsWith('/calcChain')){rels=rels.replace(rel,'');removed.add(resolve('xl/workbook.xml',attr(rel,'Target')));}
    zip.file('xl/workbook.xml',workbook);zip.file('xl/_rels/workbook.xml.rels',rels);zip.file(stylePath,styles.xml);
    // Remove only parts detached from the replaced summary; retain unrelated native parts.
    const reachable=new Set();
    for(const rel of tags(await read(zip,'_rels/.rels'),'Relationship'))if(attr(rel,'TargetMode')!=='External')await descendants(zip,resolve('',attr(rel,'Target')),reachable);
    for(const p of removed)if(!reachable.has(p))zip.remove(p);
    types=types.replace(tagRE('Override'),node=>removed.has((attr(node,'PartName')||'').replace(/^\//,''))&&!zip.file((attr(node,'PartName')||'').replace(/^\//,''))?'':node);zip.file('[Content_Types].xml',types);
    if(zip.file('docProps/app.xml')){
      let app=await read(zip,'docProps/app.xml'),titles=first(app,'TitlesOfParts');
      if(titles){const vector=first(titles,'vector');if(vector&&tags(content(vector),'lpstr').length===names.length){const p=prefix(first(vector,'lpstr')||'<vt:lpstr>');app=app.replace(vector,setAttr(opening(vector),'size',order.length)+order.map(name=>`<${p}lpstr>${escape(name)}</${p}lpstr>`).join('')+`</${prefix(vector)}vector>`);}}
      const pairs=first(app,'HeadingPairs');if(pairs){const variants=tags(pairs,'variant');for(let i=0;i<variants.length-1;i++)if(decode(content(first(variants[i],'lpstr')||'<lpstr/>'))==='Worksheets'){const old=first(variants[i+1],'i4');if(old)app=app.replace(old,opening(old)+order.length+`</${prefix(old)}i4>`);}}
      zip.file('docProps/app.xml',app);
    }
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  };
});
