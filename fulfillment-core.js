(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FulfillmentCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const SHEETS=['徐英','操思敏','邢芳芳','李奇'];
const COLUMNS=['采购下单人','事业部','销售分渠道','运营','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','未生产预计交付时间','是否正常履约','下单月份','OA备货单号','采购订单号','下单日期'];
const QUANTITY_FIELDS=COLUMNS.slice(10,16);
const SUMMARY_COLUMNS=['事业部','销售分渠道','运营','下单月份','OA流程号','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','为生产预计交付时间','是否正常履约'];
const DATE_FORMAT='yyyy"年"m"月"d"日"';
const MONTH_FORMAT='yyyy"年"m"月"';
const quantityText=value=>value.toLocaleString('en-US',{useGrouping:false,maximumSignificantDigits:21});
function formatCell(cell){
 // The bundled formatter misrenders tiny decimals such as 1e-7 with fixed decimal formats.
 if(cell?.t==='n'&&/^0(?:\.0+)?$/.test(cell.z||''))return quantityText(cell.v);
 return cell?.z?X.SSF.format(cell.z,cell.v):String(cell?.v??'');
}
function summaryRows(rows){return rows.map(row=>SUMMARY_COLUMNS.map(key=>{
 const sourceKey=key==='OA流程号'?'OA备货单号':key==='为生产预计交付时间'?'未生产预计交付时间':key;
 const cell=defaultCell(sourceKey,row[COLUMNS.indexOf(sourceKey)]);
 if(sourceKey==='下单月份'){const date=dateCell(cell);return date.invalid?cell:date.t==='n'?{...date,z:MONTH_FORMAT}:date;}
 if(QUANTITY_FIELDS.includes(sourceKey)&&cell.t==='n'&&Number.isFinite(cell.v)){
  const precision=(quantityText(cell.v).split('.')[1]||'').length;
  return {t:'n',v:cell.v,z:precision>250?'0.0000000000000000E+00':precision?'0.'+'0'.repeat(precision):'0'};
 }
 return cell;
}));}
const norm=v=>String(v??'').replace(/\s/g,'');
function currentYear(){return Number(new Intl.DateTimeFormat('en-US',{timeZone:'Asia/Shanghai',year:'numeric'}).format(new Date()));}
function quantityCell(sheet,r,c){
 const source=sheet[X.utils.encode_cell({r,c})],cell={t:source?.t||'s',v:source?.v??''},merged=(sheet['!merges']||[]).find(m=>r>=m.s.r&&r<=m.e.r&&c>=m.s.c&&c<=m.e.c);
 if(merged&&merged.s.c!==merged.e.c)return {...cell,invalid:'数量单元格跨列合并，无法确定数量所属字段'};
 if(merged&&r>merged.s.r)return {t:'n',v:0};
 if(source?.f||cell.t!=='s'||String(cell.v).trim()==='')return cell;
 const matches=String(cell.v).match(/[+-]?(?:\d{1,3}(?:[,，]\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/g)||[];
 if(matches.length!==1)return {...cell,invalid:matches.length?'数量文本包含多个数字，无法确定取值':'数量文本中没有可提取的数字'};
 const value=Number(matches[0].replace(/[,，]/g,''));
 if(!Number.isFinite(value)||Math.abs(value)>Number.MAX_SAFE_INTEGER)return {...cell,invalid:'数量超出可精确核验范围'};
 return {t:'n',v:value};
}
function refreshQuantities(result,book){
 const maps=new Map();
 for(const source of result.sources||[]){if(maps.has(source.sheet))continue;const sheet=book.Sheets[source.sheet];if(!sheet)continue;const range=X.utils.decode_range(sheet['!ref']||'A1');
  for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){const map={};for(let c=range.s.c;c<=range.e.c;c++){const key=norm(sheet[X.utils.encode_cell({r,c})]?.v);if(key)(map[key]??=[]).push(c);}if(map['事业部']&&map['物料编码']){maps.set(source.sheet,map);break;}}
 }
 return {...result,rows:result.rows.map((row,i)=>{const source=result.sources?.[i],sheet=book.Sheets[source?.sheet],map=maps.get(source?.sheet);if(!sheet||!map)return row;
  return row.map((cell,col)=>{const field=COLUMNS[col];if(!QUANTITY_FIELDS.includes(field)||map[field]?.length!==1)return cell;const r=source.row-1,c=map[field][0],original=sheet[X.utils.encode_cell({r,c})];
   // Upgrade only the unchanged source value in an old cache; altered quantities must still fail validation.
   if(original?.f||cell?.v!==(original?.v??'')||cell?.t!==(typeof original?.v==='number'?'n':'s'))return cell;const normalized=quantityCell(sheet,r,c);return normalized.invalid?cell:normalized;
  });
 })};
}
function defaultCell(key,cell={t:'s',v:''}){
 if(key==='OA备货单号'&&['','/'].includes(norm(cell.v)))return {t:'s',v:'无'};
 if(key==='下单月份'&&norm(cell.v)==='/')return {...dateCell({v:'2026-04-01'}),z:'yyyy"年"m"月"'};
 if(['下单日期','下单月份','合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)&&(norm(cell.v)==='历史订单'||/^\d{4}[年/.-]\d{1,2}月?$/.test(norm(cell.v)))){const date=dateCell(cell);if(!date.invalid)return date;}
 return cell;
}
function dateCell(cell,date1904,defaultYear=currentYear()){
 const v=cell?.v;if(v==null||v==='')return {t:'s',v:''};if(String(v).trim()==='/')return {t:'s',v:'/'};
 if(norm(v)==='历史订单')return {t:'n',v:(Date.UTC(2025,0,1)-Date.UTC(1899,11,30))/86400000,z:'yyyy"年"m"月"'};
 let y,m,d,monthOnly=false;
 if(typeof v==='number'){const p=X.SSF.parse_date_code(v,{date1904});if(p){y=p.y;m=p.m;d=p.d;}}
 else {const match=String(v).trim().match(/^(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);if(match)[,y,m,d]=match.map(Number);
 else {const month=norm(v).match(/^(\d{4})[年/.-](\d{1,2})月?$/);if(month){y=Number(month[1]);m=Number(month[2]);d=1;monthOnly=true;}
 else {const monthly=norm(v).match(/^(?:(\d{4})年)?(\d{1,2})月陆续交货$/);if(monthly){y=monthly[1]?Number(monthly[1]):defaultYear;m=Number(monthly[2]);if(m>=1&&m<=12)d=new Date(Date.UTC(y,m,0)).getUTCDate();}}}
 }
 const utc=Date.UTC(y,m-1,d),check=new Date(utc);
 if(!y||y<1900||y>9999||check.getUTCFullYear()!==y||check.getUTCMonth()!==m-1||check.getUTCDate()!==d)return {t:'s',v:String(v),invalid:true};
 let serial=(utc-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:monthOnly?'yyyy"年"m"月"':DATE_FORMAT};
}
function merge(book,filename=''){
 const yearMatch=String(filename).match(/(?:^|[^0-9])((?:19|20)\d{2})(?=[年./-])/),defaultYear=yearMatch?Number(yearMatch[1]):currentYear();
 const selected=SHEETS.map(person=>{const matches=Object.keys(book.Sheets).filter(name=>new RegExp('^'+person+'(?:[0-9]{4}|[0-9]{8})?$').test(norm(name)));if(!matches.length)throw Error('缺少工作表：'+person);if(matches.length>1)throw Error(person+'：匹配到多个工作表（'+matches.join('、')+'），请仅保留本次需要合并的版本');return matches[0];});
 const rows=[],sources=[],counts=[],warnings=[];
 for(const name of selected){const sheet=book.Sheets[name],range=X.utils.decode_range(sheet['!ref']||'A1');let hr=-1,map;
 for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){const h={};for(let c=range.s.c;c<=range.e.c;c++){let key=norm(sheet[X.utils.encode_cell({r,c})]?.v);if(['OA备货流程号','OA流程号'].includes(key))key='OA备货单号';if(key==='为生产预计交付时间')key='未生产预计交付时间';if(key)(h[key]??=[]).push(c);}if(h['物料编码']&&h['事业部']){hr=r;map=h;break;}}
 if(hr<0)throw Error(name+'：找不到明细表头');
 for(const key of COLUMNS){if(!map[key]){if(COLUMNS.indexOf(key)>=20){if(key!=='下单日期')warnings.push(name+'：缺少列“'+key+'”，'+(key==='OA备货单号'?'已填“无”':'已留空'));continue;}throw Error(name+'：缺少列“'+key+'”');}if(map[key].length!==1)throw Error(name+'：重复列“'+key+'”');}
 let count=0;
 for(let r=hr+1;r<=range.e.r;r++){let nonempty=false;for(let c=range.s.c;c<=range.e.c;c++){const cell=sheet[X.utils.encode_cell({r,c})];if(cell?.v!=null&&String(cell.v).trim()!==''){nonempty=true;break;}}if(!nonempty)continue;
 const record=COLUMNS.map((key,i)=>{if(!map[key])return defaultCell(key);const address=X.utils.encode_cell({r,c:map[key][0]}),cell=sheet[address];let output;
 if(QUANTITY_FIELDS.includes(key))output=quantityCell(sheet,r,map[key][0]);
 else if(['下单日期','合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)){output=dateCell(cell,!!book.Workbook?.WBProps?.date1904,defaultYear);if(output.invalid)warnings.push(`${name}!${address}：日期“${cell.v}”无法识别，保留原文`);}
 else if(cell?.t==='e'||(cell?.f&&cell.v==null)){warnings.push(`${name}!${address}：公式结果不可用，请在 Excel 中重算后上传`);output={t:'s',v:cell.w||'#公式结果缺失'};}
 else if(key==='下单月份'){output=dateCell(cell,!!book.Workbook?.WBProps?.date1904,defaultYear);if(output.z)output.z=MONTH_FORMAT;if(output.invalid)warnings.push(`${name}!${address}：日期“${cell?.v}”无法识别，保留原文`);}
 else if(['物料编码','SKU','OA备货单号','采购订单号'].includes(key))output={t:'s',v:cell?String(cell.w??cell.v??''):''};
 else output={t:typeof cell?.v==='number'?'n':'s',v:cell?.v??''};
 delete output.invalid;return defaultCell(key,output);});rows.push(record);sources.push({sheet:name,row:r+1});count++;}
 counts.push({sheet:name,count});}
 return {rows,sources,counts,warnings};
}

function currentScope(result){
 if(!result)return result;
 const excluded=name=>/^孙立柱(?:[0-9]{4}|[0-9]{8})?$/.test(norm(name));
 const keep=result.rows.map((_,i)=>!excluded(result.sources?.[i]?.sheet));
 const rows=result.rows.filter((_,i)=>keep[i]).map(row=>COLUMNS.map((key,col)=>defaultCell(key,row[col])));
 return {...result,rows,sources:(result.sources||[]).filter((_,i)=>keep[i]),counts:(result.counts||[]).filter(x=>!excluded(x.sheet)),warnings:(result.warnings||[]).filter(w=>{if(excluded(w.split('!')[0]))return false;const value=w.match(/：日期“(.+)”无法识别，保留原文$/s)?.[1];return value==null||dateCell({v:value}).invalid;})};
}
const DIFF_COLUMNS=['采购下单人','供应商','物料编码','SKU','物料名称','未交付数量'];
function aggregate(result){
 const groups=new Map();
 for(const [i,row] of result.rows.entries()){
  const code=String(row[7].v??'').trim(),raw=String(row[10].v??'').trim();
  const at=result.sources?.[i],where=at?`${at.sheet}第 ${at.row} 行`:`第 ${i+1} 条`;
  const value=raw===''?'0':raw;
  if(!/^[+-]?\d+(?:\.\d+)?$/.test(value))throw Error(`${where}：未交付数量“${raw}”不是有效数字，未更新历史数据`);
  const parts=value.replace(/^[+-]/,'').split('.'),scale=(parts[1]||'').length;
  const units=BigInt(parts.join(''))*(value.startsWith('-')?-1n:1n);
  if(!code){if(units!==0n)throw Error(`${where}：物料编码为空，无法对比，未更新历史数据`);continue;}
  if(!groups.has(code))groups.set(code,{units:0n,scale:0,values:[new Set(),new Set(),new Set(),new Set(),new Set()]});
  const g=groups.get(code),nextScale=Math.max(g.scale,scale);
  g.units=g.units*10n**BigInt(nextScale-g.scale)+units*10n**BigInt(nextScale-scale);g.scale=nextScale;
  [0,4,7,8,9].forEach((col,j)=>{const v=String(row[col].v??'').trim();if(v)g.values[j].add(v);});
 }
 return groups;
}
function difference(previous,current){
 const before=previous?aggregate(previous):new Map(),after=aggregate(current),rows=[];
 if(!previous)return rows;
 for(const code of new Set([...after.keys(),...before.keys()])){
  const a=before.get(code),b=after.get(code),scale=Math.max(a?.scale||0,b?.scale||0);
  const delta=(b?b.units*10n**BigInt(scale-b.scale):0n)-(a?a.units*10n**BigInt(scale-a.scale):0n);
  if(delta===0n)continue;
  const number=Number(delta)/10**scale;
  if(!Number.isFinite(number)||Math.abs(Number(delta))>Number.MAX_SAFE_INTEGER)throw Error(`物料 ${code}：数量超出精确计算范围，未更新历史数据`);
  const values=(b||a).values.map(v=>({t:'s',v:[...v].join(' / ')}));
  rows.push([...values,{t:'n',v:number}]);
 }
 return rows;
}
function advance(history,current){
 aggregate(current.result);
 if(history&&JSON.stringify(history.current.result.rows)===JSON.stringify(current.result.rows)){
  const old=history.current.sourceBytes,next=current.sourceBytes,oldView=old?new Uint8Array(old):null;
  if(next&&(!old||old.byteLength!==next.byteLength||new Uint8Array(next).some((v,i)=>v!==oldView[i])))return {...history,current};
  return history;
 }
 const previous=history?.current||null;
 difference(previous?.result,current.result);
 return {version:1,previous,current};
}

function columnWidth(key){return key==='物料名称'?42:['下单日期','合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)?23:key==='SKU'?26:20;}
function exportFilename(input){const base=String(input??'').trim().replace(/\.xlsx?$/i,'').trim()||'订单履约表';return (/更新版$/.test(base)?base.replace(/更新版$/,'汇总版'):/汇总版$/.test(base)?base:base+'汇总版')+'.xlsx';}
function workbook(result){const rows=summaryRows(result.rows);const s=X.utils.aoa_to_sheet([SUMMARY_COLUMNS]);rows.forEach((row,r)=>row.forEach((cell,c)=>s[X.utils.encode_cell({r:r+1,c})]={...cell}));s['!ref']=`A1:${X.utils.encode_col(SUMMARY_COLUMNS.length-1)}${rows.length+1}`;s['!autofilter']={ref:s['!ref']};s['!cols']=SUMMARY_COLUMNS.map((key,col)=>({wch:[key,...rows.map(row=>formatCell(row[col]))].reduce((max,text)=>Math.max(max,[...text].reduce((width,ch)=>width+(ch.charCodeAt(0)>255?2:1),0)+4),columnWidth(key))}));const w=X.utils.book_new();X.utils.book_append_sheet(w,s,'汇总表');if(result.differences){const d=X.utils.aoa_to_sheet([DIFF_COLUMNS]);result.differences.forEach((row,r)=>row.forEach((cell,c)=>d[X.utils.encode_cell({r:r+1,c})]={...cell}));d['!ref']=`A1:F${result.differences.length+1}`;d['!autofilter']={ref:d['!ref']};d['!cols']=DIFF_COLUMNS.map((key,col)=>({wch:Math.max(20,...[key,...result.differences.map(r=>String(r[col].v))].map(v=>[...v].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),4)))}));X.utils.book_append_sheet(w,d,'差异');}return w;}
return {SUMMARY_COLUMNS,summaryRows,formatCell,currentScope,DIFF_COLUMNS,aggregate,difference,advance,SHEETS,COLUMNS,QUANTITY_FIELDS,quantityCell,refreshQuantities,DATE_FORMAT,columnWidth,dateCell,merge,workbook,exportFilename};
});
