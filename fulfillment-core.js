(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FulfillmentCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const SHEETS=['徐英','李奇','操思敏','邢芳芳','孙立柱'];
const COLUMNS=['采购下单人','事业部','销售分渠道','运营','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','未生产预计交付时间','是否正常履约','下单月份','OA备货单号','采购订单号'];
const DATE_FORMAT='yyyy"年"m"月"d"日"';
const norm=v=>String(v??'').replace(/\s/g,'');
function currentYear(){return Number(new Intl.DateTimeFormat('en-US',{timeZone:'Asia/Shanghai',year:'numeric'}).format(new Date()));}
function dateCell(cell,date1904,defaultYear=currentYear()){
 const v=cell?.v;if(v==null||v==='')return {t:'s',v:''};if(String(v).trim()==='/')return {t:'s',v:'/'};
 let y,m,d;
 if(typeof v==='number'){const p=X.SSF.parse_date_code(v,{date1904});if(p){y=p.y;m=p.m;d=p.d;}}
 else {const match=String(v).trim().match(/^(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);if(match)[,y,m,d]=match.map(Number);
 else {const monthly=norm(v).match(/^(?:(\d{4})年)?(\d{1,2})月陆续交货$/);if(monthly){y=monthly[1]?Number(monthly[1]):defaultYear;m=Number(monthly[2]);if(m>=1&&m<=12)d=new Date(Date.UTC(y,m,0)).getUTCDate();}}
 }
 const utc=Date.UTC(y,m-1,d),check=new Date(utc);
 if(!y||y<1900||y>9999||check.getUTCFullYear()!==y||check.getUTCMonth()!==m-1||check.getUTCDate()!==d)return {t:'s',v:String(v),invalid:true};
 let serial=(utc-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:DATE_FORMAT};
}
function merge(book,filename=''){
 const yearMatch=String(filename).match(/(?:^|[^0-9])((?:19|20)\d{2})(?=[年./-])/),defaultYear=yearMatch?Number(yearMatch[1]):currentYear();
 const selected=SHEETS.map(person=>{const matches=Object.keys(book.Sheets).filter(name=>new RegExp('^'+person+'(?:[0-9]{4}|[0-9]{8})?$').test(norm(name)));if(!matches.length)throw Error('缺少工作表：'+person);if(matches.length>1)throw Error(person+'：匹配到多个工作表（'+matches.join('、')+'），请仅保留本次需要合并的版本');return matches[0];});
 const rows=[],sources=[],counts=[],warnings=[];
 for(const name of selected){const sheet=book.Sheets[name],range=X.utils.decode_range(sheet['!ref']||'A1');let hr=-1,map;
 for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){const h={};for(let c=range.s.c;c<=range.e.c;c++){let key=norm(sheet[X.utils.encode_cell({r,c})]?.v);if(key==='OA备货流程号')key='OA备货单号';if(key==='为生产预计交付时间')key='未生产预计交付时间';if(key)(h[key]??=[]).push(c);}if(h['物料编码']&&h['事业部']){hr=r;map=h;break;}}
 if(hr<0)throw Error(name+'：找不到明细表头');
 for(const key of COLUMNS){if(!map[key]){if(COLUMNS.indexOf(key)>=20){warnings.push(name+'：缺少列“'+key+'”，已留空');continue;}throw Error(name+'：缺少列“'+key+'”');}if(map[key].length!==1)throw Error(name+'：重复列“'+key+'”');}
 let count=0;
 for(let r=hr+1;r<=range.e.r;r++){let nonempty=false;for(let c=range.s.c;c<=range.e.c;c++){const cell=sheet[X.utils.encode_cell({r,c})];if(cell?.v!=null&&String(cell.v).trim()!==''){nonempty=true;break;}}if(!nonempty)continue;
 const record=COLUMNS.map((key,i)=>{if(!map[key])return {t:'s',v:''};const address=X.utils.encode_cell({r,c:map[key][0]}),cell=sheet[address];let output;
 if(['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)){output=dateCell(cell,!!book.Workbook?.WBProps?.date1904,defaultYear);if(output.invalid)warnings.push(`${name}!${address}：日期“${cell.v}”无法识别，保留原文`);}
 else if(cell?.t==='e'||(cell?.f&&cell.v==null)){warnings.push(`${name}!${address}：公式结果不可用，请在 Excel 中重算后上传`);output={t:'s',v:cell.w||'#公式结果缺失'};}
 else if(key==='下单月份'&&typeof cell?.v==='number'){output=dateCell(cell,!!book.Workbook?.WBProps?.date1904,defaultYear);if(output.z)output.z='yyyy"年"m"月"';}
 else if(['物料编码','SKU','OA备货单号','采购订单号'].includes(key))output={t:'s',v:cell?String(cell.w??cell.v??''):''};
 else output={t:typeof cell?.v==='number'?'n':'s',v:cell?.v??''};
 delete output.invalid;return output;});rows.push(record);sources.push({sheet:name,row:r+1});count++;}
 counts.push({sheet:name,count});}
 return {rows,sources,counts,warnings};
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

function columnWidth(key){return key==='物料名称'?42:['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(key)?23:key==='SKU'?26:20;}
function exportFilename(input){const base=String(input??'').trim().replace(/\.xlsx?$/i,'').trim()||'订单履约表';return (/更新版$/.test(base)?base.replace(/更新版$/,'汇总版'):/汇总版$/.test(base)?base:base+'汇总版')+'.xlsx';}
function workbook(result){const s=X.utils.aoa_to_sheet([COLUMNS]);result.rows.forEach((row,r)=>row.forEach((cell,c)=>s[X.utils.encode_cell({r:r+1,c})]={...cell}));s['!ref']=`A1:${X.utils.encode_col(COLUMNS.length-1)}${result.rows.length+1}`;s['!autofilter']={ref:s['!ref']};s['!cols']=COLUMNS.map((key,col)=>({wch:[key,...result.rows.map(row=>row[col]?.z?X.SSF.format(row[col].z,row[col].v):String(row[col]?.v??''))].reduce((max,text)=>Math.max(max,[...text].reduce((width,ch)=>width+(ch.charCodeAt(0)>255?2:1),0)+4),columnWidth(key))}));const w=X.utils.book_new();X.utils.book_append_sheet(w,s,'汇总表');if(result.differences){const d=X.utils.aoa_to_sheet([DIFF_COLUMNS]);result.differences.forEach((row,r)=>row.forEach((cell,c)=>d[X.utils.encode_cell({r:r+1,c})]={...cell}));d['!ref']=`A1:F${result.differences.length+1}`;d['!autofilter']={ref:d['!ref']};d['!cols']=DIFF_COLUMNS.map((key,col)=>({wch:Math.max(20,...[key,...result.differences.map(r=>String(r[col].v))].map(v=>[...v].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),4)))}));X.utils.book_append_sheet(w,d,'差异');}return w;}
return {DIFF_COLUMNS,aggregate,difference,advance,SHEETS,COLUMNS,DATE_FORMAT,columnWidth,dateCell,merge,workbook,exportFilename};
});
