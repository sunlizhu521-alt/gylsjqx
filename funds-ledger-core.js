(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FundsLedgerCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const SHEET_NAMES=['已申请OA付款','未申请OA付款'];
const HEADERS=SHEET_NAMES.map((_,i)=>['供应商','付款主体','对公付款OA流程号','申请付款金额','摘要',i?'预估付款时间':'应付款时间','是否提交付款审单']);
const FIELDS={
 supplier:['供应商名称','供应商'],entity:['付款主体'],oa:['对公付款0A流程号','对公付款OA流程号','OA'],
 amount:['申请付款金额','本次申请付款金额'],summary:['摘要'],due:['应付款时间','应付时间'],estimate:['预估付款时间','预估付款月份'],
 submitted:['是否提交付款申请单','是否提交付款审单'],settled:['是否结清'],negotiated:['是否谈判付款'],month:['所属月份','月份','登记月份'],terms:['账期']
};
const REQUIRED=['supplier','oa','amount','settled','negotiated'];
const norm=v=>String(v??'').normalize('NFKC').replace(/\s+/g,'');
const text=c=>String(c?.v??'').trim();
const blank=()=>({t:'s',v:''});
const aliases=new Map(Object.entries(FIELDS).flatMap(([key,names])=>names.map(name=>[norm(name),key])));
function valid(cell,label){
 if(cell?.t==='e')throw Error(label+'：Excel 错误 '+(cell.w||cell.v)+'，请修正源表后重试');
 if(cell?.f&&(cell.v==null||cell.t==='z'||cell.t==='n'&&typeof cell.v!=='number'))throw Error(label+'：公式缺少有效缓存结果，请在 Excel/WPS 中重新计算并保存');
 if(typeof cell?.v==='number'&&!Number.isFinite(cell.v))throw Error(label+'：数值不是有限数');
 return cell;
}
function copy(cell){
 if(!cell||cell.v==null||cell.t==='z')return blank();
 const result={t:cell.t==='str'?'s':cell.t||'s',v:cell.v instanceof Date?new Date(cell.v.getTime()):cell.v};
 if(cell.z)result.z=cell.z;
 return result;
}
function money(cell,label){
 valid(cell,label);if(!cell||cell.v==null||text(cell)==='')return blank();
 let value=cell.v;
 if(typeof value!=='number'){
  let s=String(value).normalize('NFKC').trim().replace(/^(?:人民币|RMB|CNY|[¥￥])\s*/i,'').replace(/\s*元$/,'');
  if(/^\(.*\)$/.test(s))s='-'+s.slice(1,-1);
  if(!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(s))throw Error(label+'：金额无法识别“'+String(value)+'”，未按零处理');
  s=s.replace(/,/g,'');
  if(s.replace(/^[+-]/,'').replace('.','').replace(/^0+/,'').replace(/0+$/,'').length>15)throw Error(label+'：金额精度超过 Excel 安全范围');
  value=Number(s);
 }
 if(!Number.isFinite(value)||Math.abs(value)>Number.MAX_SAFE_INTEGER/100)throw Error(label+'：金额超出安全范围');
 return {t:'n',v:value,z:'#,##0.00########'};
}
function date(cell,date1904,label,warnings,monthPrecision=false){
 valid(cell,label);if(!cell||cell.v==null||text(cell)==='')return blank();
 const value=cell.v,sourceFormat=typeof cell.z==='string'&&X.SSF.is_date(cell.z)?cell.z:'';
 const monthFormat=sourceFormat&&!/d/i.test(sourceFormat.replace(/"[^"]*"/g,''))?sourceFormat:'yyyy-mm';
 if(typeof value==='number'&&X.SSF.parse_date_code(value,{date1904}))return {t:'n',v:value+(date1904?1462:0),z:monthPrecision?monthFormat:sourceFormat||'yyyy-mm-dd'};
 if(value instanceof Date&&Number.isNaN(value.getTime()))throw Error(label+'：日期值无效，请核实源表');
 let year,month,day,monthOnly=false,time=0;
 if(value instanceof Date&&!Number.isNaN(value.getTime())){
  year=value.getUTCFullYear();month=value.getUTCMonth()+1;day=value.getUTCDate();
  time=(value.getUTCHours()*3600+value.getUTCMinutes()*60+value.getUTCSeconds()+value.getUTCMilliseconds()/1000)/86400;
 }else if(typeof value==='string'){
  const full=value.trim().match(/^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?$/),partial=value.trim().match(/^(\d{4})(?:[年./-])(\d{1,2})月?$/);
  if(full){year=+full[1];month=+full[2];day=+full[3];}
  else if(partial){year=+partial[1];month=+partial[2];day=1;monthOnly=true;}
 }
 const stamp=Date.UTC(year,month-1,day),check=new Date(stamp);
 if(year>=1900&&year<=9999&&check.getUTCFullYear()===year&&check.getUTCMonth()+1===month&&check.getUTCDate()===day){
  let serial=(stamp-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
  return {t:'n',v:serial+time,z:monthPrecision||monthOnly?monthFormat:sourceFormat||'yyyy-mm-dd'};
 }
 warnings.push(label+'：日期无法识别，已保留原值“'+String(value)+'”');return copy(cell);
}
function estimatedDate(monthCell,termsCell,date1904,monthLabel,termsLabel,warnings){
 valid(monthCell,monthLabel);valid(termsCell,termsLabel);
 const parsed=date(monthCell,date1904,monthLabel,[]),month=parsed.t==='n'?X.SSF.parse_date_code(parsed.v):null;
 if(!month||month.y<1900||month.y>9999){warnings.push(monthLabel+'：无预估付款时间列，所属月份为空或无法识别，未推算预估付款时间');return blank();}
 const terms=text(termsCell).normalize('NFKC'),durations=[];
 if(/^\d+$/.test(terms))durations.push(Number(terms));
 else for(const match of terms.matchAll(/(\d+(?:\.\d+)?)\s*(个月|天|日)/g))durations.push(Number(match[1])*(match[2]==='个月'?30:1));
 const days=[...new Set(durations)];
 if(days.length!==1||!Number.isSafeInteger(days[0])||days[0]<0||/[+＋~～至]|[-—−]\s*\d|\d\s*(?:\/|或者?)\s*\d|约|左右|至少|最多|以上|以内/.test(terms)){
  warnings.push(termsLabel+'：无预估付款时间列，账期为空、存在冲突或无法识别“'+terms+'”，未推算预估付款时间');return blank();
 }
 const end=Date.UTC(month.y,month.m,0),stamp=end+(days[0]+7)*86400000;
 if(!Number.isFinite(stamp)||new Date(stamp).getUTCFullYear()>9999){warnings.push(termsLabel+'：推算日期超出有效范围，预估付款时间保留空白');return blank();}
 let serial=(stamp-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:'yyyy-mm-dd'};
}
function table(sheet,sheetName,filename){
 if(!sheet)throw Error(filename+' / '+sheetName+'：工作表不可读取');
 let lastRow=0,lastCol=0;
 for(const address of Object.keys(sheet)){
  if(!/^[A-Z]+\d+$/.test(address))continue;const cell=sheet[address];
  if((!text(cell)||cell?.t==='z')&&!cell?.f)continue;
  const p=X.utils.decode_cell(address);lastRow=Math.max(lastRow,p.r);lastCol=Math.max(lastCol,p.c);
 }
 if(lastRow>100000||lastCol>300)throw Error(filename+' / '+sheetName+'：数据范围过大，请检查源表');
 let header=-1,map={},best=0;
 for(let r=0;r<=Math.min(lastRow,99);r++){
  const found={};
  for(let c=0;c<=lastCol;c++){const key=aliases.get(norm(sheet[X.utils.encode_cell({r,c})]?.v));if(key)(found[key]??=[]).push(c);}
  const score=REQUIRED.filter(key=>found[key]).length;
  if(score>best){best=score;header=r;map=found;}
  if(score===REQUIRED.length)break;
 }
 if(header<0)throw Error(filename+' / '+sheetName+'：未找到资金台账表头');
 const missing=REQUIRED.filter(key=>!map[key]);
 if(missing.length)throw Error(filename+' / '+sheetName+'：缺少必要表头 '+missing.map(key=>FIELDS[key][0]).join('、'));
 for(const [key,columns]of Object.entries(map))if(columns.length>1)throw Error(filename+' / '+sheetName+'：字段“'+FIELDS[key][0]+'”重复，无法确定来源');
 return {header,map:Object.fromEntries(Object.entries(map).map(([key,cols])=>[key,cols[0]])),lastRow};
}
function build(book,filename='资金台账.xlsx'){
 if(!book?.SheetNames||!book?.Sheets)throw Error('请先上传可读取的资金台账');
 const selected=new Set();
 for(const person of PEOPLE){
  const pattern=new RegExp('^对账明细[-－—]?'+person+'(?:[-_（(]?(?:\\d{4}|\\d{8}|\\d{4}[-./]\\d{1,2}[-./]\\d{1,2})[）)]?)?$');
  const names=book.SheetNames.filter(name=>pattern.test(norm(name)));
  if(names.length!==1)throw Error(filename+'：'+person+'的对账明细'+(names.length?'存在多个版本（'+names.join('、')+'）':'缺失')+'，请保留唯一对应工作表');
  selected.add(names[0]);
 }
 const sheets=SHEET_NAMES.map((name,i)=>({name,headers:[...HEADERS[i]],rows:[],sources:[]})),counts=[],warnings=[];
 const date1904=book.Workbook?.WBProps?.date1904===true||book.Workbook?.WBProps?.date1904===1;
 for(const name of book.SheetNames.filter(name=>selected.has(name))){
  const sheet=book.Sheets[name],{header,map,lastRow}=table(sheet,name,filename),count={sheet:name,applied:0,unapplied:0,excluded:0};
  const absent=['entity','summary','submitted'].filter(key=>map[key]===undefined);
  if(absent.length)warnings.push(filename+' / '+name+'：缺少 '+absent.map(key=>FIELDS[key][0]).join('、')+'，对应输出保留空白');
  for(let r=header+1;r<=lastRow;r++){
   const at=key=>map[key]===undefined?undefined:sheet[X.utils.encode_cell({r,c:map[key]})];
   const label=key=>filename+' / '+name+'!'+(map[key]===undefined?'第'+(r+1)+'行':X.utils.encode_cell({r,c:map[key]}))+'（'+FIELDS[key][0]+'）';
   if(Object.keys(map).every(key=>!at(key)?.f&&!text(at(key))))continue;
   valid(at('settled'),label('settled'));valid(at('negotiated'),label('negotiated'));
   if(text(at('settled'))!=='否'||text(at('negotiated'))!=='否'){count.excluded++;continue;}
   valid(at('oa'),label('oa'));const oa=text(at('oa')),index=!oa||norm(oa)==='未提交OA'?1:0,time=index?'estimate':'due';
   const read=key=>{if(map[key]===undefined)return blank();return copy(valid(at(key),label(key)));};
   const monthPrecision=index&&map[time]!==undefined&&norm(sheet[X.utils.encode_cell({r:header,c:map[time]})]?.v)==='预估付款月份';
   const timeCell=map[time]!==undefined?date(at(time),date1904,label(time),warnings,monthPrecision):index?estimatedDate(at('month'),at('terms'),date1904,label('month'),label('terms'),warnings):blank();
   const oaCell={t:'s',v:typeof at('oa')?.v==='number'?formatCell(at('oa')):String(at('oa')?.v??''),z:'@'};
   sheets[index].rows.push([read('supplier'),read('entity'),oaCell,money(at('amount'),label('amount')),read('summary'),timeCell,read('submitted')]);
   sheets[index].sources.push({sheet:name,row:r+1});count[index?'unapplied':'applied']++;
  }
  counts.push(count);
 }
 return {sheets,counts,warnings};
}
function workbook(result){
 const book=X.utils.book_new();
 result.sheets.forEach((source,i)=>{
  const sheet=X.utils.aoa_to_sheet([source.headers,...source.rows.map(row=>row.map(copy))]);
  const width=value=>String(value).split(/\r?\n/).reduce((max,line)=>Math.max(max,Array.from(line).reduce((n,ch)=>n+(ch.charCodeAt(0)>255?2:1),0)),0);
  const limits=[50,40,52,28,72,34,40];
  sheet['!cols']=source.headers.map((header,c)=>{
   const contentWidth=source.rows.reduce((max,row)=>Math.max(max,width(formatCell(row[c]))),0);
   return {wch:Math.max(12,width(header)+4,Math.min(limits[c],contentWidth+2))};
  });
  for(let r=0;r<=source.rows.length;r++){const cell=sheet[X.utils.encode_cell({r,c:4})];if(cell)cell.s={alignment:{wrapText:true,vertical:'top'}};}
  sheet['!autofilter']={ref:'A1:G'+(source.rows.length+1)};
  X.utils.book_append_sheet(book,sheet,SHEET_NAMES[i]);
 });return book;
}
function formatCell(cell){return !cell||cell.v==null?'':X.utils.format_cell(copy(cell));}
function exportFilename(filename){return String(filename||'资金台账').replace(/^.*[\\/]/,'').replace(/\.[^.]+$/,'')+'－资金台账整理.xlsx';}
return {SHEET_NAMES,HEADERS,build,workbook,formatCell,exportFilename};
});
