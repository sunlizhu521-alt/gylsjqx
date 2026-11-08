(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FundsLedgerCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const HEADERS=['供应商','付款主体','对公付款OA流程号','申请付款金额','摘要','应付款时间','当前节点','未操作者','流程状态','是否已付款','需要提交付款审单的日期','是否提交纸质付款审单'];
const FIELDS={
 supplier:['供应商名称','供应商'],entity:['付款主体'],oa:['对公付款0A流程号','对公付款OA流程号','OA'],
 amount:['申请付款金额','本次申请付款金额'],summary:['摘要'],due:['应付款时间','应付时间'],estimate:['预估付款时间','预估付款月份'],
 negotiated:['是否谈判付款'],month:['所属月份','月份','登记月份'],terms:['账期']
};
const REQUIRED=['supplier','oa','amount','negotiated'];
const OA_FIELDS={oa:['流程编号','流程编码'],node:['当前节点'],pending:['未操作者'],state:['流程状态']};
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
 warnings.push(label+'：日期无法识别“'+String(value)+'”，未归入任何月份');return copy(cell);
}
function estimatedDate(monthCell,termsCell,date1904,monthLabel,termsLabel,warnings){
 valid(monthCell,monthLabel);valid(termsCell,termsLabel);
 const parsed=date(monthCell,date1904,monthLabel,[]),month=parsed.t==='n'?X.SSF.parse_date_code(parsed.v):null;
 if(!month||month.y<1900||month.y>9999){warnings.push(monthLabel+'：应付及预估时间为空，所属月份为空或无法识别，未推算付款月份');return blank();}
 const terms=text(termsCell).normalize('NFKC'),durations=[];
 if(/^\d+$/.test(terms))durations.push(Number(terms));
 else for(const match of terms.matchAll(/(\d+(?:\.\d+)?)\s*(个月|天|日)/g))durations.push(Number(match[1])*(match[2]==='个月'?30:1));
 const days=[...new Set(durations)];
 if(days.length!==1||!Number.isSafeInteger(days[0])||days[0]<0||/[+＋~～至]|[-—−]\s*\d|\d\s*(?:\/|或者?)\s*\d|约|左右|至少|最多|以上|以内/.test(terms)){
  warnings.push(termsLabel+'：应付及预估时间为空，账期为空、存在冲突或无法识别“'+terms+'”，未推算付款月份');return blank();
 }
 const end=Date.UTC(month.y,month.m,0),stamp=end+(days[0]+7)*86400000;
 if(!Number.isFinite(stamp)||new Date(stamp).getUTCFullYear()>9999){warnings.push(termsLabel+'：推算日期超出有效范围，预估付款时间保留空白');return blank();}
 let serial=(stamp-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:'yyyy-mm-dd'};
}
function table(sheet,sheetName,filename,fields=FIELDS,required=REQUIRED,probe=false){
 if(!sheet)throw Error(filename+' / '+sheetName+'：工作表不可读取');
 let lastRow=0,lastCol=0;
 for(const address of Object.keys(sheet)){
  if(!/^[A-Z]+\d+$/.test(address))continue;const cell=sheet[address];
  if((!text(cell)||cell?.t==='z')&&!cell?.f)continue;
  const p=X.utils.decode_cell(address);lastRow=Math.max(lastRow,p.r);lastCol=Math.max(lastCol,p.c);
 }
 if(lastRow>100000||lastCol>300)throw Error(filename+' / '+sheetName+'：数据范围过大，请检查源表');
 const fieldAliases=fields===FIELDS?aliases:new Map(Object.entries(fields).flatMap(([key,names])=>names.map(name=>[norm(name),key])));
 let header=-1,map={},best=0;
 for(let r=0;r<=Math.min(lastRow,99);r++){
  const found={};
  for(let c=0;c<=lastCol;c++){const key=fieldAliases.get(norm(sheet[X.utils.encode_cell({r,c})]?.v));if(key)(found[key]??=[]).push(c);}
  const score=required.filter(key=>found[key]).length;
  if(score>best){best=score;header=r;map=found;}
  if(score===required.length)break;
 }
 if(header<0){if(probe)return null;throw Error(filename+' / '+sheetName+'：未找到资金台账表头');}
 const missing=required.filter(key=>!map[key]);
 if(missing.length){if(probe)return null;throw Error(filename+' / '+sheetName+'：缺少必要表头 '+missing.map(key=>fields[key][0]).join('、'));}
 for(const [key,columns]of Object.entries(map))if(columns.length>1)throw Error(filename+' / '+sheetName+'：字段“'+fields[key][0]+'”重复，无法确定来源');
 return {header,map:Object.fromEntries(Object.entries(map).map(([key,cols])=>[key,cols[0]])),lastRow};
}
function build(book,filename='资金台账.xlsx',options={}){
 if(!book?.SheetNames||!book?.Sheets)throw Error('请先上传可读取的资金台账');
 const today=options.today===undefined?new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()):options.today;
 if(typeof today!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(today))throw Error('申请日期基准必须是有效日期（YYYY-MM-DD）');
 const baseline=date({t:'s',v:today},false,'申请日期基准',[]);
 if(baseline.t!=='n')throw Error('申请日期基准必须是有效日期（YYYY-MM-DD）');
 const thisMonday=mondayStamp(X.SSF.parse_date_code(baseline.v));
 const month=options.month===undefined?today.slice(0,7):options.month;
 if(!/^(?:19\d{2}|[2-9]\d{3})-(?:0[1-9]|1[0-2])$/.test(month))throw Error('请选择有效月份（YYYY-MM）');
 const selected=new Set();
 for(const person of PEOPLE){
  const pattern=new RegExp('^对账明细[-－—]?'+person+'(?:[-_（(]?(?:\\d{4}|\\d{8}|\\d{4}[-./]\\d{1,2}[-./]\\d{1,2})[）)]?)?$');
  const names=book.SheetNames.filter(name=>pattern.test(norm(name)));
  if(names.length!==1)throw Error(filename+'：'+person+'的对账明细'+(names.length?'存在多个版本（'+names.join('、')+'）':'缺失')+'，请保留唯一对应工作表');
  selected.add(names[0]);
 }
 const rows=[],counts=[],warnings=[],date1904=book.Workbook?.WBProps?.date1904===true||book.Workbook?.WBProps?.date1904===1;
 for(const name of book.SheetNames.filter(name=>selected.has(name))){
  const sheet=book.Sheets[name],{header,map,lastRow}=table(sheet,name,filename),count={sheet:name,included:0,excluded:0,undated:0};
  const absent=['entity','summary'].filter(key=>map[key]===undefined);
  if(absent.length)warnings.push(filename+' / '+name+'：缺少 '+absent.map(key=>FIELDS[key][0]).join('、')+'，对应输出保留空白');
  for(let r=header+1;r<=lastRow;r++){
   const at=key=>map[key]===undefined?undefined:sheet[X.utils.encode_cell({r,c:map[key]})];
   const label=key=>filename+' / '+name+'!'+(map[key]===undefined?'第'+(r+1)+'行':X.utils.encode_cell({r,c:map[key]}))+'（'+FIELDS[key][0]+'）';
   if(Object.keys(map).every(key=>!at(key)?.f&&!text(at(key))))continue;
   valid(at('negotiated'),label('negotiated'));
   if(text(at('negotiated'))!=='否'){count.excluded++;continue;}
   const warningCount=warnings.length;let timeCell,inferred=false;
   valid(at('due'),label('due'));
   if(text(at('due')))timeCell=date(at('due'),date1904,label('due'),warnings);
   else{
    valid(at('estimate'),label('estimate'));
    if(text(at('estimate'))){
     const monthPrecision=norm(sheet[X.utils.encode_cell({r:header,c:map.estimate})]?.v)==='预估付款月份';
     timeCell=date(at('estimate'),date1904,label('estimate'),warnings,monthPrecision);
    }else{
     inferred=true;
     timeCell=estimatedDate(at('month'),at('terms'),date1904,label('month'),label('terms'),warnings);
    }
   }
   const parts=timeCell.t==='n'?X.SSF.parse_date_code(timeCell.v):null;
   if(!parts||parts.y<1900||parts.y>9999){
    count.undated++;if(warnings.length===warningCount)warnings.push(filename+' / '+name+' 第'+(r+1)+'行：应付款时间及可用预估付款时间为空，未归入任何月份');continue;
   }
   const paymentMonth=String(parts.y).padStart(4,'0')+'-'+String(parts.m).padStart(2,'0');
   if(paymentMonth!==month){count.excluded++;continue;}
   const read=key=>copy(valid(at(key),label(key)));valid(at('oa'),label('oa'));
   const oa={t:'s',v:typeof at('oa')?.v==='number'?formatCell(at('oa')):String(at('oa')?.v??''),z:'@'};
   const monthOnly=timeCell.z&&!/d/i.test(timeCell.z.replace(/"[^"]*"/g,''));
   timeCell.z=monthOnly?'yyyy"年"mm"月"':'yyyy"年"mm"月"dd"日"';
   let submissionDate=blank();
   if(inferred||monthOnly)warnings.push(filename+' / '+name+' 第'+(r+1)+'行：应付款时间'+(inferred?'为“未对账待定”':'只有年月')+'，需要提交付款审单的日期留空；请补充具体付款日期');
   else{
    const stamp=Math.max(thisMonday,mondayStamp(parts)-7*86400000);
    submissionDate=date({t:'d',v:new Date(stamp)},false,'需要提交付款审单的日期',[]);
    submissionDate.z='yyyy"年"mm"月"dd"日"';
   }
   rows.push({row:[read('supplier'),read('entity'),oa,money(at('amount'),label('amount')),read('summary'),inferred?{t:'s',v:'未对账待定',z:'@'}:timeCell,...Array.from({length:4},blank),submissionDate,blank()],source:{sheet:name,row:r+1},day:monthOnly?0:Math.floor(timeCell.v),order:rows.length});
   count.included++;
  }
  counts.push(count);
 }
 rows.sort((a,b)=>a.day-b.day||a.order-b.order);
 if(options.oa){
  const parsed=parseOA(options.oa.book,options.oa.filename);warnings.push(...parsed.warnings);
  for(const record of rows){
   const label=filename+' / '+record.source.sheet+' 第'+record.source.row+'行',key=oaKey(record.row[2]),matches=parsed.byOA.get(key)||[];
   record.row[9]={t:'s',v:'待核对'};
   if(!key){warnings.push(label+'：没有可唯一匹配的 OA 编号，是否已付款待核对');continue;}
   if(!matches.length){warnings.push(label+'：OA“'+key+'”在 '+parsed.filename+' / '+parsed.sheetName+' 未匹配，是否已付款待核对');continue;}
   const unique=new Set(matches.map(match=>JSON.stringify(['node','pending','state'].map(field=>String(match.values[field].v??'').normalize('NFKC').trim()))));
   const sources=matches.map(match=>parsed.filename+' / '+parsed.sheetName+' 第'+match.row+'行').join('、');
   if(unique.size!==1){warnings.push(label+'：OA“'+key+'”的流程信息冲突，未任选记录；来源：'+sources);continue;}
   const values=matches[0].values;for(const [column,field]of [[6,'node'],[7,'pending'],[8,'state']])record.row[column]=copy(values[field]);
   const state=norm(values.state.v);
   if(['归档','已归档'].includes(state))record.row[9]={t:'s',v:'已付款'};
   else if(state==='未归档')record.row[9]={t:'s',v:'未付款'};
   else warnings.push(label+'：OA“'+key+'”流程状态为空或无法识别，是否已付款待核对；来源：'+sources);
  }
 }
 if(options.applications){
  const parsed=parseApplications(options.applications.book,options.applications.filename);warnings.push(...parsed.warnings);
  for(const record of rows){
   const key=oaKey(record.row[2]);
   record.row[11]={t:'s',v:key?(parsed.codes.has(key)?'已提交':'未提交'):'待核对'};
  }
 }
 const name=month.slice(0,4)+'年'+Number(month.slice(5))+'月正常付款';
 return {month,sheets:[{name,headers:[...HEADERS],rows:rows.map(record=>record.row),sources:rows.map(record=>record.source)}],counts,warnings};
}
function mondayStamp(parts){
 const stamp=Date.UTC(parts.y,parts.m-1,parts.d);
 return stamp-((new Date(stamp).getUTCDay()+6)%7)*86400000;
}
function oaKey(cell){
 const raw=String(typeof cell?.v==='number'?formatCell(cell):cell?.v??'').normalize('NFKC').trim();
 if(/[,;、|/\r\n]/.test(raw)||raw.split(/\s+/).filter(part=>/\d/.test(part)&&/^[A-Z0-9_-]+$/i.test(part)).length>1)return '';
 const value=norm(raw).toUpperCase();
 return /\d/.test(value)&&/^[A-Z0-9]+(?:[-_][A-Z0-9]+)*$/.test(value)?value:'';
}
function parseOA(book,filename='对公付款OA.xlsx'){
 if(!book?.SheetNames||!book?.Sheets)throw Error(filename+'：OA 文件不可读取');
 const candidates=[];
 for(const name of book.SheetNames){const found=table(book.Sheets[name],name,filename,OA_FIELDS,Object.keys(OA_FIELDS),true);if(found)candidates.push({name,...found});}
 if(candidates.length!==1)throw Error(filename+'：'+(candidates.length?'找到多个 OA 工作表（'+candidates.map(candidate=>candidate.name).join('、')+'），无法唯一确定来源':'未找到包含流程编号/流程编码、当前节点、未操作者、流程状态的 OA 工作表'));
 const {name:sheetName,header,map,lastRow}=candidates[0],sheet=book.Sheets[sheetName],byOA=new Map(),warnings=[];let count=0;
 for(let r=header+1;r<=lastRow;r++){
  const at=key=>sheet[X.utils.encode_cell({r,c:map[key]})];
  if(Object.keys(map).every(key=>!at(key)?.f&&!text(at(key))))continue;
  const values=Object.fromEntries(Object.keys(OA_FIELDS).map(key=>[key,copy(valid(at(key),filename+' / '+sheetName+'!'+X.utils.encode_cell({r,c:map[key]})+'（'+OA_FIELDS[key][0]+'）'))]));
  const key=oaKey(values.oa);count++;
  if(!key){warnings.push(filename+' / '+sheetName+' 第'+(r+1)+'行：流程编号为空或不是单一编号，未用于匹配');continue;}
  if(!byOA.has(key))byOA.set(key,[]);byOA.get(key).push({values,row:r+1});
 }
 return {filename,sheetName,count,warnings,byOA};
}
function validateOA(book,filename){const parsed=parseOA(book,filename);return {sheetName:parsed.sheetName,count:parsed.count,warnings:parsed.warnings};}
function parseApplications(book,filename='付款申请单汇总.xlsx'){
 if(!book?.SheetNames||!book?.Sheets)throw Error(filename+'：付款申请单汇总不可读取');
 const names=book.SheetNames.includes('付款申请汇总')?['付款申请汇总']:book.SheetNames,candidates=[];
 for(const name of names){
  // Choose a column for the whole sheet. A blank newer field never falls back to 合同号.
  for(const heading of ['流程编码','流程编号','合同号']){
   const found=table(book.Sheets[name],name,filename,{oa:[heading]},['oa'],true);
   if(found){candidates.push({name,heading,...found});break;}
  }
 }
 if(candidates.length!==1)throw Error(filename+'：'+(candidates.length?'付款申请来源不明确，找到多个候选工作表（'+candidates.map(candidate=>candidate.name).join('、')+'）':'未在'+(names.length===1?'“'+names[0]+'”':'工作表中')+'找到流程编码／流程编号／合同号表头'));
 const {name:sheetName,heading,header,map}=candidates[0],sheet=book.Sheets[sheetName],activeRows=new Set(),codes=new Set(),warnings=[];
 for(const address of Object.keys(sheet)){
  if(!/^[A-Z]+\d+$/.test(address))continue;
  const cell=sheet[address],r=X.utils.decode_cell(address).r;
  if(r>header&&(cell?.f||cell?.t!=='z'&&text(cell)))activeRows.add(r);
 }
 for(const r of [...activeRows].sort((a,b)=>a-b)){
  const address=X.utils.encode_cell({r,c:map.oa}),label=filename+' / '+sheetName+'!'+address+'（'+heading+'）';
  const key=oaKey(valid(sheet[address],label));
  if(key)codes.add(key);
  else warnings.push(label+'：编号为空或不是单一有效编号，未用于纸质付款审单匹配，请核实该行编号');
 }
 return {sheetName,count:activeRows.size,warnings,codes};
}
function validateApplications(book,filename){const parsed=parseApplications(book,filename);return {sheetName:parsed.sheetName,count:parsed.count,warnings:parsed.warnings};}
function workbook(result){
 const book=X.utils.book_new();
 for(const source of result.sheets){
  const sheet=X.utils.aoa_to_sheet([source.headers,...source.rows.map(row=>row.map(copy))]);
  const width=value=>String(value).split(/\r?\n/).reduce((max,line)=>Math.max(max,Array.from(line).reduce((n,ch)=>n+(ch.charCodeAt(0)>255?2:1),0)),0);
  const limits=[50,40,52,28,72,34,42,42,32,32,34,32];
  sheet['!cols']=source.headers.map((header,c)=>{
   const contentWidth=source.rows.reduce((max,row)=>Math.max(max,width(formatCell(row[c]))),0);
   return {wch:Math.max(12,width(header)+4,Math.min(limits[c],contentWidth+2))};
  });
  for(let r=0;r<=source.rows.length;r++)for(const c of [4,6,7,8,9]){const cell=sheet[X.utils.encode_cell({r,c})];if(cell)cell.s={alignment:{wrapText:true,vertical:'top'}};}
  sheet['!autofilter']={ref:'A1:'+X.utils.encode_col(source.headers.length-1)+(source.rows.length+1)};
  X.utils.book_append_sheet(book,sheet,source.name);
 }
 return book;
}
function formatCell(cell){return !cell||cell.v==null?'':X.utils.format_cell(copy(cell));}
function exportFilename(date,version=1){return `资金台账整理-${date}-V${version}.xlsx`;}
return {HEADERS,build,workbook,formatCell,exportFilename,validateOA,validateApplications};
});
