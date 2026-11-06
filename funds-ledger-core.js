(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FundsLedgerCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const SHEET_NAMES=['已申请OA付款','未申请OA付款'];
const EXTRA_HEADERS=['申请日期','是否已签字','当前节点','未操作者','流程状态','申请人','是否已付款'];
const HEADERS=SHEET_NAMES.map((_,i)=>['供应商','付款主体','对公付款OA流程号','申请付款金额','摘要',i?'预估付款时间':'应付款时间','是否提交付款审单',...EXTRA_HEADERS]);
const FIELDS={
 supplier:['供应商名称','供应商'],entity:['付款主体'],oa:['对公付款0A流程号','对公付款OA流程号','OA'],
 amount:['申请付款金额','本次申请付款金额'],summary:['摘要'],due:['应付款时间','应付时间'],estimate:['预估付款时间','预估付款月份'],
 submitted:['是否提交付款申请单','是否提交付款审单'],settled:['是否结清'],negotiated:['是否谈判付款'],month:['所属月份','月份','登记月份'],terms:['账期']
};
const REQUIRED=['supplier','oa','amount','settled','negotiated'];
const SUPPLEMENTS={
 applications:{sheetName:'付款申请汇总',required:['supplier','oa','amount'],preferred:{oa:'流程编码'},fields:{supplier:['供应商全称','供应商名称','供应商'],entity:['付款主体'],oa:['流程编码','合同号'],amount:['本次申请付款金额','申请付款金额'],appliedDate:['申请日期'],signature:['是否签字完成','是否已签字']}},
 progress:{sheetName:'付款进度整理',required:['supplier','oa'],fields:{supplier:['供应商名称','供应商全称','供应商'],entity:['付款主体'],oa:['对公付款OA号','对公付款0A流程号','对公付款OA流程号','OA流程号'],amount:['申请付款金额','本次申请付款金额'],node:['当前节点'],pending:['未操作者'],state:['流程状态'],applicant:['申请人'],paid:['是否已付款']}}
};
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
function table(sheet,sheetName,filename,fields=FIELDS,required=REQUIRED,preferred={}){
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
  for(const [key,label]of Object.entries(preferred)){const primary=(found[key]||[]).filter(c=>norm(sheet[X.utils.encode_cell({r,c})]?.v)===label);if(primary.length)found[key]=primary;}
  const score=required.filter(key=>found[key]).length;
  if(score>best){best=score;header=r;map=found;}
  if(score===required.length)break;
 }
 if(header<0)throw Error(filename+' / '+sheetName+'：未找到资金台账表头');
 const missing=required.filter(key=>!map[key]);
 if(missing.length)throw Error(filename+' / '+sheetName+'：缺少必要表头 '+missing.map(key=>fields[key][0]).join('、'));
 for(const [key,columns]of Object.entries(map))if(columns.length>1)throw Error(filename+' / '+sheetName+'：字段“'+fields[key][0]+'”重复，无法确定来源');
 return {header,map:Object.fromEntries(Object.entries(map).map(([key,cols])=>[key,cols[0]])),lastRow};
}
function build(book,filename='资金台账.xlsx',options={}){
 if(!book?.SheetNames||!book?.Sheets)throw Error('请先上传可读取的资金台账');
 const selected=new Set();
 for(const person of PEOPLE){
  const pattern=new RegExp('^对账明细[-－—]?'+person+'(?:[-_（(]?(?:\\d{4}|\\d{8}|\\d{4}[-./]\\d{1,2}[-./]\\d{1,2})[）)]?)?$');
  const names=book.SheetNames.filter(name=>pattern.test(norm(name)));
  if(names.length!==1)throw Error(filename+'：'+person+'的对账明细'+(names.length?'存在多个版本（'+names.join('、')+'）':'缺失')+'，请保留唯一对应工作表');
  selected.add(names[0]);
 }
 const sheets=SHEET_NAMES.map((name,i)=>({name,headers:[...HEADERS[i]],rows:[],sources:[]})),counts=[],warnings=[],identities=[];
 const date1904=book.Workbook?.WBProps?.date1904===true||book.Workbook?.WBProps?.date1904===1;
 for(const name of book.SheetNames.filter(name=>selected.has(name))){
  const sheet=book.Sheets[name],{header,map,lastRow}=table(sheet,name,filename),count={sheet:name,applied:0,unapplied:0,excluded:0};
  const absent=['entity','summary','submitted'].filter(key=>map[key]===undefined);
  if(absent.length)warnings.push(filename+' / '+name+'：缺少 '+absent.map(key=>FIELDS[key][0]).join('、')+'，对应输出保留空白');
  for(let r=header+1;r<=lastRow;r++){
   const at=key=>map[key]===undefined?undefined:sheet[X.utils.encode_cell({r,c:map[key]})];
   const label=key=>filename+' / '+name+'!'+(map[key]===undefined?'第'+(r+1)+'行':X.utils.encode_cell({r,c:map[key]}))+'（'+FIELDS[key][0]+'）';
   if(Object.keys(map).every(key=>!at(key)?.f&&!text(at(key))))continue;
   if((options.applications||options.progress)&&['supplier','oa','amount'].some(key=>at(key)?.f||at(key)?.t==='e'||text(at(key))))identities.push(ledgerIdentity(at,{file:filename,sheet:name,row:r+1}));
   valid(at('settled'),label('settled'));valid(at('negotiated'),label('negotiated'));
   if(!['否',''].includes(text(at('settled')))||text(at('negotiated'))!=='否'){count.excluded++;continue;}
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
 const result={sheets,counts,warnings};enrich(result,filename,options,identities);return result;
}
function oaKey(cell){
 const raw=String(typeof cell?.v==='number'?formatCell(cell):cell?.v??'').normalize('NFKC').trim();
 if(/[,;、|/\r\n]/.test(raw)||raw.split(/\s+/).filter(part=>/\d/.test(part)&&/^[A-Z0-9_-]+$/i.test(part)).length>1)return '';
 const value=norm(raw).toUpperCase();
 return /\d/.test(value)&&/^[A-Z0-9]+(?:[-_][A-Z0-9]+)*$/.test(value)?value:'';
}
function cents(cell){
 if(cell?.t!=='n'||typeof cell.v!=='number')return null;
 const value=cell.v*100,rounded=Math.round(value);
 return Number.isSafeInteger(rounded)&&Math.abs(value-rounded)<0.00001?String(rounded):null;
}
function ledgerIdentity(at,source){
 const values={};
 for(const key of ['supplier','entity','oa','amount']){
  try{values[key]=key==='amount'?money(at(key),''):copy(valid(at(key),''));}catch{values[key]=null;}
 }
 return {source,supplier:values.supplier===null?null:norm(values.supplier.v),entity:values.entity===null?null:norm(values.entity.v),oa:values.oa===null?null:oaKey(values.oa),cents:cents(values.amount),fallback:values.oa===null?null:!norm(values.oa.v)||norm(values.oa.v)==='未提交OA'};
}
const location=record=>record.source.file+' / '+record.source.sheet+' 第'+record.source.row+'行';
function supplement(book,role,filename){
 const spec=SUPPLEMENTS[role];if(!spec)throw Error('未知补充来源类型：'+role);
 filename=filename||spec.sheetName+'.xlsx';
 if(!book?.SheetNames?.includes(spec.sheetName)||!book.Sheets?.[spec.sheetName])throw Error(filename+'：未找到“'+spec.sheetName+'”工作表');
 const sheet=book.Sheets[spec.sheetName],{header,map,lastRow}=table(sheet,spec.sheetName,filename,spec.fields,spec.required,spec.preferred),warnings=[],records=[];
 const missing=Object.keys(spec.fields).filter(key=>map[key]===undefined);
 if(missing.length)warnings.push(filename+' / '+spec.sheetName+'：缺少 '+missing.map(key=>spec.fields[key][0]).join('、')+'，对应信息保留空白或待核对');
 const date1904=book.Workbook?.WBProps?.date1904===true||book.Workbook?.WBProps?.date1904===1;
 for(let r=header+1;r<=lastRow;r++){
  const at=key=>map[key]===undefined?undefined:sheet[X.utils.encode_cell({r,c:map[key]})];
  const label=key=>filename+' / '+spec.sheetName+'!'+(map[key]===undefined?'第'+(r+1)+'行':X.utils.encode_cell({r,c:map[key]}))+'（'+spec.fields[key][0]+'）';
  if(Object.keys(map).every(key=>!at(key)?.f&&!text(at(key))))continue;
  const values=Object.fromEntries(Object.keys(spec.fields).map(key=>[key,copy(valid(at(key),label(key)))]));
  if(map.amount!==undefined)values.amount=money(at('amount'),label('amount'));
  if(role==='applications')values.appliedDate=date(at('appliedDate'),date1904,label('appliedDate'),warnings);
  const record={values,source:{file:filename,sheet:spec.sheetName,row:r+1},order:records.length};
  record.oa=oaKey(values.oa);record.supplier=norm(values.supplier.v);record.entity=norm(values.entity.v);record.cents=cents(values.amount);records.push(record);
 }
 return {sheetName:spec.sheetName,records,warnings};
}
function validateSupplement(book,role,filename){const parsed=supplement(book,role,filename);return {sheetName:parsed.sheetName,count:parsed.records.length,warnings:parsed.warnings};}
function enrich(result,filename,options,identities){
 const applications=options.applications?supplement(options.applications.book,'applications',options.applications.filename):null;
 const progress=options.progress?supplement(options.progress.book,'progress',options.progress.filename):null;
 for(const source of [applications,progress])if(source)result.warnings.push(...source.warnings);
 const targets=result.sheets.flatMap(sheet=>sheet.rows.map((row,i)=>({row,source:{file:filename,...sheet.sources[i]},supplier:norm(row[0].v),entity:norm(row[1].v),oa:oaKey(row[2]),cents:cents(row[3]),fallback:!norm(row[2].v)||norm(row[2].v)==='未提交OA'})));
 const targetMap=new Map(targets.map(target=>[JSON.stringify([target.source.sheet,target.source.row]),target]));
 for(const owner of identities)owner.target=targetMap.get(JSON.stringify([owner.source.sheet,owner.source.row]));
 const warn=(target,message,records=[])=>result.warnings.push(location(target)+'：'+message+(records.length?'；来源：'+records.map(location).join('、'):''));
 const identity=(target,record)=>!!target.supplier&&target.supplier===record.supplier&&!(target.entity&&record.entity&&target.entity!==record.entity);
 const fallbackIdentity=(target,record)=>!!target.supplier&&!!target.entity&&target.cents!==null&&target.supplier===record.supplier&&target.entity===record.entity&&target.cents===record.cents;
 const possibleIdentity=(owner,record)=>(!owner.supplier||!record.supplier||owner.supplier===record.supplier)&&!(owner.entity&&record.entity&&owner.entity!==record.entity);
 const indexed=records=>{const map=new Map();for(const record of records)if(record.oa){if(!map.has(record.oa))map.set(record.oa,[]);map.get(record.oa).push(record);}return map;};
 const applicationIndex=indexed(applications?.records||[]),progressIndex=indexed(progress?.records||[]);
 const applicationOwners=new Map((applications?.records||[]).map(record=>{
  const exact=record.oa?identities.filter(owner=>(owner.oa===record.oa||owner.oa===null)&&possibleIdentity(owner,record)):[];
  const owners=exact.length>1?exact.filter(owner=>owner.cents===null||record.cents===null||owner.cents===record.cents):exact.length?exact:identities.filter(owner=>owner.fallback!==false&&possibleIdentity(owner,record)&&(owner.cents===null||record.cents===null||owner.cents===record.cents));
  return [record,owners];
 }));
 for(const target of targets){
  target.row.push(...EXTRA_HEADERS.map(blank));let matched=[];
  if(applications){
   if(target.oa){
    const candidates=applicationIndex.get(target.oa)||[],compatible=candidates.filter(record=>identity(target,record));
    matched=compatible.filter(record=>{const owners=applicationOwners.get(record);return owners.length===1&&owners[0].target===target;});
    const ambiguous=compatible.filter(record=>!matched.includes(record));if(ambiguous.length)warn(target,'同 OA 的付款申请不能唯一归属本条台账，未重复补入',ambiguous);
    const conflicts=candidates.filter(record=>!identity(target,record));if(conflicts.length)warn(target,'付款申请的供应商或付款主体冲突，冲突记录未补入',conflicts);
    if(!compatible.length&&!conflicts.length)warn(target,'付款申请未找到同一 OA 与供应商的记录');
   }else if(target.fallback){
    const candidates=applications.records.filter(record=>fallbackIdentity(target,record));
    if(candidates.length===1){
     const owners=applicationOwners.get(candidates[0]);
     if(owners.length===1&&owners[0].target===target)matched=candidates;else warn(target,'付款申请已由其他 OA 台账认领或可能对应多条台账记录（含筛选排除行、身份不确定行），未按供应商、主体和金额补入',[...candidates,...owners]);
    }else warn(target,candidates.length?'付款申请有多条同供应商、主体和金额候选，未任选一条补入':'付款申请缺少唯一的供应商、主体和金额匹配',candidates);
   }else warn(target,'台账 OA 为占位文字或包含多个编号，未按金额替代匹配付款申请');
   matched.sort((a,b)=>{const av=a.values.appliedDate,bv=b.values.appliedDate,ad=av.t==='n'?av.v:Infinity,bd=bv.t==='n'?bv.v:Infinity;return ad-bd||a.order-b.order;});
   if(matched.length===1){target.row[7]=copy(matched[0].values.appliedDate);target.row[8]=copy(matched[0].values.signature);}
   else if(matched.length>1)for(const [column,key]of [[7,'appliedDate'],[8,'signature']])target.row[column]={t:'s',v:matched.map((record,i)=>(i+1)+'、'+(formatCell(record.values[key])||'空白')).join('\n')};
  }
  target.matched=matched;target.queryOA=target.oa||(target.fallback&&matched.length===1?matched[0].oa:'');
 }
 if(!progress)return;
 for(const target of targets){
  target.row[13]={t:'s',v:'待核对'};const queryOA=target.queryOA;
  if(queryOA&&!target.oa)warn(target,'仅借用付款申请 OA“'+queryOA+'”查询付款进度，台账 OA 与付款分类保持原值',target.matched);
  if(!queryOA){warn(target,'缺少可唯一使用的 OA，付款进度待核对');continue;}
  const candidates=progressIndex.get(queryOA)||[],compatible=candidates.filter(record=>identity(target,record));
  if(compatible.length&&compatible.length<candidates.length)warn(target,'同 OA 的其他付款进度存在供应商或付款主体冲突，冲突记录未补入',candidates.filter(record=>!identity(target,record)));
  if(!compatible.length){warn(target,candidates.length?'付款进度的供应商或付款主体冲突，未补入':'未找到该 OA 的付款进度',candidates);continue;}
  const unique=[...new Map(compatible.map(record=>[JSON.stringify([record.supplier,record.entity,record.cents,...['node','pending','state','applicant','paid'].map(key=>record.values[key].v)]),record])).values()];
  if(unique.length>1&&unique.some(record=>record.cents===null)){warn(target,'同 OA 的多条付款进度有金额不完整记录，不能唯一归属本笔台账',unique);continue;}
  let selected=unique;
  if(unique.length>1||unique[0].cents!==null&&target.cents!==null&&unique[0].cents!==target.cents)selected=unique.filter(record=>target.cents!==null&&record.cents===target.cents);
  if(selected.length!==1){warn(target,selected.length?'同 OA、同金额的付款进度信息冲突，未合并或任选':'付款进度金额与本笔台账不能唯一对应，未按分批付款判断整笔已付',selected.length?selected:unique);continue;}
  const record=selected[0],owners=identities.filter(owner=>((owner.target?owner.target.queryOA:owner.oa)===queryOA||owner.oa===null)&&possibleIdentity(owner,record)&&(owner.cents===null||record.cents===null||owner.cents===record.cents));
  if(owners.length!==1){warn(target,'同一付款进度可能归属多条台账记录，未重复判断已付款',[record]);continue;}
  for(const [column,key]of [[9,'node'],[10,'pending'],[11,'state'],[12,'applicant']])target.row[column]=copy(record.values[key]);
  const explicit=norm(record.values.paid.v),state=norm(record.values.state.v);
  if(['重新申请','待核对'].includes(explicit))target.row[13]=copy(record.values.paid);
  else if(record.cents===null||target.cents===null)warn(target,'金额为空或不能按分核实，流程信息已显示，是否已付款待核对',[record]);
  else if(['归档','流程归档','已归档'].includes(state))target.row[13]={t:'s',v:'已付款'};
  else if(state==='未归档')target.row[13]={t:'s',v:'未付款'};
  else warn(target,'流程状态为空或无法识别，是否已付款待核对',[record]);
 }
}
function workbook(result){
 const book=X.utils.book_new();
 result.sheets.forEach((source,i)=>{
  const sheet=X.utils.aoa_to_sheet([source.headers,...source.rows.map(row=>row.map(copy))]);
  const width=value=>String(value).split(/\r?\n/).reduce((max,line)=>Math.max(max,Array.from(line).reduce((n,ch)=>n+(ch.charCodeAt(0)>255?2:1),0)),0);
  const limits=[50,40,52,28,72,34,40,36,40,42,42,32,32,32];
  sheet['!cols']=source.headers.map((header,c)=>{
   const contentWidth=source.rows.reduce((max,row)=>Math.max(max,width(formatCell(row[c]))),0);
   return {wch:Math.max(12,width(header)+4,Math.min(limits[c],contentWidth+2))};
  });
  for(let r=0;r<=source.rows.length;r++)for(const c of [4,7,8,9,10,11,12,13]){const cell=sheet[X.utils.encode_cell({r,c})];if(cell)cell.s={alignment:{wrapText:true,vertical:'top'}};}
  sheet['!autofilter']={ref:'A1:'+X.utils.encode_col(source.headers.length-1)+(source.rows.length+1)};
  X.utils.book_append_sheet(book,sheet,SHEET_NAMES[i]);
 });return book;
}
function formatCell(cell){return !cell||cell.v==null?'':X.utils.format_cell(copy(cell));}
function exportFilename(filename){return String(filename||'资金台账').replace(/^.*[\\/]/,'').replace(/\.[^.]+$/,'')+'－资金台账整理.xlsx';}
return {SHEET_NAMES,HEADERS,build,workbook,formatCell,exportFilename,validateSupplement};
});
