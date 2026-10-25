(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FundsCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const NORMAL_HEADERS=['登记月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单','提交时间汇总','当前节点','未操作者','流程状态','申请人','流程编号','是否对账','付款自然周'];
const ADDED_HEADERS=['提交付款申请单日期','提交时间汇总','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号','付款自然周'];
const SUPPLEMENT_HEADERS=['对公付款0A流程号','付款主体','单据号【财务应付单】'];
const REVIEW_HEADERS=['付款类别','供应商名称','付款主体','申请付款金额','是否提交付款申请单','核对说明','候选付款申请','来源文件','来源工作表','来源行号'];
const FLOW_FIELDS=['当前节点','未操作者','流程状态','申请人'];
const OA_FIELDS=['流程编号','收款单位','付款金额',...FLOW_FIELDS];
const DATE_FORMAT='yyyy"年"m"月"d"日"';
const norm=v=>String(v??'').normalize('NFKC').replace(/\s+/g,'').trim();
const str=v=>String(v??'').trim();
const empty=()=>({t:'s',v:''});
const cell=v=>({t:typeof v==='number'?'n':'s',v:v??''});
const clone=c=>c?{...c}:empty();
const aliases={月份:'登记月份',所属月份:'登记月份',供应商:'供应商名称',供应商全称:'供应商名称',对公付款OA流程号:'对公付款0A流程号',付款OA号:'对公付款0A流程号',应付时间:'应付款时间',已付款金额:'已付金额',剩余付款金额:'剩余金额',合同号:'OA合同编号',合同金额:'合同总金额',本次申请付款金额:'申请付款金额',付款事由:'摘要'};
const canonical=v=>aliases[norm(v)]||norm(v);
const BATCH_HEADERS=['谈判明细','分批总金额','核对'];
const relevant=new Set([...NORMAL_HEADERS,...ADDED_HEADERS,...BATCH_HEADERS,'付款分类','申请日期','备注','对接人','创建人','流程编号','流程编码']);
const entity=v=>({'浙江迈德斯特医疗器械科技有限公司':'迈德斯特','杭州国源养老科技有限公司':'国源'}[norm(v)]||norm(v));
function money(c){
 if(c?.t==='e'||(c?.f&&(c.v==null||c.v==='')))return {invalid:true};
 let v=c?.v;if(v==null||v==='')return {blank:true};
 if(typeof v==='number'){
  if(!Number.isFinite(v)||Math.abs(v)>Number.MAX_SAFE_INTEGER/100)return {invalid:true};
  const n=Math.round(v*100);if(Math.abs(v*100-n)>0.00001)return {invalid:true};return {cents:BigInt(n)};
 }
 let text=str(v).replace(/[￥¥,，\s]/g,'');if(/^\(.*\)$/.test(text))text='-'+text.slice(1,-1);
 if(!/^[+-]?\d+(?:\.\d{1,2})?$/.test(text))return {invalid:true};
 const negative=text.startsWith('-'),parts=text.replace(/^[+-]/,'').split('.');const cents=(BigInt(parts[0])*100n+BigInt((parts[1]||'').padEnd(2,'0')))*(negative?-1n:1n);
 return cents>BigInt(Number.MAX_SAFE_INTEGER)||cents< -BigInt(Number.MAX_SAFE_INTEGER)?{invalid:true}:{cents};
}
const amountCell=n=>({t:'n',v:Number(n)/100,z:'#,##0.00'});
const centsText=n=>{const a=n<0n?-n:n;return (n<0n?'-':'')+(a/100n)+'.'+String(a%100n).padStart(2,'0');};
const amountText=c=>{const n=money(c).cents;return n===undefined?str(c?.w??c?.v)||'金额为空':centsText(n);};
function codes(v,ap=false){const s=str(v).toUpperCase();return [...new Set(s.match(ap?/AP\d+/g:/[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)*\d[A-Z0-9-]*/g)||[])];}
const identifiersConflict=(a,b)=>a.length>0&&b.length>0&&(a.length!==b.length||a.some(v=>!b.includes(v)));
function settled(v){const n=norm(v);return ['是','结清','已结清'].includes(n)?'是':['否','未结清'].includes(n)?'否':n;}
function date(c,date1904,month=false){
 if(!c||c.v==null||c.v==='')return empty();
 if(c.t==='e'||(c.f&&c.v==null))return {...cell(c.w||'公式结果不可用'),invalid:true};
 let parts;
 if(typeof c.v==='number')parts=X.SSF.parse_date_code(c.v,{date1904});
 else if(c.v instanceof Date)parts={y:c.v.getFullYear(),m:c.v.getMonth()+1,d:c.v.getDate()};
 else{const m=str(c.v).match(/^(\d{4})[年./-](\d{1,2})(?:[月./-](\d{1,2})日?)?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);if(m&& (m[3]||month))parts={y:+m[1],m:+m[2],d:+(m[3]||1)};}
 if(!parts)return {...cell(c.v),invalid:true};
 const t=Date.UTC(parts.y,parts.m-1,parts.d),d=new Date(t);
 if(parts.y<1900||parts.y>9999||d.getUTCFullYear()!==parts.y||d.getUTCMonth()+1!==parts.m||d.getUTCDate()!==parts.d)return {...cell(c.v),invalid:true};
 let serial=(t-Date.UTC(1899,11,30))/86400000;if(serial<61)serial--;
 return {t:'n',v:serial,z:month?'yyyy"年"m"月"':DATE_FORMAT};
}
function paymentWeek(c,date1904){
 const parsed=date(c,date1904);
 if(parsed.invalid||parsed.v===''||c?.t==='e')return cell('日期待确认');
 const parts=X.SSF.parse_date_code(parsed.v),start=new Date(Date.UTC(parts.y,parts.m-1,parts.d));
 start.setUTCDate(start.getUTCDate()-(start.getUTCDay()+6)%7);
 return cell(String(start.getUTCFullYear()).padStart(4,'0')+'/'+String(start.getUTCMonth()+1).padStart(2,'0')+'/'+String(start.getUTCDate()).padStart(2,'0')+'周');
}
function plannedPaymentWeeks(t,row){
 const weeks=new Set();let invalid=false;
 for(let c=0;c<=t.range.e.c;c++){
  const header=t.sheet[X.utils.encode_cell({r:t.headerRow,c})],parsed=date(header,t.date1904);
  const valid=parsed.v!==''&&!parsed.invalid&&header?.t!=='e';
  // Incomplete date headings must not inherit a year from today or a neighbouring plan.
  if(!valid&&!/^(?:\d{4}[年./-]\d{1,2}(?:[月./-]\d{1,2}日?)?|\d{1,2}月\d{1,2}日?)$/.test(str(header?.v)))continue;
  const amount=money(t.sheet[X.utils.encode_cell({r:row,c})]);
  if(amount.blank||amount.cents===0n)continue;
  if(amount.invalid||amount.cents<0n||!valid){invalid=true;continue;}
  weeks.add(paymentWeek(header,t.date1904).v);
 }
 const labels=[...weeks].sort();if(invalid)labels.push('付款计划待确认');
 return cell(labels.join('；')||'日期待确认');
}
function submissionDates(c,date1904){
 if(c?.t==='e'||(c?.f&&(c.v==null||c.v==='')))return null;
 if(!c||c.v==null||c.v==='')return [];
 const parts=typeof c.v==='string'?c.v.split(/[、,，;；\n]+/).map(str).filter(Boolean):[c.v];
 const values=[];
 for(const part of parts){const d=date({v:part},date1904);if(d.invalid||d.v==='')return null;values.push(X.SSF.format('yyyy-mm-dd',d.v));}
 return values;
}
function submissionSummary(matched){
 return cell(matched.map((r,i)=>{
  const d=date(r.values['申请日期'],r.date1904),valid=!d.invalid&&d.v!=='';
  const label=valid?X.SSF.format('yyyy-mm-dd',d.v):str(r.values['申请日期']?.w??r.values['申请日期']?.v)||'日期为空';
  return {sort:valid?label:'\uffff',i,text:label+'-'+amountText(r.values['申请付款金额'])};
 }).sort((a,b)=>a.sort.localeCompare(b.sort)||a.i-b.i).map(r=>r.text).join('\n'));
}
function table(input,sheetName,required){
 const sheet=input.book.Sheets[sheetName];if(!sheet)throw Error(input.name+'：缺少工作表 '+sheetName);
 // WPS can extend !ref to XFD solely because of formatting. Keep all real values/formulas.
 const points=Object.keys(sheet).filter(a=>/^[A-Z]+[1-9]\d*$/.test(a)&&(sheet[a]?.f||sheet[a]?.v!==null&&sheet[a]?.v!==undefined&&sheet[a]?.v!=='')).map(a=>X.utils.decode_cell(a));
 const range=points.reduce((range,p)=>({s:{r:Math.min(range.s.r,p.r),c:Math.min(range.s.c,p.c)},e:{r:Math.max(range.e.r,p.r),c:Math.max(range.e.c,p.c)}}),{s:points[0]||{r:0,c:0},e:points[0]||{r:0,c:0}});
 if((range.e.r-range.s.r+1)*(range.e.c-range.s.c+1)>2000000)throw Error(sheetName+'：数据范围过大');
 let headerRow=-1,map,headers;
 for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){
  const h={},hs=[];for(let c=0;c<=range.e.c;c++){const cc=sheet[X.utils.encode_cell({r,c})],key=canonical(cc?.v);hs.push(str(cc?.w??cc?.v));if((required.includes('流程编号')?OA_FIELDS.includes(key):relevant.has(key)))(h[key]??=[]).push(c);}
  if(required.every(k=>h[k])){headerRow=r;map=h;headers=hs;break;}
 }
 if(headerRow<0)throw Error(input.name+' / '+sheetName+'：找不到表头，必须包含 '+required.join('、'));
 for(const [key,cols] of Object.entries(map))if(cols.length>1)throw Error(sheetName+'：字段“'+key+'”重复，无法确定来源');
 map=Object.fromEntries(Object.entries(map).map(([k,v])=>[k,v[0]]));
 const records=[];
 for(let r=headerRow+1;r<=range.e.r;r++){
  const values={};for(const [k,c] of Object.entries(map))values[k]=clone(sheet[X.utils.encode_cell({r,c})]);
  const supplier=str(values['供应商名称']?.v),oa=str(values['对公付款0A流程号']?.v),amount=values['申请付款金额']?.v;
  if(required.includes('流程编号')){if(!str(values['流程编号']?.v))continue;}
  else{if(/^(?:合计|总计|小计)$/.test(supplier))continue;if(!supplier&&!oa&&!str(values['流程编码']?.v)&&!str(values['OA合同编号']?.v)&&(amount==null||amount===''))continue;}
  records.push({id:sheetName+'!'+(r+1),r,values,source:{file:input.name,sheet:sheetName,row:r+1},date1904:!!input.book.Workbook?.WBProps?.date1904,owner:sheetName.replace(/^对账明细[-－—]?/,'')});
 }
 return {sheetName,sheet,range,headerRow,map,headers,records,name:input.name,date1904:!!input.book.Workbook?.WBProps?.date1904};
}
function oneTable(input,required){
 const found=[];for(const name of input.book.SheetNames){try{found.push(table(input,name,required));}catch(e){if(!e.message.includes('找不到表头'))throw e;}}
 if(found.length!==1)throw Error(input.name+'：'+(found.length?'存在多个可用明细工作表，请只保留本次明细':'未找到包含 '+required.join('、')+' 的工作表'));
 return found[0];
}
const value=(r,k)=>r.values[k]?.v;
const oas=r=>codes(value(r,'对公付款0A流程号'));
const hasPaymentOA=r=>r.values['对公付款0A流程号']?.t!=='e'&&norm(value(r,'对公付款0A流程号'))!=='';
const recordLabel=r=>`${r.source.file} / ${r.source.sheet} 第 ${r.source.row} 行；${str(value(r,'供应商名称'))}；${str(value(r,'付款主体'))}；申请 ${str(value(r,'申请付款金额'))}；${str(value(r,'对公付款0A流程号')||value(r,'OA合同编号'))}；${str(value(r,'摘要'))}；财务单 ${str(value(r,'单据号【财务应付单】')||value(r,'备注'))}；合同 ${str(value(r,'OA合同编号'))}`+(r.values['流程编码']?'；流程编码 '+str(value(r,'流程编码')):'');
function negotiatedSheets(book){return (book?.SheetNames||[]).filter(n=>n.includes('谈判付款'));}
function build(inputs,decisions={}){
 for(const role of ['ledger','applications','negotiated','oa'])if(!inputs[role]?.book)throw Error('请上传全部四份文件');
 const candidates=negotiatedSheets(inputs.negotiated.book),selected=inputs.negotiated.sheetName||(candidates.length===1?candidates[0]:'');
 if(!selected||!candidates.includes(selected))throw Error(candidates.length?'请选择包含“谈判付款”的工作表':'未找到名称包含“谈判付款”的工作表');
 const negotiated=table(inputs.negotiated,selected,['供应商名称','申请付款金额']);
 const applications=oneTable(inputs.applications,['供应商名称','申请付款金额']);
 const flow=oneTable(inputs.oa,['流程编号']);
 const ledgerTables=PEOPLE.map(person=>{
  const names=inputs.ledger.book.SheetNames.filter(n=>new RegExp('^对账明细[-－—]?'+person+'(?:\\d{4}|\\d{8})?$').test(norm(n)));
  if(names.length!==1)throw Error('资金台账中'+person+'的对账明细'+(names.length?'存在多个版本':'缺失'));
  return table(inputs.ledger,names[0],['供应商名称','申请付款金额']);
 });
 const ledger=ledgerTables.flatMap(t=>t.records),issues=[],updateMap=new Map(),reviewRows=[];
 function issue(id,kind,message,rec,choices=[{value:'keep',label:'保留原值'}],multiple=false){
  const requested=Array.isArray(decisions[id])?decisions[id]:[],allowed=new Set(choices.map(c=>c.value));
  const valid=requested.length>0&&requested.every(v=>allowed.has(v))&&new Set(requested).size===requested.length&&(multiple||requested.length===1)&&(!(requested.includes('keep')||requested.includes('not-submitted'))||requested.length===1);
  const answer=valid?requested:[];
  issues.push({id,kind,message,source:rec.source,targetLabel:recordLabel(rec),choices,multiple,resolved:valid,selected:answer});return answer;
 }
 const headerRecord=t=>({source:{file:t.name,sheet:t.sheetName,row:t.headerRow+1},values:{}});
 for(const t of ledgerTables){const absent=NORMAL_HEADERS.slice(0,17).filter(k=>k!=='是否谈判付款'&&t.map[k]===undefined);if(absent.length)issue('headers:'+t.sheetName,'missing-header','缺少字段，输出保留空白：'+absent.join('、'),headerRecord(t));}
 const absentFlow=['收款单位','付款金额',...FLOW_FIELDS].filter(k=>flow.map[k]===undefined);
 if(absentFlow.length)issue('headers:'+flow.sheetName,'missing-header','缺少字段，无法据此自动匹配或更新：'+absentFlow.join('、'),headerRecord(flow));
 const targetCols={...negotiated.map},negHeaders=[...negotiated.headers];
 const batchLayout=BATCH_HEADERS.every(h=>targetCols[h]!==undefined);
 const addedHeaders=batchLayout?ADDED_HEADERS.filter(h=>!['提交付款申请单日期','提交时间汇总','付款自然周'].includes(h)):ADDED_HEADERS;
 for(const h of [...SUPPLEMENT_HEADERS,'应付款时间',...addedHeaders])if(targetCols[h]===undefined){targetCols[h]=negHeaders.length;negHeaders.push(h);updateMap.set(X.utils.encode_cell({r:negotiated.headerRow,c:targetCols[h]}),cell(h));}
 const classification=r=>{
  const supplier=r.values['供应商名称'],amount=money(r.values['申请付款金额']);
  const name=supplier?.t==='e'||(supplier?.f&&(supplier.v==null||supplier.v===''))?'':norm(supplier?.v);
  const cents=amount.cents!==undefined&&amount.cents>=0n?amount.cents:undefined;
  return {name,cents,key:name&&cents!==undefined?JSON.stringify([name,cents.toString()]):''};
 };
 // Prefer the new column even when its value is blank; old files may still use "合同号".
 const applicationOAField=applications.map['流程编码']===undefined?'OA合同编号':'流程编码';
 const applicationIds=record=>{const reference=record.values[applicationOAField];return reference?.t==='e'?[]:codes(reference?.v);};
 const applicationNegotiatedIds=new Set(applications.records.filter(r=>r.values['付款分类']?.t!=='e'&&norm(value(r,'付款分类'))==='谈判付款').flatMap(applicationIds));
 const isApplicationNegotiated=r=>oas(r).some(id=>applicationNegotiatedIds.has(id));
 const negotiatedKeys=new Set(),uncertainNegotiated=[];
 for(const r of negotiated.records){const info=classification(r);if(info.key)negotiatedKeys.add(info.key);else uncertainNegotiated.push({record:r,...info});}
 const normal=[];let negotiatedLedgerRows=0,excludedNormalRows=0;
 for(const r of ledger){
  const info=classification(r);
  if(isApplicationNegotiated(r)||(info.key&&negotiatedKeys.has(info.key))){negotiatedLedgerRows++;continue;}
  if(!hasPaymentOA(r)){excludedNormalRows++;continue;}
  if(!info.key){normal.push(r);continue;}
  const uncertain=uncertainNegotiated.filter(n=>(!n.name||n.name===info.name)&&(n.cents===undefined||n.cents===info.cents));
  if(uncertain.length){
   const message='谈判表存在可能对应的记录，但供应商名称或申请付款金额不完整，无法确认未命中。\n谈判来源：\n'+uncertain.map(n=>recordLabel(n.record)).join('\n');
   const answer=issue('classification:'+r.id,'classification',message,r,[{value:'normal',label:'核实后确认正常付款，纳入正常付款表'},{value:'keep',label:'暂不分类，本次不纳入正常付款；保留源文件原值'}])[0];
   if(answer!=='normal')continue;
  }
  normal.push(r);
 }
 const blank=c=>!c?.f&&c?.t!=='e'&&!norm(c?.v);
 const usable=c=>c?.t!=='e'&&norm(c?.v)!=='';
 const compatible=(a,b,key)=>{
  const x=a.values[key],y=b.values[key];
  if((!blank(x)&&!usable(x))||(!blank(y)&&!usable(y)))return false;
  if(key==='付款主体')return blank(x)||blank(y)||entity(x.v)===entity(y.v);
  const left=codes(x?.v,key==='单据号【财务应付单】'),right=codes(y?.v,key==='单据号【财务应付单】');
  if((!blank(x)&&!left.length)||(!blank(y)&&!right.length))return false;
  return !identifiersConflict(left,right);
 };
 const normalObjects=normal.map((r,i)=>({record:r,normal:true,index:i,values:{...r.values}}));
 const negotiatedObjects=negotiated.records.map(r=>{
  const target={record:r,normal:false,values:{...r.values}},key=classification(r).key;
  const candidates=key?ledger.filter(l=>classification(l).key===key&&[...SUPPLEMENT_HEADERS,'OA合同编号'].every(h=>compatible(r,l,h))):[];
  if(candidates.length===1)for(const h of SUPPLEMENT_HEADERS){
   const source=candidates[0].values[h],c=targetCols[h];
   const merged=(negotiated.sheet['!merges']||[]).some(m=>r.r>=m.s.r&&r.r<=m.e.r&&c>=m.s.c&&c<=m.e.c&&(r.r!==m.s.r||c!==m.s.c));
   if(blank(target.values[h])&&usable(source)&&!merged)set(target,h,cell(source.v));
  }
  target.record={...r,values:{...target.values}};return target;
 });
 function applyDueDate(target){
  const key=classification(target.record).key;if(!key)return;
  const matches=ledger.filter(r=>classification(r).key===key);if(!matches.length)return;
  const candidates=matches.map(r=>({record:r,date:date(r.values['应付款时间'],r.date1904)}));
  if(candidates.every(c=>c.date.v===''&&!c.date.invalid))return;
  const dates=new Map();for(const c of candidates)if(!c.date.invalid&&c.date.v!=='')dates.set(c.date.v,c.date);
  const column=targetCols['应付款时间'],row=target.record.r;
  const merged=(negotiated.sheet['!merges']||[]).some(m=>row>=m.s.r&&row<=m.e.r&&column>=m.s.c&&column<=m.e.c&&(m.s.r!==m.e.r||m.s.c!==m.e.c));
  if(merged){issue('due-date:'+target.record.id,'due-date','应付款时间位于合并单元格中，保留原值及合并结构，请核对。',target.record);return;}
  let chosen=dates.size===1&&candidates.every(c=>!c.date.invalid&&c.date.v!=='')?[...dates.values()][0]:null;
  if(!chosen){
   const choices=[{value:'keep',label:'保留原应付款时间'},...[...dates].map(([serial,d])=>({value:'date:'+serial,label:'采用 '+X.SSF.format(DATE_FORMAT,d.v)}))];
   const sources=candidates.map(c=>recordLabel(c.record)+'；应付款时间 '+(str(c.record.values['应付款时间']?.w??value(c.record,'应付款时间'))||'空白')).join('\n');
   const answer=issue('due-date:'+target.record.id,'due-date','按供应商名称＋申请付款金额匹配到的应付款时间不一致、为空或无法识别，保留原值，请核对。\n'+sources,target.record,choices)[0];
   if(answer?.startsWith('date:'))chosen=dates.get(Number(answer.slice(5)));
  }
  if(chosen)set(target,'应付款时间',{...chosen,v:chosen.v-(negotiated.date1904?1462:0)});
 }
 negotiatedObjects.forEach(applyDueDate);
 const targets=[...negotiatedObjects,...normalObjects];
 function set(target,key,c){if(!target.normal&&targetCols[key]===undefined)return;target.values[key]=clone(c);if(!target.normal)updateMap.set(X.utils.encode_cell({r:target.record.r,c:targetCols[key]}),clone(c));}
 function applyFinancial(target){
  const src=target.record;
  const a=money(target.values['申请付款金额']),p=money(src.values['已付金额']);
  if(p.blank||blank(src.values['已付金额']))return;
  if(a.cents===undefined||p.cents===undefined){issue('amount:'+target.record.id,'amount','申请金额或已付金额不是可用的分精度数值（含公式结果缺失），保留原付款情况。',target.record);return;}
  const balance=a.cents-p.cents;
  if(balance<0n||a.cents<0n||p.cents<0n){issue('amount:'+target.record.id,'amount','金额为负或已付超过申请金额，不能自动判断结清；保留原付款情况。',target.record);return;}
  const expected=balance===0n?'是':'否',declared=settled(value(src,'是否结清'));
  const retained=settled(target.values['是否结清']?.v);
  if((declared&&declared!==expected)||(!declared&&retained&&retained!==expected)){const chosen=issue('settlement:'+target.record.id,'settlement','台账结清状态与申请金额减已付金额矛盾。',target.record,[{value:'keep',label:'保留原付款情况'},{value:'recalculate',label:'确认以明确金额重算并更新结清状态'}])[0];if(chosen!=='recalculate')return;}
  set(target,'已付金额',amountCell(p.cents));set(target,'是否结清',cell(expected));
  const r=target.index+2;
  set(target,'剩余金额',{...amountCell(balance),f:'H'+r+'-M'+r});
 }
 normalObjects.forEach(applyFinancial);
 const applicationsByKey=new Map();
 for(const record of applications.records){const key=classification(record).key;if(key){if(!applicationsByKey.has(key))applicationsByKey.set(key,[]);applicationsByKey.get(key).push(record);}}
 function matchingApplications(target,preferOA=false){
  if(preferOA){
   const ids=oas(target.record),matched=applications.records.filter(r=>applicationIds(r).some(id=>ids.includes(id)));
   if(matched.length)return matched;
  }
  return applicationsByKey.get(classification(target.record).key)||[];
 }
 function reviewApplication(target,message,candidates){
  if(isApplicationNegotiated(target.record))return;
  const source=target.record.source,amount=money(target.values['申请付款金额']);
  const display=c=>c?.t==='e'?c.w||'Excel错误':c?.f&&(c.v==null||c.v==='')?'公式结果不可用':c?.v??'';
  const status=matchingApplications(target,true).length?'已提交':'未申请';
  reviewRows.push([cell(target.normal?'正常付款':'谈判付款'),cell(display(target.values['供应商名称'])),cell(display(target.values['付款主体'])),amount.cents!==undefined?amountCell(amount.cents):cell(display(target.values['申请付款金额'])),cell(status),cell(message),cell(candidates.map(a=>{const d=date(a.values['申请日期'],a.date1904),label=!d.invalid&&d.v!==''?X.SSF.format('yyyy-mm-dd',d.v):str(value(a,'申请日期'))||'日期为空';return label+'－'+amountText(a.values['申请付款金额'])+'；流程编码 '+(str(value(a,applicationOAField))||'空')+'；'+a.source.file+'/'+a.source.sheet+'!'+a.source.row;}).join('\n')||'无匹配申请'),cell(source.file),cell(source.sheet),cell(source.row)]);
 }
 function applyApplications(target){
  const matched=matchingApplications(target);
  set(target,'是否提交付款申请单',cell(matched.length?'已提交':'未申请'));
  set(target,'提交时间汇总',submissionSummary(matched));
 }
 function applySubmissionDetails(target,matched){
  const dates=[];let invalid=false;
  for(const a of matched){const d=date(a.values['申请日期'],a.date1904);if(d.invalid||d.v===''){invalid=true;continue;}dates.push(X.SSF.format('yyyy-mm-dd',d.v));}
  const historical=submissionDates(target.values['提交付款申请单日期'],target.record.date1904);
  if(invalid||historical===null){reviewApplication(target,'申请或提交日期为空/异常，请核对。',matched);return;}
  set(target,'提交付款申请单日期',cell([...new Set([...dates,...historical])].sort().join('\n')));
 }
 const flowId=r=>norm(value(r,'流程编号')).toUpperCase();
 const payee=r=>r.values['收款单位']?.t==='e'?'':norm(value(r,'收款单位'));
 const flowProblems=r=>OA_FIELDS.filter(f=>{const c=r.values[f];return c?.t==='e'||(c?.f&&(c.v==null||c.v===''));});
 const flowText=(r,f)=>{const c=r.values[f];return c?.t==='e'?c.w||'Excel错误':c?.f&&(c.v==null||c.v==='')?'公式结果不可用':str(c?.v);};
 const flowLabel=r=>`${r.source.file} / ${r.source.sheet} 第 ${r.source.row} 行；流程编号 ${flowText(r,'流程编号')}；收款单位 ${flowText(r,'收款单位')}；付款金额 ${flowText(r,'付款金额')}；`+FLOW_FIELDS.map(f=>f+'='+flowText(r,f)).join('；');
 const flows=new Map();for(const r of flow.records){const id=flowId(r);if(!flows.has(id))flows.set(id,[]);flows.get(id).push(r);}
 // OA exports repeat the process total on invoice detail rows; never sum these amounts.
 for(const [id,rows] of flows)flows.set(id,[...new Map(rows.map(r=>{
  const amount=money(r.values['付款金额']);
  return [JSON.stringify([payee(r),amount.cents?.toString()??['invalid',r.values['付款金额']],...FLOW_FIELDS.map(f=>str(value(r,f))),flowProblems(r)]),r];
 })).values()]);
 const flowRecords=[...flows.values()].flat();let oaConsistent=0,oaMatched=0;
 function applyFlow(target){
  const ids=oas(target.record),supplier=norm(value(target.record,'供应商名称')),amount=money(target.record.values['申请付款金额']);
  const valid=supplier&&target.record.values['供应商名称']?.t!=='e'&&amount.cents!==undefined&&amount.cents>=0n;
  const problems=target.normal?flowProblems:r=>flowProblems(r).filter(f=>f==='流程编号'||FLOW_FIELDS.includes(f));
  const exact=target.normal?(valid?flowRecords.filter(r=>payee(r)===supplier&&money(r.values['付款金额']).cents===amount.cents):[]):
   [...new Map(ids.flatMap(id=>flows.get(id)||[]).map(r=>[JSON.stringify([flowId(r),...FLOW_FIELDS.map(f=>str(value(r,f))),problems(r)]),r])).values()];
  if(!exact.length)return;
  const invalid=exact.filter(r=>problems(r).length);
  const conflictIds=[...new Set(exact.map(flowId))].filter(id=>(target.normal?flows.get(id):exact.filter(r=>flowId(r)===id)).length>1);
  const consistent=rows=>ids.length>0&&ids.length===rows.length&&rows.every(r=>ids.includes(flowId(r)));
  let chosen=exact;
  if(conflictIds.length||invalid.length){
   let message=invalid.length?'OA 字段含 Excel 错误或公式结果不可用，不能写入异常流程信息：'+[...new Set(invalid.flatMap(problems))].join('、')+'。':target.normal?'同一流程编号的收款单位、付款金额或进度信息存在冲突，请核对。':'同一流程编号的进度信息存在冲突，请核对。';
   message+=target.normal?' 流程编号仅以供应商名称与金额完全一致的 OA 记录为准，原 OA 号保留，不作为匹配依据。':' 谈判付款按原对公付款 OA 流程号与 OA 表流程编号匹配，原 OA 号保留。';
   const evidence=conflictIds.length?(target.normal?conflictIds.flatMap(id=>flows.get(id)):exact.filter(r=>conflictIds.includes(flowId(r)))):invalid;
   if(evidence.length)message+='\nOA 来源记录：\n'+evidence.map(flowLabel).join('\n');
   const answer=issue('oa:'+target.record.id,'oa',message,target.record,[{value:'keep',label:'保留原流程编号及流程信息'},...exact.filter(r=>!problems(r).length).map(r=>({value:'match:'+r.id,label:'确认匹配：'+flowLabel(r)}))],true);
   if(!answer.length||answer.includes('keep'))return;
   chosen=exact.filter(r=>answer.includes('match:'+r.id)||(!conflictIds.includes(flowId(r))&&!problems(r).length));
   const chosenIds=new Set(chosen.map(flowId));
   if(chosenIds.size!==chosen.length||exact.some(r=>!chosenIds.has(flowId(r)))){const current=issues[issues.length-1];current.resolved=false;current.message+=' 每个有冲突或异常的流程编号需要选择一个有效版本，其余无冲突流程自动保留。';return;}
  }
  oaMatched++;if(consistent(chosen))oaConsistent++;
  set(target,'流程编号',cell(chosen.map(r=>str(value(r,'流程编号'))).join('\n')));
  for(const f of FLOW_FIELDS){if(flow.map[f]===undefined)continue;set(target,f,cell(chosen.map(r=>(chosen.length>1?flowId(r)+'：':'')+str(value(r,f))).join('\n')));}
 }
 targets.forEach(applyFlow);
 normalObjects.forEach(applyApplications);
 const submissions=new Map(negotiatedObjects.map(t=>[t,{rows:[],ambiguous:[],reasons:[]}]));
 const references=new Map(negotiatedObjects.map(t=>[t,oas(t.record)]));
 for(const application of applications.records){
  const ids=applicationIds(application);
  const possible=negotiatedObjects.filter(t=>ids.some(id=>references.get(t).includes(id)));
  if(possible.length>1){for(const t of possible){const group=submissions.get(t);group.ambiguous.push(application);group.reasons.push('同一流程对应多笔付款，请核对归属。');}continue;}
  if(possible.length===1){
   const target=possible[0],group=submissions.get(target);
   if(ids.length>1&&ids.some(id=>!references.get(target).includes(id))){group.ambiguous.push(application);group.reasons.push('申请含多个流程，请核对金额归属。');}
   else group.rows.push(application);
  }
 }
 for(const target of negotiatedObjects){
  const group=submissions.get(target),matched=group.rows,info=classification(target.record);
  if(!batchLayout)set(target,'付款自然周',plannedPaymentWeeks(negotiated,target.record.r));
  if(str(target.values['是否提交付款申请单']?.v)==='已提交付款申请单')set(target,'是否提交付款申请单',cell('已提交'));
  if(group.ambiguous.length){
   if(blank(target.values['是否提交付款申请单']))set(target,'是否提交付款申请单',cell('待核对'));
   reviewApplication(target,[...new Set(group.reasons)].join('\n'),[...matched,...group.ambiguous]);continue;
  }
  set(target,'提交时间汇总',submissionSummary(matched));
  if(!matched.length){
   set(target,'是否提交付款申请单',cell('未申请'));
   continue;
  }
  const amounts=matched.map(r=>money(r.values['申请付款金额']).cents);
  if(info.cents===undefined||amounts.some(n=>n===undefined||n<0n)){
   set(target,'是否提交付款申请单',cell('待核对'));
   reviewApplication(target,'申请金额为空或异常，请核对。',matched);
  }else{
   const total=amounts.reduce((a,b)=>a+b,0n);
   set(target,'是否提交付款申请单',cell(total>info.cents?'待核对':total===info.cents?'已全部提交':'已部分提交'));
   if(total>info.cents)reviewApplication(target,'申请超额：累计 '+centsText(total)+'，应申请 '+amountText(target.values['申请付款金额'])+'。',matched);
  }
  applySubmissionDetails(target,matched);
 }
 function normalCell(target,key){
  if(key==='付款自然周')return paymentWeek(target.values['应付款时间'],target.record.date1904);
  if(key==='是否对账'){
   return cell(hasPaymentOA(target.record)?'已对账':'未对账');
  }
  const original=target.values[key];if(!original)return empty();
  if(key==='剩余金额'&&original.f&&target.values['已付金额']!==target.record.values['已付金额'])return clone(original);
  if(original.t==='e'||(original.f&&(original.v==null||original.v===''))){issue('cell:'+target.record.id+':'+key,'formula','“'+key+'”公式结果不可用，输出保留错误提示，请在源表重算后上传或确认保留。',target.record);return cell(original.w||'公式结果不可用');}
  if(['登记月份','开票时间','应付款时间'].includes(key)){
   const d=date(original,target.record.date1904,key==='登记月份');if(d.invalid&&key==='开票时间')issue('date:'+target.record.id+':'+key,'date','“'+key+'”日期无法识别，保留原文。',target.record);delete d.invalid;return d;
  }
  if(['开票金额','申请付款金额','合同总金额','已付金额','剩余金额'].includes(key)){
   const amount=money(original);if(amount.cents!==undefined)return amountCell(amount.cents);if(amount.invalid)issue('cell:'+target.record.id+':'+key,'amount','“'+key+'”不是可用金额，保留原文。',target.record);
  }
  return {t:typeof original.v==='number'?'n':'s',v:original.v??'',...(original.z?{z:original.z}:{})};
 }
 const normalRows=normalObjects.map(t=>NORMAL_HEADERS.map(h=>h==='是否谈判付款'?cell('否'):normalCell(t,h)));
 const rows=[];for(let r=negotiated.headerRow+1;r<=negotiated.range.e.r;r++)rows.push(negHeaders.map((_,c)=>clone(updateMap.get(X.utils.encode_cell({r,c}))||negotiated.sheet[X.utils.encode_cell({r,c})])));
 const updates=[...updateMap].map(([address,c])=>({address,cell:c}));
 const columnOrder=negHeaders.map((_,i)=>i),dateColumn=targetCols['提交付款申请单日期'],summaryColumn=targetCols['提交时间汇总'];
 if(dateColumn!==undefined&&summaryColumn!==undefined&&Math.abs(dateColumn-summaryColumn)!==1){columnOrder.splice(summaryColumn,1);columnOrder.splice(columnOrder.indexOf(dateColumn)+1,0,summaryColumn);}
 return {normalHeaders:[...NORMAL_HEADERS],normalRows,reviewHeaders:[...REVIEW_HEADERS],reviewRows,negotiated:{sheetName:selected,headerRow:negotiated.headerRow,sourceRange:negotiated.range,batchLayout,headers:negHeaders,rows,updates,columnOrder},issues,stats:{normalRows:normalRows.length,negotiatedRows:negotiated.records.length,ledgerRows:ledger.length,negotiatedLedgerRows,excludedNormalRows,unclassifiedLedgerRows:ledger.length-negotiatedLedgerRows-normalRows.length-excludedNormalRows,applications:applications.records.length,oaFlows:flows.size,oaConsistent,oaMatched,pending:issues.filter(i=>!i.resolved).length}};
}
return {PEOPLE,NORMAL_HEADERS,ADDED_HEADERS,FLOW_FIELDS,DATE_FORMAT,negotiatedSheets,build,money};
});
