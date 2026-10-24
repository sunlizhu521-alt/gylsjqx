(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'));else root.FundsCore=factory(root.XLSX);})(typeof globalThis!=='undefined'?globalThis:this,function(X){
'use strict';
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const NORMAL_HEADERS=['登记月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单','当前节点','未操作者','流程状态','申请人'];
const ADDED_HEADERS=['提交付款申请单日期','是否提交付款申请单','当前节点','未操作者','流程状态','申请人'];
const FLOW_FIELDS=['当前节点','未操作者','流程状态','申请人'];
const DATE_FORMAT='yyyy"年"m"月"d"日"';
const norm=v=>String(v??'').normalize('NFKC').replace(/\s+/g,'').trim();
const str=v=>String(v??'').trim();
const empty=()=>({t:'s',v:''});
const cell=v=>({t:typeof v==='number'?'n':'s',v:v??''});
const clone=c=>c?{...c}:empty();
const aliases={月份:'登记月份',所属月份:'登记月份',供应商:'供应商名称',供应商全称:'供应商名称',对公付款OA流程号:'对公付款0A流程号',付款OA号:'对公付款0A流程号',应付时间:'应付款时间',已付款金额:'已付金额',剩余付款金额:'剩余金额',合同号:'OA合同编号',合同金额:'合同总金额',本次申请付款金额:'申请付款金额',付款事由:'摘要'};
const canonical=v=>aliases[norm(v)]||norm(v);
const relevant=new Set([...NORMAL_HEADERS,...ADDED_HEADERS,'申请日期','备注','对接人','创建人','流程编号']);
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
function codes(v,ap=false){const s=str(v).toUpperCase();return [...new Set(s.match(ap?/AP\d+/g:/[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)*\d[A-Z0-9-]*/g)||[])];}
const shared=(a,b)=>a.some(v=>b.includes(v));
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
function submissionDates(c,date1904){
 if(c?.t==='e'||(c?.f&&(c.v==null||c.v==='')))return null;
 if(!c||c.v==null||c.v==='')return [];
 const parts=typeof c.v==='string'?c.v.split(/[、,，;；\n]+/).map(str).filter(Boolean):[c.v];
 const values=[];
 for(const part of parts){const d=date({v:part},date1904);if(d.invalid||d.v==='')return null;values.push(X.SSF.format('yyyy-mm-dd',d.v));}
 return values;
}
function table(input,sheetName,required){
 const sheet=input.book.Sheets[sheetName];if(!sheet)throw Error(input.name+'：缺少工作表 '+sheetName);
 const range=X.utils.decode_range(sheet['!ref']||'A1');if((range.e.r-range.s.r+1)*(range.e.c-range.s.c+1)>2000000)throw Error(sheetName+'：数据范围过大');
 let headerRow=-1,map,headers;
 for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++){
  const h={},hs=[];for(let c=0;c<=range.e.c;c++){const cc=sheet[X.utils.encode_cell({r,c})],key=canonical(cc?.v);hs.push(str(cc?.w??cc?.v));if((required.includes('流程编号')?['流程编号',...FLOW_FIELDS].includes(key):relevant.has(key)))(h[key]??=[]).push(c);}
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
  else{if(/^(?:合计|总计|小计)$/.test(supplier))continue;if(!supplier&&!oa&&(amount==null||amount===''))continue;}
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
const identity=(a,b)=>norm(value(a,'供应商名称'))!==''&&norm(value(a,'供应商名称'))===norm(value(b,'供应商名称'))&&entity(value(a,'付款主体'))!==''&&entity(value(a,'付款主体'))===entity(value(b,'付款主体'));
const equalAmount=(a,b)=>{const x=money(a.values['申请付款金额']),y=money(b.values['申请付款金额']);return x.cents!==undefined&&x.cents===y.cents;};
const aps=r=>codes(value(r,'单据号【财务应付单】')||value(r,'备注'),true);
const oas=r=>codes(value(r,'对公付款0A流程号'));
const contracts=r=>codes(value(r,'OA合同编号'));
const reasonEqual=(a,b)=>norm(value(a,'摘要'))!==''&&norm(value(a,'摘要'))===norm(value(b,'摘要'));
const recordLabel=r=>`${r.source.file} / ${r.source.sheet} 第 ${r.source.row} 行；${str(value(r,'供应商名称'))}；${str(value(r,'付款主体'))}；申请 ${str(value(r,'申请付款金额'))}；${str(value(r,'对公付款0A流程号')||value(r,'OA合同编号'))}；${str(value(r,'摘要'))}；财务单 ${str(value(r,'单据号【财务应付单】')||value(r,'备注'))}；合同 ${str(value(r,'OA合同编号'))}`;
function negotiatedSheets(book){return (book?.SheetNames||[]).filter(n=>n.includes('谈判付款'));}
function build(inputs,decisions={}){
 for(const role of ['ledger','applications','negotiated','oa'])if(!inputs[role]?.book)throw Error('请上传全部四份文件');
 const candidates=negotiatedSheets(inputs.negotiated.book),selected=inputs.negotiated.sheetName||(candidates.length===1?candidates[0]:'');
 if(!selected||!candidates.includes(selected))throw Error(candidates.length?'请选择包含“谈判付款”的工作表':'未找到名称包含“谈判付款”的工作表');
 const negotiated=table(inputs.negotiated,selected,['供应商名称','申请付款金额']);
 const applications=oneTable(inputs.applications,['供应商名称','申请付款金额','申请日期']);
 const flow=oneTable(inputs.oa,['流程编号']);
 const ledgerTables=PEOPLE.map(person=>{
  const names=inputs.ledger.book.SheetNames.filter(n=>new RegExp('^对账明细[-－—]?'+person+'(?:\\d{4}|\\d{8})?$').test(norm(n)));
  if(names.length!==1)throw Error('资金台账中'+person+'的对账明细'+(names.length?'存在多个版本':'缺失'));
  return table(inputs.ledger,names[0],['供应商名称','申请付款金额']);
 });
 const ledger=ledgerTables.flatMap(t=>t.records),issues=[],updateMap=new Map();
 function issue(id,kind,message,rec,choices=[{value:'keep',label:'保留原值'}],multiple=false){
  const requested=Array.isArray(decisions[id])?decisions[id]:[],allowed=new Set(choices.map(c=>c.value));
  const valid=requested.length>0&&requested.every(v=>allowed.has(v))&&new Set(requested).size===requested.length&&(multiple||requested.length===1)&&(!(requested.includes('keep')||requested.includes('not-submitted'))||requested.length===1);
  const answer=valid?requested:[];
  issues.push({id,kind,message,source:rec.source,targetLabel:recordLabel(rec),choices,multiple,resolved:valid,selected:answer});return answer;
 }
 const headerRecord=t=>({source:{file:t.name,sheet:t.sheetName,row:t.headerRow+1},values:{}});
 for(const t of ledgerTables){const absent=NORMAL_HEADERS.slice(0,17).filter(k=>t.map[k]===undefined);if(absent.length)issue('headers:'+t.sheetName,'missing-header','缺少字段，输出保留空白：'+absent.join('、'),headerRecord(t));}
 for(const [t,fields] of [[applications,['付款主体','OA合同编号','摘要']],[flow,FLOW_FIELDS],[negotiated,['付款主体','对公付款0A流程号','已付金额','是否结清','剩余金额']]]){
  const absent=fields.filter(k=>t.map[k]===undefined);if(absent.length)issue('headers:'+t.sheetName,'missing-header','缺少字段，无法据此自动匹配或更新：'+absent.join('、'),headerRecord(t));
 }
 const targetCols={...negotiated.map},negHeaders=[...negotiated.headers];
 for(const h of ['已付金额','是否结清','剩余金额',...ADDED_HEADERS])if(targetCols[h]===undefined){const label=h==='剩余金额'?'剩余付款金额':h;targetCols[h]=negHeaders.length;negHeaders.push(label);updateMap.set(X.utils.encode_cell({r:negotiated.headerRow,c:targetCols[h]}),cell(label));}
 const normal=[];
 for(const r of ledger){
  const flag=norm(value(r,'是否谈判付款'));
  let include=flag==='否';
  if(flag!=='是'&&flag!=='否')include=issue('classification:'+r.id,'classification','是否谈判付款为空或不是“是/否”，请确认是否纳入正常付款。',r,[{value:'normal',label:'纳入正常付款（否）'},{value:'keep',label:'保留原分类，本次不纳入正常付款'}])[0]==='normal';
  if(include)normal.push(r);
 }
 const apConflict=(a,b)=>identifiersConflict(aps(a),aps(b));
 const oaConflict=(a,b)=>identifiersConflict(oas(a),oas(b));
 const contractConflict=(a,b)=>identifiersConflict(contracts(a),contracts(b));
 const strongPayment=(a,b)=>identity(a,b)&&equalAmount(a,b)&&!apConflict(a,b)&&!oaConflict(a,b)&&!contractConflict(a,b)&&(shared(oas(a),oas(b))||shared(aps(a),aps(b)));
 const paymentCandidates=new Map(negotiated.records.map(r=>[r.id,ledger.filter(l=>strongPayment(r,l))]));
 const claims=new Map();for(const [target,rows] of paymentCandidates)for(const r of rows){if(!claims.has(r.id))claims.set(r.id,[]);claims.get(r.id).push(target);}
 const used=new Set();
 function selectPayment(target){
  const exact=paymentCandidates.get(target.id)||[];
  if(exact.length===1&&claims.get(exact[0].id)?.length===1&&!normal.includes(exact[0])&&!used.has(exact[0].id)){used.add(exact[0].id);return exact[0];}
  const possible=ledger.filter(l=>norm(value(l,'供应商名称'))===norm(value(target,'供应商名称'))||shared(oas(l),oas(target))||shared(aps(l),aps(target)));
  const choices=[{value:'keep',label:'保留谈判表原付款情况'},...possible.filter(l=>!normal.includes(l)&&!used.has(l.id)).map(l=>({value:'match:'+l.id,label:recordLabel(l)}))];
  const answer=issue('payment-match:'+target.id,'payment-match',exact.some(l=>normal.includes(l))?'对应资金来源已在正常付款输出，不能重复分配已付金额；请保留原值或选择其他未使用来源。':exact.length?'存在重复来源，请核对，避免重复分配已付金额。':'没有唯一且编号、主体、金额一致的付款记录，请核对来源。',target,choices)[0];
  if(answer?.startsWith('match:')){const r=possible.find(l=>'match:'+l.id===answer);if(r&&!normal.includes(r)&&!used.has(r.id)){used.add(r.id);return r;}}
  return null;
 }
 const normalObjects=normal.map((r,i)=>({record:r,normal:true,index:i,values:{...r.values}}));
 const negotiatedObjects=negotiated.records.map(r=>({record:r,normal:false,values:{...r.values},payment:selectPayment(r)}));
 const targets=[...negotiatedObjects,...normalObjects];
 for(const target of targets)if(!norm(value(target.record,'供应商名称'))||!norm(value(target.record,'付款主体')))issue('identity:'+target.record.id,'identity','供应商或付款主体为空，保留源值并核对，不按空值匹配。',target.record);
 function set(target,key,c){target.values[key]=clone(c);if(!target.normal)updateMap.set(X.utils.encode_cell({r:target.record.r,c:targetCols[key]}),clone(c));}
 function applyFinancial(target){
  const src=target.normal?target.record:target.payment;if(!src)return;
  const a=money(target.values['申请付款金额']),p=money(src.values['已付金额']);
  if(a.cents===undefined||p.cents===undefined){issue('amount:'+target.record.id,'amount',p.blank?'已付金额为空，不能当作零；保留原付款情况。':'申请金额或已付金额不是可用的分精度数值（含公式结果缺失），保留原付款情况。',target.record);return;}
  const balance=a.cents-p.cents;
  if(balance<0n||a.cents<0n||p.cents<0n){issue('amount:'+target.record.id,'amount','金额为负或已付超过申请金额，不能自动判断结清；保留原付款情况。',target.record);return;}
  const expected=balance===0n?'是':'否',declared=settled(value(src,'是否结清'));
  const retained=settled(target.values['是否结清']?.v);
  if((declared&&declared!==expected)||(!declared&&retained&&retained!==expected)){const chosen=issue('settlement:'+target.record.id,'settlement','台账结清状态与申请金额减已付金额矛盾。',target.record,[{value:'keep',label:'保留原付款情况'},{value:'recalculate',label:'确认以明确金额重算并更新结清状态'}])[0];if(chosen!=='recalculate')return;}
  set(target,'已付金额',amountCell(p.cents));set(target,'是否结清',cell(expected));
  const r=target.normal?target.index+2:target.record.r+1,ac=target.normal?'H':X.utils.encode_col(targetCols['申请付款金额']),pc=target.normal?'M':X.utils.encode_col(targetCols['已付金额']);
  set(target,'剩余金额',{...amountCell(balance),f:ac+r+'-'+pc+r});
 }
 targets.forEach(applyFinancial);
 const appStrong=(t,a)=>{const r=t.record,bridge=t.payment||r;return identity(bridge,a)&&equalAmount(bridge,a)&&!apConflict(bridge,a)&&!oaConflict(bridge,a)&&!contractConflict(bridge,a)&&(shared(aps(bridge),aps(a))||(shared(contracts(bridge),contracts(a))&&reasonEqual(bridge,a)));};
 const appClaims=new Map();for(const target of targets)for(const a of applications.records)if(appStrong(target,a)){if(!appClaims.has(a.id))appClaims.set(a.id,[]);appClaims.get(a.id).push(target);}
 function applyApplications(target){
  const bridge=target.payment||target.record;
  const exact=applications.records.filter(a=>appStrong(target,a));
  let matched=exact.length&&exact.every(a=>appClaims.get(a.id).length===1)?exact:[];
  if(!matched.length){
   const possible=applications.records.filter(a=>norm(value(a,'供应商名称'))===norm(value(bridge,'供应商名称'))||shared(aps(bridge),aps(a))||shared(contracts(bridge),contracts(a)));
   const answer=issue('application:'+target.record.id,'application','未找到唯一完整匹配的付款申请；可能未提交、分批申请或对应多笔台账，请核对。',target.record,[{value:'keep',label:'保留原状态，空白标记待核对'},{value:'not-submitted',label:'确认未提交，标记为否'},...possible.map(a=>({value:'match:'+a.id,label:recordLabel(a)+'；申请日期 '+str(value(a,'申请日期'))}))],true);
   if(answer[0]==='not-submitted'){set(target,'是否提交付款申请单',cell('否'));if(!target.normal)set(target,'提交付款申请单日期',empty());return;}
   matched=possible.filter(a=>answer.includes('match:'+a.id));
   if(!matched.length){if(!str(target.values['是否提交付款申请单']?.v))set(target,'是否提交付款申请单',cell('待核对'));return;}
  }
  const dates=[];let invalid=false;
  for(const a of matched){const d=date(a.values['申请日期'],a.date1904);if(d.invalid||d.v===''){invalid=true;continue;}dates.push(X.SSF.format('yyyy-mm-dd',d.v));}
  if(invalid){issue('application-date:'+target.record.id,'application-date','已匹配申请的日期为空或无法识别；保留原提交状态和日期。',target.record);return;}
  if(!target.normal){
   const historical=submissionDates(target.values['提交付款申请单日期'],target.record.date1904);
   if(historical===null){issue('application-history:'+target.record.id,'application-history','原提交付款申请单日期无法完整识别，保留原日期和提交状态，不以本次申请覆盖历史。',target.record);return;}
   dates.push(...historical);
   set(target,'提交付款申请单日期',cell([...new Set(dates)].sort().join('\n')));
  }
  set(target,'是否提交付款申请单',cell('是'));
 }
 targets.forEach(applyApplications);
 const flows=new Map();for(const r of flow.records){const key=norm(value(r,'流程编号')).toUpperCase();if(!flows.has(key))flows.set(key,[]);flows.get(key).push(r);}
 function applyFlow(target){
  const ids=oas(target.record);if(!ids.length){issue('oa:'+target.record.id,'oa','付款 OA 号为空或不是有效流程编号，保留原流程信息。',target.record);return;}
  const collected=[];
  for(const id of ids){
   const rows=flows.get(id)||[],unique=[...new Map(rows.map(r=>[JSON.stringify(FLOW_FIELDS.map(f=>str(value(r,f)))),r])).values()];
   if(unique.length===1){collected.push({id,record:unique[0]});continue;}
   const answer=issue('oa:'+target.record.id+':'+id,'oa',unique.length?'同一流程编号的进度信息不一致，请选择对应记录或保留原值。':'OA 流程表未找到 '+id+'，保留原流程信息。',target.record,[{value:'keep',label:'保留原流程信息'},...unique.map(r=>({value:'match:'+r.id,label:`${r.source.sheet} 第 ${r.source.row} 行：`+FLOW_FIELDS.map(f=>f+'='+str(value(r,f))).join('；')}))])[0];
   const chosen=unique.find(r=>'match:'+r.id===answer);if(chosen)collected.push({id,record:chosen});else return;
  }
  for(const f of FLOW_FIELDS){if(flow.map[f]===undefined)continue;set(target,f,cell(collected.map(x=>(ids.length>1?x.id+'：':'')+str(value(x.record,f))).join('\n')));}
 }
 targets.forEach(applyFlow);
 function normalCell(target,key){
  const original=target.values[key];if(!original)return empty();
  if(key==='剩余金额'&&original.f&&target.values['已付金额']!==target.record.values['已付金额'])return clone(original);
  if(original.t==='e'||(original.f&&(original.v==null||original.v===''))){issue('cell:'+target.record.id+':'+key,'formula','“'+key+'”公式结果不可用，输出保留错误提示，请在源表重算后上传或确认保留。',target.record);return cell(original.w||'公式结果不可用');}
  if(['登记月份','开票时间','应付款时间'].includes(key)){
   const d=date(original,target.record.date1904,key==='登记月份');if(d.invalid)issue('date:'+target.record.id+':'+key,'date','“'+key+'”日期无法识别，保留原文。',target.record);delete d.invalid;return d;
  }
  if(['开票金额','申请付款金额','合同总金额','已付金额','剩余金额'].includes(key)){
   const amount=money(original);if(amount.cents!==undefined)return amountCell(amount.cents);if(amount.invalid)issue('cell:'+target.record.id+':'+key,'amount','“'+key+'”不是可用金额，保留原文。',target.record);
  }
  return {t:typeof original.v==='number'?'n':'s',v:original.v??'',...(original.z?{z:original.z}:{})};
 }
 const normalRows=normalObjects.map(t=>NORMAL_HEADERS.map(h=>h==='是否谈判付款'?cell('否'):normalCell(t,h)));
 const rows=[];for(let r=negotiated.headerRow+1;r<=negotiated.range.e.r;r++)rows.push(negHeaders.map((_,c)=>clone(updateMap.get(X.utils.encode_cell({r,c}))||negotiated.sheet[X.utils.encode_cell({r,c})])));
 const updates=[...updateMap].map(([address,c])=>({address,cell:c}));
 return {normalHeaders:[...NORMAL_HEADERS],normalRows,negotiated:{sheetName:selected,headerRow:negotiated.headerRow,headers:negHeaders,rows,updates},issues,stats:{normalRows:normalRows.length,negotiatedRows:negotiated.records.length,ledgerRows:ledger.length,applications:applications.records.length,oaFlows:flows.size,pending:issues.filter(i=>!i.resolved).length}};
}
return {PEOPLE,NORMAL_HEADERS,ADDED_HEADERS,FLOW_FIELDS,DATE_FORMAT,negotiatedSheets,build,money};
});
