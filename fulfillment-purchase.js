(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'),require('./fulfillment-core.js'),require('./fulfillment-validation.js'),require('./fulfillment-check-core.js'));else root.FulfillmentPurchase=factory(root.XLSX,root.FulfillmentCore,root.FulfillmentValidation,root.FulfillmentCheckCore);})(typeof globalThis!=='undefined'?globalThis:this,function(X,C,V,Check){
'use strict';
const norm=Check.norm,text=v=>String(v??'').trim(),cell=v=>({t:'s',v,z:'@'}),key=values=>JSON.stringify(values.map(text));
function dateRange(filename){
 const matches=[...String(filename).matchAll(/((?:19|20)\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?\s*[-—–~至]\s*(?:((?:19|20)\d{2})[./年-])?(\d{1,2})[./月-](\d{1,2})日?/g)];
 if(matches.length!==1)return null;
 const [,ys,ms,ds,ye,me,de]=matches[0],startYear=Number(ys),endYear=ye?Number(ye):startYear+(Number(me)<Number(ms)?1:0);
 const iso=(y,m,d)=>`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`,start=iso(startYear,ms,ds),end=iso(endYear,me,de),a=C.dateCell({v:start}),b=C.dateCell({v:end});
 if(a.t!=='n'||b.t!=='n'||b.v<a.v)return null;
 const label=(startYear!==endYear?startYear+'.':'')+Number(ms)+'.'+Number(ds)+'-'+(startYear!==endYear?endYear+'.':'')+Number(me)+'.'+Number(de)+'新增采购订单';
 return {start,end,startSerial:a.v,endSerial:b.v,label};
}
const fields={date:['采购日期','下单日期'],order:['单据编号','采购订单号'],supplier:['供应商','供应商全称','供应商名称'],code:['物料编码'],qty:['采购数量'],department:['事业部'],buyer:['采购下单人'],channel:['销售分渠道','销售分渠道部'],short:['供应商简称'],series:['系列','销售系列']};
const receiptFields={...fields,date:['入库日期'],order:['订单单号','采购订单号'],document:['单据编号','入库单号'],qty:['实收数量','入库数量']};
function tables(input,aliases,required){
 const found=[];
 for(const name of input.book.SheetNames){
  const sheet=input.book.Sheets[name],rows=new Map();
  for(const [address,c]of Object.entries(sheet))if(/^[A-Z]+[1-9]\d*$/.test(address)&&(c.v!=null&&c.v!==''||c.f)){const p=X.utils.decode_cell(address);if(!rows.has(p.r))rows.set(p.r,[]);rows.get(p.r).push([p.c,c]);}
  for(const r of [...rows.keys()].filter(r=>r<30).sort((a,b)=>a-b)){
   const map={};for(const [field,names]of Object.entries(aliases)){const columns=rows.get(r).filter(([,c])=>names.includes(norm(c.v)));if(columns.length)map[field]=columns.map(([col])=>col);}
   if(!required.every(f=>map[f]))continue;
   for(const [field,columns]of Object.entries(map))if(columns.length>1)throw Error(`${input.filename} / ${name}：${aliases[field].join('／')}表头重复`);
   found.push({name,sheet,header:r,map:Object.fromEntries(Object.entries(map).map(([k,v])=>[k,v[0]])),rows:[...rows.keys()].filter(n=>n>r).sort((a,b)=>a-b)});break;
  }
 }
 return found;
}
function reader(input){
 const read=V.createReader(input.book);
 return (table,r,field)=>{
  const col=table.map[field];if(col==null)return '';
  let point={r,c:col};const merge=(table.sheet['!merges']||[]).find(m=>m.s.r<=r&&r<=m.e.r&&m.s.c<=col&&col<=m.e.c);
  if(merge){if(field==='qty'&&r>merge.s.r)return 0;point=merge.s;}
  const address=X.utils.encode_cell(point),c=table.sheet[address];if(!c)return '';
  if(c.t==='e'||c.f&&(c.t==='z'||c.v==null||c.v===''))throw Error(`${table.name}!${address}：错误值或公式结果缺失`);
  if(c.f)return read(table.name,address);
  return ['code','order','document'].includes(field)&&typeof c.v==='number'&&c.z&&c.z!=='General'?X.utils.format_cell(c):c.v??'';
 };
}
function parse(input,kind='purchase'){
 const receipt=kind==='receipt',aliases=receipt?receiptFields:fields,quantity=receipt?'实收数量':'采购数量',dateLabel=receipt?'入库日期':'采购日期';
 const selected=tables(input,aliases,['date','order','supplier','code','qty']);
 if(!selected.length)throw Error(receipt?'采购入库单缺少入库日期、订单单号、供应商、物料编码或实收数量表头':'采购订单缺少采购日期、单据编号、供应商、物料编码或采购数量表头');
 const records=[],read=reader(input),date1904=!!input.book.Workbook?.WBProps?.date1904;
 for(const table of selected)for(const r of table.rows){
  const values={},issues=[];
  for(const field of Object.keys(aliases))try{values[field]=read(table,r,field);}catch(e){values[field]='';issues.push(e.message);}
  if(!Object.values(values).some(v=>norm(v))&&!issues.length)continue;
  if(!norm(values.code)&&!norm(values.order)&&[values.supplier,values.date,values.buyer].some(v=>/^(合计|总计|小计)$/.test(norm(v))))continue;
  const source=input.filename+' / '+table.name+'!'+X.utils.encode_cell({r,c:table.map.qty});
  const date=C.dateCell({v:values.date},date1904);if(date.t!=='n'||!norm(values.date))issues.push(dateLabel+'为空或无效');
  let qty=null;try{
   qty=Check.number(values.qty);if(qty<0)throw Error(quantity+'为负数');
   if(typeof values.qty==='string'){
    const a=decimal(values.qty.normalize('NFKC').trim().replace(/(?<=\d),(?=\d{3}(?:\D|$))/g,'')),b=decimal(qty),scale=Math.max(a.scale,b.scale);
    if(a.units*10n**BigInt(scale-a.scale)!==b.units*10n**BigInt(scale-b.scale))throw Error(quantity+'超出精确计算范围');
   }
  }catch(e){issues.push(e.message);qty=null;}
  if(!norm(values.code))issues.push('物料编码为空');if(!norm(values.supplier))issues.push('供应商为空');
  records.push({...Object.fromEntries(Object.entries(values).map(([k,v])=>[k,text(v)])),qty,date:date.t==='n'?Math.floor(date.v):null,source,issues});
 }
 return {filename:input.filename,records};
}
function enrich({filename,comparison=[],current,previous,purchase,directory,kind='purchase'}){
 const receipt=kind==='receipt',range=dateRange(filename),label=receipt?'采购入库单数量':'新增采购订单';
 if(receipt&&range)range.label=range.label.replace(/新增采购订单$/,label);
 if(!range&&!purchase)return {columns:C.SERIES_COLUMNS,rows:comparison,issues:[],stats:{},range};
 const columns=[...C.SERIES_COLUMNS.slice(0,8),range?.label||label+'（日期待确认）',...C.SERIES_COLUMNS.slice(8)];
 const issues=[],stats={periodRows:0,matchedRows:0,unmatchedRows:0,uncertainRows:0},totals=new Map(),uncertain=new Set();
 const finish=placeholder=>({columns,rows:comparison.map(row=>[...row.slice(0,8),placeholder?cell(placeholder):uncertain.has(key(row.slice(0,5).map(c=>c.v)))?cell('待核对'):numeric(totals.get(key(row.slice(0,5).map(c=>c.v)))),...row.slice(8)]),issues,stats,range});
 if(!range){issues.push({source:filename,reason:'文件名中没有唯一有效日期区间，请使用例如2026.11.13-11.19'});return finish('日期待确认');}
 if(!purchase)return finish(receipt?'未上传采购入库单':'未上传采购订单');
 if(purchase.error){issues.push({source:purchase.filename,reason:purchase.error});return finish(receipt?'入库文件待核对':'采购文件待核对');}
 let catalog,owners=new Map();
 try{
  catalog=Check.supplierDirectory(directory);
  if(directory){
   const aliases={supplier:fields.supplier,code:['正确物料编码','物料编码'],buyer:fields.buyer},read=reader(directory);
   for(const table of tables(directory,aliases,['supplier','code','buyer']))for(const r of table.rows){
    const supplier=norm(read(table,r,'supplier')),code=norm(read(table,r,'code')),buyer=text(read(table,r,'buyer'));if(!supplier||!code||!buyer)continue;
    const k=key([supplier,code]);if(!owners.has(k))owners.set(k,new Set());owners.get(k).add(buyer);
   }
  }
 }catch(e){issues.push({source:directory?.filename||'供应商分工明细',reason:e.message});return finish('供应商分工待核对');}
 const identity=v=>catalog.identity(v)||norm(v),sameSupplier=(a,b)=>identity(a)===identity(b);
 const collect=snapshot=>(snapshot?.rows||[]).map((row,i)=>({labels:[0,1,2,4,6].map(c=>text(row[c]?.v)),code:norm(row[7]?.v),orders:Check.orders(row[22]?.v).ids,source:snapshot.sources?.[i]}));
 const now=collect(current),before=collect(previous),all=[...now,...before];
 const department=(value,labels)=>{const v=norm(value),division=norm(labels[1]),channel=norm(labels[2]);return v&&v!=='NA'&&(v===division||v===channel||channel===division&&v.startsWith(division+'*'));};
 for(const record of purchase.records){
  if(record.date!==null&&(record.date<range.startSerial||record.date>range.endSerial))continue;
  stats.periodRows++;
  const supplierOwners=owners.get(key([identity(record.supplier),norm(record.code)]));
  const consistent=e=>sameSupplier(record.supplier,e.labels[3])&&(!record.buyer||norm(record.buyer)===norm(e.labels[0]))&&(!supplierOwners||supplierOwners.has(e.labels[0]))&&(!record.series||norm(record.series)===norm(e.labels[4]))&&(!record.channel||norm(record.channel)===norm(e.labels[2]));
  const ids=Check.orders(record.order).ids;
  const exact=scope=>scope.filter(e=>e.code===norm(record.code)&&ids.length===1&&e.orders.includes(ids[0]));
  let matches=exact(now);if(!matches.length)matches=exact(before);
  let reason='';
  if(matches.length){
   if(matches.some(e=>!sameSupplier(record.supplier,e.labels[3])||record.department&&norm(record.department)!=='NA'&&!department(record.department,e.labels)||record.buyer&&norm(record.buyer)!==norm(e.labels[0])||record.channel&&norm(record.channel)!==norm(e.labels[2])||record.series&&norm(record.series)!==norm(e.labels[4])||supplierOwners&&!supplierOwners.has(e.labels[0]))){reason='同单号和物料的供应商、采购下单人、事业部／渠道或系列信息冲突或对应关系未确认，请核实分工明细';}
  }else{
   const complete=record.buyer&&record.channel&&record.series;
   matches=all.filter(e=>(!receipt&&complete||e.code===norm(record.code))&&consistent(e)&&(receipt&&!record.department||department(record.department,e.labels)));
   if(supplierOwners?.size>1)reason='供应商分工明细对应多个采购下单人，归属待核对';
  }
  const candidates=[...new Set(matches.map(e=>key(e.labels)))];
  if(candidates.length!==1||record.issues.length||reason){
   for(const k of candidates)uncertain.add(k);
   if(candidates.length)stats.uncertainRows++;else stats.unmatchedRows++;
   issues.push({source:record.source,order:record.order,...(receipt?{document:record.document}:{}),code:record.code,supplier:record.supplier,qty:record.qty,reason:[...record.issues,reason||(!candidates.length?'未找到供应商、物料和事业部／渠道一致的系列分组':'对应多个系列分组，未分配数量')].filter(Boolean).join('；')});continue;
  }
  const k=candidates[0];
  try{add(totals,k,record.qty);numeric(totals.get(k));stats.matchedRows++;}catch(e){uncertain.add(k);stats.uncertainRows++;issues.push({source:record.source,...(receipt?{document:record.document,order:record.order,supplier:record.supplier,qty:record.qty}:{}),code:record.code,reason:e.message});}
 }
 return finish();
}
function decimal(value){
 const text=typeof value==='number'?value.toLocaleString('en-US',{useGrouping:false,maximumSignificantDigits:21}):value,parts=text.split('.');
 return {scale:(parts[1]||'').length,units:BigInt(parts.join(''))};
}
function add(totals,key,value){
 const {scale,units}=decimal(value);
 const old=totals.get(key)||{units:0n,scale:0},next=Math.max(scale,old.scale);
 totals.set(key,{units:old.units*10n**BigInt(next-old.scale)+units*10n**BigInt(next-scale),scale:next});
}
function numeric(value){
 let {units=0n,scale=0}=value||{};while(scale&&units%10n===0n){units/=10n;scale--;}
 if(units>BigInt(Number.MAX_SAFE_INTEGER))throw Error('数量合计超出精确计算范围');
 return {t:'n',v:Number(units.toString()+'e-'+scale),z:scale>250?'0.0000000000000000E+00':scale?'0.'+'0'.repeat(scale):'0'};
}
return {dateRange,parse,enrich};
});
