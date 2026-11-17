(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./vendor/xlsx.full.min.js'),require('./fulfillment-validation.js'));else root.FulfillmentCheckCore=factory(root.XLSX,root.FulfillmentValidation);})(typeof globalThis!=='undefined'?globalThis:this,function(X,V){
'use strict';
const PEOPLE=['徐英','操思敏','邢芳芳','李奇'];
const SUMMARY=['供应商','采购剩余入库数量有效值合计','履约未交付数量有效值合计','已匹配差额','一致项数','数量差异项数','待核对项数','仅采购项数','仅履约项数','已关闭异常项数'];
const DETAIL=['核对结果','供应商','履约供应商原值','物料编码','SKU','物料名称','采购单据编号','履约采购订单号','采购订单组成','采购剩余入库数量','履约未交付数量','差额','关闭状态','采购来源','履约来源','匹配依据','需要处理'];
const norm=value=>String(value??'').normalize('NFKC').replace(/\s/g,'').trim();
const key=(...parts)=>JSON.stringify(parts);
const unique=values=>[...new Set(values.filter(v=>v!==''&&v!=null))];
function sum(values){let n=0,c=0;for(const v of values){const y=v-c,t=n+y;c=(t-n)-y;n=t;}return Number(n.toPrecision(15));}
function orders(value){
 const s=norm(value).toUpperCase();if(!s||['/','无','未下单','未提交'].includes(s))return {ids:[],error:'采购订单号为空或未填写有效编号'};
 const ids=s.match(/CGDD\d+/g);
 if(ids){const rest=s.replace(/CGDD\d+/g,'').replace(/[+、,;；/|&\-()]/g,'');return {ids:unique(ids).sort(),error:rest?'订单号中含无法识别的内容：'+rest:''};}
 const parts=s.split(/[+、,;；/|&]/);if(parts.every(p=>/^[A-Z0-9][A-Z0-9_-]*$/.test(p)))return {ids:unique(parts).sort(),error:''};
 return {ids:[],error:'采购订单号无法识别：'+value};
}
function number(value){
 if(value==null||String(value).trim()==='')throw Error('数量为空，未按零处理');
 if(typeof value==='number'){if(Number.isFinite(value)&&Math.abs(value)<=Number.MAX_SAFE_INTEGER)return value;throw Error('数量超出有效范围');}
 const s=String(value).normalize('NFKC').trim().replace(/(?<=\d),(?=\d{3}(?:\D|$))/g,'');
 if(/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(s))return number(Number(s));
 throw Error('数量不是明确数值：'+value);
}
const aliases={supplier:['供应商','供应商名称','供应商简称'],code:['物料编码'],order:['采购订单号','单据编号'],qty:['剩余入库数量','未交付数量'],closed:['关闭状态'],sku:['SKU'],name:['物料名称']};
function parse(input,kind){
 if(!input?.book)throw Error('请上传'+(kind==='po'?'采购订单':'订单履约表'));
 const book=input.book,read=V.createReader(book),rows=[],problems=[];
 let selected=book.SheetNames;
 if(kind==='fulfillment')selected=PEOPLE.map(person=>{
  const matches=book.SheetNames.filter(n=>new RegExp('^'+person+'(?:[0-9]{4}|[0-9]{8})?$').test(norm(n)));
  if(matches.length!==1)throw Error(input.filename+'：'+person+(matches.length?'匹配到多张人员表，请保留唯一来源':'工作表缺失'));
  return matches[0];
 });
 let found=0;
 for(const name of selected){
  const sheet=book.Sheets[name],cells=Object.entries(sheet).filter(([a])=>/^[A-Z]+[1-9]\d*$/.test(a));
  const contents=new Map();for(const [a,c] of cells){const r=X.utils.decode_cell(a).r;if(!contents.has(r))contents.set(r,[]);contents.get(r).push(norm(c.v));}
  const headerRows=new Map();for(const [a,c] of cells){const p=X.utils.decode_cell(a);if(p.r>29||c.v==null)continue;if(!headerRows.has(p.r))headerRows.set(p.r,new Map());headerRows.get(p.r).set(p.c,norm(c.v));}
  let header,fields;
  for(const [r,columns] of [...headerRows].sort((a,b)=>a[0]-b[0])){
   const f={};for(const [field,names] of Object.entries(aliases)){
    const allowed=field==='qty'?[kind==='po'?'剩余入库数量':'未交付数量']:names;
    const matches=[...columns].filter(([,v])=>allowed.includes(v));if(matches.length===1)f[field]=matches[0][0];else if(matches.length>1)f[field]=-1;
   }
   if(f.code!==undefined&&f.qty!==undefined){header=r;fields=f;break;}
  }
  if(header===undefined){if(kind==='fulfillment')throw Error(input.filename+' / '+name+'：找不到物料编码、未交付数量表头');continue;}
  found++;
  for(const field of ['supplier','code','qty',...(kind==='po'?['order','closed']:[])])if(fields[field]==null||fields[field]<0)throw Error(input.filename+' / '+name+'：'+aliases[field].join('／')+'表头缺失或重复');
  const rowNumbers=unique(cells.filter(([,c])=>c.v!=null&&c.v!==''||c.f).map(([a])=>X.utils.decode_cell(a).r)).filter(r=>r>header).sort((a,b)=>a-b);
  for(const r of rowNumbers){
   const issues=[],address=f=>fields[f]>=0?X.utils.encode_cell({r,c:fields[f]}):'',get=f=>{
    const a=address(f);if(!a)return '';
    let cell=sheet[a];const merge=(sheet['!merges']||[]).find(m=>m.s.c<=fields[f]&&fields[f]<=m.e.c&&m.s.r<=r&&r<=m.e.r);
    if(merge){if(f==='qty'&&r>merge.s.r)return 0;cell=sheet[X.utils.encode_cell(merge.s)];}
    if(!cell)return '';
    try{if(cell.f||cell.F||cell.D)return read(name,merge?X.utils.encode_cell(merge.s):a);if(cell.t==='e')throw Error('错误值 '+(cell.w||cell.v));return ['code','order'].includes(f)&&typeof cell.v==='number'&&cell.z&&cell.z!=='General'?X.utils.format_cell(cell):cell.v??'';}
    catch(e){issues.push(a+'：'+e.message);return '';}
   };
   const rawSupplier=String(get('supplier')).trim(),code=norm(get('code')),rawOrder=String(get('order')).trim(),rawQty=get('qty');
   const line=contents.get(r)||[];
   if(!code&&!rawOrder&&(!rawSupplier||/^(合计|总计|小计)$/.test(rawSupplier))&&line.some(s=>/^(合计|总计|小计)$/.test(s)))continue;
   if(!rawSupplier&&!code&&!rawOrder&&(rawQty===''||rawQty==null))continue;
   if(!rawSupplier)issues.push('供应商为空');if(!code)issues.push('物料编码为空');
   const parsed=orders(rawOrder);if(parsed.error)issues.push(parsed.error);
   if(kind==='po'&&parsed.ids.length!==1)issues.push('采购来源单据编号不是一个明确编号');
   let qty=null;try{qty=number(rawQty);if(qty<0)issues.push('数量为负数，请核实源记录');}catch(e){issues.push((address('qty')||'数量列')+'：'+e.message);}
   const closed=kind==='po'?norm(get('closed')):'';
   if(kind==='po'&&!['未关闭','已关闭'].includes(closed))issues.push('关闭状态为空或不是未关闭／已关闭');
   rows.push({id:kind+':'+name+':'+(r+1),kind,sheet:name,row:r+1,rawSupplier,supplier:norm(rawSupplier),code,rawOrder,ids:parsed.ids,qty,closed,sku:String(get('sku')),name:String(get('name')),issues,
    source:input.filename+' / '+name+' 第 '+(r+1)+' 行（'+address('qty')+'）'});
  }
 }
 if(!found)throw Error(input.filename+'：未找到含物料编码、剩余入库数量的采购明细表');
 return {rows,problems};
}
function build({purchase,fulfillment,decisions={}}){
 const p=parse(purchase,'po').rows,f=parse(fulfillment,'fulfillment').rows;
 const suppliers=new Map();for(const row of p)if(row.supplier&&!suppliers.has(row.supplier))suppliers.set(row.supplier,row.rawSupplier);
 const byOrderCode=new Map();for(const row of p)for(const id of row.ids){const k=key(id,row.code);if(!byOrderCode.has(k))byOrderCode.set(k,[]);byOrderCode.get(k).push(row);}
 const aliasesMap=new Map(),choices=[];
 for(const supplier of unique(f.map(r=>r.supplier))){
  if(!supplier)continue;
  const source=f.filter(r=>r.supplier===supplier),evidence=unique(source.flatMap(r=>r.ids.flatMap(id=>(byOrderCode.get(key(id,r.code))||[]).map(p=>p.supplier))));
  const contains=[...suppliers.keys()].filter(n=>n.includes(supplier));
  const candidates=unique([...evidence,...contains]);let canonical='',basis='';
  if(decisions[supplier]&&suppliers.has(decisions[supplier])){canonical=decisions[supplier];basis='人工确认供应商';}
  else if(suppliers.has(supplier)){canonical=supplier;basis='供应商名称一致';}
  else if(evidence.length===1){canonical=evidence[0];basis='单据编号＋物料编码唯一对应供应商';}
  else if(!evidence.length&&contains.length===1){canonical=contains[0];basis='简称唯一包含于采购供应商全称';}
  if(!canonical)choices.push({supplier,label:source[0].rawSupplier,candidates:candidates.map(n=>({value:n,label:suppliers.get(n),sources:unique(p.filter(r=>r.supplier===n&&source.some(s=>s.code===r.code&&s.ids.some(id=>r.ids.includes(id)))).map(r=>r.source))})),sources:source.map(r=>r.source),reason:evidence.length>1?'同编号与物料对应多个供应商':'没有唯一供应商对应关系'});
  aliasesMap.set(supplier,{canonical,basis});
 }
 for(const r of f){const a=aliasesMap.get(r.supplier);r.canonical=a?.canonical||r.supplier;r.basis=a?.basis||'';if(!a?.canonical&&r.supplier)r.issues.push('供应商对应关系待确认');}
 const groups=new Map();for(const r of f){const k=r.ids.length&&r.code&&r.supplier?key(r.canonical,r.code,r.ids):r.id;if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);}
 const work=[...groups.values()].map(rows=>{
  const first=rows[0],candidates=unique(first.ids.flatMap(id=>(byOrderCode.get(key(id,first.code))||[]).filter(r=>r.supplier===first.canonical)));
  return {rows,first,candidates,issues:unique(rows.flatMap(r=>r.issues))};
 });
 const claims=new Map();for(const g of work)for(const r of g.candidates){if(!claims.has(r.id))claims.set(r.id,[]);claims.get(r.id).push(g);}
 const used=new Set(),records=[];
 const aggregate=rows=>rows.length&&rows.every(r=>r.qty!==null)?sum(rows.map(r=>r.qty)):null;
 const add=(status,prs,frs,extra={})=>{
  const first=frs[0]||prs[0],supplier=frs[0]?.canonical||first?.supplier||'',pq=extra.pqty??null,fq=aggregate(frs);
  const record={status,supplier,supplierName:suppliers.get(supplier)||first?.rawSupplier||'（供应商为空）',purchase:prs,fulfillment:frs,pqty:pq,fqty:fq,difference:extra.difference??null,basis:frs[0]?.basis||'',action:extra.action||'',...extra};records.push(record);return record;
 };
 for(const g of work){
  const {rows,first,candidates}=g,issues=[...g.issues];
  const overlap=candidates.some(r=>claims.get(r.id).length>1);
  if(overlap)issues.push('订单集合与其他履约记录重叠，请明确数量归属，未分摊采购数量');
  const complete=first.ids.every(id=>candidates.some(r=>r.ids.includes(id)));
  if(candidates.length&&!complete)issues.push('部分采购订单编号未找到，请补全来源订单');
  if(!candidates.length&&first.ids.some(id=>(byOrderCode.get(key(id,first.code))||[]).length))issues.push('同编号、物料的供应商不一致，请核实双方供应商');
  const valid=candidates.filter(r=>r.closed==='未关闭');
  // Closed rows provide identity only; their cached quantity is never counted.
  for(const r of candidates.filter(r=>r.closed!=='已关闭'))issues.push(...r.issues);
  if(!overlap)candidates.forEach(r=>used.add(r.id));
  if(issues.length){add('待核对',candidates,rows,{pqty:overlap?null:aggregate(valid),action:unique(issues).join('；')});continue;}
  if(!candidates.length){add('仅履约',[],rows,{action:'采购订单未找到该供应商、物料及订单，请核实订单编号或导出范围'});continue;}
  const fq=aggregate(rows);
  if(candidates.every(r=>r.closed==='已关闭')){add(fq!==0?'已关闭异常':'一致',candidates,rows,{pqty:0,difference:fq===0?0:null,action:fq!==0?'采购订单已关闭，但履约仍有未交付数量；请确认关闭是否正确及履约是否需要更新':''});continue;}
  const pq=aggregate(valid),diff=sum([fq,-pq]);
  add(diff===0?'一致':'数量差异',candidates,rows,{pqty:pq,difference:diff,action:diff?'核对双方数量；采购只计算未关闭订单，差额＝履约未交付－采购剩余入库':''});
 }
 const remaining=new Map();for(const r of p){if(used.has(r.id)||r.closed==='已关闭')continue;const k=r.ids.length===1?key(r.supplier,r.code,r.ids):r.id;if(!remaining.has(k))remaining.set(k,[]);remaining.get(k).push(r);}
 for(const rows of remaining.values()){
  const issues=unique(rows.flatMap(r=>r.issues));if(rows.some(r=>claims.has(r.id)))issues.push('关联的履约订单集合重叠，请明确对应行，未重复分配数量');
  add(issues.length?'待核对':'仅采购',rows,[],{pqty:aggregate(rows.filter(r=>r.closed==='未关闭')),action:issues.join('；')||'履约四个人员表中未找到对应订单，请核实是否遗漏或由其他人员负责'});
 }
 const totals=new Map(),get=supplier=>{if(!totals.has(supplier))totals.set(supplier,{supplier,name:suppliers.get(supplier)||f.find(r=>r.canonical===supplier)?.rawSupplier||'（供应商为空）',p:[],f:[],diff:[],counts:{}});return totals.get(supplier);};
 for(const r of p)if(r.closed!=='已关闭'){const g=get(r.supplier);if(r.closed==='未关闭'&&r.qty!==null)g.p.push(r.qty);}
 for(const r of f){const g=get(r.canonical);if(r.qty!==null)g.f.push(r.qty);}
 for(const r of records){const g=get(r.supplier);g.counts[r.status]=(g.counts[r.status]||0)+1;if(r.difference!==null)g.diff.push(r.difference);}
 const summary=[...totals.values()].map(g=>[g.name,g.p.length?sum(g.p):null,g.f.length?sum(g.f):null,g.diff.length?sum(g.diff):null,...['一致','数量差异','待核对','仅采购','仅履约','已关闭异常'].map(s=>g.counts[s]||0)]);
 const detail=records.filter(r=>r.status!=='一致').map(r=>{
  const all=[...r.fulfillment,...r.purchase],join=field=>unique(all.map(x=>x[field])).join('、');
  return [r.status,r.supplierName,unique(r.fulfillment.map(x=>x.rawSupplier)).join('、'),join('code'),join('sku'),join('name'),unique(r.purchase.flatMap(x=>x.ids)).join('＋'),unique(r.fulfillment.map(x=>x.rawOrder)).join('；'),
   r.purchase.map(x=>x.rawOrder+'：'+(x.qty===null?'数量空白或异常':x.qty)+'（'+x.closed+'；'+x.sheet+'!'+x.row+'）').join('\n'),r.pqty,r.fqty,r.difference,unique(r.purchase.map(x=>x.closed)).join('、'),r.purchase.map(x=>x.source).join('\n'),r.fulfillment.map(x=>x.source).join('\n'),r.basis,r.action];
 });
 return {summary,detail,records,choices,purchaseRows:p,fulfillmentRows:f,stats:{purchase:p.length,open:p.filter(r=>r.closed==='未关闭').length,closed:p.filter(r=>r.closed==='已关闭').length,fulfillment:f.length,matched:records.filter(r=>r.status==='一致').length},canExport:true};
}
return {build,parse,orders,number,norm,SUMMARY,DETAIL};
});
