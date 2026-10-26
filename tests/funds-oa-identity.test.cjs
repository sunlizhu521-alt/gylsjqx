const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');

const OLD='MDST-DGFK20990101001',CANDIDATE='MDST-DGFK20990101002',HEALTHY='MDST-DGFK20990101003';
const SUPPLIER='合成供应商甲',OTHER='合成供应商乙',CONTROL='合成正常供应商';
const FLOW_FIELDS=['流程编号','当前节点','未操作者','流程状态','申请人'];
const HEADERS=['登记月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单','提交付款申请单日期','提交时间汇总',...FLOW_FIELDS];
const APP_HEADERS=['申请日期','供应商全称','付款主体','流程编码','本次申请付款金额','付款分类'];
const OA_HEADERS=['流程编号','收款单位','付款金额','当前节点','未操作者','流程状态','申请人'];
const retained={'流程编号':'原流程信息','当前节点':'原节点','未操作者':'原操作者','流程状态':'原状态','申请人':'原申请人','提交付款申请单日期':'2026-09-01','提交时间汇总':'历史提交信息'};

function book(name,headers,records){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...records.map(r=>headers.map(h=>r[h]??''))]),name);return b;}
function payment(supplier,id,amount,extra={}){return {'登记月份':'2026-10-01','供应商名称':supplier,'付款主体':'合成采购主体','对公付款0A流程号':id,'开票时间':'2026-10-01','开票金额':amount,'摘要':'合成货款','申请付款金额':amount,'应付款时间':'2026-11-02','单据号【财务应付单】':'','OA合同编号':'','合同总金额':amount,'已付金额':0,'是否结清':'否','剩余金额':amount,'是否谈判付款':'否','是否提交付款申请单':'未申请',...extra};}
function flow(id,supplier,amount,extra={}){return {'流程编号':id,'收款单位':supplier,'付款金额':amount,'当前节点':'新节点','未操作者':'新操作者','流程状态':'新状态','申请人':'新申请人',...extra};}
function application(id,supplier,amount,type){return {'申请日期':'2026-10-25','供应商全称':supplier,'付款主体':'合成采购主体','流程编码':id,'本次申请付款金额':amount,'付款分类':type};}
function fixture(kind,{oldAmount=70,candidate=true,targetOA=OLD,targetSupplier=SUPPLIER,oldSupplier=OTHER}={}){
 const target=payment(targetSupplier,targetOA,100,{...retained,'是否提交付款申请单':kind==='谈判付款'?'未提交':'未申请'});
 const control=payment(CONTROL,HEALTHY,30);
 const ledger=book('对账明细-李奇',HEADERS,kind==='正常付款'?[target,control]:[control]);
 for(const person of C.PEOPLE.slice(1))X.utils.book_append_sheet(ledger,X.utils.aoa_to_sheet([HEADERS]),'对账明细-'+person);
 const oa=[flow(OLD,oldSupplier,oldAmount),flow(HEALTHY,CONTROL,30)];if(candidate)oa.push(flow(CANDIDATE,targetSupplier,100));
 const apps=[application(OLD,oldSupplier,oldAmount,kind),application(CANDIDATE,targetSupplier,100,kind),application(HEALTHY,CONTROL,30,'正常付款')];
 return {ledger:{book:ledger,name:'合成资金台账.xlsx'},negotiated:{book:book('合成谈判付款',HEADERS,kind==='谈判付款'?[target]:[]),name:'合成谈判.xlsx'},applications:{book:book('申请汇总',APP_HEADERS,apps),name:'合成申请.xlsx'},oa:{book:book('Sheet0',OA_HEADERS,oa),name:'合成OA.xlsx'}};
}
function output(result,kind,supplier=SUPPLIER){const headers=kind==='谈判付款'?result.negotiated.headers:result.normalHeaders,rows=kind==='谈判付款'?result.negotiated.rows:result.normalRows,row=rows.find(r=>r[headers.indexOf('供应商名称')]?.v===supplier);assert(row,'target record must remain in output');return Object.fromEntries(headers.map((h,i)=>[h,row[i]]));}
function append(input,role,headers,row){const b=input[role].book;X.utils.sheet_add_aoa(b.Sheets[b.SheetNames[0]],[headers.map(h=>row[h]??'')],{origin:-1});}
function set(input,role,header,row,value){const b=input[role].book,s=b.Sheets[b.SheetNames[0]],headers=X.utils.sheet_to_json(s,{header:1})[0];s[X.utils.encode_cell({r:row-1,c:headers.indexOf(header)})]={t:typeof value==='number'?'n':'s',v:value};}
function assertConflict(input,kind,{supplier=SUPPLIER,oldOA=OLD}={}){
 const before=JSON.stringify(input),result=C.build(input),row=output(result,kind,supplier);
 assert.equal(row['对公付款0A流程号'].v,oldOA,'an OA belonging to another supplier must never be replaced automatically');
 for(const field of FLOW_FIELDS)assert.equal(row[field].v,retained[field],field+' must retain its source value');
 assert.equal(row['是否提交付款申请单'].v,kind==='谈判付款'?'未提交':'未申请');
 assert.equal(row['提交时间汇总'].v,retained['提交时间汇总']);
 if(kind==='谈判付款')assert.equal(row['提交付款申请单日期'].v,retained['提交付款申请单日期']);
 const reviews=result.reviewRows.filter(r=>r[0].v===kind&&r[1].v===supplier),message=reviews.map(r=>r[5].v).join('\n');
 assert.match(message,/供应商.*冲突|冲突.*供应商/);assert(message.includes(OLD));assert(message.includes(supplier));assert(message.includes(OTHER));assert(message.includes('合成OA.xlsx'));assert(message.includes('Sheet0'));assert(!message.includes('已更新为'));
 const normal=output(result,'正常付款',CONTROL);assert.equal(normal['对公付款0A流程号'].v,HEALTHY);assert.equal(normal['流程编号'].v,HEALTHY);assert.equal(normal['当前节点'].v,'新节点');assert.equal(normal['是否提交付款申请单'].v,'已提交');
 assert.equal(JSON.stringify(input),before,'source cells, formulas and source order must not be changed');
 return result;
}

for(const kind of ['谈判付款','正常付款']){
 test(kind+'原OA属于其他供应商时，同金额或不同金额均保留原编号、流程及提交信息',()=>{
  for(const oldAmount of [70,100])assertConflict(fixture(kind,{oldAmount}),kind);
 });
 test(kind+'原OA跨供应商且无名称金额候选，仍须复核而不能读入其他供应商流程',()=>assertConflict(fixture(kind,{candidate:false}),kind));
 test(kind+'正确名称金额候选有重复明细，仍不能跨供应商自动更换原OA',()=>{
  const input=fixture(kind);append(input,'oa',OA_HEADERS,flow(CANDIDATE,SUPPLIER,100));assertConflict(input,kind);
 });
 test(kind+'原OA重复记录部分收款单位不同，即使其中一行匹配也必须保留复核',()=>{
  const input=fixture(kind,{oldAmount:100,candidate:false});append(input,'oa',OA_HEADERS,flow(OLD,SUPPLIER,100));assertConflict(input,kind);
 });
 test(kind+'跨供应商冲突时空白原流程字段及提交信息不得被其他供应商记录补入',()=>{
  const input=fixture(kind),role=kind==='谈判付款'?'negotiated':'ledger';for(const h of [...FLOW_FIELDS,'是否提交付款申请单','提交付款申请单日期','提交时间汇总'])set(input,role,h,2,'');
  const result=C.build(input),row=output(result,kind);assert.equal(row['对公付款0A流程号'].v,OLD);for(const h of [...FLOW_FIELDS,'是否提交付款申请单','提交时间汇总'])assert.equal(row[h].v,'',h);
  if(kind==='谈判付款')assert.equal(row['提交付款申请单日期'].v,'');assert(result.reviewRows.some(r=>/供应商.*冲突|冲突.*供应商/.test(r[5].v)));
 });
 test(kind+'原号未在OA表找到但名称金额唯一时维持原有自动修正规则',()=>{
  const input=fixture(kind,{targetOA:'MDST-DGFK20990101999'}),result=C.build(input),row=output(result,kind);assert.equal(row['对公付款0A流程号'].v,CANDIDATE);assert.equal(row['流程编号'].v,CANDIDATE);assert.equal(row['当前节点'].v,'新节点');
 });
 test(kind+'同一供应商名称空格及全半角、流程编码大小写差异仍正确匹配',()=>{
  const supplier='合成供应商Ａ',input=fixture(kind,{targetSupplier:supplier,targetOA:' ｍｄｓｔ－ｄｇｆｋ２０９９０１０１００１ ',oldSupplier:' 合成供应商 A ',oldAmount:100,candidate:false}),result=C.build(input),row=output(result,kind,supplier);
  assert.equal(row['流程编号'].v,OLD);assert.equal(row['当前节点'].v,'新节点');assert(!result.reviewRows.some(r=>/供应商.*冲突|冲突.*供应商/.test(r[5].v)));assert.equal(row['是否提交付款申请单'].v,kind==='谈判付款'?'已经提交全部':'已提交');
 });
 test(kind+'换行或空格分隔多个OA时任意编号跨供应商均须拦截',()=>{
  for(const separator of ['\n',' ']){const targetOA=OLD+separator+CANDIDATE;assertConflict(fixture(kind,{targetOA}),kind,{oldOA:targetOA});}
 });
 test(kind+'同供应商多个OA以换行或空格分隔时仍逐个读取流程及付款申请',()=>{
  for(const separator of ['\n',' ']){
   const targetOA=OLD+separator+CANDIDATE,input=fixture(kind,{targetOA,oldSupplier:SUPPLIER,oldAmount:100}),before=JSON.stringify(input),result=C.build(input),row=output(result,kind);
   assert.equal(row['对公付款0A流程号'].v,targetOA);assert.equal(row['流程编号'].v,OLD+'\n'+CANDIDATE);assert.equal(row['当前节点'].v,OLD+'：新节点\n'+CANDIDATE+'：新节点');assert.equal(row['是否提交付款申请单'].v,kind==='谈判付款'?'已经提交全部':'已提交');assert(!result.reviewRows.some(r=>r[5].v.includes('OA供应商冲突')));assert.equal(JSON.stringify(input),before);
  }
 });
}
test('谈判同供应商分批金额不同仍按原OA读取流程，申请累计只按该OA计算',()=>{
 const input=fixture('谈判付款',{oldSupplier:SUPPLIER,oldAmount:300,candidate:false});set(input,'applications','本次申请付款金额',2,40);const result=C.build(input),row=output(result,'谈判付款');
 assert.equal(row['对公付款0A流程号'].v,OLD);assert.equal(row['流程编号'].v,OLD);assert.equal(row['当前节点'].v,'新节点');assert.equal(row['是否提交付款申请单'].v,'已经提交部分');assert.match(row['提交时间汇总'].v,/40.00/);assert(!result.reviewRows.some(r=>/供应商.*冲突|冲突.*供应商/.test(r[5].v)));
});
test('谈判空OA仍可按唯一名称金额补齐，正常付款空OA仍不新增输出',()=>{
 const negotiated=C.build(fixture('谈判付款',{targetOA:''}));assert.equal(output(negotiated,'谈判付款')['对公付款0A流程号'].v,CANDIDATE);
 const normal=C.build(fixture('正常付款',{targetOA:''}));assert.equal(normal.normalRows.length,1);assert.equal(normal.normalRows[0][normal.normalHeaders.indexOf('供应商名称')].v,CONTROL);
});

test('谈判原OA空白时，台账补入的编号若属于其他供应商，应撤回补入并保留原值',()=>{
 const input=fixture('谈判付款',{targetOA:''});append(input,'ledger',HEADERS,payment(SUPPLIER,OLD,100));const before=JSON.stringify(input),result=C.build(input),row=output(result,'谈判付款');
 assert.equal(row['对公付款0A流程号'].v,'');for(const h of FLOW_FIELDS)assert.equal(row[h].v,retained[h]);assert.equal(row['是否提交付款申请单'].v,'未提交');assert.equal(row['提交时间汇总'].v,retained['提交时间汇总']);assert(result.reviewRows.some(r=>r[1].v===SUPPLIER&&/OA供应商冲突/.test(r[5].v)));assert.equal(JSON.stringify(input),before);
});
