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
 assert.equal(row['对公付款0A流程号'].v,OLD);assert.equal(row['流程编号'].v,OLD);assert.equal(row['当前节点'].v,'新节点');assert.equal(row['是否提交付款申请单'].v,'已经提交部分');assert.equal(row['提交时间汇总'].v,retained['提交时间汇总']);assert(!result.reviewRows.some(r=>/供应商.*冲突|冲突.*供应商/.test(r[5].v)));
});
test('谈判空OA仍可按唯一名称金额补齐，正常付款空OA仍不新增输出',()=>{
 const negotiated=C.build(fixture('谈判付款',{targetOA:''}));assert.equal(output(negotiated,'谈判付款')['对公付款0A流程号'].v,CANDIDATE);
 const normal=C.build(fixture('正常付款',{targetOA:''}));assert.equal(normal.normalRows.length,1);assert.equal(normal.normalRows[0][normal.normalHeaders.indexOf('供应商名称')].v,CONTROL);
});

test('谈判原OA空白时，台账补入的编号若属于其他供应商，应撤回补入并保留原值',()=>{
 const input=fixture('谈判付款',{targetOA:''});append(input,'ledger',HEADERS,payment(SUPPLIER,OLD,100));const before=JSON.stringify(input),result=C.build(input),row=output(result,'谈判付款');
 assert.equal(row['对公付款0A流程号'].v,'');for(const h of FLOW_FIELDS)assert.equal(row[h].v,retained[h]);assert.equal(row['是否提交付款申请单'].v,'未提交');assert.equal(row['提交时间汇总'].v,retained['提交时间汇总']);assert(result.reviewRows.some(r=>r[1].v===SUPPLIER&&/OA供应商冲突/.test(r[5].v)));assert.equal(JSON.stringify(input),before);
});

function sameAmountNormals(){
 const input=fixture('正常付款',{oldSupplier:SUPPLIER,oldAmount:100});
 append(input,'ledger',HEADERS,payment(SUPPLIER,CANDIDATE,100,{'应付款时间':'2026-10-26'}));
 input.applications.book=book('申请汇总',APP_HEADERS,[application(OLD,SUPPLIER,100,'正常付款')]);
 return input;
}
const normalByOA=(result,id)=>Object.fromEntries(result.normalHeaders.map((h,c)=>[h,result.normalRows.find(r=>r[result.normalHeaders.indexOf('对公付款0A流程号')].v===id)[c]]));
test('正常付款同供应商同金额按OA号区分提交状态、申请日期、摘要和整行标红',()=>{
 const input=sameAmountNormals(),before=JSON.stringify(input),result=C.build(input),submitted=normalByOA(result,OLD),unsubmitted=normalByOA(result,CANDIDATE);
 assert.equal(result.normalRows.filter(r=>r[2].v===SUPPLIER).length,2,'台账两行原样保留');
 assert.equal(submitted['流程编号'].v,OLD);assert.equal(unsubmitted['流程编号'].v,CANDIDATE);
 assert.equal(submitted['是否提交付款申请单'].v,'已提交');assert.equal(submitted['申请日期'].t,'n');assert.equal(submitted['提交时间汇总'].v,'2026-10-25-100.00');
 assert.equal(unsubmitted['是否提交付款申请单'].v,'未申请');assert.equal(unsubmitted['申请日期'].v,'');assert.equal(unsubmitted['提交时间汇总'].v,'');
 assert.deepEqual(result.highlights.normalRows,[result.normalRows.findIndex(r=>r[4].v===OLD)+2]);assert.equal(JSON.stringify(input),before);
});
test('正常付款不同人员的同供应商同额OA分别保留，流程节点及申请人不会串到另一行',()=>{
 const input=sameAmountNormals(),ledger=input.ledger.book.Sheets['对账明细-李奇'];
 const row=HEADERS.map((_,c)=>ledger[X.utils.encode_cell({r:3,c})]?.v??'');
 X.utils.sheet_add_aoa(input.ledger.book.Sheets['对账明细-徐英'],[row],{origin:'A2'});for(let c=0;c<HEADERS.length;c++)delete ledger[X.utils.encode_cell({r:3,c})];
 for(const [h,v] of [['当前节点','办理人执行'],['未操作者','人员乙'],['流程状态','归档'],['申请人','申请人乙']])set(input,'oa',h,4,v);
 const before=JSON.stringify(input),result=C.build(input),first=normalByOA(result,OLD),second=normalByOA(result,CANDIDATE);
 assert.equal(result.normalRows.filter(r=>r[2].v===SUPPLIER).length,2);assert.equal(first['跟单员'].v,'李奇');assert.equal(second['跟单员'].v,'徐英');
 assert.deepEqual(FLOW_FIELDS.map(h=>first[h].v),[OLD,'新节点','新操作者','新状态','新申请人']);
 assert.deepEqual(FLOW_FIELDS.map(h=>second[h].v),[CANDIDATE,'办理人执行','人员乙','归档','申请人乙']);
 for(const r of [first,second]){assert.equal(r['本次申请付款金额'].v,100);assert.equal(r['已付金额'].v,0);assert.equal(r['剩余金额'].v,100);}
 assert.equal(second['是否提交付款申请单'].v,'未申请');assert.equal(JSON.stringify(input),before);
});
test('同供应商同额的另一个OA有流程冲突，不影响本行正确OA的信息',()=>{
 const input=sameAmountNormals();append(input,'oa',OA_HEADERS,flow(CANDIDATE,SUPPLIER,100,{'当前节点':'冲突节点'}));
 const result=C.build(input);assert.equal(normalByOA(result,OLD)['流程编号'].v,OLD);assert.equal(normalByOA(result,OLD)['当前节点'].v,'新节点');
 assert.equal(result.issues.some(i=>i.id==='oa:对账明细-李奇!2'),false);assert.equal(result.issues.some(i=>i.id==='oa:对账明细-李奇!4'),true);
 assert.equal(result.normalRows.filter(r=>r[2].v===SUPPLIER).length,2);
});
test('正常付款同额不同OA的重复申请只归本OA；分类、供应商及金额不同不匹配',()=>{
 const input=sameAmountNormals();append(input,'applications',APP_HEADERS,{...application(OLD,SUPPLIER,'￥100.00','正常付款'),'申请日期':'2026-10-26'});
 append(input,'applications',APP_HEADERS,application(CANDIDATE,OTHER,100,'正常付款'));
 append(input,'applications',APP_HEADERS,application(CANDIDATE,SUPPLIER,99.99,'正常付款'));
 const result=C.build(input);assert.equal(normalByOA(result,CANDIDATE)['是否提交付款申请单'].v,'未申请');assert.equal(normalByOA(result,OLD)['提交时间汇总'].v,'2026-10-25-100.00\n2026-10-26-100.00');assert.equal(normalByOA(result,OLD)['申请日期'].v,'2026年10月25日；2026年10月26日');
 assert.equal(result.highlights.normalRows.length,1);assert.equal(result.reviewRows.filter(r=>r[5].v.startsWith('重复提交')).length,1);
});
test('正常付款缺申请OA仅唯一同名同额可补判，多笔时不猜测并列明补编号的来源',()=>{
 const input=sameAmountNormals();set(input,'applications','流程编码',2,'');let result=C.build(input);
 assert(result.normalRows.every(r=>r[16].v==='未申请'));assert.deepEqual(result.highlights.normalRows,[]);
 const reviews=result.reviewRows.filter(r=>r[5].v.startsWith('付款申请归属核对：'));
 assert.equal(reviews.length,2);assert(reviews.every(r=>r[5].v.includes(OLD)&&r[5].v.includes(CANDIDATE)&&r[5].v.includes('合成申请.xlsx / 申请汇总 第 2 行')&&r[5].v.includes('需要处理')));
 assert.equal(result.stats.pending,0,'归属核对不阻止下载');
 const sheet=input.ledger.book.Sheets['对账明细-李奇'];for(let c=0;c<HEADERS.length;c++)delete sheet[X.utils.encode_cell({r:3,c})];result=C.build(input);
 assert.equal(normalByOA(result,OLD)['是否提交付款申请单'].v,'已提交');assert.equal(result.highlights.normalRows.length,1);
});
test('正常付款使用核验后的OA且兼容编号标准化、旧合同号；异常编号不能作为缺号回退',()=>{
 for(const reference of [' ｍｄｓｔ－ｄｇｆｋ２０９９０１０１００１ ',OLD]){
  const input=sameAmountNormals();set(input,'applications','流程编码',2,reference);const s=input.applications.book.Sheets.申请汇总;s.D1.v='合同号';const result=C.build(input);
  assert.equal(normalByOA(result,OLD)['是否提交付款申请单'].v,'已提交');assert.equal(normalByOA(result,CANDIDATE)['是否提交付款申请单'].v,'未申请');
 }
 for(const reference of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const input=fixture('正常付款',{oldSupplier:SUPPLIER,oldAmount:100,candidate:false});input.applications.book=book('申请汇总',APP_HEADERS,[application(OLD,SUPPLIER,100,'正常付款')]);input.applications.book.Sheets.申请汇总.D2=reference;
  assert.equal(output(C.build(input),'正常付款')['是否提交付款申请单'].v,'未申请');
 }
 const input=fixture('正常付款',{targetOA:'MDST-DGFK20990101999'}),result=C.build(input);assert.equal(output(result,'正常付款')['对公付款0A流程号'].v,CANDIDATE);assert.equal(output(result,'正常付款')['是否提交付款申请单'].v,'已提交');
});
test('正常付款有编号冲突时唯一同额也不回退，跨多个OA申请不能误配单个OA',()=>{
 for(const id of [CANDIDATE,OLD+' / '+CANDIDATE]){
  const input=fixture('正常付款',{oldSupplier:SUPPLIER,oldAmount:100,candidate:false});input.applications.book=book('申请汇总',APP_HEADERS,[application(id,SUPPLIER,100,'正常付款')]);
  const result=C.build(input),row=output(result,'正常付款');assert.equal(row['是否提交付款申请单'].v,'未申请');assert.equal(row['申请日期'].v,'');assert.equal(row['提交时间汇总'].v,'');assert.deepEqual(result.highlights.normalRows,[]);
 }
});
