const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,LH,NH,AH,ID}=require('./funds-progress-fixtures.cjs');
const SUPPLIER='总额核对合成供应商',PREFIX='台账金额差异核对：';
const REVIEW_HEADERS=['付款类别','供应商名称','付款主体','申请付款金额','是否提交付款申请单','核对说明','来源文件','来源工作表','来源行号','跟单员'];
const NEG_HEADERS=[...NH,'单据号【财务应付单】','OA合同编号','是否提交付款申请单'],APP_HEADERS=[...AH,'核对说明'];
const sheet=(headers,records)=>X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(h=>record[h]??''))]);
function payment(amount=100,extra={}){
 const numeric=typeof amount==='number'?amount:100;
 return {'登记月份':46300,'供应商名称':SUPPLIER,'付款主体':'迈德斯特','对公付款0A流程号':'未提交OA','开票时间':46300,'开票金额':numeric,'摘要':'原人工摘要','申请付款金额':amount,'应付款时间':46332,'单据号【财务应付单】':'','OA合同编号':'','合同总金额':numeric,'已付金额':0,'剩余金额':numeric,'剩余付款金额':numeric,'是否结清':numeric===0?'是':'否','是否谈判付款':'是','是否提交付款申请单':'未提交','备注':'保留人工备注','谈判明细':'11月2日－'+numeric,'分批总金额':numeric,'核对':'原人工核对',46328:numeric,46335:0,...extra};
}
function inputs(negotiated,ledger,applications=[]){
 const input=fixture();
 for(const person of C.PEOPLE)input.ledger.book.Sheets['对账明细-'+person]=sheet(LH,ledger.filter(record=>(record.owner||'徐英')===person));
 input.negotiated.book.Sheets.十一月谈判付款主表=sheet(NEG_HEADERS,negotiated);
 for(const column of ['M','N'])input.negotiated.book.Sheets.十一月谈判付款主表[column+'1'].z=C.DATE_FORMAT;
 input.applications.book.Sheets.付款申请汇总=sheet(APP_HEADERS,applications);
 const name=input.oa.book.SheetNames[0],headers=X.utils.sheet_to_json(input.oa.book.Sheets[name],{header:1})[0];input.oa.book.Sheets[name]=sheet(headers,[]);
 return input;
}
const audits=result=>result.reviewRows.filter(row=>row[5].v.startsWith(PREFIX));
const auditText=result=>audits(result).map(row=>row[5].v).join('\n');
function build(input){const before=JSON.stringify(input),result=C.build(input);assert.equal(JSON.stringify(input),before,'任何上传源金额、公式、备注和原行均不改动');assert.deepEqual(result.reviewHeaders,REVIEW_HEADERS);return result;}
function noAmountMismatch(result){
 assert.equal(audits(result).length,0,'供应商总额一致不产生差异核对');
 assert(!result.reviewRows.some(row=>/^已付金额同步核对：(?:未找到 OA号、供应商和申请付款金额三项完全一致|同一匹配键对应多条台账记录)/.test(row[5].v)),'总额一致不能残留旧逐笔金额未匹配提示');
}
function difference(input,{negotiated,ledger,delta,negRows=[2],ledgerRefs=['对账明细-徐英!H2'],owner='徐英'}={}){
 const result=build(input),reviews=audits(result);assert.equal(reviews.length,1,'每个供应商仅输出一条总额核对');
 const review=reviews[0],note=review[5].v;
 assert.equal(review.length,10);assert.equal(review[0].v,'谈判付款');assert.equal(review[3].v,negotiated,'金额列展示该供应商谈判总额');
 if(Array.isArray(owner))assert.deepEqual(new Set(String(review[9].v).split(/[、，,\s]+/)),new Set(owner));else assert.equal(review[9].v,owner);
 assert.deepEqual(String(review[8].v).split('、').map(Number),negRows,'来源行号列保留该供应商全部谈判行');
 for(const token of [negotiated.toFixed(2),ledger.toFixed(2),delta,'谈判减台账','negotiated.xlsx','ledger.xlsx',...negRows.map(row=>'十一月谈判付款主表!D'+row),...ledgerRefs])assert(note.includes(token),token+' 缺少完整核对依据：'+note);
 return {result,review,note};
}

test('凯普威三条谈判2696000对六条台账总额一致，换OA、不同主体和已结清不漏期，正常款排除',()=>{
 const negotiated=[payment(960000,{'对公付款0A流程号':ID(1)}),payment(1000000,{'对公付款0A流程号':ID(2)}),payment(736000,{'对公付款0A流程号':ID(3)})];
 const installments=Array.from({length:4},(_,i)=>payment(240000,{'对公付款0A流程号':ID(i===0?99:1),'付款主体':i%2?'国源':'迈德斯特',...(i===0?{'已付金额':240000,'剩余金额':0,'是否结清':'是'}:{})}));
 const ledger=[...installments,payment(1000000,{'对公付款0A流程号':ID(2)}),payment(736000,{'对公付款0A流程号':ID(3)}),payment(110500,{'是否谈判付款':'否'}),payment(1234,{'是否谈判付款':''})];
 const input=inputs(negotiated,ledger),result=build(input);noAmountMismatch(result);assert.equal(result.negotiated.rows.length,3);assert.equal(result.normalRows.length,2);
});

test('四条同额分期不去重，少一期真实差240000，不按960000反推补齐第四期',()=>{
 const parts=Array.from({length:4},(_,i)=>payment(240000,{'对公付款0A流程号':ID(i+10)}));
 noAmountMismatch(build(inputs([payment(960000)],parts)));
 const {note}=difference(inputs([payment(960000)],parts.slice(0,3)),{negotiated:960000,ledger:720000,delta:'240000.00',ledgerRefs:[2,3,4].map(row=>'对账明细-徐英!H'+row)});
 assert(!note.includes('对账明细-徐英!H5'),'不能虚构缺少的第四期来源');
});

test('钰民跨主体十三条谈判4750864对十四条台账4750829只报35，改齐后无差异提示',()=>{
 const negotiated=[...Array.from({length:12},(_,i)=>payment(300000,{'付款主体':i%2?'国源':'迈德斯特','对公付款0A流程号':ID(i+20)})),payment(1150864,{'付款主体':'待核对'})];
 const ledger=[...Array.from({length:12},(_,i)=>payment(300000,{'付款主体':i%2?'迈德斯特':'国源','对公付款0A流程号':ID(i+40)})),payment(500000),payment(650829),payment(5342,{'是否谈判付款':'否'})];
 difference(inputs(negotiated,ledger),{negotiated:4750864,ledger:4750829,delta:'35.00',negRows:Array.from({length:13},(_,i)=>i+2),ledgerRefs:Array.from({length:14},(_,i)=>'对账明细-徐英!H'+(i+2))});
 ledger[13]=payment(650864);noAmountMismatch(build(inputs(negotiated,ledger)));
});

test('嘉兴历史一致笔加入供应商合计不掩盖792差异，原金额计划公式和人工备注全部保留',()=>{
 const history=[...Array.from({length:6},(_,i)=>payment((i+1)*1000,{'对公付款0A流程号':ID(i+60)})),payment(121200,{'付款主体':'国源'})];
 const input=inputs([...history,payment(312026,{'付款主体':'待核对'})],[...history,payment(311234),payment(312026,{'是否谈判付款':'否'}),payment(311000,{'是否谈判付款':'否'})]);
 const source=input.negotiated.book.Sheets.十一月谈判付款主表;source.M9.v=200000;source.N9.v=112026;source.K9={t:'n',v:312026,f:'SUM(M9:N9)'};source.F9={t:'n',v:312026,f:'D9-E9'};source.I9.c=[{a:'人工',t:'不能覆盖的备注'}];
 const {result}=difference(input,{negotiated:454226,ledger:453434,delta:'792.00',negRows:[2,3,4,5,6,7,8,9],ledgerRefs:[2,3,4,5,6,7,8,9].map(row=>'对账明细-徐英!H'+row)});
 assert.equal(result.negotiated.rows.length,8);assert.equal(result.normalRows.length,2);
 for(const column of ['D','F','I','J','K','M','N'])assert.deepEqual(result.negotiated.rows[7][X.utils.decode_col(column)],source[column+'9'],column+'9 原金额公式备注保留');
});

test('供应商名称按精确标准化合计，同供应商的主体或OA冲突不排除本金，其他供应商不用于对冲',()=>{
 const result=build(inputs([payment(40,{'供应商名称':' 合成供应商Ａ ','付款主体':'国源','对公付款0A流程号':ID(71)}),payment(60,{'供应商名称':'合成供应商A','付款主体':'迈德斯特','对公付款0A流程号':ID(72)})],[payment(100,{'供应商名称':'合成供应商 A','付款主体':'另一主体','对公付款0A流程号':ID(73)}),payment(99,{'供应商名称':'合成供应商A分公司'})]));
 noAmountMismatch(result);
 const separate=build(inputs([payment(100),payment(80,{'供应商名称':'独立供应商'})],[payment(90),payment(90,{'供应商名称':'独立供应商'})]));
 assert.equal(audits(separate).length,2,'不同供应商的正负10不能对冲为零');assert(auditText(separate).includes('-10.00'));
});

test('合计包含完整谈判组，不挑40加60凑100，也不因跨人员重复等额而省略来源',()=>{
 const {note}=difference(inputs([payment(100)],[payment(40),payment(60),payment(5)]),{negotiated:100,ledger:105,delta:'-5.00',ledgerRefs:[2,3,4].map(row=>'对账明细-徐英!H'+row)});
 assert(note.includes('5.00'));
 difference(inputs([payment(100)],[payment(90,{owner:'徐英'}),payment(90,{owner:'李奇'})]),{negotiated:100,ledger:180,delta:'-80.00',ledgerRefs:['对账明细-徐英!H2','对账明细-李奇!H2'],owner:['李奇','徐英']});
});

test('台账是否谈判为否或空白排除，纵向合并的有效是继承后各行均计入',()=>{
 const input=inputs([payment(100)],[payment(40),payment(60,{'是否谈判付款':''}),payment(777,{'是否谈判付款':'否'}),payment(888,{'是否谈判付款':''})]);
 input.ledger.book.Sheets['对账明细-徐英']['!merges']=[{s:{r:1,c:15},e:{r:2,c:15}}];
 noAmountMismatch(build(input));
 delete input.ledger.book.Sheets['对账明细-徐英']['!merges'];difference(input,{negotiated:100,ledger:40,delta:'60.00'});
});

test('汇总按分计算且有效零计入，分期金额为空错误无缓存或负数均保留核对不当零',()=>{
 noAmountMismatch(build(inputs([payment(0.1),payment(0.2),payment(0)],[payment(0.3),payment({t:'n',f:'0',v:0})])));
 difference(inputs([payment(0.1),payment(0.2)],[payment(0.29)]),{negotiated:0.3,ledger:0.29,delta:'0.01',negRows:[2,3]});
 for(const amount of ['',{t:'e',v:23,w:'#REF!'},{t:'n',f:'SUM(A99:A100)'},{t:'z',f:'SUM(A99:A100)',v:0},-1])for(const side of ['ledger','negotiated']){
  const result=build(inputs(side==='negotiated'?[payment(100),payment(amount)]:[payment(100)],side==='ledger'?[payment(100),payment(amount)]:[payment(100)])),note=auditText(result);
  assert.equal(audits(result).length,1,side+' '+JSON.stringify(amount));assert.match(note,/为空|异常|无效|缺失|不可用|负数/);
  assert.match(note,side==='ledger'?/对账明细-徐英!H3/:/十一月谈判付款主表!D3/);assert.doesNotMatch(note,/差额[^\n；。]*(?:\)|）|:|：)\s*-?\d+\.\d{2}/,'不完整合计不能产生确定差额');
 }
});

test('单元格金额有效但供应商总和超出安全整数分范围时不输出近似合计或静默一致',()=>{
 const max='90071992547409.91';
 for(const side of ['ledger','negotiated']){
  const result=build(inputs(side==='negotiated'?[payment(max),payment(0.01)]:[payment(max)],side==='ledger'?[payment(max),payment(0.01)]:[payment(max)])),note=auditText(result);
  assert.equal(audits(result).length,1);assert.match(note,/超出|超过|溢出/);assert.match(note,/安全|精度|范围/);assert.doesNotMatch(note,/差额[^\n；。]*(?:\)|）|:|：)\s*-?0\.01/);
 }
});

test('未提交、已提交或重新申请不改变供应商本金核对结果',()=>{
 const application={'申请日期':'2026-11-18','供应商全称':SUPPLIER,'付款主体':'迈德斯特','流程编码':'未提交OA','本次申请付款金额':100,'付款分类':'谈判付款'};
 for(const applications of [[],[application],[{...application,'核对说明':'重新申请'}]])difference(inputs([payment(100)],[payment(90)],applications),{negotiated:100,ledger:90,delta:'10.00'});
});

test('全部不支付的供应商跳过；应付与不支付混合且台账无法分离时待核对，不能直接按零凑平',()=>{
 for(const marker of ['不支付',' 不 支 付 '])assert.equal(audits(build(inputs([payment(100,{'对公付款0A流程号':marker})],[payment(100)]))).length,0);
 noAmountMismatch(build(inputs([payment(100),payment(0,{'对公付款0A流程号':'不支付'})],[payment(100)])));
 difference(inputs([payment(100),payment(0,{'对公付款0A流程号':'不支付'})],[payment(90)]),{negotiated:100,ledger:90,delta:'10.00'});
 const mixed=build(inputs([payment(100),payment(50,{'对公付款0A流程号':'不支付'})],[payment(150)])),note=auditText(mixed);
 assert.equal(audits(mixed).length,1);assert.match(note,/不支付/);assert.match(note,/待核对|待确认|无法|不能/);assert.doesNotMatch(note,/差额[^\n；。]*(?:\)|）|:|：)\s*-50\.00/);
 difference(inputs([payment(100,{'对公付款0A流程号':'不支付待确认'})],[payment(90)]),{negotiated:100,ledger:90,delta:'10.00'});
});
