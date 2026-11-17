const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,LH,NH,AH,ID}=require('./funds-progress-fixtures.cjs');
const SUPPLIER='嘉兴迈锐合成供应商',PREFIX='台账金额差异核对：';
const REVIEW_HEADERS=['付款类别','供应商名称','付款主体','申请付款金额','是否提交付款申请单','核对说明','来源文件','来源工作表','来源行号','跟单员'];
const NEG_HEADERS=[...NH,'单据号【财务应付单】','OA合同编号','是否提交付款申请单'];
const APP_HEADERS=[...AH,'核对说明'];
const sheet=(headers,records)=>X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(h=>record[h]??''))]);
function payment(amount=100,extra={}){
 const numeric=typeof amount==='number'?amount:100;
 return {'登记月份':46300,'供应商名称':SUPPLIER,'付款主体':'迈德斯特','对公付款0A流程号':'未提交OA','开票时间':46300,'开票金额':numeric,'摘要':'原人工摘要','申请付款金额':amount,'应付款时间':46332,'单据号【财务应付单】':'','OA合同编号':'','合同总金额':numeric,'已付金额':0,'剩余金额':numeric,'剩余付款金额':numeric,'是否结清':numeric===0?'是':'否','是否谈判付款':'是','是否提交付款申请单':'未提交','备注':'保留人工备注','谈判明细':'11月2日－'+numeric,'分批总金额':numeric,'核对':'原人工核对',46328:numeric,46335:0,...extra};
}
function inputs(negotiated,ledger,applications=[]){
 const input=fixture();
 for(const person of C.PEOPLE)input.ledger.book.Sheets['对账明细-'+person]=sheet(LH,ledger.filter(record=>(record.owner||'徐英')===person));
 const name=input.negotiated.book.SheetNames[0];input.negotiated.book.Sheets[name]=sheet(NEG_HEADERS,negotiated);
 for(const column of ['M','N'])input.negotiated.book.Sheets[name][column+'1'].z=C.DATE_FORMAT;
 input.applications.book.Sheets.付款申请汇总=sheet(APP_HEADERS,applications);
 const oa=input.oa.book.SheetNames[0],headers=X.utils.sheet_to_json(input.oa.book.Sheets[oa],{header:1})[0];input.oa.book.Sheets[oa]=sheet(headers,[]);
 return input;
}
const audits=result=>result.reviewRows.filter(row=>row[5].v.startsWith(PREFIX));
const auditText=result=>audits(result).map(row=>row[5].v).join('\n');
const output=(result,index=0)=>Object.fromEntries(result.negotiated.headers.map((header,c)=>[header,result.negotiated.rows[index][c]]));
function uniqueAudit(input,{amount=100,ledgerAmount=90,difference='10.00',row=2,ledgerRow=2}={}){
 const before=JSON.stringify(input),result=C.build(input),reviews=audits(result);assert.equal(reviews.length,1);
 const review=reviews[0],note=review[5].v;
 assert.equal(review.length,10);assert.deepEqual(result.reviewHeaders,REVIEW_HEADERS);assert.equal(review[0].v,'谈判付款');assert.equal(review[3].v,amount);assert.equal(review[8].v,row);assert.equal(review[9].v,'徐英');
 for(const token of [Number(amount).toFixed(2),Number(ledgerAmount).toFixed(2),difference,'谈判减台账','匹配依据','negotiated.xlsx','十一月谈判付款主表','D'+row,'ledger.xlsx','对账明细-徐英','H'+ledgerRow])assert(note.includes(token),token+' 缺少核对依据：'+note);
 assert.equal(JSON.stringify(input),before,'输入金额、公式、备注及原行均不改动');
 return {result,review,note};
}

test('同供应商八条谈判排除六条OA独占和一条主体金额独占后，核对312026与311234差792并归徐英',()=>{
 const exact=Array.from({length:6},(_,i)=>payment((i+1)*1000,{'对公付款0A流程号':ID(i+1)}));
 const negotiated=[...exact,payment(121200,{'付款主体':'国源','对公付款0A流程号':''}),payment(312026,{'付款主体':'待核对'})];
 const ledger=[...exact,payment(121200,{'付款主体':'国源','对公付款0A流程号':''}),payment(311234),payment(312026,{'是否谈判付款':'否'}),payment(311000,{'是否谈判付款':'否'})];
 const input=inputs(negotiated,ledger),source=input.negotiated.book.Sheets.十一月谈判付款主表;
 source.M9.v=200000;source.N9.v=112026;source.K9={t:'n',v:312026,f:'SUM(M9:N9)'};source.F9={t:'n',v:312026,f:'D9-E9'};source.I9.c=[{a:'人工',t:'不能覆盖的备注'}];
 const {result}=uniqueAudit(input,{amount:312026,ledgerAmount:311234,difference:'792.00',row:9,ledgerRow:9});
 assert.equal(result.negotiated.rows.length,8);assert.equal(result.normalRows.length,2);
 const preserved=output(result,7);for(const header of ['申请付款金额','剩余付款金额','分批总金额','备注','谈判明细',String(46328),String(46335)]){
  const column=NEG_HEADERS.findIndex(value=>String(value)===header),original=source[X.utils.encode_cell({r:8,c:column})],actual=result.negotiated.rows[7][column];
  assert.deepEqual(actual,original,header+' 原业务值与公式保留');
 }
 assert.equal(preserved['付款主体'].v,'待核对','未知主体不能冒充已明确身份');
});

test('OA、财务AP和合同号身份优先于金额接近，强标识匹配不受弱候选分类限制',()=>{
 for(const [field,id] of [['对公付款0A流程号',ID(11)],['单据号【财务应付单】','AP20261118001'],['OA合同编号','HT-TEST-001']]){
  const input=inputs([payment(100,{[field]:id})],[payment(90,{[field]:id}),payment(99,{[field]:field==='对公付款0A流程号'?ID(12):field==='单据号【财务应付单】'?'AP20261118002':'HT-TEST-002'})]);
  uniqueAudit(input);
 }
 uniqueAudit(inputs([payment(100,{'对公付款0A流程号':ID(11)})],[payment(90,{'对公付款0A流程号':ID(11),'是否谈判付款':'否'})]));
});

test('不同供应商、冲突独立编号或明确不同主体不能被同名近额回退凑成匹配',()=>{
 const target=payment(100,{'对公付款0A流程号':ID(21),'单据号【财务应付单】':'AP20261118001','OA合同编号':'HT-TEST-001'});
 for(const override of [{'供应商名称':SUPPLIER+'分公司'},{'对公付款0A流程号':ID(22)},{'单据号【财务应付单】':'AP20261118002'},{'OA合同编号':'HT-TEST-002'},{'付款主体':'国源'}]){
  const result=C.build(inputs([target],[{...target,'申请付款金额':90,...override}]));
  assert.doesNotMatch(auditText(result),/差额[^\n；。]*10\.00/,JSON.stringify(override));
 }
 uniqueAudit(inputs([payment(100,{'供应商名称':' 合成供应商Ａ ','付款主体':'待确认'})],[payment(90,{'供应商名称':'合成供应商 A'})]));
});

test('无独立编号的候选仅含台账是否谈判是，空白与否不能充当唯一弱匹配',()=>{
 for(const flag of ['否','','待确认',{t:'e',v:23,w:'#REF!'}]){
  const result=C.build(inputs([payment(100)],[payment(90,{'是否谈判付款':flag})]));assert.equal(audits(result).length,0,JSON.stringify(flag));
 }
 for(const entity of ['待核对','待确认','']){
  const {result}=uniqueAudit(inputs([payment(100,{'付款主体':entity})],[payment(90)]));assert.equal(output(result)['付款主体'].v,entity);
 }
});

test('同一台账被两条谈判争用时不从弱候选中剔除，多候选不按最接近金额挑选',()=>{
 const identified=payment(100,{'对公付款0A流程号':ID(31)}),input=inputs([identified,identified,payment(200)],[identified,payment(180)]),result=C.build(input);
 const weak=audits(result).find(row=>row[8].v===4);assert(weak,'弱匹配仍须列出争用候选');assert.match(weak[5].v,/待确认|归属|多条|多个/);
 for(const token of ['100.00','180.00','H2','H3'])assert(weak[5].v.includes(token),token);assert.doesNotMatch(weak[5].v,/差额[^\n；。]*20\.00/);
 const multiple=C.build(inputs([payment(100)],[payment(99.99),payment(80)])),note=auditText(multiple);
 assert.equal(audits(multiple).length,1);assert.match(note,/待确认|归属|多条|多个/);for(const token of ['99.99','80.00','H2','H3'])assert(note.includes(token),token);
 assert.doesNotMatch(note,/差额[^\n；。]*0\.01/,'不能选取金额最相近的记录');
});

test('独占匹配先整体确定，不因谈判行顺序影响剩余唯一台账金额核对',()=>{
 const fixed=payment(80,{'对公付款0A流程号':ID(41)}),weak=payment(100),ledger=[fixed,payment(90)];
 uniqueAudit(inputs([weak,fixed],ledger),{ledgerRow:3});uniqueAudit(inputs([fixed,weak],ledger),{row:3,ledgerRow:3});
});

test('有效同OA拆笔台账合计等于谈判总额不误报，各候选均同额也不产生非零差异',()=>{
 const same={'对公付款0A流程号':ID(51)},split=C.build(inputs([payment(100,same)],[payment(40,same),payment(60,same)]));assert.equal(audits(split).length,0);
 const equal=C.build(inputs([payment(100)],[payment(100),payment(100)]));
 assert(audits(equal).every(row=>!/差额[^\n；。]*[1-9]\d*\.\d{2}/.test(row[5].v)),'多条同额候选不得报告金额不等');
});

test('同OA必须核对完整拆笔组，不能挑金额子集凑总额或混入不同主体',()=>{
 const same={'对公付款0A流程号':ID(61)},result=C.build(inputs([payment(100,same)],[payment(40,same),payment(60,same),payment(5,same)]));
 assert.equal(audits(result).length,1);assert.match(auditText(result),/待确认|归属|多条|多个/);
 for(const token of ['40.00','60.00','5.00','H2','H3','H4'])assert(auditText(result).includes(token),token+' 不可从完整同OA组中丢弃');
 uniqueAudit(inputs([payment(100,same)],[payment(90,same),payment(10,{...same,'付款主体':'国源'})]));
});

test('跨人员多候选不指派唯一跟单员，身份公式错误不能抢占另一条谈判的台账候选',()=>{
 const input=inputs([payment(100)],[payment(90,{owner:'徐英'}),payment(80,{owner:'李奇'})]),result=C.build(input),review=audits(result)[0];
 assert.equal(audits(result).length,1);assert.equal(review[9].v,'');assert.match(review[5].v,/对账明细-徐英!H2/);assert.match(review[5].v,/对账明细-李奇!H2/);
 const invalid=payment(100,{'对公付款0A流程号':{t:'e',v:23,w:'#REF!'}}),next=C.build(inputs([invalid,payment(200)],[payment(100,{'对公付款0A流程号':ID(62)}),payment(180)]));
 const weak=audits(next).find(row=>row[8].v===3);assert(weak);assert.match(weak[5].v,/待确认|归属|多条|多个/);for(const token of ['H2','H3','100.00','180.00'])assert(weak[5].v.includes(token),token);
});

test('差额以分计算且零金额有效，空白错误和公式无缓存不得当零核对',()=>{
 uniqueAudit(inputs([payment(0.3)],[payment(0.29)]),{amount:0.3,ledgerAmount:0.29,difference:'0.01'});
 uniqueAudit(inputs([payment(0)],[payment(0.01)]),{amount:0,ledgerAmount:0.01,difference:'-0.01'});
 const zero=C.build(inputs([payment(0)],[payment({t:'n',f:'0',v:0})]));assert.equal(audits(zero).length,0);
 for(const amount of ['',{t:'e',v:23,w:'#REF!'},{t:'n',f:'SUM(A99:A100)'},{t:'z',f:'SUM(A99:A100)',v:0}])for(const side of ['ledger','negotiated']){
  const input=inputs([payment(side==='negotiated'?amount:100)],[payment(side==='ledger'?amount:100)]),result=C.build(input),note=auditText(result);
  assert.equal(audits(result).length,1,side+' '+JSON.stringify(amount));assert.match(note,/金额.*(?:为空|异常|无效|缺失|不可用)|(?:空白|公式|错误).*金额/);
  assert.doesNotMatch(note,/差额[^\n；。]*-?100\.00/,'没有有效金额不产生按零计算的差额');
 }
});

test('申请为空、未提交以及重新申请均不屏蔽台账金额审查，不支付记录跳过该审查',()=>{
 for(const applications of [[],[{'申请日期':'2026-11-18','供应商全称':SUPPLIER,'付款主体':'迈德斯特','流程编码':'未提交OA','本次申请付款金额':100,'付款分类':'谈判付款'}],[{'申请日期':'2026-11-18','供应商全称':SUPPLIER,'付款主体':'迈德斯特','流程编码':'未提交OA','本次申请付款金额':100,'付款分类':'谈判付款','核对说明':'重新申请'}]])uniqueAudit(inputs([payment(100)],[payment(90)],applications));
 for(const marker of ['不支付',' 不 支 付 '])assert.equal(audits(C.build(inputs([payment(100,{'对公付款0A流程号':marker})],[payment(90)]))).length,0);
 uniqueAudit(inputs([payment(100,{'对公付款0A流程号':'不支付待确认'})],[payment(90)]));
});
