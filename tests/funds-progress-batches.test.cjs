const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,set,append,ID,AH,OH,book}=require('./funds-progress-fixtures.cjs');
const CASHIER='[出纳执行]\n部分支付：615833\n姚懿玲 2026-09-29 16:08:44';
const row=(amount,note,date='2026-09-28')=>[date,'合成供应商1','迈德斯特',ID(1),amount,'谈判付款',note];
function batches(rows=[row(615833,'分批付款第1批'),row(700000,'分批付款第2批')],opinion=CASHIER){
 const i=fixture();
 i.applications.book=book('付款申请汇总',[[...AH,'核对说明'],...rows]);
 i.oa.book=book('Sheet0',[[...OH,'业务所属公司'],[ID(1),'合成供应商1',1986240,'归档抄送','合成审核人','归档','OA申请人',opinion,'浙江迈德斯特医疗器械科技有限公司']]);
 for(const h of ['申请付款金额','剩余金额'])set(i,'ledger',h,2,1986240,'对账明细-李奇');
 set(i,'negotiated','申请付款金额',2,1986240);set(i,'negotiated','剩余付款金额',2,{t:'n',v:1986240,f:'D2-E2'});
 return i;
}
const states=r=>r.progressRows.map(row=>row[r.progressHeaders.indexOf('是否已付款')].v);
const progressReviews=r=>r.reviewRows.filter(row=>String(row[r.reviewHeaders.indexOf('核对说明')]?.v||'').startsWith('付款进度核对：'));
const reviewText=r=>progressReviews(r).map(row=>row[r.reviewHeaders.indexOf('核对说明')].v).join('\n');

test('归档分批OA只凭唯一出纳实付金额确认对应申请，其他批保留待核对',()=>{
 const i=batches(),before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(states(r),['已付款','待核对']);assert.equal(r.progressRows.length,2);
 assert.equal(r.progressRows[0][r.progressHeaders.indexOf('申请付款金额')].v,615833);
 assert.match(reviewText(r),/700000|700,000/);assert.match(reviewText(r),/需要处理/);
 assert.equal(r.progressEvidence.length,r.progressRows.length);assert.match(JSON.stringify(r.progressEvidence[0]),/615833|615,833/);
 assert.match(JSON.stringify(r.progressEvidence[0]),/2026-09-29|2026年9月29日/);
 assert.equal(JSON.stringify(i),before,'不改动申请、OA和台账源值');
});

test('同额不同批次不能按出纳事件时间排序分配，即使事件数量足够也待核对',()=>{
 const opinions=[CASHIER,'[出纳执行]\n部分支付：615833\n姚懿玲 2026-10-29 16:08:44'].join('\n');
 const r=C.build(batches([row(615833,'分批付款第1批'),row(615833,'分批付款第2批')],opinions));
 assert.deepEqual(states(r),['待核对','待核对']);assert.match(reviewText(r),/批|归属/);
});

test('同批重复申请均保留并共享同一证据，不能将证据复用于另一个同额批次',()=>{
 const duplicates=[row(615833,'分批付款第1批'),row(615833,'分批付款第1批','2026-09-30')];
 let r=C.build(batches(duplicates));assert.deepEqual(states(r),['已付款','已付款']);assert.equal(r.progressRows.length,2);
 r=C.build(batches([...duplicates,row(615833,'分批付款第2批','2026-10-30')]));
 assert.deepEqual(states(r),['待核对','待核对','待核对']);assert.equal(r.progressRows.length,3);
});

test('未填写批次的相同OA供应商主体金额重复申请仍按同一组保留',()=>{
 const r=C.build(batches([row(615833,''),row(615833,'','2026-09-30')]));
 assert.deepEqual(states(r),['已付款','已付款']);assert.equal(r.progressRows.length,2);
});

test('OA导出重复行和同时间空出纳意见不会制造新的付款批次',()=>{
 const i=batches([row(615833,'分批付款第1批'),row(615833,'分批付款第2批')],CASHIER+'\n[出纳执行]\n姚懿玲 2026-09-29 16:08:44');
 const source=X.utils.sheet_to_json(i.oa.book.Sheets.Sheet0,{header:1})[1];append(i,'oa',source);append(i,'oa',source);
 const r=C.build(i);assert.deepEqual(states(r),['待核对','待核对']);
 const j=batches([row(615833,'分批付款第1批')],source[7]);append(j,'oa',source);
 assert.deepEqual(states(C.build(j)),['已付款'],'相同事件重复导出不应导致唯一实付证据失效');
});

test('申请和审批的支付文字不能冒充实付证据或与出纳金额累计',()=>{
 const opinions='[申请]\n本次申请支付700000元\n李奇 2026-09-28 12:00:00\n[财务总监审批]\n同意支付700000元\n杜海宁 2026-09-28 14:00:00\n'+CASHIER;
 assert.deepEqual(states(C.build(batches(undefined,opinions))),['已付款','待核对']);
 const onlyApproval='[财务总监审批]\n已支付615833元\n杜海宁 2026-09-29 16:08:44';
 assert.deepEqual(states(C.build(batches([row(615833,'分批付款第1批')],onlyApproval))),['待核对']);
});

test('否定、计划、累计、零和负数支付不作为单批实付证据',()=>{
 for(const text of ['未支付：615833','尚未支付615833元','计划支付：615833','拟支付615833元','累计支付：615833','已累计支付615833元','部分支付：0','部分支付：-615833']){
  const opinion='[出纳执行]\n'+text+'\n姚懿玲 2026-09-29 16:08:44';
  const r=C.build(batches([row(615833,'分批付款第1批')],opinion));
  assert.deepEqual(states(r),['待核对'],text);assert.match(reviewText(r),/需要处理/,text);
 }
});

test('OA收款单位或付款主体冲突时不能凭同编号同额意见标为已付款',()=>{
 for(const [header,value] of [['收款单位','另一个供应商'],['业务所属公司','杭州国源养老科技有限公司']]){
  const i=batches([row(615833,'分批付款第1批')]);set(i,'oa',header,2,value);
  const r=C.build(i);assert.deepEqual(states(r),['待核对'],header);assert.match(reviewText(r),/供应商|收款|主体|公司/,header);
 }
 const i=batches([row(615833,'分批付款第1批')]),duplicate=X.utils.sheet_to_json(i.oa.book.Sheets.Sheet0,{header:1})[1];duplicate[1]='另一个供应商';append(i,'oa',duplicate);
 assert.deepEqual(states(C.build(i)),['待核对'],'相同编号下的身份冲突不能取第一条');
});

test('明确分批但没有出纳金额时，归档和办理人执行都不能判为已付款',()=>{
 for(const node of ['归档抄送','办理人执行'])for(const opinion of ['', '[出纳执行]\n同意\n姚懿玲 2026-09-29 16:08:44']){
  const i=batches([row(615833,'分批付款第1批')],opinion);set(i,'oa','当前节点',2,node);
  const r=C.build(i);assert.deepEqual(states(r),['待核对']);assert.match(reviewText(r),/需要处理/);
 }
});

test('实付早于补录申请日期仍允许唯一金额匹配，不用日期先后推断批次',()=>{
 const r=C.build(batches([row(615833,'分批付款第1批','2026-09-30')]));assert.deepEqual(states(r),['已付款']);
});

test('分批金额按分匹配，超过OA总额或意见日期无效时待核对',()=>{
 const opinion='[出纳执行]\n部分支付：1,183.85元\n姚懿玲 2026-09-29 16:08:44';
 assert.deepEqual(states(C.build(batches([row(1183.85,'第1批')],opinion))),['已付款']);
 assert.deepEqual(states(C.build(batches([row(1183.86,'第1批')],opinion))),['待核对']);
 const i=batches([row(1183.85,'第1批')],opinion);set(i,'oa','付款金额',2,1183);
 assert.deepEqual(states(C.build(i)),['待核对']);assert.match(reviewText(C.build(i)),/超过OA付款金额/);
 for(const time of ['2026-02-30 16:08:44','2026-09-29 25:00:00'])assert.deepEqual(states(C.build(batches([row(1183.85,'第1批')],opinion.replace('2026-09-29 16:08:44',time)))),['待核对']);
});

test('重新申请优先于分批实付证据，并继续排除本行付款进度核对',()=>{
 const i=batches([row(615833,'分批付款第1批，重新申请'),row(700000,'分批付款第2批，重新申请')]);
 const r=C.build(i);assert.deepEqual(states(r),['重新申请','重新申请']);assert.equal(progressReviews(r).length,0);
});

test('非分批整笔申请保留原归档、办理人执行与未归档状态规则',()=>{
 for(const [node,status,wanted] of [['财务审批','归档','已付款'],['办理人执行','未归档','已付款'],['财务审批','未归档','未付款']]){
  const i=fixture();i.applications.book=book('付款申请汇总',[AH,['2026-11-02','合成供应商2','迈德斯特',ID(2),200,'正常付款']]);
  set(i,'oa','当前节点',3,node);set(i,'oa','流程状态',3,status);assert.deepEqual(states(C.build(i)),[wanted]);
 }
});

test('分批实付状态不读取台账已付金额，也不改变原申请金额及提交标红',()=>{
 const i=batches(),baseline=C.build(i);
 for(const amount of ['',0,1986240,615833,-1]){
  const j=batches();set(j,'ledger','已付金额',2,amount,'对账明细-李奇');const r=C.build(j);
  assert.deepEqual(states(r),['已付款','待核对'],String(amount));assert.deepEqual(r.highlights,baseline.highlights);
  assert.deepEqual(r.progressRows.map(row=>row[r.progressHeaders.indexOf('申请付款金额')]),baseline.progressRows.map(row=>row[baseline.progressHeaders.indexOf('申请付款金额')]));
 }
});

test('有效出纳支付后出现退汇撤回冲正作废等反证时，不沿用原已付款结果',()=>{
 for(const text of ['银行退汇615833元','付款已撤回','支付冲正615833元','支付不成功','付款作废','款项未到账']){
  const opinion=CASHIER+'\n[出纳执行]\n'+text+'\n姚懿玲 2026-09-30 16:08:44';
  const r=C.build(batches([row(615833,'分批付款第1批')],opinion));
  assert.deepEqual(states(r),['待核对'],text);assert.match(reviewText(r),/需要处理/,text);
  const missingTime=opinion.replace('姚懿玲 2026-09-30 16:08:44','');
  assert.deepEqual(states(C.build(batches([row(615833,'分批付款第1批')],missingTime))),['待核对'],'缺时间的退汇意见也不能被忽略');
 }
});

test('OA总额异常或同编号金额冲突时，不凭单条出纳实付意见绕过核对',()=>{
 const paid40='[出纳执行]\n部分支付40元\n姚懿玲 2026-09-29 16:08:44';
 for(const value of ['',0,{t:'e',v:23,w:'#REF!'},{t:'n',f:'SUM(Z1:Z3)'}]){
  const i=batches([row(40,'分批付款第1批')],paid40);set(i,'oa','付款金额',2,value);
  const r=C.build(i);assert.deepEqual(states(r),['待核对'],JSON.stringify(value));assert.match(reviewText(r),/金额|总额/);
 }
 const i=batches([row(40,'分批付款第1批')],paid40);set(i,'oa','付款金额',2,100);
 const duplicate=X.utils.sheet_to_json(i.oa.book.Sheets.Sheet0,{header:1})[1];duplicate[2]=200;append(i,'oa',duplicate);
 const r=C.build(i);assert.deepEqual(states(r),['待核对']);assert.match(reviewText(r),/金额|总额/);
});
