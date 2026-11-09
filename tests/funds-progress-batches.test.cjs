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

function redBatch(){
 const i=batches(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];
 s.M2={t:'n',v:615833};s.N2={t:'n',v:700000,s:{patternType:'solid',fgColor:{rgb:'FF0000'}}};
 return i;
}
const CONFIRMED='11月3日-OA已评论\n11月5日核实-已支付\n申请日期（付款节奏跟进）：2026年11月2日\n11月5日核实-已支付';
function commentBatch(text=CONFIRMED){const i=redBatch();i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]].N2.c=[{a:'人工核实',t:text}];return i;}
test('唯一分批批注明确核实已支付时确认本批，保留金额公式与人工内容',()=>{
 for(const red of [true,false]){
  const i=commentBatch();if(!red)delete i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]].N2.s;
  const before=JSON.stringify(i),r=C.build(i),baseline=C.build(redBatch());
  assert.deepEqual(states(r),['已付款','已付款']);assert.equal(progressReviews(r).length,0);
  assert.match(r.progressEvidence[1],/谈判批注付款依据.*11月5日核实-已支付.*negotiated.xlsx.*第 2 行.*N2/);
  assert.equal((r.progressEvidence[1].match(/11月5日核实-已支付/g)||[]).length,1);
  for(const key of ['normalRows','highlights','applicationComments'])assert.deepEqual(r[key],baseline[key]);
  assert.deepEqual(r.negotiated.rows.map(row=>row.map(({v,f})=>({v,f}))),baseline.negotiated.rows.map(row=>row.map(({v,f})=>({v,f}))));assert.equal(JSON.stringify(i),before);
 }
 for(const text of ['已付款','2026年11月5日-确认：已付清','11/5 出纳核实-本批已支付700,000.00元'])assert.deepEqual(states(C.build(commentBatch(text))),['已付款','已付款'],text);
});
test('评论、申请日期、疑问、计划与否定文字不冒充支付确认，最新批注移除后重新判断',()=>{
 for(const text of ['11月3日-OA已评论','申请日期（付款节奏跟进）：2026年11月2日','申请日期（付款节奏跟进）：已支付','11月5日核实-未支付','计划已支付','预计已支付700000元','是否已支付','已支付？','如果已支付请反馈','已申请付款','已部分支付211820元'])assert.equal(states(C.build(commentBatch(text)))[1],'待核对',text);
 const i=commentBatch();assert.equal(states(C.build(i))[1],'已付款');delete i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]].N2.c;
 assert.equal(states(C.build(i))[1],'待核对');
});
test('批注金额不一致、退汇或相互矛盾时继续核对，不被红色或OA实付证据覆盖',()=>{
 for(const text of [CONFIRMED+'\n11月6日未支付',CONFIRMED+'\n11月6日银行退汇','已支付699999.99元','已支付0元','已支付-700000元','已支付700000.001元','部分支付700000元']){
  const i=commentBatch(text);set(i,'oa','签字意见',2,CASHIER+'\n[出纳执行]\n部分支付700000元\n姚懿玲 2026-11-05 16:00:00');
  const r=C.build(i);assert.equal(states(r)[1],'待核对',text);assert.match(reviewText(r),/谈判批注/);assert.match(reviewText(r),/N2/);
 }
});
test('批注只作用于对应OA供应商主体和唯一金额批次，不借用其他批次或普通付款',()=>{
 for(const change of [
  i=>{const s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];s.M2.c=s.N2.c;delete s.N2.c;},
  i=>{const s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];s.M2.v=700000;},
  i=>append(i,'applications',row(700000,'分批付款第3批')),
  i=>set(i,'applications','付款分类',3,'正常付款'),
  i=>set(i,'negotiated','付款主体',2,'国源'),
  i=>set(i,'negotiated','供应商名称',2,'其他供应商'),
  i=>{set(i,'negotiated','对公付款0A流程号',2,ID(99));const flow=X.utils.sheet_to_json(i.oa.book.Sheets.Sheet0,{header:1})[1];flow[0]=ID(99);append(i,'oa',flow);},
  i=>set(i,'oa','付款金额',2,0),
  i=>set(i,'oa','业务所属公司',2,'杭州国源养老科技有限公司'),
  i=>set(i,'oa','签字意见',2,CASHIER+'\n[出纳执行]\n付款已撤回\n姚懿玲 2026-11-06 16:00:00')
 ]){const i=commentBatch();change(i);assert.equal(states(C.build(i))[1],'待核对');}
});
test('同批重复申请共享批注依据，重复申请核对保留；重新申请优先',()=>{
 const i=commentBatch();append(i,'applications',row(700000,'分批付款第2批','2026-11-03'));
 const r=C.build(i);assert.deepEqual(states(r),['已付款','已付款','已付款']);assert(r.reviewRows.some(row=>row[5].v.includes('重复提交')));
 set(i,'applications','核对说明',3,'重新申请');assert.equal(states(C.build(i))[1],'重新申请');
});
test('同额多批以原批注申请日期唯一定位，日期缺失或重复不猜测批次',()=>{
 const i=commentBatch();set(i,'applications','申请日期',3,'2026-11-02');
 const s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];
 X.utils.sheet_add_aoa(s,[[46342],[700000]],{origin:'O1'});s.O1.z=C.DATE_FORMAT;
 let r=C.build(i);assert.equal(states(r)[1],'已付款');assert.match(r.progressEvidence[1],/N2/);
 s.O2.c=[{a:'人工',t:'申请日期（付款节奏跟进）：2026年11月2日'}];
 assert.equal(states(C.build(i))[1],'待核对','同日期指向两批无法唯一定位');
 delete s.O2.c;set(i,'applications','申请日期',3,'2026-11-16');
 assert.equal(states(C.build(i))[1],'待核对','另一批申请不能复用11月2日的确认');
 set(i,'applications','申请日期',3,'');assert.equal(states(C.build(i))[1],'待核对');
 s.N2.c=[{a:'人工',t:'11月5日核实-已支付'}];set(i,'applications','申请日期',3,'2026-11-02');
 assert.equal(states(C.build(i))[1],'待核对','没有原申请日期批注时不按最早金额猜测已付');
});
test('源谈判分批已红且本次有对应申请时不重复核对，付款状态和其他结果保持不变',()=>{
 const i=redBatch(),before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(states(r),['已付款','待核对']);assert.equal(progressReviews(r).length,0);
 assert.match(r.progressEvidence[1],/已标红.*对应申请/);assert.equal(JSON.stringify(i),before);
 const j=redBatch();delete j.negotiated.book.Sheets[j.negotiated.book.SheetNames[0]].N2.s;
 const baseline=C.build(j);assert.match(reviewText(baseline),/700000\.00/);
 for(const key of ['normalRows','highlights','applicationComments'])assert.deepEqual(r[key],baseline[key]);
 assert.deepEqual(r.negotiated.rows.map(row=>row.map(({v,f})=>({v,f}))),baseline.negotiated.rows.map(row=>row.map(({v,f})=>({v,f}))));
 append(i,'applications',row(700000,'分批付款第2批','2026-11-03'));
 const repeated=C.build(i);assert.equal(progressReviews(repeated).length,0);assert.deepEqual(states(repeated),['已付款','待核对','待核对']);
 assert(repeated.reviewRows.some(row=>row[5].v.includes('重复提交')),'重复申请核对独立保留');
});
test('新增标红、红色在其他批次或其他OA、无申请和普通付款不豁免出纳核对',()=>{
 for(const change of [
  i=>{delete i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]].N2.s;},
  i=>{const s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];s.M2.s=s.N2.s;delete s.N2.s;},
  i=>{set(i,'negotiated','对公付款0A流程号',2,ID(99));const flow=X.utils.sheet_to_json(i.oa.book.Sheets.Sheet0,{header:1})[1];flow[0]=ID(99);append(i,'oa',flow);},
  i=>{set(i,'applications','付款分类',3,'正常付款');},
  i=>{set(i,'negotiated','付款主体',2,'国源');},
  i=>{set(i,'negotiated','供应商名称',2,'其他供应商');}
 ]){const i=redBatch();change(i);const r=C.build(i);assert.match(reviewText(r),/700000\.00/);assert.deepEqual(states(r),['已付款','待核对']);}
 const i=redBatch();i.applications.book=book('付款申请汇总',[[...AH,'核对说明'],row(615833,'第1批')]);
 const r=C.build(i);assert.equal(r.progressRows.length,1);assert(!r.highlights.negotiatedCells.includes('N2'));
});
test('已有红色及对应申请不能绕过主体、金额、多批、退汇等实质冲突',()=>{
 for(const change of [
  i=>set(i,'oa','业务所属公司',2,'杭州国源养老科技有限公司'),
  i=>set(i,'oa','付款金额',2,0),
  i=>append(i,'applications',row(700000,'分批付款第3批')),
  i=>set(i,'oa','签字意见',2,CASHIER+'\n[出纳执行]\n付款已撤回\n姚懿玲 2026-10-30 16:08:44'),
  i=>{const s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];s.M2={...s.N2};s.M1={...s.N1};}
 ]){const i=redBatch();change(i);assert.match(reviewText(C.build(i)),/700000\.00/);}
 const i=redBatch();set(i,'negotiated','应付款时间',2,'无法识别');
 const r=C.build(i);assert.match(reviewText(r),/应付时间/);assert.doesNotMatch(reviewText(r),/未找到与本笔申请金额相等/);
});

test('归档分批OA只凭唯一出纳实付金额确认对应申请，其他批保留待核对',()=>{
 const i=batches(),before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(states(r),['已付款','待核对']);assert.equal(r.progressRows.length,2);
 assert.equal(r.progressRows[0][r.progressHeaders.indexOf('申请付款金额')].v,615833);
 assert.match(reviewText(r),/700000|700,000/);assert.match(reviewText(r),/需要处理/);
 assert.equal(r.progressEvidence.length,r.progressRows.length);assert.match(JSON.stringify(r.progressEvidence[0]),/615833|615,833/);
 assert.match(JSON.stringify(r.progressEvidence[0]),/2026-09-29|2026年9月29日/);
 assert.equal(JSON.stringify(i),before,'不改动申请、OA和台账源值');
});

test('申请1090000但出纳仅支付211820时提醒确认分批，台账未匹配不附加缺失字段提示',()=>{
 const i=batches([row(1090000,'分批付款第2批')],'[出纳执行]\n部分支付 211820\n姚懿玲 2026-10-08 15:43:37');
 const id='MDST-DGFK20260924004';set(i,'applications','流程编码',2,id);set(i,'oa','流程编号',2,id);set(i,'negotiated','对公付款0A流程号',2,id);
 const r=C.build(i),note=reviewText(r);
 assert.deepEqual(states(r),['待核对']);assert.match(note,/请先确认是否为分批付款/);assert.match(note,/1090000\.00/);assert.match(note,/211820/);
 assert.match(note,/2026-10-08 15:43:37/);assert.match(note,/Sheet0 第 2 行/);assert.match(note,/本次申请对应的批次及是否实际支付/);
 assert.doesNotMatch(note,/跟单员\/资金台账|应付时间未填充|需要处理：分批付款核对/);
 assert.doesNotMatch(C.formatReviewText(note),/^\d+、[。；]+$/m);
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
