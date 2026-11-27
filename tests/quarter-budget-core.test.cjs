const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../quarter-budget-core.js');
const H=C.HEADERS,serial=s=>(Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const names={negotiated:'谈判付款付款表付款节奏',ledger:'正常付款',budget:'预估预算',manualBudget:'Sheet1'};
const source=(role,rows,name=role)=>({book:{SheetNames:[names[role]],Sheets:{[names[role]]:X.utils.aoa_to_sheet([H,...rows])}},filename:name+'.xlsx',fingerprint:name});
const row=(change={})=>{const o={month:serial('2026-12-01'),supplier:'合成供应商',subject:'合成主体',oa:'MDST-DGFK20261110001',summary:'合成货款',amount:100,status:'待核对',notes:'',category:'',owner:'合成跟单员',excluded:'否',...change};return [o.month,o.supplier,o.subject,o.oa,o.summary,o.amount,o.status,o.notes,o.category,o.owner,o.excluded];};
const inputs=(n=[],l=[],b=[],m)=>({negotiated:source('negotiated',n),ledger:source('ledger',l),budget:source('budget',b),...(m?{manualBudget:source('manualBudget',m)}:{})});
const build=(i,o={})=>C.build(i,{budgetStartMonth:'2026-12',...o});
const roleRows=(result,role)=>result.records.filter(r=>r.source.role===role);
const blocking=result=>result.issues.filter(i=>i.blocking&&!i.resolved);
function schedule(rows){
 const H=['供应商','付款主体','对公付款0A流程号','谈判明细',serial('2026-12-07'),serial('2026-12-21'),'摘要'];
 const sheet=X.utils.aoa_to_sheet([H,...rows]);sheet.E1.z=sheet.F1.z='yyyy-mm-dd';
 return {book:{SheetNames:['谈判付款'],Sheets:{'谈判付款':sheet}},filename:'schedule.xlsx',fingerprint:'schedule'};
}

test('预算起月优先人工选择、其次手工唯一月份，否则按北京时间下月且跨年',()=>{
 const i=inputs([],[],[],[row({month:serial('2027-02-15')})]);
 assert.equal(C.resolveBudgetStartMonth(i),'2027-02');assert.equal(C.resolveBudgetStartMonth(i,{budgetStartMonth:'2026-12'}),'2026-12');
 assert.equal(C.resolveBudgetStartMonth({}, {now:'2026-12-31T16:01:00Z'}),'2027-02');
 assert.equal(C.resolveBudgetStartMonth({}, {now:'2026-11-30T15:59:00Z'}),'2026-12');
 assert.throws(()=>C.resolveBudgetStartMonth(i,{budgetStartMonth:'2026-13'}),/预算起月/);
 assert.throws(()=>C.resolveBudgetStartMonth({}, {now:'invalid'}),/日期无效/);
});
test('四槽契约拒绝旧模式和旧独立依据，三个主来源必需',()=>{
 assert.throws(()=>build({}),/上传/);
 assert.throws(()=>build({...inputs(),existing:source('manualBudget',[])}),/已取消/);
 assert.throws(()=>build({...inputs(),evidence:source('ledger',[])}),/已取消/);
 assert.throws(()=>build(inputs(),{mode:'existing'}),/已取消/);
 assert.equal(build(inputs()).canExport,true);assert.equal(build(inputs()).comparison,null);
});
test('跨来源按OA、供应商、主体、年月、金额逐笔去重，保留最多来源的分期次数和零金额',()=>{
 const result=build(inputs([row(),row({amount:0,oa:'ZERO001'})],[row(),row(),row({amount:0,oa:'ZERO001'})],[row(),row(),row()]));
 assert.equal(result.records.length,4);assert.equal(result.deleted.length,4);assert.deepEqual(result.records.map(r=>r.cents).sort((a,b)=>a-b),[0,10000,10000,10000]);
 assert.equal(result.summary.totals.totalCents,30000);assert.equal(roleRows(result,'negotiated').length,2);
 assert.equal(new Set(result.records.map(r=>r.id)).size,4);
});
test('相同OA与金额但月份、供应商、主体或明确摘要不同不自动删除',()=>{
 for(const change of [{month:serial('2027-01-01')},{supplier:'不同供应商'},{subject:'不同主体'},{summary:'另一笔货款'}]){
  const result=build(inputs([row()],[row(change)]));assert.equal(result.deleted.length,0,JSON.stringify(change));assert.equal(result.records.length,2);
 }
 const result=build(inputs([row({month:null})],[row({month:null})]));assert.equal(result.deleted.length,0,'未知月份不能作为相同年月');
});
test('主体全称与简称可核对，不跨供应商合并，去重明细保留双方来源',()=>{
 const result=build(inputs([row({subject:'浙江迈德斯特医疗器械科技有限公司'})],[row({subject:'迈德斯特'})]));
 assert.equal(result.deleted.length,1);assert.equal(result.deleted[0].kept.length,1);assert.equal(result.deleted[0].record.source.role,'ledger');
 assert.match(result.records[0].notes,/来源：negotiated.xlsx/);assert.match(result.records[0].notes,/第2行/);
});
test('同一物理文件上传不同槽位不会因记录ID相同误删保留项',()=>{
 const one=source('budget',[row()],'same-file'),i=inputs();i.negotiated=one;i.budget=one;
 const result=build(i);assert.equal(result.records.length,1);assert.equal(result.deleted.length,1);assert.notEqual(result.records[0].id,result.deleted[0].record.id);assert.equal(result.records[0].source.role,'negotiated');
});
test('手工普通记录仅作对比、预估分类或明确备注才纳入；原分类及来源保留',()=>{
 const manual=[row({amount:500,status:'已支付',category:'正常付款-正常对账表'}),row({oa:'',amount:2000000,supplier:'多家供应商',category:'正常付款-预估付款'}),row({oa:'',amount:14892,supplier:'预估供应商',category:'正常付款-正常对账表',notes:'此款为预估'})];
 const input=inputs([],[],[],manual),before=JSON.stringify(input),result=build(input);
 assert.equal(result.records.length,2);assert.equal(result.summary.totals.totalCents,201489200);assert.equal(result.summary.totals.paidCents,0);
 assert(result.records.every(r=>r.category==='正常付款-预估付款'));assert.match(result.records[1].notes,/原分类：正常付款-正常对账表/);assert.match(result.records[1].notes,/manualBudget.xlsx.*Sheet1.*第4行/);
 assert.equal(result.comparison.totals.manualCents,201539200);assert.equal(result.comparison.totals.systemCents,201489200);assert.equal(result.comparison.totals.differenceCents,-50000);assert.equal(JSON.stringify(input),before);
});
test('无OA补充预估重复必须逐笔确认，仍保留多笔须说明理由',()=>{
 const i=inputs([],[],[row({oa:'',amount:14892})],[row({oa:'',amount:14892,category:'正常付款-正常对账表',notes:'此款为预估'})]);
 let result=build(i);assert.equal(result.records.length,2);assert.equal(result.deleted.length,0);assert.equal(result.canExport,false);assert.equal(blocking(result).filter(i=>i.type==='duplicate').length,1);
 const [a,b]=result.records;result=build(i,{decisions:{[a.id]:{excluded:'否'},[b.id]:{excluded:'是',reason:'与预估模板同笔'}}});
 assert.equal(result.canExport,true);assert.equal(result.summary.totals.totalCents,1489200);assert.equal(result.comparison.totals.differenceCents,0);assert.match(result.records[1].notes,/与预估模板同笔/);
 result=build(i,{decisions:{[a.id]:{excluded:'否'},[b.id]:{excluded:'否'}}});assert.equal(result.canExport,false);
 result=build(i,{decisions:{[a.id]:{excluded:'否',reason:'第一笔独立订单'},[b.id]:{excluded:'否',reason:'第二笔独立订单'}}});assert.equal(result.canExport,true);assert.equal(result.summary.totals.totalCents,2978400);
 result=build(i);assert.equal(result.canExport,false,'清空选择重新阻止下载');
});
test('不同已知月份的无OA同额款不产生重复确认；零金额仍正常显示',()=>{
 const result=build(inputs([],[],[row({oa:'',amount:0})],[row({oa:'',amount:0,month:serial('2027-01-01'),category:'正常付款-预估付款'})]));
 assert.equal(result.records.length,2);assert.equal(result.canExport,true);assert(!result.issues.some(i=>i.type==='duplicate'));assert.equal(result.summary.totals.totalCents,0);
});
test('空金额、错误缓存不按0统计，可明确剔除并说明原因；金额issue不重复',()=>{
 for(const cell of [{t:'s',v:''},{t:'n',f:'SUM(Z1:Z2)'},{t:'e',v:15,w:'#VALUE!'}]){
  const i=inputs([],[],[row({amount:10})]);i.budget.book.Sheets.预估预算.F2=cell;
  let result=build(i);const r=result.records[0];assert.equal(result.canExport,false);assert.equal(result.summary.totals.totalCents,null);assert.equal(blocking(result).filter(i=>i.type==='amount').length,1);assert.equal(blocking(result).filter(i=>i.type==='error').length,0);assert.match(blocking(result)[0].message,/F2/);
  result=build(i,{decisions:{[r.id]:{excluded:'是'}}});assert.equal(result.canExport,false);
  result=build(i,{decisions:{[r.id]:{excluded:'是',reason:'源金额待补，本次暂不纳入'}}});assert.equal(result.canExport,true);assert.equal(result.summary.totals.totalCents,0);
 }
});
test('手工对比基准缺金额保持null且阻止下载，系统侧剔除不能修改手工基准',()=>{
 const i=inputs([],[],[],[row({amount:null,category:'正常付款-预估付款'})]);let result=build(i),r=result.records[0];
 assert.equal(result.comparison.totals.manualCents,null);assert.equal(result.comparison.rows[0].differenceCents,null);assert.equal(result.comparison.rows[0].status,'金额待核对');assert.equal(result.canExport,false);
 result=build(i,{decisions:{[r.id]:{excluded:'是',reason:'暂不计入系统预算'}}});assert.equal(result.comparison.totals.systemCents,0);assert.equal(result.comparison.totals.manualCents,null);assert.equal(result.canExport,false);
});
test('对比按月份供应商主体分类合计，差额系统减手工，缺组0且包含一致组',()=>{
 const n=[row({supplier:'甲',amount:0.1,category:'正常付款-谈判付款表'}),row({supplier:'甲',amount:0.2,category:'正常付款-谈判付款表'})];
 const l=[row({supplier:'乙',amount:20,category:'正常付款-正常对账表'})];
 const m=[row({supplier:'甲',amount:0.3,category:'正常付款-谈判付款表'}),row({supplier:'丙',amount:5,category:'谈判付款-谈判付款表'})];
 const result=build(inputs(n,l,[],m)),comparison=result.comparison;
 assert.equal(comparison.rows.length,3);assert.deepEqual(comparison.totals,{systemCents:2030,manualCents:530,differenceCents:1500});
 const a=comparison.rows.find(r=>r.supplier==='甲'),b=comparison.rows.find(r=>r.supplier==='乙'),c=comparison.rows.find(r=>r.supplier==='丙');
 assert.deepEqual([a.systemCents,a.manualCents,a.differenceCents,a.status],[30,30,0,'一致']);assert.equal(a.systemSources.length,2);assert.equal(a.manualSources.length,1);
 assert.deepEqual([b.systemCents,b.manualCents,b.status],[2000,0,'仅系统有']);assert.deepEqual([c.systemCents,c.manualCents,c.status],[0,500,'仅手工有']);assert.equal(comparison.rows.at(-1).status,'一致');
});
test('对比主体别名一致但分类独立、其他月份及未知月份提示且不混入本月',()=>{
 const result=build(inputs([row({subject:'浙江迈德斯特医疗器械科技有限公司',category:'正常付款-谈判付款表'})],[],[],[
  row({subject:'迈德斯特',category:'正常付款-正常对账表'}),row({month:serial('2027-01-01'),category:'正常付款-正常对账表'}),row({month:null,category:'正常付款-正常对账表'})
 ]));
 assert.equal(result.comparison.rows.length,2);assert.equal(result.comparison.totals.manualCents,10000);assert.equal(result.comparison.warnings.length,2);assert(result.comparison.rows.every(r=>r.subject==='迈德斯特'));
});
test('仅月份按预算起月顺延，明确年份不变，空月份仍计入单列而不当零',()=>{
 const result=build(inputs([],[],[row({month:'1月',oa:'A001'}),row({month:'2026.4月',oa:'A002'}),row({month:null,oa:'A003'})]));
 assert.deepEqual(result.records.map(r=>r.monthKey),['2027-01','2026-04','']);assert.equal(result.summary.totals.totalCents,30000);assert(result.summary.monthly.some(r=>r.monthKey===''&&r.totalCents===10000));assert.equal(result.canExport,true);
});
test('人工剔除同步影响系统对比，手工原金额与备注保持不变',()=>{
 const i=inputs([row({category:'谈判付款-谈判付款表'})],[],[],[row({category:'谈判付款-谈判付款表',notes:'原手工备注'})]);const first=build(i),id=first.records[0].id;
 const result=build(i,{decisions:{[id]:{excluded:'是',reason:'不再支付'}}});assert.equal(result.comparison.totals.systemCents,0);assert.equal(result.comparison.totals.manualCents,10000);assert.equal(result.comparison.totals.differenceCents,-10000);assert.equal(i.manualBudget.book.Sheets.Sheet1.H2.v,'原手工备注');
});
test('同原表相同OA同月同额各分期按自己的金额格状态保留',()=>{
 const n=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','12月10日-100；12月25日-100',100,100,'合成货款']]);n.book.Sheets.谈判付款.E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const result=build({...inputs(),negotiated:n});assert.deepEqual(result.records.map(r=>r.status),['已支付','未支付']);assert.equal(result.summary.totals.paidCents,10000);
});
test('同一已付依据不能同时支付两条没有独立依据的记录',()=>{
 const result=build(inputs([],[row({oa:'',status:'已支付'})],[row({oa:''}),row({oa:''})]));
 assert.deepEqual(roleRows(result,'budget').map(r=>r.status),['待核对','待核对']);assert.equal(result.summary.totals.paidCents,10000);
});
test('无月份的台账汇总不能作为同一排期之外的第二份付款证据',()=>{
 const n=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','12月10日-100',100,null,'合成货款']]);n.book.Sheets.谈判付款.E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const result=build({...inputs([],[row({month:null,status:'已支付'})]),negotiated:n});assert.equal(result.records.length,2);assert.equal(result.summary.totals.paidCents,0);assert(result.records.every(r=>r.status==='待核对'));
});
test('弱待核对状态不能覆盖本行明确已支付，手工对比的已付也不是状态依据',()=>{
 let result=build(inputs([row({status:'已支付'})],[row()]));assert.equal(result.records[0].status,'已支付');
 result=build(inputs([row()],[],[],[row({status:'已支付',category:'谈判付款-谈判付款表'})]));assert.equal(result.records[0].status,'待核对');assert.equal(result.summary.totals.paidCents,0);
});
test('台账明确已付与同月未付排期冲突仍待核对',()=>{
 const n=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','12月10日-100',100,null,'合成货款']]);
 const result=build({...inputs([],[row({status:'已支付'})]),negotiated:n});assert.equal(result.records[0].status,'待核对');assert.match(result.records[0].notes,/台账.*冲突/);
});
test('失效付款依据选择提示且不能沿用旧选择',()=>{
 const i=inputs([row()]),r=build(i).records[0];const result=build(i,{matches:{[r.id]:'stale-id'}});assert.equal(result.records[0].status,'待核对');assert(result.warnings.some(w=>w.includes('失效')));
});
test('合并金额包含已付及未付部分时不能认定整笔已付',()=>{
 const result=build(inputs([],[row({amount:100,status:'已支付'}),row({amount:200,status:'未支付'})],[row({oa:'',amount:300,notes:'100.00+200.00',status:'未支付'})]));
 const combined=roleRows(result,'budget')[0];assert.equal(combined.status,'待核对');assert.match(combined.notes,/合并金额/);
});
test('其他月份已付分期不会支付本月同额记录，有无OA均适用',()=>{
 for(const oa of ['MDST-DGFK20261110001','']){
  const result=build(inputs([row({oa,status:'已支付'})],[],[row({oa,month:serial('2027-01-01')})]));const target=roleRows(result,'budget')[0];assert.equal(target.status,'待核对');assert.match(target.notes,/月份不同/);
 }
});
test('明确不同摘要的同OA同额分期不能借用其他分期已付状态',()=>{
 const n=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','12月10日-100',100,null,'第一笔']]);
 const result=build({...inputs([],[row({summary:'第二笔',status:'已支付'})]),negotiated:n});assert.equal(result.records.length,2);assert.deepEqual(result.records.map(r=>r.status),['未支付','已支付']);assert.equal(result.summary.totals.paidCents,10000);
});
test('明确单笔状态不按网页行序交换，重复同额各有独立依据',()=>{
 const result=build(inputs([row({summary:'第二笔',status:'未支付'}),row({summary:'第一笔',status:'已支付'})]));
 assert.deepEqual(result.records.map(r=>[r.summary,r.status]),[['第二笔','未支付'],['第一笔','已支付']]);
});
test('来源追加不覆盖原备注，重复计算不累积备注',()=>{
 const i=inputs([row({notes:'保留原说明\n分期付款条件'})]);const a=build(i),b=build(i);assert.equal(a.records[0].notes,b.records[0].notes);assert.match(a.records[0].notes,/^保留原说明\n分期付款条件\n来源：/);assert.equal(i.negotiated.book.Sheets[names.negotiated].H2.v,'保留原说明\n分期付款条件');
});
test('汇总及对比不会静默超出安全整数分，空结果为明确零',()=>{
 assert.throws(()=>C.summarize([{monthKey:'2026-12',category:'测试',cents:Number.MAX_SAFE_INTEGER,status:'未支付'},{monthKey:'2026-12',category:'测试',cents:1,status:'未支付'}]),/范围/);
 const result=build(inputs([],[],[],[]));assert.deepEqual(result.comparison.totals,{systemCents:0,manualCents:0,differenceCents:0});assert.equal(result.canExport,true);
});
test('页面分组与预算对比共用年月、供应商、主体及分类归一化口径',()=>{
 const record={monthKey:'2026-12',supplier:' 合成 供应商 ',subject:'浙江迈德斯特医疗器械科技有限公司',category:'正常付款-预估付款'};
 const group={...record,supplier:'合成供应商',subject:'迈德斯特'};
 assert.equal(C.comparisonKey(record),C.comparisonKey(group));
 for(const changed of [{monthKey:'2027-01'},{supplier:'另一供应商'},{subject:'国源'},{category:'谈判付款-谈判付款表'}])assert.notEqual(C.comparisonKey(record),C.comparisonKey({...group,...changed}));
});

test('同金额两来源一行确认合并，只计选择来源一次且保留原记录与备注',()=>{
 const i=inputs([],[],[row({oa:'',notes:'原预估备注'})],[row({oa:'',category:'正常付款-预估付款',notes:'手工原备注'})]),result=build(i),before=JSON.stringify(result);
 const groups=C.duplicateReviewGroups(result);assert.equal(groups.length,1);assert.equal(groups[0].records.length,2);assert.equal(groups[0].canMerge,true);assert.equal(groups[0].cents,10000);
 const keep=groups[0].records[1],decisions=C.applyDuplicateReviewDecision(result,{},groups[0].id,'keep:'+keep.id),next=build(i,{decisions});
 assert.deepEqual(C.applyDuplicateReviewDecision(result,decisions,groups[0].id,'keep:'+keep.id),decisions);
 assert.equal(next.canExport,true);assert.equal(next.summary.totals.totalCents,10000);assert.equal(next.records.filter(r=>r.excluded!=='是')[0].id,keep.id);assert.equal(next.records.length,2);
 assert.match(decisions[groups[0].records[0].id].reason,/manualBudget.xlsx.*Sheet1.*第2行/);assert.match(next.records[1].notes,/手工原备注/);assert.match(next.records[0].notes,/原预估备注/);assert.equal(JSON.stringify(result),before);
});
test('同金额三来源跨分类也一行，可任选一个来源保留',()=>{
 const i=inputs([],[row({oa:'',category:'正常付款-正常对账表'})],[row({oa:'',category:'正常付款-预估付款'})],[row({oa:'',category:'正常付款-预估付款'})]),result=build(i),group=C.duplicateReviewGroups(result)[0];
 assert.equal(C.duplicateReviewGroups(result).length,1);assert.equal(group.records.length,3);assert.equal(group.canMerge,true);
 for(const kept of group.records){
  const decisions=C.applyDuplicateReviewDecision(result,{},group.id,'keep:'+kept.id),next=build(i,{decisions});
  assert.equal(next.canExport,true);assert.equal(next.summary.totals.totalCents,10000);assert.equal(next.records.filter(r=>r.excluded!=='是')[0].id,kept.id);
 }
});
test('不合并全部保留，已有原因保留且重复选择不累积确认文本；重置不影响其他组',()=>{
 const i=inputs([],[],[row({oa:''})],[row({oa:'',category:'正常付款-预估付款'})]),result=build(i),group=C.duplicateReviewGroups(result)[0];
 const old={other:{excluded:'否',reason:'其他组原选择'},[group.records[0].id]:{excluded:'否',reason:' 已经人工写下的原因 '}};
 const oldSnapshot=JSON.stringify(old),next=C.applyDuplicateReviewDecision(result,old,group.id,'separate');
 assert.equal(JSON.stringify(old),oldSnapshot);assert.notEqual(next,old);assert.notEqual(next.other,old.other);assert.deepEqual(next.other,old.other);
 assert.equal(next[group.records[0].id].reason,' 已经人工写下的原因 \n人工确认不同款项，分别保留');assert(group.records.every(r=>next[r.id].excluded==='否'));
 assert.deepEqual(C.applyDuplicateReviewDecision(result,next,group.id,'separate'),next);assert.equal(build(i,{decisions:next}).summary.totals.totalCents,20000);assert.equal(build(i,{decisions:next}).canExport,true);
 const reset=C.applyDuplicateReviewDecision(result,next,group.id,'');assert.deepEqual(reset,{other:old.other});assert.equal(build(i,{decisions:reset}).canExport,false);
});
test('一个大型关联问题按金额拆行，单桶确认不会自动确认其他金额',()=>{
 const i=inputs([],[row({oa:'',amount:403})],[row({oa:'',amount:403}),row({oa:'',amount:1134})]),result=build(i),groups=C.duplicateReviewGroups(result);
 assert.equal(result.issues.filter(i=>i.type==='duplicate').length,1);assert.equal(groups.length,2);
 const small=groups.find(g=>g.cents===40300),single=groups.find(g=>g.cents===113400);assert.equal(small.records.length,2);assert.equal(small.canMerge,true);assert.equal(single.records.length,1);assert.equal(single.canMerge,false);assert.match(single.restriction,/仅一条/);
 let decisions=C.applyDuplicateReviewDecision(result,{},small.id,'keep:'+small.records[0].id);assert(!decisions[single.records[0].id]);assert.equal(build(i,{decisions}).canExport,false);
 decisions=C.applyDuplicateReviewDecision(result,decisions,single.id,'separate');assert.equal(build(i,{decisions}).canExport,true);assert.equal(build(i,{decisions}).summary.totals.totalCents,153700);
 assert.deepEqual(C.duplicateReviewGroups(build(i,{decisions})).map(g=>g.id),groups.map(g=>g.id));
 const reset=C.applyDuplicateReviewDecision(result,decisions,small.id,'');assert.equal(reset[single.records[0].id].excluded,'否');assert.equal(build(i,{decisions:reset}).canExport,false);
});
test('快捷分组归一供应商主体但不混入不同供应商主体月份或金额，分类不拆组',()=>{
 const base=build(inputs([],[],[row({oa:''})])).records[0];
 const records=[{...base,id:'a',supplier:' 合成 供应商 ',subject:'浙江迈德斯特医疗器械科技有限公司',source:{...base.source,role:'ledger'}},{...base,id:'b',subject:'迈德斯特',category:'另一分类'},
  {...base,id:'c',supplier:'另一供应商'},{...base,id:'d',subject:'国源'},{...base,id:'e',monthKey:'2027-01'},{...base,id:'f',cents:20000}];
 const result={records,issues:[{id:'mixed',type:'duplicate',recordIds:records.map(r=>r.id)}]},groups=C.duplicateReviewGroups(result);
 assert.equal(groups.length,5);assert.deepEqual(groups[0].records.map(r=>r.id),['a','b']);assert.equal(groups[0].canMerge,true);assert(groups.slice(1).every(g=>!g.canMerge));
});
test('零金额允许快捷合并且仍按0显示',()=>{
 const i=inputs([],[],[row({oa:'',amount:0})],[row({oa:'',amount:0,category:'正常付款-预估付款'})]),result=build(i),group=C.duplicateReviewGroups(result)[0];
 assert.equal(group.cents,0);assert.equal(group.canMerge,true);const decisions=C.applyDuplicateReviewDecision(result,{},group.id,'keep:'+group.records[0].id);assert.equal(build(i,{decisions}).canExport,true);assert.equal(build(i,{decisions}).summary.totals.totalCents,0);
});
test('空金额只能逐笔核对，不合并选择不会绕过金额门禁',()=>{
 const i=inputs([],[row({oa:'',amount:null})],[row({oa:'',amount:null})]),result=build(i),group=C.duplicateReviewGroups(result)[0];
 assert.equal(group.cents,null);assert.equal(group.canMerge,false);assert.match(group.restriction,/金额未确定/);assert.throws(()=>C.applyDuplicateReviewDecision(result,{},group.id,'keep:'+group.records[0].id),/金额未确定/);
 const decisions=C.applyDuplicateReviewDecision(result,{},group.id,'separate');assert.equal(build(i,{decisions}).canExport,false);
});
test('不同有效OA与同来源多笔均禁止快捷合并，确认不合并不改金额',()=>{
 for(const [i,pattern,count] of [[inputs([],[row({oa:'OA001'})],[row({oa:'OA002'})]),/不同OA/,2],[inputs([],[row({oa:''})],[row({oa:''}),row({oa:''})]),/同一来源/,3]]){
  const result=build(i),group=C.duplicateReviewGroups(result)[0];assert.equal(group.canMerge,false);assert.match(group.restriction,pattern);assert.throws(()=>C.applyDuplicateReviewDecision(result,{},group.id,'keep:'+group.records[0].id),pattern);
  const decisions=C.applyDuplicateReviewDecision(result,{},group.id,'separate'),next=build(i,{decisions});assert.equal(next.summary.totals.totalCents,count*10000);assert.equal(next.canExport,true);
 }
});
test('未知身份、月份及来源类型不能快捷合并',()=>{
 const base=build(inputs([],[],[row({oa:''})])).records[0];
 for(const override of [{supplier:' '},{supplier:'待核对'},{supplier:'多家供应商'},{supplier:'多个供应商'},{subject:'/'},{monthKey:''},{monthKey:'2026-13'},{source:{...base.source,role:'unknown'}}]){
  const a={...base,id:'a',...override},b={...base,id:'b',...override,source:{...base.source,role:'manualBudget'}};
  const result={records:[a,b],issues:[{id:'missing',type:'duplicate',recordIds:['a','b']}]},group=C.duplicateReviewGroups(result)[0];assert.equal(group.canMerge,false,JSON.stringify(override));assert.throws(()=>C.applyDuplicateReviewDecision(result,{},group.id,'keep:a'));
 }
});
test('快捷选择校验组与记录归属，错误输入不修改旧决策与原结果',()=>{
 const i=inputs([],[row({oa:'',amount:403})],[row({oa:'',amount:403}),row({oa:'',amount:1134})]),result=build(i),groups=C.duplicateReviewGroups(result),old={other:{excluded:'是',reason:'原原因'}};
 const beforeResult=JSON.stringify(result),beforeOld=JSON.stringify(old);
 assert.throws(()=>C.applyDuplicateReviewDecision(result,old,'stale','separate'),/分组已变化/);
 assert.throws(()=>C.applyDuplicateReviewDecision(result,old,groups[0].id,'keep:'+groups[1].records[0].id),/不属于/);
 for(const choice of ['merge',null,{}])assert.throws(()=>C.applyDuplicateReviewDecision(result,old,groups[0].id,choice),/请选择/);
 assert.equal(JSON.stringify(result),beforeResult);assert.equal(JSON.stringify(old),beforeOld);assert.deepEqual(C.duplicateReviewGroups(null),[]);
});
