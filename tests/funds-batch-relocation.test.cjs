const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {sample,ID,prefixes}=require('./funds-batch-relocation-fixtures.cjs');
const neg=i=>i.negotiated.book.Sheets.谈判付款,apps=i=>i.applications.book.Sheets.付款申请汇总;
const match=r=>r.applicationBatchMatches['付款申请汇总!3'];
const notes=r=>r.reviewRows.map(row=>row[5].v).filter(t=>t.startsWith('申请批次OA号核对'));
test('旧OA指向15230，新批49990唯一对应另一OA：匹配新批、留原编号及历史批注',()=>{
 const i=sample(),before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(match(r),{address:'M3',unique:true});assert.deepEqual(r.highlights.negotiatedCells,['L2','M3']);assert.deepEqual(r.highlights.negotiatedPaidCells,['L2']);
 assert.equal(r.negotiated.rows[0][7].v,'已经提交全部');assert.equal(r.negotiated.rows[1][7].v,'已经提交全部');
 assert.equal(r.batchComments.L2[1],prefixes[1]+'2026年10月21日');assert.equal(r.batchComments.M3[1],prefixes[1]+'2026年11月19日');assert.equal(r.batchComments.M3[3],prefixes[3]+'未记录');
 assert.equal(r.negotiated.rows[1][2].v,ID(2));assert.equal(r.progressRows[1][r.progressHeaders.indexOf('对公付款OA号')].v,ID(1));
 assert.equal(notes(r).length,1);for(const text of ['49990.00',ID(1),ID(2),'谈判付款!M3','第 3 行','已按供应商'])assert(notes(r)[0].includes(text));
 assert(!r.reviewRows.some(row=>/申请标红核对|申请超额/.test(row[5].v)));assert.equal(JSON.stringify(i),before);
});
test('同额不同批次：旧红格已有原申请日期，新申请只匹配唯一后续格',()=>{
 const i=sample(120000,120000),r=C.build(i);
 assert.deepEqual(match(r),{address:'M3',unique:true});assert.deepEqual(r.highlights.negotiatedPaidCells,['L2']);
 assert.equal(r.batchComments.L2[1],prefixes[1]+'2026年10月21日');assert(!r.reviewRows.some(row=>row[5].v.startsWith('重复提交')));
});
test('两个同额后续批次不能仅按日期早晚跨OA猜配，保持原批注和颜色',()=>{
 const i=sample(120000,120000,true),before=JSON.stringify(i),r=C.build(i);
 assert.equal(match(r),undefined);assert.deepEqual(r.highlights.negotiatedCells,['L2']);assert.equal(notes(r).length,1);
 for(const text of ['多个同额候选','M3','N4','确认本次申请对应哪一批'])assert(notes(r)[0].includes(text));
 assert(!r.reviewRows.some(row=>/申请标红核对|申请超额/.test(row[5].v)));assert.equal(JSON.stringify(i),before);
});
test('重新上传用原申请日期批注明确批次，保留另一个同额未申请批次',()=>{
 const i=sample(120000,120000,true);neg(i).M3.c=[{a:'工具',t:prefixes[1]+'2026年11月19日'}];neg(i).M3.s={fgColor:{rgb:'FFFF00'}};
 const r=C.build(i);assert.deepEqual(match(r),{address:'M3',unique:true});assert(!r.highlights.negotiatedCells.includes('N4'));
 for(const [address,lines]of Object.entries(r.batchComments))neg(i)[address].c=[{a:'工具',t:lines.join('\n')}];
 const again=C.build(i);assert.deepEqual(again.highlights,r.highlights);assert.deepEqual(match(again),match(r));
});
test('原红黄底或付款批注不能被另一申请覆盖；同一申请日期才可重用',()=>{
 for(const change of [c=>c.s={fgColor:{rgb:'FF0000'}},c=>c.s={fgColor:{rgb:'FFFFFF00'}},c=>c.c=[{t:'11月5日-已支付'}],c=>c.c=[{t:prefixes[3]+'2026年11月5日'}],c=>c.c=[{t:prefixes[1]+'2026年11月18日'}]]){
  const i=sample();change(neg(i).M3);const before=JSON.stringify(i),r=C.build(i);assert.equal(match(r),undefined);assert(!r.highlights.negotiatedCells.includes('M3'));assert.equal(JSON.stringify(i),before);
 }
});
test('供应商主体缺失或冲突、申请日期无效、金额零异常、计划公式异常不跨号猜配',()=>{
 for(const change of [i=>apps(i).B3.v='其他供应商',i=>apps(i).C3.v='',i=>apps(i).C3.v='迈德斯特',i=>apps(i).A3.v='',i=>apps(i).A3.v='不明日期',i=>apps(i).E3.v=0,i=>apps(i).E3.v='金额未知',i=>neg(i).M3={t:'n',f:'A99'},i=>neg(i).M3={t:'e',v:23,f:'#REF!'},i=>neg(i)['!merges']=[X.utils.decode_range('M3:N3')]]){
  const i=sample();change(i);const r=C.build(i);assert.equal(match(r),undefined);assert(!r.highlights.negotiatedCells.includes('M3'));
 }
});
test('正常台账类别与多OA申请不启用跨号批次匹配',()=>{
 for(const change of [i=>apps(i).F3.v='正常付款',i=>apps(i).D3.v=ID(1)+'、'+ID(2)]){const i=sample();change(i);assert.equal(match(C.build(i)),undefined);}
});
test('候选已被正确OA的申请占用时，不分配给另一旧OA申请',()=>{
 const i=sample(),row=X.utils.sheet_to_json(apps(i),{header:1})[2];row[3]=ID(2);X.utils.sheet_add_aoa(apps(i),[row],{origin:-1});
 const r=C.build(i);assert.equal(match(r),undefined);assert.deepEqual(r.applicationBatchMatches['付款申请汇总!4'],{address:'M3',unique:true});
});
test('多笔不同OA竞争同一新批次时均不自动分配；同OA重复申请仍只计一次',()=>{
 const i=sample(),s=neg(i),n=X.utils.sheet_to_json(s,{header:1})[1];n[2]=ID(3);n[3]=n[4]=n[9]=n[11]=100;X.utils.sheet_add_aoa(s,[n],{origin:-1});
 const row=X.utils.sheet_to_json(apps(i),{header:1})[2];row[3]=ID(3);X.utils.sheet_add_aoa(apps(i),[row],{origin:-1});
 const r=C.build(i);assert.equal(match(r),undefined);assert.equal(r.applicationBatchMatches['付款申请汇总!4'],undefined);assert(!r.highlights.negotiatedCells.includes('M3'));assert(notes(r).every(t=>t.includes('多笔不同申请')));
 const j=sample(),copy=X.utils.sheet_to_json(apps(j),{header:1})[2];X.utils.sheet_add_aoa(apps(j),[copy],{origin:-1});const duplicate=C.build(j);
 assert.deepEqual(duplicate.highlights.negotiatedCells,['L2','M3']);assert.equal(duplicate.reviewRows.filter(row=>row[5].v.startsWith('重复提交')).length,1);
});
