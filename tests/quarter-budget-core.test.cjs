const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../quarter-budget-core.js');
const H=C.HEADERS,serial=s=>(Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const source=(role,rows,name=role)=>({book:{SheetNames:['明细'],Sheets:{'明细':X.utils.aoa_to_sheet([H,...rows])}},filename:name+'.xlsx',fingerprint:name});
const row=(change={})=>{const o={month:serial('2026-11-01'),supplier:'合成供应商',subject:'合成主体',oa:'MDST-DGFK20261110001',summary:'合成货款',amount:100,status:'待核对',notes:'',category:'自定义分类',owner:'合成跟单员',excluded:'否',...change};return [o.month,o.supplier,o.subject,o.oa,o.summary,o.amount,o.status,o.notes,o.category,o.owner,o.excluded];};
function newInputs(n=[],l=[],b=[]){return {negotiated:source('negotiated',n),ledger:source('ledger',l),budget:source('budget',b)};}
function schedule(rows){
 const H=['供应商','付款主体','对公付款0A流程号','谈判明细',new Date('2026-11-02T00:00:00Z'),new Date('2026-11-16T00:00:00Z')];
 const sheet=X.utils.aoa_to_sheet([H,...rows]);sheet.E1={t:'n',v:serial('2026-11-02'),z:'yyyy-mm-dd'};sheet.F1={t:'n',v:serial('2026-11-16'),z:'yyyy-mm-dd'};
 return {book:{SheetNames:['谈判付款'],Sheets:{'谈判付款':sheet}},filename:'schedule.xlsx',fingerprint:'schedule'};
}
test('cross-source OA+amount removes lower priority but keeps same-source installments and zero',()=>{
 const result=C.build(newInputs([row({status:'未支付'}),row({month:serial('2026-12-01'),status:'未支付'}),row({oa:'MDST-DGFK20261110002',amount:0})],[row(),row({oa:'MDST-DGFK20261110002',amount:0})],[row()]));
 assert.equal(result.records.length,3);assert.equal(result.deleted.length,3);assert.equal(result.records[0].category,'谈判付款');assert.equal(result.summary.totals.totalCents,20000);
});
test('empty OA and amount do not auto-delete; suspicious pairs are grouped, retained and red',()=>{
 const result=C.build(newInputs([row({oa:'',summary:'不同摘要'}),row({oa:'',amount:null})],[],[row({oa:'',summary:'不同摘要'}),row({oa:'',amount:null})]));
 assert.equal(result.deleted.length,0);assert.equal(result.records.length,4);assert.ok(result.records.every(r=>r.excluded==='否'));assert.ok(result.records.some(r=>r.related));
 assert.equal(result.summary.totals.totalCents,20000);assert.ok(result.warnings.some(x=>x.includes('金额')));
});
test('manual exclusion updates summary and remark, keep restores original financial scope',()=>{
 const inputs=newInputs([row({oa:''})],[],[row({oa:''})]);let result=C.build(inputs);const id=result.records[1].id;
 result=C.build(inputs,{decisions:{[id]:{excluded:'是',reason:'已核对重复'}}});
 assert.equal(result.records[1].excluded,'是');assert.match(result.records[1].notes,/已核对重复/);assert.equal(result.summary.totals.totalCents,10000);
 result=C.build(inputs,{decisions:{[id]:{excluded:'否'}}});assert.equal(result.summary.totals.totalCents,20000);
});
test('existing mode preserves rows, classifications, exact amounts and uses actual calendar months',()=>{
 const inputs={existing:source('existing',[row({category:'暂停付款',status:'已支付',amount:123.45}),row({month:serial('2026-11-24'),category:'暂停付款',amount:10}),row({month:'待定',amount:2,excluded:'是'}),row({month:'12月',amount:3})])};
 const result=C.build(inputs,{mode:'existing'});
 assert.equal(result.records.length,4);assert.equal(result.records[0].category,'暂停付款');assert.equal(result.records[0].amount,123.45);assert.equal(result.summary.monthly.length,2);assert.deepEqual(result.summary.totals,{totalCents:13645,paidCents:12345,remainingCents:1300});
 assert.equal(result.summary.monthly[1].monthKey,'');
});
test('same OA/month/amount separate installments follow distinct payment cells',()=>{
 const e=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','1、11月10日-100\n2、11月25日-100',100,100]]);e.book.Sheets['谈判付款'].E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const result=C.build({existing:source('existing',[row({status:'已支付'}),row({status:'未支付'})]),evidence:e},{mode:'existing'});
 assert.deepEqual(result.records.map(r=>r.status),['已支付','未支付']);assert.equal(result.summary.totals.paidCents,10000);
});
test('one paid evidence cannot pay two distinct active records',()=>{
 const evidence=source('evidence',[row({status:'已支付'})]);const result=C.build({existing:source('existing',[row(),row()]),evidence},{mode:'existing'});
 assert.ok(result.records.every(r=>r.status==='待核对'));assert.equal(result.summary.totals.paidCents,0);assert.equal(result.issues.filter(i=>i.type==='evidence').length,2);
});
test('a weak pending process does not downgrade an existing paid row',()=>{
 const result=C.build({existing:source('existing',[row({status:'已支付'})]),evidence:source('evidence',[row({status:'待核对'})])},{mode:'existing'});
 assert.equal(result.records[0].status,'已支付');
});
test('explicit contrary evidence flags payment state conflict',()=>{
 const result=C.build({existing:source('existing',[row({status:'已支付'})]),evidence:source('evidence',[row({status:'未支付'})])},{mode:'existing'});
 assert.equal(result.records[0].status,'待核对');assert.match(result.records[0].notes,/冲突|原记录/);
});
test('invalid manual evidence choice cannot be applied after source changes',()=>{
 const existing=source('existing',[row()]);let result=C.build({existing},{mode:'existing'});const id=result.records[0].id;
 result=C.build({existing},{mode:'existing',matches:{[id]:'stale-id'}});assert.equal(result.records[0].status,'待核对');assert.ok(result.warnings.some(w=>w.includes('失效')));
});
test('combined paid and unpaid amounts remain reviewable, not wholly paid',()=>{
 const result=C.build({existing:source('existing',[row({oa:'',amount:300,notes:'100.00+200.00',status:'未支付'})]),evidence:source('evidence',[row({amount:100,status:'已支付'}),row({amount:200,status:'未支付'})])},{mode:'existing'});
 assert.equal(result.records[0].status,'待核对');assert.match(result.records[0].notes,/合并金额/);
});
test('missing required slots and invalid financial formula prevent usable results',()=>{
 assert.throws(()=>C.build({}),/上传/);
 const existing=source('existing',[row({amount:100})]);existing.book.Sheets['明细'].F2={t:'n',f:'SUM(Z1:Z2)'};
 const result=C.build({existing},{mode:'existing'});assert.equal(result.canExport,false);assert.ok(result.warnings.length);
});
test('summary does not silently overflow integer cents',()=>{
 assert.throws(()=>C.summarize([{monthKey:'2026-11',category:'测试',cents:Number.MAX_SAFE_INTEGER,status:'未支付'},{monthKey:'2026-11',category:'测试',cents:1,status:'未支付'}]),/范围/);
});
test('a paid installment from another month cannot pay this row, with or without OA',()=>{
 for(const oa of ['MDST-DGFK20261110001','']){
  const result=C.build({existing:source('existing',[row({oa,month:serial('2026-12-01')})]),evidence:source('evidence',[row({oa,status:'已支付'})])},{mode:'existing'});
  assert.equal(result.records[0].status,'待核对');assert.equal(result.summary.totals.paidCents,0);assert.match(result.records[0].notes,/月份不同/);
 }
});
test('reordered manual installments use unique summary instead of edited row order',()=>{
 const evidence=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','11月10日-100',100],['合成供应商','合成主体','MDST-DGFK20261110001','11月25日-100',null,100]]);
 const sheet=evidence.book.Sheets['谈判付款'];sheet.G1={t:'s',v:'摘要'};sheet.G2={t:'s',v:'第一笔'};sheet.G3={t:'s',v:'第二笔'};sheet['!ref']='A1:G3';sheet.E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const existing=source('existing',[row({summary:'第二笔',status:'未支付'}),row({summary:'第一笔',status:'已支付'})]);
 const result=C.build({existing,evidence},{mode:'existing'});
 assert.deepEqual(result.records.map(r=>r.status),['未支付','已支付']);
});
test('indistinguishable mixed-state installments are not assigned by row order',()=>{
 const evidence=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','11月10日-100；11月25日-100',100,100]]);
 evidence.book.Sheets['谈判付款'].E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const result=C.build({existing:source('existing',[row(),row()]),evidence},{mode:'existing'});
 assert.deepEqual(result.records.map(r=>r.status),['待核对','待核对']);assert.equal(result.summary.totals.paidCents,0);
 const preserved=C.build({existing:source('existing',[row({status:'未支付'}),row({status:'已支付'})]),evidence},{mode:'existing'});
 assert.deepEqual(preserved.records.map(r=>r.status),['未支付','已支付']);
});
test('undated ledger summary is not a second payment beside a schedule cell',()=>{
 const evidence=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','11月10日-100',100]]);
 evidence.book.Sheets['谈判付款'].E2.s={fill:{patternType:'solid',fgColor:{rgb:'FFFF0000'}}};
 const ledger=source('ledger',[row({month:null,status:'已支付'})]);
 const result=C.build({existing:source('existing',[row(),row({month:serial('2026-12-01')})]),ledger,evidence},{mode:'existing'});
 assert.deepEqual(result.records.map(r=>r.status),['已支付','待核对']);assert.equal(result.summary.totals.paidCents,10000);
});
test('explicit same-month ledger and schedule conflicts remain reviewable',()=>{
 const evidence=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','11月10日-100',100]]);
 const result=C.build({existing:source('existing',[row()]),ledger:source('ledger',[row({status:'已支付'})]),evidence},{mode:'existing'});
 assert.equal(result.records[0].status,'待核对');assert.match(result.records[0].notes,/台账.*冲突/);
});
test('manual keep explanation is retained in exported notes',()=>{
 const inputs={existing:source('existing',[row()])};const r=C.build(inputs,{mode:'existing'}).records[0];
 const result=C.build(inputs,{mode:'existing',decisions:{[r.id]:{excluded:'否',reason:'经核对是不同分期'}}});
 assert.match(result.records[0].notes,/经核对是不同分期/);
});
test('same-amount ledger evidence for another explicit summary is a separate installment',()=>{
 const evidence=schedule([['合成供应商','合成主体','MDST-DGFK20261110001','11月10日-100',100]]);
 const sheet=evidence.book.Sheets['谈判付款'];sheet.G1={t:'s',v:'摘要'};sheet.G2={t:'s',v:'第一笔'};sheet['!ref']='A1:G2';
 const result=C.build({existing:source('existing',[row({summary:'第一笔',status:'未支付'}),row({summary:'第二笔',status:'已支付'})]),ledger:source('ledger',[row({summary:'第二笔',status:'已支付'})]),evidence},{mode:'existing'});
 assert.deepEqual(result.records.map(r=>r.status),['未支付','已支付']);assert.equal(result.summary.totals.paidCents,10000);
});
test('continuing an exported detail does not add prefixes to manual remarks',()=>{
 const notes='保留原说明\n分期付款条件';
 const first=C.build({existing:source('existing',[row({notes})])},{mode:'existing'});
 assert.equal(first.records[0].notes,notes);
 const second=C.build({existing:source('existing',first.records.map(C.toRow))},{mode:'existing'});
 assert.equal(second.records[0].notes,notes);
});
