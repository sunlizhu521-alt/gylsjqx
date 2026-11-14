const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../product-dimension-core.js');
const H=['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组',...C.FIELD_HEADERS];
const M=['编码','领星SKU','SKU','名称','创建日期','物料分组'];
function book(headers,rows,date1904=false){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows]),'数据');b.Workbook={WBProps:{date1904}};return b;}
function old(code,sku,extra={}){const values={物料编码:code,领星SKU:sku,SKU:sku,金蝶名称:code,金蝶创建时间:'2024/01/01',物料分组:'测试组',销售产品线:'产品线甲',销售系列:'系列甲',销售产品分类:'全新品',场景分组:'场景甲',一级品类:'一级甲',二级分类:'二级甲',三级分类:'三级甲',型号:'型号-'+code,销售区域:'中国',财务加权平均价:0,...extra};return H.map(h=>values[h]??'');}
const material=(code,sku,name='新商品',group='测试组',date='2026/01/15')=>[code,sku,sku,name,date,group];
const analyze=(history,incoming,opts={})=>C.analyze(book(H,history),book(M,incoming),{asOfDate:'2026-11-14',...opts});
const get=(r,h)=>r.values[H.indexOf(h)];
test('exact SKU fills only new fields, does not copy old prices, records evidence and preserves old values',()=>{
 const source=old('旧编码','S-1',{品牌:'迈德斯特',是否正常备货:'正常备货',当前日期:1,生命周期:'原状态'}),a=analyze([source],[material('新增','S-1')]),r=C.resolve(a);
 assert.equal(r.canExport,true);assert.equal(r.pendingRows,0);assert.deepEqual(r.rows[0].values,source);
 const n=r.rows[1];assert.equal(get(n,'销售产品线'),'产品线甲');assert.equal(get(n,'财务加权平均价'),'');assert.equal(get(n,'不含税结算价'),'');assert.equal(get(n,'品牌'),'迈德斯特');assert.equal(get(n,'是否正常备货'),'正常备货');
 assert.equal(n.fields.销售产品线.method,'history-sku');assert.deepEqual(n.fields.销售产品线.sourceRows,[2]);
 assert(C.EDITABLE_HEADERS.every(h=>n.fields[h]));assert.equal(get(n,'当前日期'),46340);assert.equal(get(n,'上架时间（月）'),'9个月');assert.equal(get(n,'生命周期'),'稳定期');
});
test('a stronger matching conflict stops inference instead of falling back to a convenient group',()=>{
 const a=analyze([old('a','S-1',{物料分组:'组甲',销售产品线:'甲'}),old('b','S-1',{物料分组:'组乙',销售产品线:'乙'}),old('c','S-2',{物料分组:'组三',销售产品线:'丙'})],[material('新','S-1','新商品','组三')]);
 const r=C.resolve(a),n=r.rows.at(-1);assert.equal(r.canExport,false);assert.equal(get(n,'销售产品线'),'待确认');assert.deepEqual(n.fields.销售产品线.candidates.map(v=>v.value),['甲','乙']);assert(r.pendingFields.some(p=>p.code==='新'&&p.field==='销售产品线'));assert(!r.issues.some(i=>i.severity==='error'));
});
test('polluted numeric category values are not learned and zero/placeholder SKU never match',()=>{
 const a=analyze([old('a','S',{场景分组:1628.318}),old('b','S',{场景分组:'有效场景'}),old('zero',0,{物料分组:'不同组',销售产品线:'错误关联'})],[material('有SKU','S'),material('无SKU',0,'未知商品','未知组')]);
 const r=C.resolve(a);assert.equal(get(r.rows.find(x=>x.code==='有SKU'),'场景分组'),'有效场景');assert.equal(get(r.rows.find(x=>x.code==='无SKU'),'销售产品线'),'待确认');
});
test('manual changes survive recalculation; empty is deliberate and invalid prices remain editable',()=>{
 const a=analyze([old('a','S')],[material('新','S')]);
 const r=C.resolve(a,{'新':{销售产品线:'',财务加权平均价:'abc',金蝶创建时间:'2026-06-14',当前日期:'1999-01-01'}}),n=r.rows.at(-1);
 assert.equal(get(n,'销售产品线'),'');assert.equal(n.fields.销售产品线.method,'manual');assert.equal(get(n,'财务加权平均价'),'abc');assert.equal(get(n,'上架时间（月）'),'5个月');assert.equal(get(n,'当前日期'),46340);assert.equal(r.pendingRows,1);assert(r.pendingFields.some(p=>p.field==='财务加权平均价'));
 const fixed=C.resolve(a,{'新':{销售产品线:'产品线甲',财务加权平均价:'0',金蝶创建时间:'2026-06-14'}});assert.equal(fixed.canExport,true);assert.equal(get(fixed.rows.at(-1),'财务加权平均价'),0);
});
test('manual product line and series recalculate dependent classifications',()=>{
 const a=analyze([old('a','S',{销售产品线:'升降椅',销售系列:'AX25',场景分组:'九大产品线',一级品类:'升降椅',二级分类:'升降椅'})],[material('新','','未知','无历史组')]);
 const r=C.resolve(a,{'新':{销售产品线:'升降椅',销售系列:'AX25'}}),n=r.rows.at(-1);assert.equal(get(n,'场景分组'),'九大产品线');assert.equal(get(n,'一级品类'),'升降椅');assert.equal(get(n,'二级分类'),'升降椅');
});
test('explicit country labels fill country/region while group alone cannot invent brand or stock',()=>{
 const a=analyze([old('旧','S',{品牌:'迈德斯特',国家:'英国',是否正常备货:'正常备货'})],[material('美','','新品（美国）'),material('英','','新品（英国）'),material('中','','新品国内')]),r=C.resolve(a);
 for(const [code,country,area]of [['美','美国','美国'],['英','英国','欧洲'],['中','中国','中国']]){const n=r.rows.find(x=>x.code===code);assert.equal(get(n,'国家'),country);assert.equal(get(n,'销售区域'),area);assert.equal(get(n,'品牌'),'待确认');assert.equal(get(n,'是否正常备货'),'待确认');}
});
test('return state reuses the exact original SKU and validates model transformation with two pairs',()=>{
 const history=[old('b1','M-BASE1',{型号:'MD-BASE1'}),old('b2','M-BASE2',{型号:'MD-BASE2'}),old('b3','M-BASE3',{型号:'MD-BASE3'}),old('r1','RE-M-BASE1',{型号:'RE-MD-BASE1',销售产品分类:'退回分类',三级分类:'退回三级'}),old('r2','RE-M-BASE2',{型号:'RE-MD-BASE2',销售产品分类:'退回分类',三级分类:'退回三级'})];
 let a=analyze(history,[material('新','RE-M-BASE3','全新换包装新品国内')]),r=C.resolve(a),n=r.rows.at(-1);
 assert.equal(r.canExport,true);assert.equal(get(n,'型号'),'RE-MD-BASE3');assert.equal(get(n,'销售产品分类'),'退回分类');assert.equal(get(n,'三级分类'),'退回三级');assert.equal(get(n,'财务加权平均价'),'');
 a=analyze(history.slice(0,-1),[material('新','RE-M-BASE3')]);n=C.resolve(a).rows.at(-1);assert.equal(get(n,'型号'),'待确认');assert.match(n.fields.型号.reason,/不足两组/);
 history[4][H.indexOf('型号')]='RE-错误型号';a=analyze(history,[material('新','RE-M-BASE3')]);n=C.resolve(a).rows.at(-1);assert.equal(get(n,'型号'),'待确认');assert.match(n.fields.型号.reason,/反例/);
});
test('state classification conflicts remain pending and optional model does not gate export',()=>{
 const history=[old('b','BASE'),old('r1','RE-X',{销售产品分类:'甲',三级分类:'退回'}),old('r2','RE-Y',{销售产品分类:'乙',三级分类:'退回'})];
 const a=analyze(history,[material('新','RE-BASE')]),r=C.resolve(a);assert.equal(r.canExport,false);assert(r.pendingFields.some(p=>p.field==='销售产品分类'));assert(!r.pendingFields.some(p=>p.field==='型号'));
});
test('complete months and lifecycle boundaries use the fixed processing date',()=>{
 for(const [date,months,stage]of [['2026-05-14',6,'新品期'],['2026-05-15',5,'新品期'],['2024-11-14',24,'稳定期'],['2024-10-14',25,'衰弱期']]){
  const r=C.resolve(analyze([old('a','S')],[material('新','S','新','测试组',date)]));assert.equal(r.canExport,true);assert.equal(get(r.rows.at(-1),'上架时间（月）'),months+'个月');assert.equal(get(r.rows.at(-1),'生命周期'),stage);
 }
 const leap=C.resolve(analyze([old('a','S')],[material('新','S','新','测试组','2024-02-29')],{asOfDate:'2025-02-28'}));assert.equal(get(leap.rows.at(-1),'上架时间（月）'),'11个月');
 for(const date of ['', '2026-02-30','2026-11-15']){const r=C.resolve(analyze([old('a','S')],[material('新','S','新','测试组',date)]));assert.equal(r.canExport,false);assert(r.pendingFields.some(p=>p.field==='金蝶创建时间'));assert.equal(r.rows.length,2);assert(!r.issues.some(i=>i.severity==='error'));}
});
test('1904 dates are normalized and validation independently rejects forged completion',()=>{
 const d=book(H,[old('a','S')],true),m=book(M,[material('新','S','新','测试组','2026-01-15')]);
 const a=C.analyze(d,m,{asOfDate:'2026-11-14'}),r=C.resolve(a);assert.equal(r.canExport,true);assert.equal(get(r.rows.at(-1),'当前日期'),46340-1462);assert.deepEqual(C.validateResult(a,r),[]);
 const fake=structuredClone(r);fake.canExport=true;fake.rows.at(-1).isNew=false;fake.rows.at(-1).values[H.indexOf('销售区域')]='';assert(C.validateResult(a,fake).some(e=>e.field==='销售区域'));
 for(const field of ['金蝶创建时间','当前日期']){const bad=structuredClone(r);bad.rows.at(-1).values[H.indexOf(field)]='2026-01-15';assert(C.validateResult(a,bad).some(e=>e.field===field));}
 const bad=structuredClone(r);bad.rows.at(-1).values[H.indexOf('上架时间（月）')]='999个月';assert(C.validateResult(a,bad).some(e=>e.field==='上架时间（月）'));
});
test('unmatched, indistinguishable and check-needed placeholders never satisfy a mandatory field',()=>{
 for(const placeholder of ['未匹配','无法区分','待核对']) {
  const a=analyze([old('旧','S',{销售区域:placeholder})],[material('新','S')]),r=C.resolve(a);
  assert.equal(get(r.rows.at(-1),'销售区域'),'待确认');assert(r.pendingFields.some(p=>p.field==='销售区域'));
  const manual=C.resolve(a,{'新':{销售区域:placeholder}});assert.equal(manual.canExport,false);assert(manual.pendingFields.some(p=>p.field==='销售区域'));
  const forged=structuredClone(r);forged.canExport=true;forged.rows.at(-1).values[H.indexOf('销售区域')]=placeholder;assert(C.validateResult(a,forged).some(p=>p.field==='销售区域'));
 }
});
test('explicit other country names are facts, while regions still require unique historical evidence',()=>{
 for(const country of ['加拿大','俄罗斯','波兰','日本']) {
  const a=analyze([old('旧','S',{国家:'英国'})],[material('新','S','新商品（'+country+'）')]),n=C.resolve(a).rows.at(-1);
  assert.equal(get(n,'国家'),country);assert.equal(n.fields.国家.method,'derived');
 }
});
test('new prices stay empty even for exact SKU and near-zero old prices; whitespace is an intentional clear',()=>{
 const a=analyze([old('旧','S',{财务加权平均价:0,不含税结算价:1e-9,结算价:100})],[material('新','S')]);
 const r=C.resolve(a),n=r.rows.at(-1);for(const h of ['财务加权平均价','不含税结算价','结算价'])assert.equal(get(n,h),'');
 const cleared=C.resolve(a,{'新':{财务加权平均价:'   '}});assert.equal(cleared.canExport,true);assert.equal(cleared.pendingRows,0);assert.equal(get(cleared.rows.at(-1),'财务加权平均价'),'');
});
test('return fields and stock never borrow normal-state data when return history is absent',()=>{
 const a=analyze([old('base','BASE',{金蝶名称:'相同名称',是否正常备货:'正常备货'})],[material('新','RE-BASE','相同名称')]),r=C.resolve(a),n=r.rows.at(-1);
 assert.equal(get(n,'销售产品线'),'产品线甲');for(const h of ['销售产品分类','三级分类','型号','是否正常备货'])assert.equal(get(n,h),'待确认');
 assert(r.pendingFields.some(p=>p.field==='销售产品分类'));assert(r.pendingFields.some(p=>p.field==='三级分类'));
 const normal=analyze([old('returned','RE-S',{金蝶名称:'同名',销售产品分类:'退回',三级分类:'退回',型号:'RE-M',是否正常备货:'正常备货'})],[material('普通','S','同名')]);
 const output=C.resolve(normal).rows.at(-1);for(const h of ['销售产品分类','三级分类','型号','是否正常备货'])assert.equal(get(output,h),'待确认');
});
test('restoring product line and series keeps precise SKU evidence above broader dependencies',()=>{
 const a=analyze([old('甲','SKU-A',{三级分类:'细类甲'}),old('乙','SKU-B',{三级分类:'细类乙'})],[material('新','SKU-A')]);
 assert.equal(get(C.resolve(a).rows.at(-1),'三级分类'),'细类甲');
 const restored=C.resolve(a,{'新':{销售产品线:'产品线甲',销售系列:'系列甲'}});
 assert.equal(restored.canExport,true);assert.equal(get(restored.rows.at(-1),'三级分类'),'细类甲');assert.equal(restored.rows.at(-1).fields.三级分类.method,'history-sku');
 const changed=C.resolve(a,{'新':{销售产品线:'不同产品线'}});assert.equal(changed.canExport,false);assert(changed.pendingFields.some(p=>p.field==='三级分类'));
});
test('all named countries are collected before selecting a country or sales region',()=>{
 for(const name of ['商品美国、加拿大','商品英国、德国']) {
  const a=analyze([old('旧','S',{国家:'英国',销售区域:'欧洲'})],[material('新','S',name)]),r=C.resolve(a),n=r.rows.at(-1);
  assert.equal(get(n,'国家'),'待确认');assert.equal(get(n,'销售区域'),'待确认');assert(r.pendingFields.some(p=>p.field==='销售区域'));assert.match(n.fields.销售区域.reason,/多个国家/);
 }
});
test('Indonesia aliases do not match India and explicit countries never fall back to a generic group region',()=>{
 const indian=old('印度旧','IND',{金蝶名称:'印度商品',销售区域:'印度'});
 for(const name of ['商品印度尼西亚','商品印尼']) {
  let a=analyze([indian],[material('新','NEW',name)]),r=C.resolve(a),n=r.rows.at(-1);
  assert.equal(get(n,'国家'),'印度尼西亚');assert.equal(get(n,'销售区域'),'待确认');assert(r.pendingFields.some(p=>p.field==='销售区域'));
  a=analyze([indian,old('印尼旧','IDN',{金蝶名称:'印尼商品',销售区域:'东南亚'})],[material('新','NEW',name)]);n=C.resolve(a).rows.at(-1);
  assert.equal(get(n,'国家'),'印度尼西亚');assert.equal(get(n,'销售区域'),'东南亚');
 }
});
test('negative or paused normal-stock phrases cannot be overridden by historical stock matches',()=>{
 for(const phrase of ['不再正常备货','暂停正常备货','停止正常备货','暂不正常备货','非正常备货','正常备货已停止']) {
  const name='新品'+phrase,a=analyze([old('旧','S',{金蝶名称:name,是否正常备货:'正常备货'})],[material('新','S',name)]),n=C.resolve(a).rows.at(-1);
  assert.equal(get(n,'是否正常备货'),'待确认');assert.equal(n.fields.是否正常备货.method,'pending');assert.match(n.fields.是否正常备货.reason,/否定|暂停/);
 }
});
test('completion placeholders and numeric classifications are blocked but numeric models and price zero remain valid',()=>{
 for(const value of ['待补全','--','—','－－','0',0,'123','1e3']) {
  const a=analyze([old('旧','S',{销售区域:value})],[material('新','S')]),r=C.resolve(a);assert.equal(r.canExport,false);assert.equal(get(r.rows.at(-1),'销售区域'),'待确认');
  const manual=C.resolve(a,{'新':{销售区域:value}});assert.equal(manual.canExport,false);assert(manual.pendingFields.some(p=>p.field==='销售区域'));assert(C.validateResult(a,manual).some(p=>p.field==='销售区域'));
 }
 const a=analyze([old('旧','S')],[material('新','S')]);
 for(const field of ['销售产品线','品牌']){const r=C.resolve(a,{'新':{[field]:'123'}});assert.equal(r.canExport,false);assert(r.pendingFields.some(p=>p.field===field));assert(C.validateResult(a,r).some(p=>p.field===field));}
 const allowed=C.resolve(a,{'新':{型号:117,财务加权平均价:0}});assert.equal(allowed.canExport,true);assert.equal(get(allowed.rows.at(-1),'型号'),'117');assert.equal(get(allowed.rows.at(-1),'财务加权平均价'),0);
});
