const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,set,append,ID,LH,AH,OH,book}=require('./funds-progress-fixtures.cjs');
const records=r=>r.progressRows.map(row=>Object.fromEntries(r.progressHeaders.map((h,c)=>[h,row[c]])));
const review=r=>r.reviewRows.filter(row=>row[5].v.startsWith('付款进度核对：'));
const display=c=>c.t==='n'&&c.z?X.SSF.format(c.z,c.v):c.v;
test('重新申请不进入付款申请核对，仍保留进度、提交状态、金额及标红',()=>{
 const i=fixture();set(i,'applications','流程编码',5,ID(99));
 const before=C.build(i);assert(review(before).some(row=>row[8].v===5));
 X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],['需重新申请'],[''],['重新申请'],['退回后重新申请']],{origin:'G1'});
 const source=JSON.stringify(i),r=C.build(i);
 assert.equal(r.reviewRows.length,0);assert.deepEqual(records(r).map(row=>row['是否已付款'].v),['重新申请','未付款','重新申请','重新申请']);
 for(const key of ['normalRows','negotiated','highlights','applicationComments','stats'])assert.deepEqual(r[key],before[key],key);
 assert.equal(JSON.stringify(i),source);
 set(i,'applications','核对说明',5,'');assert(review(C.build(i)).some(row=>row[8].v===5),'重新上传取消说明后恢复核对');
});
test('重复核对仅统计未标重新申请的记录；普通记录及异常说明不误排除',()=>{
 const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],['重新申请']],{origin:'G1'});
 let r=C.build(i);assert(!r.reviewRows.some(row=>row[5].v.startsWith('重复提交')));assert.equal(review(r).length,0);
 append(i,'applications',['2026-11-05','合成供应商1','迈德斯特',ID(1),40,'谈判付款']);
 r=C.build(i);const duplicates=r.reviewRows.filter(row=>row[5].v.startsWith('重复提交'));assert.equal(duplicates.length,1);assert.match(duplicates[0][5].v,/共 2 次/);assert.equal(duplicates[0][8].v,4);
 for(const note of ['', '需核实', {t:'e',v:23,w:'重新申请'}, {t:'s',f:'A99'}]){
  const j=fixture();set(j,'applications','流程编码',5,ID(99));X.utils.sheet_add_aoa(j.applications.book.Sheets.付款申请汇总,[['核对说明']],{origin:'G1'});set(j,'applications','核对说明',5,note);
  assert(review(C.build(j)).some(row=>row[8].v===5));
 }
});
test('重新申请不掩盖台账或OA自身核对',()=>{
 const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],['重新申请'],['重新申请'],['重新申请'],['重新申请']],{origin:'G1'});
 set(i,'oa','收款单位',3,'不同供应商');const r=C.build(i);
 assert(r.reviewRows.some(row=>row[5].v.startsWith('OA供应商冲突')));assert.equal(review(r).length,0);
});
test('每笔申请一行、重复及未进入付款表申请保留，来源日期金额类型正确，归档仅改变新表状态',()=>{
 const i=fixture(),before=JSON.stringify(i),r=C.build(i),p=records(r);
 assert.equal(p.length,4);assert.deepEqual(p.map(row=>row['申请付款金额'].v),[40,200,40,0]);
 assert.deepEqual(p.map(row=>row['是否已付款'].v),['已付款','未付款','已付款','未付款']);
 assert.deepEqual(p.map(row=>row['跟单员'].v),['李奇','徐英','李奇','']);
 assert.equal(display(p[0]['应付时间']),'2026年11月6日');assert.equal(display(p[1]['应付时间']),'2026年11月9日');
 assert.equal(display(p[0]['申请日期']),'2026年11月1日');assert.equal(p[0]['申请日期'].t,'n');assert.equal(p[3]['申请付款金额'].t,'n');assert.equal(p[0]['对公付款OA号'].t,'s');
 assert.equal(p[0]['付款主体'].v,'浙江迈德斯特医疗器械科技有限公司');assert.equal(p[0]['申请人'].v,'OA申请人');
 assert.equal(r.stats.pending,0);assert.equal(review(r).length,0);
 assert.deepEqual(r.highlights.negotiatedCells,['M2']);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][5].f,'D2-E2');assert.equal(JSON.stringify(i),before);
});
test('付款进度仅缺资金台账来源时不生成核对，跟单员应付时间留空并保留金额状态',()=>{
 for(const amount of [0,123.45]){
  const i=fixture(),id='MDST-DGFK20260509004';
  set(i,'applications','流程编码',5,id);set(i,'applications','本次申请付款金额',5,amount);
  set(i,'oa','流程编号',5,id);set(i,'oa','付款金额',5,amount);
  const before=JSON.stringify(i),r=C.build(i),p=records(r)[3];
  assert.equal(p['跟单员'].v,'');assert.equal(p['应付时间'].v,'');assert.equal(p['申请付款金额'].v,amount);
  assert.equal(p['对公付款OA号'].v,id);assert.equal(p['是否已付款'].v,'未付款');
  assert(!r.reviewRows.some(row=>row[6].v===i.applications.name&&row[8].v===5));
  assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('付款进度仍提示真实的多笔来源冲突、来源日期异常和缺少OA流程',()=>{
 const i=fixture(),source=X.utils.sheet_to_json(i.ledger.book.Sheets['对账明细-徐英'],{header:1})[1];
 append(i,'ledger',source,'对账明细-操思敏');
 const r=C.build(i);assert.match(review(r).find(row=>row[8].v===3)[5].v,/对应多笔记录/);assert.equal(records(r)[1]['跟单员'].v,'');
 const j=fixture();set(j,'ledger','应付款时间',2,'日期未知','对账明细-徐英');
 assert.match(review(C.build(j)).find(row=>row[8].v===3)[5].v,/应付时间为空或无法识别/);
 const k=fixture();set(k,'applications','流程编码',5,ID(99));
 const note=review(C.build(k)).find(row=>row[8].v===5)[5].v;
 assert.match(note,/OA流程表未找到/);assert.doesNotMatch(note,/跟单员\/资金台账|应付时间未填充/);
});
test('分批申请优先编号名称主体、候选多笔才用金额区分，不按同名猜测',()=>{
 const i=fixture(),source=i.ledger.book.Sheets['对账明细-李奇'],original=X.utils.sheet_to_json(source,{header:1})[1];
 append(i,'ledger',original.map((v,c)=>c===7?40:v));let p=records(C.build(i));assert.equal(p[0]['跟单员'].v,'李奇');
 append(i,'ledger',original.map((v,c)=>c===7?40:v),'对账明细-操思敏');p=records(C.build(i));assert.equal(p[0]['跟单员'].v,'');
 set(i,'applications','付款主体',2,'国源');p=records(C.build(i));assert.equal(p[0]['跟单员'].v,'');assert.equal(p[0]['应付时间'].v,'');
 set(i,'applications','供应商全称',3,'另一个供应商');assert.equal(records(C.build(i))[1]['跟单员'].v,'');
});
test('没有OA按名称金额匹配来源，错误OA不降级名称匹配，缺失流程一律待核对',()=>{
 for(const code of ['',ID(99),'未提交OA']){
  const i=fixture();set(i,'applications','流程编码',3,code);const r=C.build(i),p=records(r)[1];
  assert.equal(p['是否已付款'].v,'待核对');assert.equal(p['跟单员'].v,code===ID(99)?'':'徐英');assert.equal(p['对公付款OA号'].v,code);
  assert(review(r).some(row=>row[8].v===3));assert.equal(r.stats.pending,0);
 }
});
test('OA同号信息一致可复用，信息冲突不选第一行；多个OA逐号列出且不拆金额',()=>{
 const i=fixture();append(i,'oa',[ID(1),'合成供应商1',100,'另一节点','合成审核人','归档','OA申请人']);
 let r=C.build(i),p=records(r)[0];assert.equal(p['是否已付款'].v,'待核对');assert.equal(p['当前节点'].v,'');assert.match(review(r)[0][5].v,/进度信息冲突.*Sheet0 第/);
 const j=fixture();set(j,'applications','流程编码',2,ID(1)+' / '+ID(2));r=C.build(j);p=records(r)[0];
 assert.equal(p['是否已付款'].v,'待核对');assert.equal(p['申请付款金额'].v,40);assert.equal(r.progressRows.length,4);assert.match(p['流程状态'].v,new RegExp(ID(1)+'：归档.*'+ID(2)+'：未归档','s'));assert.equal(p['跟单员'].v,'');
});
test('非办理人执行节点按归档与未归档判断，空白异常错误单元格及无缓存公式不会判成已付款',()=>{
 for(const status of ['', '已归档', '归档中', '已撤销', {t:'e',v:23,w:'#REF!'}, {t:'s',f:'A1'}]){
  const i=fixture();set(i,'oa','流程状态',3,status);const r=C.build(i);assert.equal(records(r)[1]['是否已付款'].v,'待核对');assert(review(r).some(row=>row[8].v===3));
 }
 const i=fixture();set(i,'oa','当前节点',3,{t:'e',v:23});assert.equal(records(C.build(i))[1]['是否已付款'].v,'待核对');
});
test('当前节点为办理人执行时进度显示已付款，包括未归档或空白流程状态',()=>{
 for(const status of ['未归档','归档','']){
  const i=fixture(),before=C.build(i);set(i,'oa','当前节点',3,' 办理人执行\n');set(i,'oa','流程状态',3,status);
  const source=JSON.stringify(i),r=C.build(i),p=records(r)[1];
  assert.equal(p['是否已付款'].v,'已付款');assert.equal(p['流程状态'].v,status);
  assert.deepEqual(r.highlights,before.highlights);assert.deepEqual(r.negotiated,before.negotiated);
  for(const h of ['已付金额','剩余金额','是否结清'])assert.deepEqual(r.normalRows.map(row=>row[C.NORMAL_HEADERS.indexOf(h)]),before.normalRows.map(row=>row[C.NORMAL_HEADERS.indexOf(h)]));
  assert.equal(JSON.stringify(i),source,'不修改输入台账或OA记录');
 }
 for(const node of ['', '财务审批','待办理人执行','办理人执行、财务审批']){
  const i=fixture();set(i,'oa','当前节点',3,node);assert.equal(records(C.build(i))[1]['是否已付款'].v,'未付款');
 }
});
test('办理人执行仍保留重新申请优先、OA冲突和异常字段待核对规则',()=>{
 const i=fixture();set(i,'oa','当前节点',3,'办理人执行');X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],[''],['重新申请']],{origin:'G1'});
 assert.equal(records(C.build(i))[1]['是否已付款'].v,'重新申请');
 for(const status of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A1'}]){
  const j=fixture();set(j,'oa','当前节点',3,'办理人执行');set(j,'oa','流程状态',3,status);assert.equal(records(C.build(j))[1]['是否已付款'].v,'待核对');
 }
 const conflict=fixture();set(conflict,'oa','当前节点',2,'办理人执行');assert.equal(records(C.build(conflict))[0]['是否已付款'].v,'待核对','同OA多个节点冲突不按一行判断');
 const multi=fixture();set(multi,'oa','当前节点',3,'办理人执行');set(multi,'applications','流程编码',3,ID(1)+' / '+ID(2));assert.equal(records(C.build(multi))[1]['是否已付款'].v,'待核对');
});
test('空白、零值、金额异常、日期异常及1904日期分别保留，不隐式变零',()=>{
 for(const value of ['',0,'0.001','异常',-1,{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'applications','本次申请付款金额',5,value);const r=C.build(i),c=records(r)[3]['申请付款金额'];
  assert.equal(c.v,typeof value==='object'?'#REF!':value);assert.equal(c.t,typeof value==='number'?'n':'s');assert.equal(r.progressRows.length,4);
 }
 const i=fixture();set(i,'applications','申请日期',3,'日期未知');let r=C.build(i);assert.equal(records(r)[1]['申请日期'].v,'日期未知');assert.match(review(r).find(row=>row[8].v===3)[5].v,/申请日期为空或无法识别/);
 i.applications.book.Workbook={WBProps:{date1904:true}};set(i,'applications','申请日期',3,46328-1462);r=C.build(i);assert.equal(display(records(r)[1]['申请日期']),'2026年11月2日');
});
test('旧合同号兼容；新流程编码列存在且为空时不回退；不依据OA申请人填跟单员',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;s.D1.v='合同号';assert.equal(records(C.build(i))[0]['是否已付款'].v,'已付款');
 X.utils.sheet_add_aoa(s,[['流程编码'],['']],{origin:'G1'});let p=records(C.build(i))[0];assert.equal(p['对公付款OA号'].v,'');assert.equal(p['是否已付款'].v,'待核对');assert.equal(p['跟单员'].v,'');
});
test('原三表金额、提交标红及公式与上一版一致，进度核对不增加下载阻断',()=>{
 const i=fixture(),first=C.build(i),again=C.build(i);assert.deepEqual(again,first);
 assert.deepEqual(first.normalRows.map(row=>row[C.NORMAL_HEADERS.indexOf('供应商名称')].v),['合成供应商2','合成供应商3']);
 assert.equal(first.normalRows[0][C.NORMAL_HEADERS.indexOf('申请付款的日期')].v,'2026/11/02周');
 assert.equal(first.applicationComments.M2,'申请日期（资金整理）：2026年11月1日；2026年11月3日');
 set(i,'oa','流程状态',2,'未归档');set(i,'oa','流程状态',6,'未归档');assert.equal(records(C.build(i))[0]['是否已付款'].v,'已付款','明确出纳支付记录不由流程状态替代');
 assert.deepEqual(C.build(i).highlights,first.highlights);
});
test('没有申请时仍输出固定12列表头与空进度表',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',[AH]);const r=C.build(i);assert.deepEqual(r.progressRows,[]);assert.deepEqual(r.progressHeaders,['申请日期','跟单员','供应商名称','付款主体','对公付款OA号','申请付款金额','应付时间','当前节点','未操作者','流程状态','申请人','是否已付款']);
 assert.equal(C.OUTPUT_SHEETS.length,4);
});

test('付款分类沿用台账谈判优先规则，应付时间不会误取台账日期',()=>{
 const i=fixture();set(i,'applications','付款分类',2,'正常付款');const r=C.build(i),p=records(r)[0];
 assert.equal(display(p['应付时间']),'2026年11月6日');assert.equal(r.normalRows.some(row=>row[2].v==='合成供应商1'),false);assert.match(review(r).find(row=>row[8].v===2)[5].v,/付款分类/);
 const j=fixture();delete j.applications.book.Sheets.付款申请汇总.F1;assert.equal(display(records(C.build(j))[0]['应付时间']),'2026年11月6日');
});
test('多OA换行文本统一显示，最新上传OA状态变化不沿用旧状态',()=>{
 const i=fixture();set(i,'applications','流程编码',2,ID(1)+'\n'+ID(2));const p=records(C.build(i))[0];assert.equal(p['对公付款OA号'].v,ID(1)+'；'+ID(2));assert.equal(p['是否已付款'].v,'待核对');
});

test('核对说明含重新申请时优先显示重新申请，不改变付款状态金额和标红',()=>{
 for(const note of ['重新申请','退回后重新申请付款','重新\n申请',{t:'s',v:'重新申请',f:'"重新申请"'}]){
  const i=fixture(),before=C.build(i);X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],[note]],{origin:'G1'});
  const r=C.build(i),p=records(r);assert.equal(p[0]['是否已付款'].v,'重新申请');assert.equal(p[0]['流程状态'].v,'归档');assert.equal(p[2]['是否已付款'].v,'已付款','同OA重复申请只按本行说明判断');
  for(const key of ['normalRows','negotiated','highlights','applicationComments','issues'])assert.deepEqual(r[key],before[key]);
  assert.deepEqual(r.reviewRows,before.reviewRows.filter(row=>!row[5].v.startsWith('重复提交')));
  for(const code of ['',ID(99),ID(1)+' / '+ID(2)]){set(i,'applications','流程编码',2,code);assert.equal(records(C.build(i))[0]['是否已付款'].v,'重新申请');}
 }
 for(const note of ['',0,'核对通过',{t:'e',v:'重新申请'},{t:'s',f:'A1'}]){
  const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],[note]],{origin:'G1'});assert.equal(records(C.build(i))[0]['是否已付款'].v,'已付款');
 }
 const i=fixture();X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['核对说明'],[''],['重新申请']],{origin:'G1'});assert.equal(records(C.build(i))[1]['是否已付款'].v,'重新申请','未归档也优先采用重新申请');
});
