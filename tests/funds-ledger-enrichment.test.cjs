const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-ledger-core.js');
const LH=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否提交付款申请单','是否结清','是否谈判付款'];
const AH=['申请日期','供应商全称','付款主体','流程编码','本次申请付款金额','是否签字完成'];
const PH=['供应商名称','付款主体','对公付款OA号','申请付款金额','当前节点','未操作者','流程状态','申请人','是否已付款'];
const ledgerRow=(o={})=>({'供应商名称':'供应商甲','付款主体':'国源','对公付款0A流程号':'OA001','申请付款金额':100,'摘要':'货款','应付款时间':'2026-11-06','预估付款时间':'2026-12','是否提交付款申请单':'否','是否结清':'否','是否谈判付款':'否',...o});
const app=(o={})=>({'申请日期':'2026-11-01','供应商全称':'供应商甲','付款主体':'国源','流程编码':'OA001','本次申请付款金额':100,'是否签字完成':'签完字',...o});
const progress=(o={})=>({'供应商名称':'供应商甲','付款主体':'国源','对公付款OA号':'OA001','申请付款金额':100,'当前节点':'出纳','未操作者':'人员甲','流程状态':'归档','申请人':'人员乙','是否已付款':'已付款',...o});
function sheet(headers,rows,title=false){return X.utils.aoa_to_sheet([...(title?[['标题']]:[]),headers,...rows.map(row=>headers.map(h=>row[h]??''))]);}
function ledger(rows){const b=X.utils.book_new();['李奇','徐英','邢芳芳','操思敏','孙立柱'].forEach((p,i)=>X.utils.book_append_sheet(b,sheet(LH,i?[]:rows),'对账明细-'+p));return b;}
function source(role,rows,headers=role==='applications'?AH:PH,title=false){const b=X.utils.book_new(),name=role==='applications'?'付款申请汇总':'付款进度整理';X.utils.book_append_sheet(b,sheet(headers,rows,title),name);return {book:b,filename:role==='applications'?'申请.xlsx':'进度.xlsx'};}
const build=(rows,applications,flow)=>C.build(ledger(rows),'台账.xlsx',{...(applications?{applications}:{}),...(flow?{progress:flow}:{})});
const extra=(r,s=0,i=0)=>r.sheets[s].rows[i].slice(7).map(C.formatCell);
const old=result=>result.sheets.map(sheet=>sheet.rows.map(row=>row.slice(0,7)));
test('恒定追加7列；没上传补充来源时全空，旧列/行序/来源/数量不变',()=>{
 const l=ledger([ledgerRow(),ledgerRow({'对公付款0A流程号':''}),ledgerRow({'是否结清':'是'})]),base=C.build(l,'台账.xlsx'),a=source('applications',[app()]),p=source('progress',[progress()]),before=JSON.stringify([l,a,p]);
 const r=C.build(l,'台账.xlsx',{applications:a,progress:p});assert.deepEqual(old(r),old(base));assert.deepEqual(r.counts,base.counts);assert.deepEqual(r.sheets.map(s=>s.sources),base.sheets.map(s=>s.sources));
 assert(base.sheets.every(s=>s.rows.every(row=>row.length===14&&row.slice(7).every(c=>c.v===''))));assert.deepEqual(C.HEADERS[0].slice(7),['申请日期','是否已签字','当前节点','未操作者','流程状态','申请人','是否已付款']);assert.equal(JSON.stringify([l,a,p]),before);
 const out=C.workbook(r);assert.equal(out.Sheets[C.SHEET_NAMES[0]]['!autofilter'].ref,'A1:N2');assert.equal(out.Sheets[C.SHEET_NAMES[0]]['!cols'].length,14);
});
test('OA精确匹配隔离同供应商同金额不同流程',()=>{
 const r=build([ledgerRow(),ledgerRow({'对公付款0A流程号':'OA002'})],source('applications',[app(),app({'流程编码':'OA002','申请日期':'2026-11-02','是否签字完成':'未签字'})]),source('progress',[progress(),progress({'对公付款OA号':'OA002','流程状态':'未归档','当前节点':'财务'})]));
 assert.equal(r.sheets[0].rows.length,2);assert.deepEqual(extra(r),['2026-11-01','签完字','出纳','人员甲','归档','人员乙','已付款']);assert.equal(extra(r,0,1)[0],'2026-11-02');assert.equal(extra(r,0,1)[6],'未付款');
});
test('OA存在但匹配不到不会按金额兜底；供应商和明确主体冲突不补入',()=>{
 const l=[ledgerRow()],a=source('applications',[app({'流程编码':'OA002'})]),p=source('progress',[progress({'付款主体':'迈德斯特'})]),r=build(l,a,p);
 assert.deepEqual(extra(r),['','','','','','','待核对']);assert(r.warnings.some(w=>/台账.xlsx.*李奇.*付款申请未找到/.test(w)));assert(r.warnings.some(w=>/主体冲突.*进度.xlsx/.test(w)));
 for(const changed of [{'供应商全称':'供应商乙'},{'付款主体':'迈德斯特'}]){const actual=build(l,source('applications',[app(changed)]));assert.deepEqual(extra(actual),['','','','','','','']);assert.match(actual.warnings.join('\n'),/冲突.*申请.xlsx.*第2行/);}
});
test('同OA多次申请日期升序且同日按来源序，签字空值编号对应，不合并或丢记录',()=>{
 const r=build([ledgerRow()],source('applications',[app({'申请日期':'2026-11-03','是否签字完成':'后签'}),app({'申请日期':'2026-11-01','是否签字完成':''}),app({'申请日期':'2026-11-01','是否签字完成':'同日签'}),app({'申请日期':'','是否签字完成':'待签'})]));
 assert.equal(extra(r)[0],'1、2026-11-01\n2、2026-11-01\n3、2026-11-03\n4、空白');assert.equal(extra(r)[1],'1、空白\n2、同日签\n3、后签\n4、待签');assert.equal(r.sheets[0].rows.length,1);
});
test('唯一无OA按供应商主体金额补申请，借用申请OA查进度但不改原OA和分类',()=>{
 for(const oa of ['', ' 未提交 OA ']){
  const r=build([ledgerRow({'对公付款0A流程号':oa})],source('applications',[app()]),source('progress',[progress()]));
  assert.equal(r.sheets[0].rows.length,0);assert.equal(r.sheets[1].rows[0][2].v,oa);assert.deepEqual(extra(r,1),['2026-11-01','签完字','出纳','人员甲','归档','人员乙','已付款']);
  assert.match(r.warnings.join('\n'),/台账.xlsx.*第2行.*借用.*OA001.*申请.xlsx.*付款申请汇总 第2行/);
 }
});
test('无OA候选多条或同一申请可归属多条台账均留空，已知OA行也阻止重复借用',()=>{
 const missing=ledgerRow({'对公付款0A流程号':''}),a=source('applications',[app()]);
 for(const rows of [[missing,missing],[ledgerRow(),missing]]){const r=build(rows,a);assert.deepEqual(extra(r,1),['','','','','','','']);assert.match(r.warnings.join('\n'),/可能对应多条台账/);}
 const many=build([missing],source('applications',[app(),app({'流程编码':'OA002'})]));assert.deepEqual(extra(many,1),['','','','','','','']);assert.match(many.warnings.join('\n'),/多条.*候选/);
});
test('已有OA优先认领分批申请，不让另一无OA同额行重复借用',()=>{
 const r=build([ledgerRow(),ledgerRow({'对公付款0A流程号':'','申请付款金额':20})],source('applications',[app({'本次申请付款金额':20})]),source('progress',[progress({'申请付款金额':20})]));
 assert.equal(extra(r)[0],'2026-11-01');assert.deepEqual(extra(r,1),['','','','','','','待核对']);assert.match(r.warnings.join('\n'),/已由其他 OA 台账认领/);assert(!r.warnings.some(w=>w.includes('借用')));
});
test('筛选排除行仍保留OA申请归属，缺OA行不能借走已结清或谈判款的分批申请',()=>{
 for(const excluded of [{'是否结清':'是'},{'是否谈判付款':'是'}]){
  const rows=[ledgerRow({...excluded,'摘要':{t:'n',f:'A99'}}),ledgerRow({'对公付款0A流程号':'未提交OA','申请付款金额':20})],a=source('applications',[app({'本次申请付款金额':20})]),p=source('progress',[progress({'申请付款金额':20})]);
  const l=ledger(rows),before=JSON.stringify(l),base=C.build(l,'台账.xlsx'),r=C.build(l,'台账.xlsx',{applications:a,progress:p});
  assert.equal(r.sheets[0].rows.length,0);assert.deepEqual(extra(r,1),['','','','','','','待核对']);assert.deepEqual(old(r),old(base));assert.deepEqual(r.counts,base.counts);assert.equal(JSON.stringify(l),before);assert.match(r.warnings.join('\n'),/筛选排除行.*台账.xlsx.*第2行/);assert(!r.warnings.some(w=>w.includes('仅借用')));
 }
});
test('已结清空OA同额行及异常身份缓存均使fallback归属不确定，但不新增公式阻断',()=>{
 for(const excluded of [ledgerRow({'是否结清':'是','对公付款0A流程号':'','申请付款金额':20}),ledgerRow({'是否结清':'是','对公付款0A流程号':{t:'n',f:'A99'},'申请付款金额':20}),ledgerRow({'是否结清':'是','对公付款0A流程号':'','申请付款金额':{t:'e',v:23,w:'#REF!'}})]){
  const r=build([excluded,ledgerRow({'对公付款0A流程号':'','申请付款金额':20})],source('applications',[app({'本次申请付款金额':20})]));
  assert.equal(r.sheets[0].rows.length,0);assert.equal(r.sheets[1].rows.length,1);assert.deepEqual(extra(r,1),['','','','','','','']);assert.match(r.warnings.join('\n'),/其他 OA 台账认领|可能对应多条台账/);
 }
});
test('只有摘要或日期的空白模版行不占用申请，身份字段含公式或错误仍保留不确定性',()=>{
 const template=ledgerRow({'供应商名称':'','付款主体':'','对公付款0A流程号':'','申请付款金额':'','摘要':'填写说明','是否结清':'','是否谈判付款':''}),target=ledgerRow({'对公付款0A流程号':'','申请付款金额':20}),a=source('applications',[app({'本次申请付款金额':20})]);
 const r=build([template,target],a);assert.equal(extra(r,1)[0],'2026-11-01');assert.equal(r.sheets[1].rows.length,1);
 for(const value of [{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const uncertain=build([{...template,'申请付款金额':value},target],a);assert.equal(extra(uncertain,1)[0],'');assert.match(uncertain.warnings.join('\n'),/可能对应多条台账/);}
});
test('同OA多个台账按本笔金额唯一分配申请，金额不能分清时保持原行并留空',()=>{
 const rows=[ledgerRow({'申请付款金额':20}),ledgerRow({'申请付款金额':80})],a=source('applications',[app({'本次申请付款金额':20,'是否签字完成':'20签字'}),app({'本次申请付款金额':80,'申请日期':'2026-11-02','是否签字完成':'80签字'})]);
 const r=build(rows,a);assert.equal(extra(r)[1],'20签字');assert.equal(extra(r,0,1)[1],'80签字');
 const ambiguous=build([rows[0],rows[0]],source('applications',[app({'本次申请付款金额':20})]),source('progress',[progress({'申请付款金额':20})]));
 assert.equal(ambiguous.sheets[0].rows.length,2);assert.deepEqual(extra(ambiguous),['','','','','','','待核对']);assert.deepEqual(extra(ambiguous,0,1),extra(ambiguous));assert.match(ambiguous.warnings.join('\n'),/同一付款进度可能归属多条台账/);
});
test('无OA金额按分精确且空白不等于0，主体或金额不齐不弱化匹配',()=>{
 for(const [amount,appAmount,matched] of [[0,0,true],['',0,false],[0,'',false],[0.29,'￥0.29',true],[0.291,0.291,false]]){
  const r=build([ledgerRow({'对公付款0A流程号':'','申请付款金额':amount})],source('applications',[app({'本次申请付款金额':appAmount})]));assert.equal(!!extra(r,1)[0],matched,JSON.stringify([amount,appAmount]));
 }
 const r=build([ledgerRow({'对公付款0A流程号':'','付款主体':''})],source('applications',[app({'付款主体':''})]));assert.equal(extra(r,1)[0],'');
});
test('占位或多个OA不能当单一ID匹配，也不按金额兜底',()=>{
 for(const oa of ['未匹配','未申请OA','待补OA','OA001/OA002','OA001\nOA002','OA001 OA002','OA001、OA002']){
  const r=build([ledgerRow({'对公付款0A流程号':oa})],source('applications',[app({'流程编码':oa})]),source('progress',[progress({'对公付款OA号':oa})]));assert.deepEqual(extra(r),['','','','','','','待核对'],oa);assert.equal(r.sheets[0].rows[0][2].v,oa);
 }
 const r=build([ledgerRow({'对公付款0A流程号':''})],source('applications',[app({'流程编码':'OA001/OA002'})]),source('progress',[progress()]));assert.equal(extra(r,1)[0],'2026-11-01');assert.equal(extra(r,1)[6],'待核对');
});
test('进度完全重复只去重信息，同OA分批必须用本笔金额唯一归属，不能将部分付款当整笔',()=>{
 const flow=source('progress',[progress({'申请付款金额':40,'当前节点':'批次一'}),progress({'申请付款金额':60,'当前节点':'批次二'}),progress({'申请付款金额':60,'当前节点':'批次二'})]);
 const matched=build([ledgerRow({'申请付款金额':60})],null,flow);assert.deepEqual(extra(matched).slice(2),['批次二','人员甲','归档','人员乙','已付款']);
 for(const sourceRows of [[progress({'申请付款金额':40})],[progress({'申请付款金额':40}),progress({'申请付款金额':60})]]){
  const r=build([ledgerRow()],null,source('progress',sourceRows));assert.deepEqual(extra(r),['','','','','','','待核对']);assert.match(r.warnings.join('\n'),/金额.*不能唯一对应/);
 }
 const duplicate=build([ledgerRow()],null,source('progress',[progress(),progress()]));assert.equal(duplicate.sheets[0].rows.length,1);assert.equal(extra(duplicate)[6],'已付款');
});
test('进度同OA同金额冲突时不任选，缺金额可展示流程但不可判已付',()=>{
 const conflict=build([ledgerRow()],null,source('progress',[progress(),progress({'当前节点':'主管'})]));assert.deepEqual(extra(conflict),['','','','','','','待核对']);assert.match(conflict.warnings.join('\n'),/信息冲突/);
 const missing=build([ledgerRow()],null,source('progress',[progress({'申请付款金额':''})]));assert.equal(extra(missing)[2],'出纳');assert.equal(extra(missing)[6],'待核对');
 const uncertainBatch=build([ledgerRow()],null,source('progress',[progress(),progress({'申请付款金额':'','是否已付款':'重新申请'})]));assert.deepEqual(extra(uncertainBatch),['','','','','','','待核对']);assert.match(uncertainBatch.warnings.join('\n'),/金额不完整/);
});
test('特殊付款状态优先，普通记录按归档判定，缺状态与未找到均待核对',()=>{
 for(const [state,paid,expected]of [['归档','重新申请','重新申请'],['已归档','待核对','待核对'],['流程归档','未付款','已付款'],['未归档','已付款','未付款'],['','已付款','待核对']]){
  const r=build([ledgerRow()],null,source('progress',[progress({'流程状态':state,'是否已付款':paid})]));assert.equal(extra(r)[6],expected);
 }
 const unmatched=build([ledgerRow()],null,source('progress',[progress({'对公付款OA号':'OA002'})]));assert.deepEqual(extra(unmatched),['','','','','','','待核对']);
});
test('表头顺序/前置标题/别名兼容，流程编码表头存在时不使用合同号单元格兜底',()=>{
 const rows=[app({'合同号':'OA001','流程编码':'','是否已签字':'签完字','供应商名称':'供应商甲'})],headers=['合同号',...AH.filter(h=>!['是否签字完成','供应商全称'].includes(h)),'是否已签字','供应商名称'].reverse();
 const a=source('applications',rows,headers,true);assert.equal(C.validateSupplement(a.book,'applications',a.filename).count,1);assert.equal(extra(build([ledgerRow()],a))[0],'');
 const legacy=source('applications',rows,headers.filter(h=>h!=='流程编码'),true);assert.equal(extra(build([ledgerRow()],legacy))[0],'2026-11-01');
 const p=source('progress',[{...progress(),'对公付款0A流程号':'OA001'}],PH.map(h=>h==='对公付款OA号'?'对公付款0A流程号':h).reverse(),true);assert.equal(extra(build([ledgerRow()],legacy,p))[6],'已付款');
});
test('单次申请日期保持数值类型，1904及缓存日期、缓存零、缓存空签字正确',()=>{
 const a=source('applications',[app({'申请日期':{t:'n',v:44870,f:'DATE(2026,11,6)',z:'yyyy-mm-dd'},'本次申请付款金额':{t:'n',v:0,f:'1-1'},'是否签字完成':{t:'s',v:'',f:'""'}})]);a.book.Workbook={WBProps:{date1904:true}};
 const r=build([ledgerRow({'对公付款0A流程号':'','申请付款金额':0})],a),date=r.sheets[1].rows[0][7];assert.equal(date.t,'n');assert.equal(date.v,46332);assert.equal(C.formatCell(date),'2026-11-06');assert.equal(date.f,undefined);assert.equal(extra(r,1)[1],'');
 const out=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});assert.equal(out.Sheets[C.SHEET_NAMES[1]].H2.t,'n');
});
test('补充文件独立校验必要表头、相关重复字段、所需公式和非法金额，未使用合同号错误不阻断',()=>{
 for(const role of ['applications','progress']){
  const rows=role==='applications'?[app()]:[progress()],headers=role==='applications'?AH:PH,oa=role==='applications'?'流程编码':'对公付款OA号',missing=source(role,rows,headers.filter(h=>h!==oa));assert.throws(()=>C.validateSupplement(missing.book,role,missing.filename),/缺少必要表头/);
  for(const h of headers){const bad=source(role,[{...rows[0],[h]:{t:'n',f:'A99'}}]);assert.throws(()=>C.validateSupplement(bad.book,role,bad.filename),/\.xlsx.*![A-I]2.*公式缺少/);}
  const invalid=source(role,[{...rows[0],[role==='applications'?'本次申请付款金额':'申请付款金额']:'待核对'}]);assert.throws(()=>C.validateSupplement(invalid.book,role,invalid.filename),/金额无法识别/);
 }
 const ignored=source('applications',[app({'合同号':{t:'n',f:'A99'}})],[...AH,'合同号']);assert.equal(C.validateSupplement(ignored.book,'applications',ignored.filename).count,1);
 const duplicate=source('applications',[app({'供应商名称':'别的供应商'})],[...AH,'供应商名称']);assert.throws(()=>C.validateSupplement(duplicate.book,'applications',duplicate.filename),/字段“供应商全称”重复/);
 const wrong=X.utils.book_new();X.utils.book_append_sheet(wrong,sheet(AH,[app()]),'历史付款申请汇总');assert.throws(()=>C.validateSupplement(wrong,'applications','错误.xlsx'),/未找到“付款申请汇总”/);
});
test('可选字段缺列只提示，上传新来源重新构建不会遗留旧日期/状态',()=>{
 const a=source('applications',[app()],AH.filter(h=>h!=='是否签字完成')),validated=C.validateSupplement(a.book,'applications',a.filename);assert.match(validated.warnings.join('\n'),/缺少 是否签字完成/);assert.equal(extra(build([ledgerRow()],a))[1],'');
 const l=ledger([ledgerRow()]),first=C.build(l,'台账.xlsx',{applications:source('applications',[app()]),progress:source('progress',[progress()])}),fresh=C.build(l,'台账.xlsx',{applications:source('applications',[app({'申请日期':'2026-11-07','是否签字完成':'未签字'})]),progress:source('progress',[progress({'流程状态':'未归档'})])});
 assert.equal(extra(first)[0],'2026-11-01');assert.equal(extra(fresh)[0],'2026-11-07');assert.equal(extra(fresh)[1],'未签字');assert.equal(extra(fresh)[6],'未付款');assert.deepEqual(extra(C.build(l,'台账.xlsx')),['','','','','','','']);
});
