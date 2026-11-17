const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-check-core.js');
const ph=['单据编号','供应商','物料编码','剩余入库数量','关闭状态'];
const fh=['供应商简称','物料编码','采购订单号','未交付数量'];
function source(rows,kind,changes={}){const book=X.utils.book_new();if(kind==='po')X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([ph,...rows]),'Sheet1');else for(const person of ['徐英','操思敏','邢芳芳','李奇'])X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([fh,...(person==='徐英'?rows:[])]),person);return {book,filename:kind+'.xlsx',...changes};}
const build=(p,f,decisions={})=>C.build({purchase:source(p,'po'),fulfillment:source(f,'f'),decisions});
const po=(id='CGDD1',q=5,s='甲有限公司',code='001')=>[id,s,code,q,'未关闭'];
const fu=(id='CGDD1',q=5,s='甲',code='001')=>[s,code,id,q];
test('按订单聚合重复行，但不同订单不互相抵消，前导零保留',()=>{
 const r=build([po('CGDD1',2),po('CGDD1',3),po('CGDD2',10)],[fu('CGDD1',6),fu('CGDD2',9)]);
 assert.equal(r.records.length,2);assert.deepEqual(r.records.map(x=>x.status),['数量差异','数量差异']);assert.deepEqual(r.records.map(x=>x.difference),[1,-1]);assert.equal(r.summary[0][1],15);assert.equal(r.summary[0][2],15);assert.equal(r.detail[0][3],'001');
});
test('多编号拆号查询合计核对，不重复履约数量；同组原行仍相加',()=>{
 const r=build([po('CGDD1',3),po('CGDD2',7)],[fu('CGDD1+CGDD2',6),fu('CGDD2+CGDD1',4)]);
 assert.equal(r.records.length,1);assert.equal(r.records[0].status,'一致');assert.equal(r.records[0].fqty,10);assert.equal(r.records[0].pqty,10);
});
test('重叠集合不分摊、不重复消耗采购数量，保留待核对',()=>{
 const r=build([po('CGDD1',3),po('CGDD2',7)],[fu('CGDD1+CGDD2',10),fu('CGDD1',3)]);
 assert(r.records.every(x=>x.status==='待核对'));assert.equal(r.summary[0][1],10);assert.equal(r.summary[0][2],13);assert(r.records.filter(x=>x.fulfillment.length).every(x=>x.pqty===null));assert.equal(r.summary[0][3],null);
});
test('只计算未关闭，关闭仍有履约时提示，不把空白剩余推算为零',()=>{
 const r=build([[...po('CGDD1',999).slice(0,4),'已关闭'],po('CGDD2','')],[fu('CGDD1',5),fu('CGDD2',0)]);
 assert.equal(r.records[0].status,'已关闭异常');assert.equal(r.records[0].pqty,0);assert.equal(r.records[0].difference,null);assert.equal(r.records[1].status,'待核对');assert.equal(r.records[1].pqty,null);assert.equal(r.summary[0][1],null);
});
test('单边、空白订单号、未知关闭状态、错误数量明确留空',()=>{
 const r=build([po('CGDD1',0),['CGDD2','乙公司','002',8,'']], [fu('CGDD3',0),fu('',4)]);
 assert(r.records.some(x=>x.status==='仅采购'&&x.pqty===0&&x.fqty===null));assert(r.records.some(x=>x.status==='仅履约'&&x.pqty===null&&x.fqty===0));assert(r.records.filter(x=>x.status!=='一致').every(x=>x.difference===null));assert(r.detail.some(x=>x[16].includes('关闭状态')));
});
test('供应商相同编号证据冲突须人工确认，不被包含规则覆盖',()=>{
 const r=build([po('CGDD1',3,'甲公司'),po('CGDD2',4,'乙公司')],[fu('CGDD1',3),fu('CGDD2',4)]);
 assert.equal(r.choices.length,1);assert.equal(r.choices[0].candidates.length,2);assert(r.records.filter(x=>x.fulfillment.length).every(x=>x.status==='待核对'));
 const fixed=build([po('CGDD1',3,'甲公司'),po('CGDD2',4,'乙公司')],[fu('CGDD1',3),fu('CGDD2',4)],{'甲':'甲公司'});
 assert.equal(fixed.records[0].status,'一致');assert.equal(fixed.records[1].status,'待核对');
});
test('仅重算所需公式，忽略辅助REF错误与筛选隐藏，原公式缓存不变',()=>{
 const p=source([po('CGDD1',12)],'po'),f=source([fu('CGDD1',0)],'f'),s=f.book.Sheets.徐英;
 s.D2={t:'n',f:'SUBTOTAL(109,E2:E3)',v:0};s.E2={t:'n',v:5};s.E3={t:'n',v:7};s.F2={t:'e',f:'#REF!',v:23};s['!rows']=[{}, {hidden:true},{hidden:true}];s['!autofilter']={ref:'A1:F3',filterColumn:[{}]};s['!ref']='A1:F3';
 const r=C.build({purchase:p,fulfillment:f});assert.equal(r.records[0].status,'一致');assert.equal(r.records[0].fqty,12);assert.equal(s.D2.v,0);assert.equal(s.D2.f,'SUBTOTAL(109,E2:E3)');
 s.D2.f='UNKNOWN(E2)';const fail=C.build({purchase:p,fulfillment:f});assert.equal(fail.records[0].status,'待核对');assert.match(fail.records[0].action,/UNKNOWN/);
});
test('人员日期后缀、不同字段位置、无订单列、纵向合并数量只算首行',()=>{
 const p=source([po('CGDD1',6)],'po'),f=source([],'f');f.book.SheetNames[0]='徐英1117';delete f.book.Sheets.徐英;
 const s=X.utils.aoa_to_sheet([['未交付数量','供应商名称','物料编码','单据编号'],[6,'甲','001','CGDD1'],[null,'甲','001','CGDD1']]);s['!merges']=[{s:{r:1,c:0},e:{r:2,c:0}}];f.book.Sheets.徐英1117=s;
 f.book.Sheets.邢芳芳=X.utils.aoa_to_sheet([['供应商','物料编码','未交付数量'],['甲','002',5]]);
 const r=C.build({purchase:p,fulfillment:f});assert.equal(r.fulfillmentRows.length,3);assert.equal(r.records[0].fqty,6);assert.equal(r.records[0].status,'一致');assert.equal(r.records[1].status,'待核对');assert.match(r.records[1].action,/订单号为空/);
});
test('小数、零、异常、总计行和其他人员不混入',()=>{
 const p=source([po('CGDD1',0.1),po('CGDD1',0.2),['合计','','',0.3,'']],'po');p.book.Sheets.Sheet1.A4={t:'s',v:''};p.book.Sheets.Sheet1.B4={t:'s',v:'合计'};
 const f=source([fu('CGDD1',0.3)],'f');X.utils.book_append_sheet(f.book,X.utils.aoa_to_sheet([fh,fu('CGDD1',999)]),'孙立柱');
 const r=C.build({purchase:p,fulfillment:f});assert.equal(r.stats.purchase,2);assert.equal(r.stats.fulfillment,1);assert.equal(r.records[0].difference,0);
 assert.throws(()=>C.number('474套左右'),/明确数值/);assert.equal(C.number('1,245'),1245);assert.throws(()=>C.number(''),/为空/);
});
test('缺必要表头与多张同人员来源不猜测',()=>{
 const p=source([po()],'po'),f=source([fu()],'f');p.book.Sheets.Sheet1.B1.v='其他';assert.throws(()=>C.build({purchase:p,fulfillment:f}),/表头缺失/);
 X.utils.book_append_sheet(f.book,X.utils.aoa_to_sheet([fh]),'徐英1117');assert.throws(()=>C.build({purchase:source([po()],'po'),fulfillment:f}),/多张/);
});
module.exports={source,po,fu};
function directory(rows,headers=['供应商名称','供应商简称']){const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'名录');return {book,filename:'供应商名录.xlsx'};}
test('名录全称映射简称，汇总和两个差异供应商字段一致，未知名称保留全称',()=>{
 const p=source([po('CGDD1',3),po('CGDD2',2,'未收录有限公司')],'po'),f=source([fu('CGDD1',4,'甲有限公司'),fu('CGDD2',3,'未收录有限公司')],'f');
 const before=JSON.stringify([p,f]),r=C.build({purchase:p,fulfillment:f,directory:directory([['甲有限公司','甲供']])});
 assert.deepEqual(r.summary.map(x=>x[0]),['甲供','未收录有限公司']);assert.equal(r.detail[0][1],'甲供');assert.equal(r.detail[0][2],'甲供');assert.equal(r.detail[1][1],'未收录有限公司');assert.equal(r.detail[1][2],'未收录有限公司');assert.equal(JSON.stringify([p,f]),before);
});
test('名录支持三个全称表头、列顺序变化、空简称及相同重复项',()=>{
 for(const header of ['供应商','供应商名称','供应商全称']){
  const d=directory([['甲供','甲有限公司'],['甲供','甲有限公司'],['','乙有限公司']],['供应商简称',header]),catalog=C.supplierDirectory(d);
  assert.equal(catalog.identity(' 甲供 '),'甲有限公司');assert.equal(catalog.display('甲有限公司','原值'),'甲供');assert.equal(catalog.display('乙有限公司','原值'),'乙有限公司');assert.equal(catalog.warnings.length,0);
 }
});
test('同简称不同公司不合并，全称对应多个简称时保持全称并提示来源',()=>{
 const d=directory([['甲有限公司','同简称'],['乙有限公司','同简称'],['丙有限公司','丙一'],['丙有限公司','丙二']]),catalog=C.supplierDirectory(d);
 assert.equal(catalog.identity('同简称'),'');assert.equal(catalog.display('丙有限公司','丙'),'丙有限公司');assert.equal(catalog.warnings.length,2);assert(catalog.warnings.every(x=>x.includes('供应商名录.xlsx')));
 const r=C.build({purchase:source([po('CGDD1',3,'甲有限公司'),po('CGDD2',2,'乙有限公司')],'po'),fulfillment:source([fu('CGDD1',3,'甲有限公司'),fu('CGDD2',2,'乙有限公司')],'f'),directory:d});
 assert.equal(r.summary.length,2);assert.deepEqual(r.summary.map(x=>x.slice(0,3)),[['同简称',3,3],['同简称',2,2]]);
});
test('明确名录对应关系优先，不再以错误订单证据把简称认成另一公司',()=>{
 const r=C.build({purchase:source([po('CGDD1',5,'乙有限公司')],'po'),fulfillment:source([fu('CGDD1',5,'甲别称')],'f'),directory:directory([['甲有限公司','甲别称']])});
 assert.equal(r.records[0].supplier,'甲有限公司');assert.equal(r.records[0].status,'待核对');assert.match(r.records[0].action,/供应商不一致/);assert.equal(r.records[0].pqty,null);
});
test('采购来源本身为简称可通过名录识别，但源名称和数量不改',()=>{
 const p=source([po('CGDD1',0,'甲别称')],'po'),f=source([fu('CGDD1',0,'甲有限公司')],'f');
 const r=C.build({purchase:p,fulfillment:f,directory:directory([['甲有限公司','甲别称']])});assert.equal(r.records[0].status,'一致');assert.equal(r.summary[0][0],'甲别称');assert.equal(r.purchaseRows[0].rawSupplier,'甲别称');assert.equal(r.summary[0][1],0);
});
test('替换或移除名录重新使用当前简称，不保留旧版本映射',()=>{
 const inputs={purchase:source([po()],'po'),fulfillment:source([fu()],'f')};
 assert.equal(C.build({...inputs,directory:directory([['甲有限公司','旧简称']])}).summary[0][0],'旧简称');
 assert.equal(C.build({...inputs,directory:directory([['甲有限公司','新简称']])}).summary[0][0],'新简称');
 assert.equal(C.build(inputs).summary[0][0],'甲有限公司');
});
test('名录缺表头、重复表头及名称公式错误不能静默忽略',()=>{
 assert.throws(()=>C.supplierDirectory(directory([['甲','简称']],['其他','供应商简称'])),/找不到/);
 assert.throws(()=>C.supplierDirectory(directory([['甲','甲','简称']],['供应商','供应商名称','供应商简称'])),/重复/);
 const d=directory([['甲有限公司','']]);d.book.Sheets.名录.B2={t:'n',f:'UNKNOWN(A2)',v:0};assert.throws(()=>C.supplierDirectory(d),/名录 第 2 行（B2）/);
});
