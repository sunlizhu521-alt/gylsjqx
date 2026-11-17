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
