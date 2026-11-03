const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../internal-statement-core.js');
const headers=['供应商','采购组织','单据编号','入库日期','物料编码','物料名称','实收数量','含税单价','价税合计','金额','备注'];
const row=(overrides={})=>Object.assign({供应商:'供应商甲',采购组织:'采购组织甲',单据编号:'RK001',入库日期:'2026/6/30',物料编码:'00123',物料名称:'测试物料',实收数量:2,含税单价:0.1695,价税合计:0.34,金额:0.3,备注:'原备注'},overrides);
function book(rows){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows.map(r=>headers.map(h=>r[h]??''))]),'Sheet1');return b;}
test('按采购组织供应商月份隔离，同编码同单价精确合计，备注去重且保留源顺序与前导零',()=>{
 const b=book([row(),row({实收数量:3,价税合计:0.51,备注:'另一备注'}),row({实收数量:1,价税合计:0.17}),row({含税单价:0.2,价税合计:0.4}),row({采购组织:'采购组织乙'}),row({供应商:'供应商乙'}),row({入库日期:'2026/7/1'})]);
 const before=JSON.stringify(b),r=C.parse(b,'源.xlsx');assert.equal(r.count,7);assert.equal(r.groups.length,4);assert.equal(r.groups[0].items.length,2);const i=r.groups[0].items[0];assert.equal(i.code,'00123');assert.equal(i.quantity,'6');assert.equal(i.price,'0.1695');assert.equal(i.amount,'102');assert.equal(i.note,'原备注\n另一备注');assert.deepEqual(i.sources.map(s=>s.row),[2,3,4]);assert.equal(JSON.stringify(b),before);
});
test('金额取价税合计，不用显示单价重新乘算；零和退货负数保留',()=>{
 const r=C.parse(book([row({实收数量:8,价税合计:1.36}),row({实收数量:-1,价税合计:-0.17}),row({实收数量:0,价税合计:0})]),'源.xlsx');assert.equal(r.quantity,'7');assert.equal(r.amount,'119');assert.equal(r.count,3);
 assert.equal(C.amountText(C.decimal('(1,234.56)','金额',2)),'-1234.56');assert.equal(C.decimalText(C.decimal('1.00000001')),'1.00000001');
});
test('显示为同一两位单价的原始价格仍分别成行',()=>{
 const r=C.parse(book([row({含税单价:0.1695}),row({含税单价:0.17})]),'源.xlsx');
 assert.equal(r.groups[0].items.length,2);assert.deepEqual(r.groups[0].items.map(i=>i.price),['0.1695','0.17']);assert.equal(r.amount,'68');
});
test('按表头读取不同位置，日期带时间和1904纪元正确，合计行不重复计入',()=>{
 const b=book([row({入库日期:'2026/6/30 12:30:00'}),{供应商:'合计',实收数量:2,价税合计:0.34}]);const a=X.utils.sheet_to_json(b.Sheets.Sheet1,{header:1}).map(r=>r.reverse());b.Sheets.Sheet1=X.utils.aoa_to_sheet(a);assert.equal(C.parse(b,'源.xlsx').groups[0].month,'2026-06');
 const n=book([row({入库日期:44860})]);n.Workbook={WBProps:{date1904:true}};assert.equal(C.parse(n,'源.xlsx').groups[0].month,'2026-10');
});
test('空值、异常、无缓存公式和合计差异明确阻止，不静默丢行或转零',()=>{
 for(const h of ['供应商','采购组织','实收数量','价税合计','物料编码'])assert.throws(()=>C.parse(book([row({[h]:''})]),'源.xlsx'),/为空或异常/);
 for(const v of ['待定','12abc','1.234'])assert.throws(()=>C.parse(book([row({价税合计:v})]),'源.xlsx'),/有效数值|小数/);
 assert.throws(()=>C.parse(book([row({入库日期:'2026/2/30'})]),'源.xlsx'),/日期无法识别/);
 assert.throws(()=>C.parse(book([row(),{供应商:'合计',实收数量:3,价税合计:0.34}]),'源.xlsx'),/实收数量合计不一致/);
 const b=book([row()]);b.Sheets.Sheet1.I2={t:'n',f:'A100'};assert.throws(()=>C.parse(b,'源.xlsx'),/为空或异常|公式/);
});
test('同编码不同名称完整保留并提示，标识不安全数字不猜测',()=>{
 const r=C.parse(book([row(),row({物料名称:'另一个名称'})]),'源.xlsx');assert.equal(r.groups[0].items[0].name,'测试物料\n另一个名称');assert.equal(r.warnings.length,1);
 assert.throws(()=>C.parse(book([row({物料编码:123.4})]),'源.xlsx'),/完整整数/);
});
test('各对账月份默认次月10日，支持跨年及闰年，不依赖今天',()=>{
 for(const [month,expected]of [['2026-06','2026-07-10'],['2026-11','2026-12-10'],['2026-12','2027-01-10'],['2024-02','2024-03-10'],['2026-02','2026-03-10']])assert.equal(C.defaults(month).reconcileDate,expected);
 assert.throws(()=>C.defaults('2026-13'),/日期无法识别/);
});
test('手填底部金额空白不当零，完整填写才计算合计，金额及日期异常报错',()=>{
 const g={amount:'10001'},f=C.defaults('2026-11');let d=C.footer(g,f);assert.equal(d.total,'');assert.equal(d.balance,'');assert.equal(d.invoices[0].amount,'');
 Object.assign(f,{balance:'0',receipts:'20',afterSales:'0',returns:'10.01'});d=C.footer(g,f);assert.equal(d.total,'70.00');assert.equal(d.balance,'0.00');assert.equal(f.balance,'0');
 assert.throws(()=>C.footer(g,{...f,receipts:'金额不明'}),/有效数值/);assert.throws(()=>C.footer(g,{...f,reconcileDate:''}),/选择对账日期/);
});
module.exports={book,row,headers};
