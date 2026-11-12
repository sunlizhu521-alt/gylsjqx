const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Core = require('../negotiated-payment-summary-core.js');

const headers = ['供应商','是否谈判分批','2026年11月支付','2027年1月支付','付款主体','OA-申请付款金额'];
function source(rows=[], labels=headers) {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([labels,...rows]),'谈判付款');
  return {book,filename:'付款节奏.xlsx'};
}

test('同供应商跨主体合并，按首次出现排序，重复明细和已付数据完整累计', () => {
  const input = source([
    ['乙公司','是',100,20,'迈德斯特',999],['甲公司','否',30,null,'国源',500],
    ['乙公司','是',100,20,'国源',999],['乙公司','否',50,null,'迈德斯特',111]
  ]);
  input.book.Sheets.谈判付款.G1={t:'s',v:'是否结清'}; input.book.Sheets.谈判付款.G2={t:'s',v:'是'};
  input.book.Sheets.谈判付款.C2.s={fill:{fgColor:{rgb:'FF0000'}}};
  input.book.Sheets.谈判付款.F2={t:'e',f:'1/0',v:7};
  X.utils.book_append_sheet(input.book,X.utils.aoa_to_sheet([headers,['无关','否',999999,999999]]),'正常付款');
  const before=JSON.stringify(input),r=Core.build(input);
  assert.equal(r.canExport,true); assert.equal(r.recordCount,4);
  assert.deepEqual(r.counts,{negotiated:2,normal:2,unresolved:0});
  assert.deepEqual(r.rows.map(x=>x.supplier),['乙公司','甲公司']);
  assert.deepEqual(r.rows[0],{supplier:'乙公司',negotiatedCents:[20000,4000],normalCents:[5000,null],negotiatedTotalCents:24000,normalTotalCents:5000});
  assert.deepEqual(r.totals,{negotiatedCents:24000,normalCents:8000,unresolvedCents:0});
  assert.equal(JSON.stringify(input),before);
});

test('按表头定位，所有月份排序，只有年度月度支付列参与统计', () => {
  const input=source([[5,' 是 ',' 甲公司 ',3,1,999]],['2027年1月支付','是否谈判分批','供应商名称','2026年12月支付','2026年9月支付','2026年11月2日']);
  const r=Core.build(input);
  assert.equal(r.canExport,true);
  assert.deepEqual(r.months,[{key:'2026-09',label:'2026年9月'},{key:'2026-12',label:'2026年12月'},{key:'2027-01',label:'2027年1月'}]);
  assert.deepEqual(r.rows[0].negotiatedCents,[100,300,500]);
  assert.equal(r.rows[0].supplier,'甲公司');
});

test('供应商只去首尾空格，不模糊归并名称，不修改原数据', () => {
  const r=Core.build(source([['甲 公司','是',1],['甲公司','是',2],[' 甲公司 ','否',3]]));
  assert.equal(r.rows.length,2); assert.equal(r.rows[0].supplier,'甲 公司'); assert.equal(r.rows[1].normalTotalCents,300);
});

test('空白、零与空供应商不同；有供应商的全空月份仍保留一行', () => {
  const r=Core.build(source([['甲','是',null,0],['乙','否','',null],['丙','否',0,-2.5]]));
  assert.equal(r.canExport,true); assert.equal(r.recordCount,3);
  assert.deepEqual(r.rows[0].negotiatedCents,[null,0]);
  assert.deepEqual(r.rows[1].normalCents,[null,null]);
  assert.equal(r.rows[1].normalTotalCents,0);
  assert.deepEqual(r.rows[2].normalCents,[0,-250]);
  const missing=Core.build(source([['','是',10]]));
  assert.equal(missing.canExport,false); assert.match(missing.errors.join('\n'),/A2.*供应商/);
});

test('分类缺失或异常需明确选择，改为是或否后重算，已知分类不被覆盖', () => {
  const input=source([['甲','',100,20],['甲','待核对',30,null],['乙','是',1,2]]);
  const r=Core.build(input);
  assert.equal(r.canExport,false); assert.equal(r.errors.length,0);
  assert.deepEqual(r.counts,{negotiated:1,normal:0,unresolved:2});
  assert.deepEqual(r.totals,{negotiatedCents:300,normalCents:0,unresolvedCents:15000});
  assert.deepEqual(r.unresolved.map(x=>[x.id,x.row,x.amountCents]),[['B2',2,12000],['B3',3,3000]]);
  assert.match(r.unresolved[0].source,/付款节奏.xlsx \/ 谈判付款 \/ 第2行/);
  const resolved=Core.build(input,{B2:'是',B3:'否',B4:'否'});
  assert.equal(resolved.canExport,true);
  assert.deepEqual(resolved.totals,{negotiatedCents:12300,normalCents:3000,unresolvedCents:0});
  assert.equal(Core.build(input,{B2:'yes',B3:'否'}).unresolved.length,1);
  assert.equal(Core.build(input).unresolved.length,2);
});

test('金额按分四舍五入，标准金额文本、科学计数、负数合法', () => {
  const r=Core.build(source([['甲','是','￥1,234.56元','-1.005'],['甲','是',0.1,'1e-2'],['甲','是',0.2,1.005]]));
  assert.equal(r.canExport,true); assert.deepEqual(r.rows[0].negotiatedCents,[123486,1]);
  assert.equal(r.warnings.length,2);
  for(const value of [false,Infinity,NaN,'1,00','待付款','/','90071992547409.92']) {
    const bad=Core.build(source([['甲','是',value]]));
    assert.equal(bad.canExport,false,String(value)); assert.match(bad.errors.join('\n'),/C2.*(?:金额|错误)/);
  }
});

test('单格、同月累计、供应商合计及全表合计都保护整数分边界', () => {
  const max='90071992547409.91';
  assert.equal(Core.build(source([['甲','是',max]])).totals.negotiatedCents,Number.MAX_SAFE_INTEGER);
  for(const rows of [
    [['甲','是',max],['甲','是',0.01]],
    [['甲','是',max,0.01]],
    [['甲','是',max],['乙','是',0.01]],
    [['甲','',max,0.01]]
  ]) {
    const r=Core.build(source(rows)); assert.equal(r.canExport,false); assert.match(r.errors.join('\n'),/安全整数分/);
  }
});

test('必要字段错误或公式缺缓存阻止导出，缓存零可用；分类选择不能绕过公式错误', () => {
  for(const address of ['A2','B2','C2']) for(const value of [{t:'n',f:'1+1'},{t:'z',f:'1+1',v:0},{t:'str',f:'A1',v:''},{t:'e',v:23,f:'bad'}]) {
    const input=source([['甲','是',10]]);input.book.Sheets.谈判付款[address]=value;
    const r=Core.build(input,{B2:'是'});
    assert.equal(r.canExport,false); assert.ok(r.errors.some(e=>e.includes(address)&&e.includes('公式')));
  }
  const input=source([['甲','是',10]]);input.book.Sheets.谈判付款.C2={t:'n',f:'1-1',v:0};
  assert.equal(Core.build(input).canExport,true);assert.equal(Core.build(input).rows[0].negotiatedCents[0],0);
});

test('真实OOXML缺公式缓存被识别，显式零缓存正常',async()=>{
  const original=X.write(source([['甲','是',10]]).book,{type:'buffer',bookType:'xlsx'});
  for(const cache of ['', '<v/>', '<v>0</v>']) {
    const zip=await Zip.loadAsync(original),file='xl/worksheets/sheet1.xml';
    const xml=await zip.file(file).async('string');
    zip.file(file,xml.replace(/<c\b[^>]*\br="C2"[^>]*>[\s\S]*?<\/c>/,`<c r="C2"><f>1-1</f>${cache}</c>`));
    const input={book:X.read(await zip.generateAsync({type:'nodebuffer'}),{type:'buffer',sheetStubs:true,cellFormula:true}),filename:'公式.xlsx'};
    const r=Core.build(input);assert.equal(r.canExport,cache==='<v>0</v>');
  }
});

test('必要表头重复、缺失、月份缺年份/越界/重复都不猜测',()=>{
  const missing=Core.build({book:{Sheets:{正常付款:{}}}});assert.equal(missing.canExport,false);
  for(const h of [
    ['供应商','2026年11月支付'],
    ['供应商','是否谈判分批'],
    ['供应商','供应商名称','是否谈判分批','2026年11月支付'],
    ['供应商','是否谈判分批','是否谈判分批','2026年11月支付'],
    ['供应商','是否谈判分批','2026年11月支付','2026年011月支付'],
    ['供应商','是否谈判分批','2026年11月支付','2026年11月支付'],
    ['供应商','是否谈判分批','2026年11月支付','11月支付'],
    ['供应商','是否谈判分批','2026年13月支付']
  ]) {const r=Core.build(source([],h));assert.equal(r.canExport,false);assert.ok(r.errors.length,h.join(','));}
});

test('大格式范围不扫描，不存在的明细和其他列内容不产生供应商记录',()=>{
  const input=source([['甲','是',0]]),sheet=input.book.Sheets.谈判付款;
  sheet['!ref']='A1:XFD1048576';sheet.XFD1048576={t:'z',s:{fill:{fgColor:{rgb:'FFFFFF'}}}};
  sheet.Z1000={t:'s',v:'备注说明'};
  const r=Core.build(input);assert.equal(r.canExport,true);assert.equal(r.recordCount,1);
  assert.equal(Core.build(source()).canExport,false);
});

test('浏览器与CommonJS导出相同构建接口',()=>{
  const context={};vm.runInNewContext(fs.readFileSync(require.resolve('../negotiated-payment-summary-core.js'),'utf8'),context);
  const result=context.NegotiatedPaymentSummaryCore.build(source([['甲','否',12.34]]));
  assert.equal(result.canExport,true);assert.equal(result.totals.normalCents,1234);
});

const sample='/Users/qingtianzhu/Desktop/付款节奏跟进-2026-11-12_V3.xlsx';
test('真实样表119条47家，月度支付而非计划日期金额，待分类18万不自动归类',{skip:!fs.existsSync(sample)},()=>{
  const input={book:X.read(fs.readFileSync(sample),{type:'buffer',cellFormula:true,sheetStubs:true}),filename:'真实样表.xlsx'};
  const r=Core.build(input);
  assert.equal(r.errors.length,0,r.errors.join('\n'));assert.equal(r.recordCount,119);assert.equal(r.rows.length,47);
  assert.deepEqual(r.counts,{negotiated:56,normal:62,unresolved:1});
  assert.deepEqual(r.totals,{negotiatedCents:3462853765,normalCents:1063699503,unresolvedCents:18000000});
  assert.equal(r.canExport,false);assert.equal(r.unresolved[0].id,'G40');
  assert.deepEqual(r.months.map(m=>m.key),['2026-09','2026-10','2026-11','2026-12','2027-01','2027-02','2027-03','2027-04']);
  const selected=Core.build(input,{G40:'否'});assert.equal(selected.canExport,true);assert.equal(selected.totals.normalCents,1081699503);
});
