const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js');
const {receiptFileName}=require('../receipt-sign.js');
function sheet(name,total){return X.utils.aoa_to_sheet([['采购入库单'],['供应商：'+name],['物料编码','物料名称','数量','实收数量','金额'],['0001','材料',999,5000,25000],['合计','',999,total,25000]]);}
test('入库单文件名取实收数量合计，不重复计入明细或其他金额',()=>{
 const s=sheet('测试甲公司',5000),source=JSON.stringify(s);
 assert.equal(receiptFileName({Sheet1:s}),'入库单-测试甲公司-5000.pdf');assert.equal(JSON.stringify(s),source);
 s.D5={t:'n',f:'SUM(D4:D4)',v:5000,w:'5,000.00'};assert.equal(receiptFileName({Sheet1:s}),'入库单-测试甲公司-5000.pdf');
});
test('多个文件及工作表合计只累加一次，供应商名称去重，小数精确相加',()=>{
 assert.equal(receiptFileName({a:sheet('甲',5000),b:sheet('甲','1,000.50'),c:sheet('乙',0.1),d:sheet('乙',0.2)}),'入库单-甲、乙-6000.8.pdf');
 for(const [value,expected] of [[0,'0'],['0.000','0'],['100.00','100'],['1.2500','1.25'],[-2,'-2'],['9007199254740993.25','9007199254740993.25']])assert.equal(receiptFileName({a:sheet('甲',value)}),'入库单-甲-'+expected+'.pdf');
});
test('合并表头与合计标签兼容空格，不改变供应商及合并单元格',()=>{
 const s=sheet('甲/乙',12);s.D3.v='实 收 数 量';s.A5.v='合 计：';s['!merges']=[{s:{r:4,c:0},e:{r:4,c:1}}];
 assert.equal(receiptFileName({a:s}),'入库单-甲_乙-12.pdf');assert.equal(s['!merges'].length,1);
});
test('缺失、异常或多个合计不当成零，错误提示包含工作表',()=>{
 for(const value of ['',undefined,'待确认','1,2',{t:'e',v:23},{t:'n',f:'SUM(D4:D4)'}]){
  const s=sheet('甲',value);if(value&&typeof value==='object')s.D5=value;
  assert.throws(()=>receiptFileName({'文件.xlsx / Sheet1':s}),/文件.xlsx \/ Sheet1：实收数量合计为空或不是有效数字/);
 }
 const s=sheet('甲',5);delete s.A5;assert.throws(()=>receiptFileName({a:s}),/无法唯一识别实收数量的合计/);
 s.A5={t:'s',v:'合计'};X.utils.sheet_add_aoa(s,[['总计','','',5]],{origin:'A6'});assert.throws(()=>receiptFileName({a:s}),/无法唯一识别实收数量的合计/);
 const noHeader=sheet('甲',5);delete noHeader.C3;delete noHeader.D3;assert.throws(()=>receiptFileName({a:noHeader}),/无法唯一识别实收数量列/);
});
