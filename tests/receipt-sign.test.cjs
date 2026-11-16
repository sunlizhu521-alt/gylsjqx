const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js');
const {receiptFileName,splitReceipts}=require('../receipt-sign.js');
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

function stackedReceipts(specs){
 const source={'!ref':'A1:D1','!cols':[{wpx:120},{wpx:210},{wpx:95},{wpx:130}],'!rows':[],'!merges':[]},blocks=[];let row=0;
 for(const [i,spec] of specs.entries()){
  row+=spec.gap??1;
  const start=row,items=spec.items||[spec.total],rows=[
   [spec.title||'采购入库单'],['单据编号：',spec.number??'CGRK'+(i+1)],['供应商：',spec.supplier||'测试甲公司'],
   ['物料编码','物料名称','实收数量','金额'],
   ...items.map((n,j)=>['00'+(j+1),'材料（采购入库单明细）',n,n*100]),
   ['合计','',spec.total,spec.total*100],['创建人','原创建人','审核人','原审核人'],['创建日期','2026/11/16','审核日期','2026/11/16']
  ];
  X.utils.sheet_add_aoa(source,rows,{origin:{r:start,c:0}});
  const end=start+rows.length-1,header=start+3,total=start+4+items.length;
  source['!merges'].push({s:{r:start,c:0},e:{r:start,c:3}});
  for(let r=start;r<=end;r++)source['!rows'][r]={hpt:20+r%3};
  blocks.push({start,end,header,total,number:spec.number??'CGRK'+(i+1)});row=end+1;
 }
 return {source,blocks};
}
const receiptsByName=parts=>Object.fromEntries(parts.map((p,i)=>['第'+(i+1)+'单',p.sheet]));
test('同sheet两张入库单独立读取编号、供应商、明细和合计，源地址及内容不变',()=>{
 const {source,blocks}=stackedReceipts([{number:'CGRK42761',supplier:'浙江金匠健康科技有限公司',total:6},{number:'CGRK42753',supplier:'浙江金匠健康科技有限公司',total:6,items:[3,3]}]);
 const before=JSON.stringify(source),parts=splitReceipts(source,'套打.XLS / Sheet');
 assert.deepEqual(parts.map(p=>p.number),['CGRK42761','CGRK42753']);
 for(const [i,part] of parts.entries()){
  const {start,end}=blocks[i];assert.equal(part.sheet['!ref'],X.utils.encode_range({s:{r:start,c:0},e:{r:end,c:3}}));
  assert.deepEqual(part.sheet['!cols'],source['!cols']);assert.deepEqual(part.sheet['!rows'],source['!rows']);
  assert.deepEqual(part.sheet['!merges'],source['!merges'].filter(m=>m.s.r>=start&&m.e.r<=end));
  for(let r=start;r<=end;r++)for(let c=0;c<4;c++){const a=X.utils.encode_cell({r,c});assert.deepEqual(part.sheet[a],source[a],a+'原值与公式保留');}
  assert.equal(receiptFileName({单据:part.sheet}),'入库单-浙江金匠健康科技有限公司-6.pdf');
 }
 assert.equal(receiptFileName(receiptsByName(parts)),'入库单-浙江金匠健康科技有限公司-12.pdf');assert.equal(JSON.stringify(source),before);
});
test('多于两张按原顺序拆分，重复编号不去重，空行和不同明细长度不串单',()=>{
 const {source}=stackedReceipts([
  {number:'重复号',supplier:'甲',total:.1,gap:0},{number:'重复号',supplier:'乙',total:.2,gap:4,title:'采 购 入 库 单'},
  {supplier:'甲',total:0,gap:2,items:[0,0,0]},{supplier:'丙',total:2.05,gap:0,items:[1,1.05]},{supplier:'乙',total:3,gap:5}
 ]);
 const parts=splitReceipts(source);assert.equal(parts.length,5);assert.deepEqual(parts.map(p=>p.number),['重复号','重复号','CGRK3','CGRK4','CGRK5']);
 assert.deepEqual(parts.map(p=>receiptFileName({单据:p.sheet})),['入库单-甲-0.1.pdf','入库单-乙-0.2.pdf','入库单-甲-0.pdf','入库单-丙-2.05.pdf','入库单-乙-3.pdf']);
 assert.equal(receiptFileName(receiptsByName(parts)),'入库单-甲、乙、丙-5.35.pdf');
});
test('零或一个标题兼容原单单据，缺编号保留空，名称中含入库单不误拆',()=>{
 for(const hasTitle of [true,false]){
  const s=sheet('甲',12);if(!hasTitle)delete s.A1;s.B4.v='采购入库单专用材料';const before=JSON.stringify(s),parts=splitReceipts(s,'旧版.XLS / Sheet1');
  assert.equal(parts.length,1);assert.equal(parts[0].number,'');assert.equal(receiptFileName(receiptsByName(parts)),'入库单-甲-12.pdf');assert.equal(JSON.stringify(s),before);
 }
});
test('块首尾空白行可去除，有内容的纵向合并和原合计公式保留完整',()=>{
 const {source,blocks}=stackedReceipts([{total:6,gap:3},{total:7,gap:4}]),first=blocks[0],last=blocks[1];
 source['!merges'].push({s:{r:first.end,c:0},e:{r:first.end+2,c:0}});
 const totalAddress=X.utils.encode_cell({r:first.total,c:2});source[totalAddress]={t:'n',v:6,f:'SUM(C'+(first.header+2)+':C'+first.total+')'};
 source['!ref']='A1:D'+(last.end+5);const before=JSON.stringify(source),parts=splitReceipts(source);
 assert.equal(X.utils.decode_range(parts[0].sheet['!ref']).e.r,first.end+2);
 assert.equal(X.utils.decode_range(parts[1].sheet['!ref']).e.r,last.end);
 assert(parts[0].sheet['!merges'].some(m=>m.s.r===first.end&&m.e.r===first.end+2));assert.deepEqual(parts[0].sheet[totalAddress],source[totalAddress]);
 assert.equal(receiptFileName(receiptsByName(parts)),'入库单-测试甲公司-13.pdf');assert.equal(JSON.stringify(source),before);
});
test('空工作表、同一行多个标题和跨单合并明确失败并带来源位置',()=>{
 const origin='异常.XLS / Sheet';
 for(const s of [{},{'!ref':'A1:D9'}])assert.throws(()=>splitReceipts(s,origin),e=>e.message.includes(origin)&&/没有内容|空/.test(e.message));
 const duplicate=stackedReceipts([{total:6},{total:7}]);duplicate.source.D2={t:'s',v:'采购入库单'};
 assert.throws(()=>splitReceipts(duplicate.source,origin),e=>e.message.includes(origin)&&/A2|D2|第\s*2\s*行/.test(e.message));
 const crossing=stackedReceipts([{total:6},{total:7}]),next=crossing.blocks[1].start;
 crossing.source['!merges'].push({s:{r:next-1,c:0},e:{r:next,c:0}});
 assert.throws(()=>splitReceipts(crossing.source,origin),e=>e.message.includes(origin)&&e.message.includes('A'+next+':A'+(next+1)));
});
test('每张单据独立校验唯一数量表头和合计，不借用其他单据合计',()=>{
 for(const mode of ['missingHeader','missingTotal','duplicateHeader','duplicateTotal']){
  const {source,blocks}=stackedReceipts([{total:6},{total:7}]),last=blocks[1],header=X.utils.encode_cell({r:last.header,c:2}),total=X.utils.encode_cell({r:last.total,c:0});
  if(mode==='missingHeader')delete source[header];if(mode==='missingTotal')delete source[total];
  if(mode==='duplicateHeader')source[X.utils.encode_cell({r:last.header,c:3})]={t:'s',v:'实收数量'};
  if(mode==='duplicateTotal')source[X.utils.encode_cell({r:last.end,c:0})]={t:'s',v:'合计'};
  assert.throws(()=>splitReceipts(source,'坏单.XLS / Sheet'),e=>e.message.includes('坏单.XLS / Sheet')&&/无法唯一识别/.test(e.message)&&e.message.includes('第'+(last.start+1)+'–'+(last.end+1)+'行'),mode);
 }
});
test('后续单据异常数量、公式结果缺失或Excel错误保留原地址并阻止漏单',()=>{
 for(const bad of [{t:'s',v:''},{t:'s',v:'待核实'},{t:'n',f:'SUM(C99:C100)'},{t:'e',v:23,w:'#REF!'}]){
  const {source,blocks}=stackedReceipts([{total:6},{total:7},{total:8}]),address=X.utils.encode_cell({r:blocks[1].total,c:2});source[address]=bad;
  assert.throws(()=>splitReceipts(source,'公式.XLS / Sheet'),e=>e.message.includes('公式.XLS / Sheet')&&e.message.includes(address));
 }
 const {source,blocks}=stackedReceipts([{total:6},{total:7}]),address=X.utils.encode_cell({r:blocks[1].header+1,c:3});source[address]={t:'n',f:'C99*100'};
 assert.throws(()=>splitReceipts(source,'公式.XLS / Sheet'),e=>e.message.includes('公式.XLS / Sheet')&&e.message.includes(address));
});
