const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../table-compare-core.js');
const HEADERS=['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','备注'];
const row=(overrides={})=>({'分类':'正常付款-谈判付款表','创跟单员':'李奇','供应商':'供应商甲','付款主体':'主体甲','对公付款0A流程号':'MDST-DGFK20261101001','摘要':'货款','OA-申请付款金额':100,'备注':'原计划',...overrides});
function source(records=[],filename='旧版.xlsx',headers=HEADERS){const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(header=>record[header]??''))]),'Sheet1');return {book,filename};}
const compare=(a,b)=>C.build(source(a),source(b,'新版.xlsx'));
const values=result=>result.sheets[0].rows.map(row=>row.map(cell=>cell.v));
function set(src,address,cell){src.book.Sheets.Sheet1[address]=cell;return src;}

test('单表固定15列，按新顺序保留全部记录，原备注和空白原因分开',()=>{
 const old=[row(),row({'供应商':'移除供应商','对公付款0A流程号':'OA002','OA-申请付款金额':50})];
 const newer=[row({'供应商':'新供应商','对公付款0A流程号':'OA003','OA-申请付款金额':30}),row()];
 const result=compare(old,newer),rows=values(result);
 assert.equal(result.sheets.length,1);assert.equal(result.sheets[0].name,'比对汇总');assert.equal(C.HEADERS.length,15);
 assert.deepEqual(result.sheets[0].headers,C.HEADERS);assert.deepEqual(rows.map(r=>[r[2],r[9],r[8]]),[['新供应商','增加',30],['供应商甲','不变',0],['移除供应商','减少',-50]]);
 assert.deepEqual(result.summary,{oldAmount:150,newAmount:130,difference:-20,complete:true,counts:{增加:1,减少:1,不变:1,待核对:0}});
 assert(rows.every(row=>row[10]==='原计划'&&row[14]===''));assert.equal(rows[0][6],'');assert.equal(rows[2][7],'');
 assert.match(rows[1][12],/旧版.xlsx \/ Sheet1 第2行/);assert.match(rows[1][13],/新版.xlsx \/ Sheet1 第3行/);
});
test('有效OA号、供应商和主体定位同笔金额变化，金额按分计算',()=>{
 for(const [oldAmount,newAmount,status,delta]of [[100,120,'增加',20],[120,100,'减少',-20],[0,0,'不变',0],[-100,-20,'增加',80],[-20,-100,'减少',-80],['￥1,200.10','1200.11元','增加',0.01]]){
  const result=compare([row({'OA-申请付款金额':oldAmount})],[row({'OA-申请付款金额':newAmount})]);assert.equal(values(result)[0][9],status);assert.equal(values(result)[0][8],delta);assert.equal(result.summary.difference,delta);
 }
});
test('相同OA多笔先匹配同额，剩余唯一一对才比较金额变化',()=>{
 const result=compare([row({'OA-申请付款金额':10}),row({'OA-申请付款金额':20})],[row({'OA-申请付款金额':30}),row({'OA-申请付款金额':10})]);
 assert.deepEqual(values(result).map(r=>[r[6],r[7],r[9]]),[[20,30,'增加'],[10,10,'不变']]);
});
test('相同OA剩余多候选保留双方各行待核对，不分摊或汇总抵消',()=>{
 const result=compare([row({'OA-申请付款金额':10}),row({'OA-申请付款金额':20})],[row({'OA-申请付款金额':15}),row({'OA-申请付款金额':15})]);
 assert.equal(values(result).length,4);assert(values(result).every(r=>r[9]==='待核对'&&r[8]===''));assert.equal(result.summary.counts.待核对,4);
 assert.equal(result.summary.difference,0);assert.equal(result.summary.complete,true);assert.match(result.warnings.join('\n'),/旧版.xlsx.*第2行/);assert.match(result.warnings.join('\n'),/新版.xlsx.*第3行/);
});
test('重复记录逐笔消费不去重，不复用旧行',()=>{
 const result=compare([row(),row()],[row(),row(),row()]);assert.deepEqual(values(result).map(r=>r[9]),['不变','不变','增加']);
 assert.deepEqual(values(result).map(r=>r[12]),['旧版.xlsx / Sheet1 第2行','旧版.xlsx / Sheet1 第3行','']);
 assert.equal(result.summary.oldAmount,200);assert.equal(result.summary.newAmount,300);
});
test('缺OA号补齐按名称主体金额唯一定位，摘要变化仅说明，不误判增减',()=>{
 const result=compare([row({'对公付款0A流程号':'未提交OA','OA-申请付款金额':3400,'摘要':'原摘要'})],[row({'OA-申请付款金额':3400,'摘要':'新摘要'})]),record=values(result)[0];
 assert.equal(values(result).length,1);assert.equal(record[9],'不变');assert.equal(record[4],'MDST-DGFK20261101001');assert.equal(record[5],'新摘要');assert.match(record[11],/按供应商、付款主体和金额唯一匹配/);assert.match(record[11],/未提交OA/);assert.match(record[11],/原摘要.*新摘要/);
 for(const empty of ['','未申请OA','无','/'])assert.equal(values(compare([row({'对公付款0A流程号':empty})],[row()]))[0][9],'不变');
});
test('两边都缺号时相同金额重复不能任取配对，双方明细仍完整',()=>{
 const empty=row({'对公付款0A流程号':''}),result=compare([empty,empty],[empty,empty]);assert.equal(values(result).length,4);assert(values(result).every(r=>r[9]==='待核对'));
});
test('大组同额缺号重复保留两边100条，不生成重复匹配组合',()=>{
 const records=Array.from({length:100},()=>row({'对公付款0A流程号':''})),result=compare(records,records);
 assert.equal(values(result).length,200);assert.equal(result.summary.counts.待核对,200);assert.equal(result.summary.oldAmount,10000);assert.equal(result.summary.newAmount,10000);assert.equal(result.summary.difference,0);
 assert(values(result).every(r=>r[9]==='待核对'&&r[8]===''));assert.equal(new Set(values(result).map(r=>r[12]||r[13])).size,200);
});
test('不同有效OA号不跨号，供应商和主体不模糊合并',()=>{
 for(const overrides of [{'对公付款0A流程号':'OA002'},{'供应商':'供应商甲有限公司'},{'付款主体':'主体乙'}]){
  const result=compare([row()],[row(overrides)]);assert.deepEqual(values(result).map(r=>r[9]),['增加','减少']);
 }
});
test('供应商与主体外围空白不改源值，编号大小写全角可标准化',()=>{
 const result=compare([row({'供应商':' 供应商甲 ','付款主体':' 主体甲 ','对公付款0A流程号':'ｏａ００１'})],[row({'对公付款0A流程号':'OA001'})]);assert.equal(values(result).length,1);assert.equal(values(result)[0][9],'不变');
});
test('缺号唯一规则不会将多个旧OA候选映射到同一新行',()=>{
 const result=compare([row({'对公付款0A流程号':'OA001'}),row({'对公付款0A流程号':'OA002'})],[row({'对公付款0A流程号':'未提交OA'})]);assert.equal(values(result).length,3);assert(values(result).every(r=>r[9]==='待核对'));
});
test('空白和异常金额保留原文，合计标不完整，正常候选不误判新增',()=>{
 for(const amount of ['',null,'待确认','3,40','￥abc',true]){
  const result=compare([row({'OA-申请付款金额':amount})],[row()]),rows=values(result);assert.equal(rows.length,2);assert(rows.every(r=>r[9]==='待核对'&&r[8]===''));assert.equal(result.summary.complete,false);assert.equal(result.summary.oldAmount,0);assert.equal(result.summary.newAmount,100);
 }
 const result=compare([row({'OA-申请付款金额':''})],[row({'OA-申请付款金额':''})]);assert.equal(values(result).length,2);assert.equal(result.summary.complete,false);
 for(const amount of [Infinity,NaN]){const a=set(source([row()]),'G2',{t:'n',v:amount});assert.equal(C.build(a,source([row()])).summary.complete,false);}
});
test('缺少供应商或主体、多个或异常OA保留待核对，不以空字段匹配',()=>{
 for(const overrides of [{'供应商':''},{'付款主体':''},{'对公付款0A流程号':'OA001、OA002'},{'对公付款0A流程号':'OA001 OA002'},{'对公付款0A流程号':'待定'}]){
  const result=compare([row(overrides)],[row(overrides)]);assert.equal(values(result).length,2);assert(values(result).every(r=>r[9]==='待核对'&&r[8]===''));
 }
});
test('零和负数新增移除按存在方向标记，差额正确',()=>{
 for(const amount of [0,-10]){
  let result=compare([],[row({'OA-申请付款金额':amount})]);assert.equal(values(result)[0][9],'增加');assert.equal(values(result)[0][8],amount);
  result=compare([row({'OA-申请付款金额':amount})],[]);assert.equal(values(result)[0][9],'减少');assert.equal(values(result)[0][8],-amount);
 }
});
test('小数按分四舍五入，二进制误差和科学记数法不造成假差异',()=>{
 for(const [a,b,expected]of [[0.1+0.2,'0.30',0.3],[1.005,'1.01',1.01],[-1.005,'-1.01',-1.01],[1e-7,'0',0],['(1,234.56)','-1234.56',-1234.56]]){
  const result=compare([row({'OA-申请付款金额':a})],[row({'OA-申请付款金额':b})]);assert.equal(values(result)[0][9],'不变');assert.equal(values(result)[0][6],expected);
 }
 const result=compare([row({'OA-申请付款金额':'999999999999999999999.00'})],[]);assert.equal(result.summary.complete,false);assert.equal(values(result)[0][9],'待核对');
});
test('原备注有变化时双方都保留，新空白备注不填程序说明',()=>{
 const result=compare([row({'备注':'旧原因','分类':'谈判付款','创跟单员':'徐英'})],[row({'备注':'新原因'})]),record=values(result)[0];
 assert.equal(record[10],'旧版：旧原因\n新版：新原因');assert.equal(record[14],'');assert.match(record[11],/分类.*谈判付款/);assert.match(record[11],/采购跟单员.*徐英/);
});
test('只读取业务表头列，辅助公式错误和格式空区忽略且源工作簿不变',()=>{
 const a=source([row()]),b=source([row()],'新版.xlsx');
 a.book.Sheets.Sheet1.I1={t:'s',v:'上周五'};a.book.Sheets.Sheet1.I2={t:'e',v:23,f:'BAD()'};a.book.Sheets.Sheet1.J2={t:'s',v:'辅助'};
 a.book.Sheets.Sheet1.XFD1048576={t:'z',s:{fill:{patternType:'solid'}}};a.book.Sheets.Sheet1['!ref']='A1:XFD1048576';
 const before=JSON.stringify([a,b]),result=C.build(a,b);assert.equal(values(result).length,1);assert.equal(values(result)[0][9],'不变');assert.equal(JSON.stringify([a,b]),before);
});
test('无业务身份的明确合计小计总计行忽略，不把明细和原汇总重复计金额',()=>{
 const footer=(field,label)=>Object.fromEntries(HEADERS.map(header=>[header,header===field?label:header==='OA-申请付款金额'?100:'']));
 for(const [field,label]of [['供应商','合计'],['摘要','小计'],['创跟单员','总计'],['分类','合计']]){
  const result=compare([row(),footer(field,label)],[row(),footer(field,label)]);assert.equal(values(result).length,1);assert.equal(result.summary.oldAmount,100);assert.equal(result.summary.newAmount,100);assert.equal(result.summary.difference,0);assert.equal(result.warnings.length,2);assert.match(result.warnings.join('\n'),/第3行.*忽略.*明确合计/);
 }
 const cases=[row({'摘要':'合计'}),row({'供应商':'合计','付款主体':'实际主体','对公付款0A流程号':''}),footer('摘要','本月合计')];
 for(const record of cases){const result=compare([record],[]);assert.equal(values(result).length,1);assert.equal(result.summary.oldAmount,100);assert(!result.warnings.some(w=>w.includes('忽略')));}
 const cachedFooter=source([row(),footer('供应商','合计')]);set(cachedFooter,'G3',{t:'n',f:'SUM(G2:G2)'});assert.equal(C.build(cachedFooter,source([row()])).summary.oldAmount,100);
});
test('缺有效OA号的异常记录保护同身份候选，不能把无法核实的款项报成新增',()=>{
 for(const overrides of [{'对公付款0A流程号':'','OA-申请付款金额':''},{'对公付款0A流程号':'未提交OA','OA-申请付款金额':'待确认'},{'对公付款0A流程号':'未提交OAD','OA-申请付款金额':100}]){
  for(const reverse of [false,true]){
   const a=row(overrides),b=row(),result=compare(reverse?[b]:[a],reverse?[a]:[b]);assert.equal(values(result).length,2);assert(values(result).every(r=>r[9]==='待核对'&&r[8]===''));
  }
 }
 const unequal=compare([row({'对公付款0A流程号':'未提交OAD'})],[row({'OA-申请付款金额':120})]);assert.deepEqual(values(unequal).map(r=>r[9]),['增加','待核对']);
 const missingIdentity=compare([row({'供应商':'','对公付款0A流程号':'','OA-申请付款金额':''})],[row()]);assert.deepEqual(values(missingIdentity).map(r=>r[9]),['增加','待核对']);
 const distinctValid=compare([row({'对公付款0A流程号':'OA001','OA-申请付款金额':''})],[row({'对公付款0A流程号':'OA002'})]);assert.deepEqual(values(distinctValid).map(r=>r[9]),['增加','待核对']);
 const knownFirst=compare([row(),row({'对公付款0A流程号':'','OA-申请付款金额':''})],[row()]);assert.deepEqual(values(knownFirst).map(r=>r[9]),['不变','待核对']);
});
test('业务公式缓存有效时取值，错误或无缓存带来源单元格阻断',()=>{
 const a=source([row()]),b=source([row()],'新版.xlsx');set(a,'G2',{t:'n',v:100,f:'50+50'});assert.equal(values(C.build(a,b))[0][9],'不变');
 for(const cell of [{t:'e',v:23,w:'#REF!',f:'BAD()'},{t:'n',f:'50+50'},{t:'z',f:'50+50'},{t:'str',f:'""'},{t:'n',v:NaN,f:'BAD()'},{t:'n',v:Infinity,f:'BAD()'}]){
  set(a,'G2',cell);assert.throws(()=>C.build(a,b),/旧版.xlsx \/ Sheet1!G2.*(?:Excel 错误|公式缺少有效计算结果)/);
 }
 set(a,'G2',{t:'str',v:'',f:'""'});assert.equal(C.build(a,b).summary.complete,false);
 set(a,'C2',{t:'e',v:23});assert.throws(()=>C.build(a,b),/Sheet1!C2/);
});
test('重新排序的表头和别名按名称识别，缺失可选字段留空提示',()=>{
 const headers=['本次申请付款金额','供应商名称','对公付款OA流程号','付款主体','跟单员'],record={'本次申请付款金额':100,'供应商名称':'供应商甲','对公付款OA流程号':'MDST-DGFK20261101001','付款主体':'主体甲','跟单员':'李奇'};
 const a=source([record],'旧版.xlsx',headers),b=source([row()],'新版.xlsx'),result=C.build(a,b);assert.equal(values(result)[0][9],'不变');assert.match(result.warnings[0],/缺少 分类、摘要、原备注/);
});
test('多个有效工作表必须选择，零候选、重复表头区和重复列明确报错',()=>{
 const a=source([row()]);X.utils.book_append_sheet(a.book,X.utils.aoa_to_sheet([HEADERS,HEADERS.map(h=>row()[h])]),'第二表');
 assert.equal(C.discover(a.book).length,2);assert.throws(()=>C.build(a,source([row()])),/请选择一个有效工作表/);
 assert.equal(values(C.build({...a,sheetName:'第二表'},source([row()]))).length,1);assert.throws(()=>C.build({...a,sheetName:'不存在'},source([row()])),/请选择一个有效工作表/);
 const empty={book:{SheetNames:['无关'],Sheets:{无关:X.utils.aoa_to_sheet([['供应商']])}},filename:'无关.xlsx'};assert.deepEqual(C.discover(empty.book),[]);assert.throws(()=>C.build(empty,a),/找不到付款台账表头/);
 const duplicate=source([row()]);X.utils.sheet_add_aoa(duplicate.book.Sheets.Sheet1,[HEADERS],{origin:'A4'});assert.throws(()=>C.discover(duplicate.book),/多个业务表头.*第1、4行/);
 const dupColumn=source([row()],undefined,[...HEADERS,'供应商名称']);assert.throws(()=>C.discover(dupColumn.book),/表头重复 供应商/);
});
test('标题行和空明细可处理，工作表名称与合计不会伪造记录',()=>{
 const a=source([]);a.book.Sheets.Sheet1=X.utils.aoa_to_sheet([['标题'],[],HEADERS]);
 const result=C.build(a,source([]));assert.equal(C.discover(a.book)[0].headerRow,2);assert.deepEqual(values(result),[]);assert.deepEqual(result.summary.counts,{增加:0,减少:0,不变:0,待核对:0});
});
test('编号文本及数字格式前导零保留，日期业务字段仅输出文本',()=>{
 for(const date1904 of [false,true]){
  const a=source([row({'对公付款0A流程号':'000123456789012345'})]),b=source([row({'对公付款0A流程号':'000123456789012345'})],'新版.xlsx');
  b.book.Workbook={WBProps:{date1904}};set(b,'F2',{t:'n',v:date1904?44870:46332,z:'yyyy-mm-dd'});
  const result=C.build(a,b);assert.equal(values(result)[0][4],'000123456789012345');assert.equal(values(result)[0][5],'2026-11-06');assert(result.sheets[0].rows[0].slice(0,6).every(cell=>cell.t==='s'&&cell.z==='@'));
 }
 const a=source([row()]),b=source([row()],'新版.xlsx');for(const s of [a,b])set(s,'E2',{t:'n',v:123,z:'000000'});assert.equal(values(C.build(a,b))[0][4],'000123');
});
test('单元格输出为值，等号开头字符串不会变成公式，文件名独立版本',()=>{
 const a=source([row({'摘要':'=HYPERLINK("bad")','备注':'=1+1'})]),b=source([row({'摘要':'=HYPERLINK("bad")','备注':'=1+1'})],'新版.xlsx');
 const result=C.build(a,b);assert.equal(result.sheets[0].rows[0][5].t,'s');assert.equal(result.sheets[0].rows[0][5].f,undefined);assert.equal(C.exportFilename('2026-11-26',3),'付款台账表格比对-2026-11-26-V3.xlsx');
});
