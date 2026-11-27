const {test}=require('node:test'),assert=require('node:assert/strict');
const C=require('../payment-core.js'),X=require('../vendor/xlsx.full.min.js');
const HEADERS=['供应商','付款主体','对公付款0A流程号','OA-申请付款金额'];
const entry=(supplier='供应商甲',subject='甲公司',oa='OA001',amount=100)=>[supplier,subject,oa,amount];
function book(rows=[entry()],headers=HEADERS,name='12月预算'){
 const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows]),name);return b;
}
const parse=(b,selected='')=>C.parseBudget(b,'预算.xlsx',X,selected);
const payment=(extra={})=>({'供应商全称':'供应商甲','付款主体':'甲公司','合同编号':'OA001','本次申请付款金额':100,...extra});
const match=(rows,extra={})=>C.matchBudget(payment(extra),parse(book(rows)));

test('预算按表头解析并保留行号、分精度、重复行和源表不变',()=>{
 const b=book([entry(),entry(),entry('供应商乙','乙公司','0000123','￥1,234.56')]),original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.sheetNames,['12月预算']);assert.equal(result.sheetName,'12月预算');assert.deepEqual(result.errors,[]);assert.equal(result.rows.length,3);
 assert.deepEqual(result.rows[2],{supplier:'供应商乙',subject:'乙公司',oa:'0000123',amount:123456n,amountText:'1,234.56',source:{filename:'预算.xlsx',sheet:'12月预算',row:4}});
 assert.equal(JSON.stringify(b),original);
});
test('别名、表头位置和列位置变化仍正确，格式空列不扩展数据范围',()=>{
 const b=book([],['预算说明']),s=X.utils.aoa_to_sheet([['预算说明'],[],['申请付款金额',' 流程编码 ','付款主体','供应商名称'],[123,'OA01','甲公司','供应商甲'],[],[0,'OA02','甲公司','供应商乙']]);
 s['XFD1048576']={t:'z',s:{fill:{fgColor:{rgb:'FFFF00'}}}};s['!ref']='A1:XFD1048576';b.Sheets['12月预算']=s;
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(row=>[row.supplier,row.oa,row.amount,row.source.row]),[['供应商甲','OA01',12300n,4],['供应商乙','OA02',0n,6]]);
 for(const alias of ['对公付款OA流程号','对公付款OA号','付款OA号','流程编号','合同编号','合同号'])assert.deepEqual(parse(book([entry()],['供应商全称','付款主体',alias,'本次申请付款金额'])).errors,[]);
});
test('只一个预算候选自动选择，多个候选不任取，明确选择后只读该表',()=>{
 const b=book();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([HEADERS,entry('供应商乙','乙公司','OA002',200)]),'明年预算');X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['说明'],['无预算字段']]),'说明');
 const multiple=parse(b);assert.deepEqual(multiple.sheetNames,['12月预算','明年预算']);assert.equal(multiple.sheetName,'');assert.deepEqual(multiple.rows,[]);assert.deepEqual(multiple.errors,['请选择预算工作表']);
 const selected=parse(b,'明年预算');assert.deepEqual(selected.errors,[]);assert.equal(selected.rows.length,1);assert.equal(selected.rows[0].supplier,'供应商乙');assert.equal(selected.rows[0].source.sheet,'明年预算');
 assert.match(parse(b,'说明').errors.join(' '),/预算.xlsx \/ 说明.*缺少预算表头/);assert.equal(parse(b,'说明').rows.length,0);
 assert.match(parse(b,'不存在').errors[0],/工作表不存在/);
});
test('缺少或重复映射表头明确报错，不静默取一列',()=>{
 const missing=parse(book([entry()],HEADERS.slice(0,3)));assert.equal(missing.rows.length,0);assert.match(missing.errors[0],/必须包含.*OA-申请付款金额/);
 const b=book([[...entry(),200]],[...HEADERS,'申请付款金额']),duplicate=parse(b);assert.equal(duplicate.rows.length,0);assert.match(duplicate.errors[0],/第 1 行.*多个列/);
 assert.match(parse(X.utils.book_new()).errors[0],/未找到预算表头/);assert.match(parse(book([])).errors[0],/没有预算明细/);
});
test('必要列公式错误或缓存缺失阻断预算，空字符串有效缓存不会冒充缺失',()=>{
 for(const address of ['A2','B2','C2','D2'])for(const cell of [{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'},{t:'s',f:'A99'},{t:'z',f:'A99',v:''}]){
  const b=book();b.Sheets['12月预算'][address]=cell;const result=parse(b);assert.equal(result.rows.length,0,address);assert(result.errors.some(e=>e.includes('预算.xlsx / 12月预算!'+address)&&/公式错误或缺少有效计算结果/.test(e)));
 }
 const b=book();b.Sheets['12月预算'].D2={t:'n',f:'1-1',v:0};b.Sheets['12月预算'].C2={t:'s',f:'IF(TRUE,"","OA001")',v:''};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows[0].amount,0n);assert.equal(result.rows[0].oa,'');
});
test('不按零替代空白、非法或超精度金额；异常所选表不返回可用部分',()=>{
 for(const amount of ['',null,'/',-1,'1.001','1,23',true,Infinity,'9999999999999']){
  const result=parse(book([entry(),entry('供应商乙','乙公司','OA002',amount)]));assert.equal(result.rows.length,0,String(amount));assert(result.errors.some(e=>e.includes('!D3')));
 }
});
test('文本前导零和数字显示前导零保留，长数字编号精度异常必须修复',()=>{
 const b=book([entry('供应商甲','甲公司','00000123'),entry('供应商乙','乙公司',123)]);b.Sheets['12月预算'].C3.z='00000000';const original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>r.oa),['00000123','00000123']);assert.equal(JSON.stringify(b),original);
 for(const oa of [1000000000000000,Number.MAX_SAFE_INTEGER+1,12.3]){const result=parse(book([entry('供应商甲','甲公司',oa)]));assert.equal(result.rows.length,0);assert.match(result.errors[0],/!C2.*精度/);}
 assert.equal(match([entry('供应商甲','甲公司','00000123')],{合同编号:'00000123'}).status,'matched');assert.equal(match([entry('供应商甲','甲公司','00000123')],{合同编号:'123'}).status,'review');
});
test('有效OA优先且精确按分匹配；标准化仅影响OA与已存在主体别名，不改原记录',()=>{
 const budget=parse(book([entry('供应商甲','浙江迈德斯特医疗器械科技有限公司','ＭＤＳＴ－ＤＧＦＫ２０２６１２０１００１','0.10')])),record=payment({付款主体:'迈德斯特',合同编号:' mdst - dgfk20261201001 ',本次申请付款金额:.1}),source=JSON.stringify(record);
 const result=C.matchBudget(record,budget);assert.equal(result.status,'matched');assert.equal(result.method,'oa');assert.equal(result.candidates[0],budget.rows[0]);assert.equal(JSON.stringify(record),source);
 assert.equal(match([entry(' 供应商甲 ',' 甲公司 ','OA001',100)]).status,'matched');assert.equal(match([entry('供应商 甲')]).status,'review');
});
test('OA相同但供应商或付款主体不符时核对，禁止按其他名称匹配回退',()=>{
 for(const [supplier,subject]of [['另一供应商','甲公司'],['供应商甲','另一公司']]){
  const result=match([entry(supplier,subject),entry('供应商甲','甲公司','',100)]);assert.equal(result.status,'review');assert.equal(result.method,'oa');assert.equal(result.candidates.length,1);assert.match(result.reason,/OA号相同.*不一致/);
 }
});
test('双方有效OA冲突不回退，即使供应商主体金额完全相同',()=>{
 const result=match([entry('供应商甲','甲公司','OA002')]);assert.equal(result.status,'review');assert.equal(result.candidates.length,1);assert.match(result.reason,/OA001.*OA002.*不能按名称回退/);
});
test('任一侧缺号时仅供应商与主体唯一且金额相同才通过',()=>{
 for(const missing of ['', ' ', '/', '未提交OA','未申请OA','未交OA']){
  for(const [budgetOA,detailOA]of [[missing,'OA001'],['OA001',missing],[missing,missing]]){
   const result=match([entry('供应商甲','甲公司',budgetOA)],{合同编号:detailOA});assert.equal(result.status,'matched',JSON.stringify([budgetOA,detailOA]));assert.equal(result.method,'supplier-subject');
  }
 }
});
test('同OA多个预算候选必须核对，不按名称或金额消歧，不合并金额',()=>{
 for(const rows of [[entry(),entry('供应商乙','乙公司','OA001',200)],[entry('供应商甲','甲公司','OA001',40),entry('供应商甲','甲公司','OA001',60)],[entry(),entry()]]){
  const result=match(rows);assert.equal(result.status,'review');assert.equal(result.candidates.length,2);assert.equal(result.method,'oa');assert.match(result.reason,/同一OA号.*多条预算/);
 }
});
test('无OA候选时名称主体多条也不以金额或仅一条缺号筛选',()=>{
 for(const codes of [['OA001','OA002'],['','OA002'],['','']]){
  const result=match([entry('供应商甲','甲公司',codes[0]),entry('供应商甲','甲公司',codes[1],200)],{合同编号:codes[0]==='OA001'?'':'OA003'});assert.equal(result.status,'review');assert.equal(result.candidates.length,2);assert.match(result.reason,/同一供应商及付款主体.*多条预算/);
 }
});
test('分批金额不等于整笔预算时核对，不按小于预算或累计相等通过',()=>{
 const budget=parse(book([entry('供应商甲','甲公司','OA001',100)]));
 for(const amount of [40,60,100.01]){const result=C.matchBudget(payment({本次申请付款金额:amount}),budget);assert.equal(result.status,'review');assert.match(result.reason,/申请金额.*预算金额.*不一致/);}
 assert.equal(C.matchBudget(payment(),budget).status,'matched');assert.equal(C.matchBudget(payment(),budget).status,'matched');assert.equal(budget.rows[0].amount,10000n);
});
test('金额零有效，非法申请金额必须核对',()=>{
 assert.equal(match([entry('供应商甲','甲公司','OA001',0)],{本次申请付款金额:0}).status,'matched');
 for(const amount of ['',null,'/',-1,'0.001']){const result=match([entry()],{本次申请付款金额:amount});assert.equal(result.status,'review');assert.match(result.reason,/申请金额异常/);}
});
test('空白或占位供应商与主体、包含多家供应商不能自动通过',()=>{
 for(const field of ['supplier','subject'])for(const value of ['', '/', '待核对', '多家供应商', '供应商甲、供应商乙']){
  const budget=parse(book([entry(field==='supplier'?value:'供应商甲',field==='subject'?value:'甲公司')]));assert.deepEqual(budget.errors,[]);
  assert.equal(C.matchBudget(payment(),budget).status,'review');
  const request=payment(field==='supplier'?{'供应商全称':value}:{'付款主体':value});assert.equal(C.matchBudget(request,budget).status,'review');
 }
});
test('多个或异常OA号不得猜测匹配，保留预算来源用于核对',()=>{
 for(const oa of ['OA001;OA002','OA001、OA002','OA001 OA002','MDST-DGFK20261201001\nMDST-DGFK20261201002','OA001/002','未知']){
  const budget=parse(book([entry('供应商甲','甲公司',oa)]));assert.deepEqual(budget.errors,[]);const result=C.matchBudget(payment({合同编号:''}),budget);assert.equal(result.status,'review',oa);assert.match(result.reason,/OA号.*多个编号或格式异常/);assert.equal(result.candidates[0].source.row,2);
  const request=match([entry()],{合同编号:oa});assert.equal(request.status,'review');assert.match(request.reason,/付款明细OA号/);
 }
});
test('未读到预算、无候选或预算错误不能返回已匹配',()=>{
 assert.equal(C.matchBudget(payment(),null).status,'review');assert.equal(C.matchBudget(payment(),{errors:['错误'],rows:[]}).status,'review');
 const result=match([entry('无关供应商','其他公司','OA002')]);assert.equal(result.status,'review');assert.equal(result.candidates.length,0);assert.match(result.reason,/未找到对应付款/);
});
