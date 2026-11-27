const {test}=require('node:test'),assert=require('node:assert/strict');
const C=require('../payment-core.js'),X=require('../vendor/xlsx.full.min.js');
const HEADERS=['流程编号','付款金额','未操作者'];
function book(rows=[['OA001',100,'人员甲']],headers=HEADERS,name='OA明细'){
 const result=X.utils.book_new();X.utils.book_append_sheet(result,X.utils.aoa_to_sheet([headers,...rows]),name);return result;
}
const parse=(value,selected='')=>C.parsePaymentOA(value,'对公付款.xlsx',X,selected);
const payment=(extra={})=>({合同编号:'OA001',本次申请付款金额:100,供应商全称:'供应商甲',付款主体:'甲公司',...extra});
const match=(rows,extra={})=>C.matchPaymentOA(payment(extra),parse(book(rows)));

test('OA表按必需表头读取，保留所有来源与金额分精度且不修改源表',()=>{
 const b=book([['OA001','￥1,234.56','人员甲'],['OA001','1234.56','人员甲'],['000123',0,'']]),original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.errors,[]);assert.deepEqual(result.sheetNames,['OA明细']);assert.equal(result.sheetName,'OA明细');assert.equal(result.rows.length,3);
 assert.deepEqual(result.rows[0],{oa:'OA001',amount:123456n,amountText:'1,234.56',unoperated:'人员甲',source:{filename:'对公付款.xlsx',sheet:'OA明细',row:2}});
 assert.equal(result.rows[1].source.row,3);assert.equal(result.rows[2].oa,'000123');assert.equal(result.rows[2].amount,0n);assert.equal(result.rows[2].unoperated,'');assert.equal(JSON.stringify(b),original);
});
test('OA动态列位置、标题行、表头空格与格式空列兼容，不读取无关业务字段',()=>{
 const b=book([]),s=X.utils.aoa_to_sheet([['导出说明'],[],['未操作\n者','收款单位',' 流程编码 ','付款金额','业务所属公司'],['人员甲','不同供应商','OA001',100,'不同主体'],[],['','其他供应商','OA002',0,'其他主体']]);
 s.XFD1048576={t:'z',s:{fill:{fgColor:{rgb:'FFFF00'}}}};s['!ref']='A1:XFD1048576';s.B4={t:'e',v:23,f:'#REF!'};b.Sheets.OA明细=s;
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>[r.oa,r.amount,r.source.row]),[['OA001',10000n,4],['OA002',0n,6]]);
 const matched=C.matchPaymentOA(payment({供应商全称:'',付款主体:''}),result);assert.equal(matched.status,'matched');assert.equal(matched.unoperated,'人员甲');
});
test('流程编码优先流程编号，空单元格不回退；未用的旧编号公式错误不影响新编号',()=>{
 const b=book([['旧001','OA001',100,'人员甲'],['旧002','',200,'人员乙']],['流程编号','流程编码','付款金额','未操作者']);b.Sheets.OA明细.A2={t:'e',f:'#REF!',v:23};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>r.oa),['OA001']);assert.equal(result.warnings.length,1);assert.match(result.warnings[0],/第 3 行.*流程号为空/);
 assert.equal(C.matchPaymentOA(payment(),result).status,'matched');assert.equal(C.matchPaymentOA(payment({合同编号:'旧002',本次申请付款金额:200}),result).status,'review');
 b.Sheets.OA明细.B3={t:'n',f:'A3'};assert.match(parse(b).errors.join(' '),/!B3.*公式错误或缺少有效计算结果/);
});
test('多个有效OA工作表必须选择，选中后不读其他表；没有有效表时明确缺失字段',()=>{
 const b=book();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([HEADERS,['OA002',200,'人员乙']]),'OA第二表');X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['流程编号'],['OA003']]),'说明');
 const result=parse(b);assert.deepEqual(result.sheetNames,['OA明细','OA第二表']);assert.equal(result.sheetName,'');assert.deepEqual(result.errors,['请选择对公付款OA工作表']);assert.deepEqual(result.rows,[]);
 const second=parse(b,'OA第二表');assert.deepEqual(second.errors,[]);assert.equal(second.rows.length,1);assert.equal(second.rows[0].oa,'OA002');assert.equal(second.rows[0].source.sheet,'OA第二表');
 assert.match(parse(b,'说明').errors[0],/对公付款.xlsx \/ 说明.*付款金额.*未操作者/);assert.match(parse(b,'不存在').errors[0],/不存在/);
 assert.match(parse(book([['OA001',100]],['流程编号','付款金额'])).errors[0],/必须包含.*未操作者/);assert.match(parse(X.utils.book_new()).errors[0],/未找到OA表头/);assert.match(parse(book([])).errors[0],/没有OA明细/);
});
test('重复的有效编号或付款金额表头报错，不任取一列',()=>{
 for(const headers of [['流程编码','流程编码','付款金额','未操作者'],['流程编号','流程编号','付款金额','未操作者'],['流程编码','付款金额','付款金额','未操作者'],['流程编号','付款金额','未操作者','未操作者']]){
  const result=parse(book([['OA001',100,'人员甲','其他值']],headers));assert.equal(result.rows.length,0);assert.match(result.errors[0],/第 1 行.*多个列/);
 }
 const prioritized=parse(book([['旧1','旧2','OA001',100,'人员甲']],['流程编号','流程编号','流程编码','付款金额','未操作者']));assert.deepEqual(prioritized.errors,[]);assert.equal(prioritized.rows[0].oa,'OA001');
});
test('必要列的公式错误、缺缓存阻断全部OA记录；有效零和空未操作者缓存保留',()=>{
 for(const address of ['A3','B3','C3'])for(const cell of [{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'},{t:'s',f:'A99'},{t:'z',f:'A99',v:''}]){
  const b=book([['OA001',100,'人员甲'],['OA002',200,'人员乙']]);b.Sheets.OA明细[address]=cell;const result=parse(b);assert.equal(result.rows.length,0);assert(result.errors.some(e=>e.includes('对公付款.xlsx / OA明细!'+address)&&/公式错误或缺少有效计算结果/.test(e)));
 }
 const b=book();b.Sheets.OA明细.B2={t:'n',f:'1-1',v:0};b.Sheets.OA明细.C2={t:'s',f:'IF(TRUE,"","人员甲")',v:''};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows[0].amount,0n);assert.equal(result.rows[0].unoperated,'');assert.equal(C.matchPaymentOA(payment({本次申请付款金额:0}),result).status,'matched');
});
test('空白或异常OA付款金额不能当零，任何异常都不返回可用部分',()=>{
 for(const amount of ['',null,'/',-1,'1.001','1,23',true,Infinity,'9999999999999']){
  const result=parse(book([['OA001',100,'人员甲'],['OA002',amount,'人员乙']]));assert.equal(result.rows.length,0,String(amount));assert(result.errors.some(e=>e.includes('!B3')));
 }
});
test('无流程号的草稿只提示来源，不参加匹配且不读取其金额或未操作者异常',()=>{
 const b=book([['OA001',100,'人员甲'],['','','人员乙'],['未提交OA',0,'人员丙'],['/',-1,'人员丁'],['','',''],['',100,'人员戊']]);
 b.Sheets.OA明细.B3={t:'e',v:23,f:'#REF!'};b.Sheets.OA明细.C3={t:'n',f:'A99'};b.Sheets.OA明细.A7={t:'s',f:'IF(TRUE,"","OA002")',v:''};
 const original=JSON.stringify(b),result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows.length,1);assert.equal(result.rows[0].oa,'OA001');assert.equal(result.warnings.length,4);
 for(const row of [3,4,5,7])assert(result.warnings.some(w=>w.includes('对公付款.xlsx / OA明细 第 '+row+' 行：流程号为空，未参与OA核对')));
 assert.equal(C.matchPaymentOA(payment(),result).status,'matched');assert.equal(JSON.stringify(b),original);
 const onlyDrafts=parse(book([['',100,'人员甲']]));assert.equal(onlyDrafts.rows.length,0);assert.equal(onlyDrafts.warnings.length,1);assert.match(onlyDrafts.errors[0],/没有OA明细/);
});
test('无流程号也不能跳过编号单元格公式错误或缓存缺失，真实全空行才不提示',()=>{
 for(const cell of [{t:'e',v:'',f:'#REF!'},{t:'s',f:'A99'},{t:'z',f:'A99',v:''}]){
  const b=book([['OA001',100,'人员甲'],['','','']]);b.Sheets.OA明细.A3=cell;const result=parse(b);assert.equal(result.rows.length,0);assert.match(result.errors[0],/!A3.*公式错误或缺少有效计算结果/);assert.deepEqual(result.warnings,[]);
 }
 const blank=parse(book([['OA001',100,'人员甲'],['','','']]));assert.deepEqual(blank.errors,[]);assert.deepEqual(blank.warnings,[]);assert.equal(blank.rows.length,1);
});
test('文本及显示前导零保留，数字OA精度异常有源定位并阻断',()=>{
 const b=book([['00000123',100,'人员甲'],[123,100,'人员甲']]);b.Sheets.OA明细.A3.z='00000000';const original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>r.oa),['00000123','00000123']);assert.equal(JSON.stringify(b),original);
 assert.equal(C.matchPaymentOA(payment({合同编号:'00000123'}),result).status,'matched');assert.equal(C.matchPaymentOA(payment({合同编号:'123'}),result).status,'review');
 for(const value of [1000000000000000,Number.MAX_SAFE_INTEGER+1,12.3]){const result=parse(book([[value,100,'人员甲']]));assert.equal(result.rows.length,0);assert.match(result.errors[0],/!A2.*精度/);assert.match(match([['OA001',100,'人员甲']],{合同编号:value}).reason,/精度/);}
});
test('OA编号规范化后精确匹配，金额按分比对且不受供应商付款主体影响',()=>{
 const record=payment({合同编号:' mdst - dgfk20261201001 ',本次申请付款金额:.1,供应商全称:'不同供应商',付款主体:'不同公司'}),before=JSON.stringify(record),result=C.matchPaymentOA(record,parse(book([['ＭＤＳＴ－ＤＧＦＫ２０２６１２０１００１','0.10','人员甲\n人员乙']])));
 assert.equal(result.status,'matched');assert.equal(result.unoperated,'人员甲\n人员乙');assert.equal(result.candidates.length,1);assert.equal(JSON.stringify(record),before);
 assert.equal(match([['OA001',100,'人员甲']],{本次申请付款金额:'100.00'}).status,'matched');
});
test('缺号、多号、格式异常和未找到仅核对，不按名称或金额回退',()=>{
 for(const oa of ['', ' ', '/', '未提交OA','未申请OA','未交OA']){const result=match([['OA001',100,'人员甲']],{合同编号:oa});assert.equal(result.status,'review');assert.match(result.reason,/缺少付款OA号/);assert.deepEqual(result.candidates,[]);assert.equal(result.unoperated,'');}
 for(const oa of ['OA001;OA002','OA001 OA002','MDST-DGFK20261201001\nMDST-DGFK20261201002','OA001/002','待核对']){const result=match([['OA001',100,'人员甲']],{合同编号:oa});assert.equal(result.status,'review');assert.match(result.reason,/多个编号或格式异常/);assert.deepEqual(result.candidates,[]);}
 const unmatched=match([['OA001',100,'人员甲']],{合同编号:'OA002'});assert.equal(unmatched.status,'review');assert.match(unmatched.reason,/未找到.*OA002/);assert.equal(unmatched.candidates.length,0);
 for(const invalidSource of ['', '/', '未提交OA','OA001;OA002']){const result=match([[invalidSource,100,'人员甲']]);assert.equal(result.status,'review');assert.equal(result.candidates.length,0);}
});
test('同号重复且付款金额及未操作者相同视为一次匹配，仍保留全部来源',()=>{
 const rows=[['OA001',100,'人员甲'],[' oa001 ','100.00',' 人员甲 '],['OA001',100,'人员甲']],parsed=parse(book(rows));
 const result=C.matchPaymentOA(payment(),parsed);assert.equal(result.status,'matched');assert.equal(result.unoperated,'人员甲');assert.equal(result.candidates.length,3);assert.deepEqual(result.candidates.map(r=>r.source.row),[2,3,4]);assert.equal(result.candidates[0].amount,10000n);assert.equal(parsed.rows.length,3);
});
test('同号金额或未操作者冲突时核对，不按与申请金额相同的那行选取',()=>{
 for(const rows of [[['OA001',100,'人员甲'],['OA001',200,'人员甲']],[['OA001',100,'人员甲'],['OA001',100,'人员乙']],[['OA001',100,''],['OA001',100,'人员甲']],[['OA001',100,'甲\n乙'],['OA001',100,'乙\n甲']]]){
  const result=match(rows);assert.equal(result.status,'review');assert.equal(result.unoperated,'');assert.equal(result.candidates.length,2);assert.match(result.reason,/付款金额或未操作者不一致/);
 }
});
test('空未操作者不阻止匹配，同金额不同OA分别读取，原申请重复不累计',()=>{
 const parsed=parse(book([['OA001',100,''],['OA002',100,'人员乙']]));
 const first=C.matchPaymentOA(payment(),parsed),second=C.matchPaymentOA(payment({合同编号:'OA002'}),parsed);
 assert.equal(first.status,'matched');assert.equal(first.unoperated,'');assert.equal(first.candidates.length,1);assert.equal(second.status,'matched');assert.equal(second.unoperated,'人员乙');assert.equal(second.candidates[0].source.row,3);assert.deepEqual(C.matchPaymentOA(payment(),parsed),first);
});
test('本次申请与OA金额不等时核对，不按小于或累计金额通过',()=>{
 for(const amount of [40,60,99.99,100.01]){const result=match([['OA001',100,'人员甲']],{本次申请付款金额:amount});assert.equal(result.status,'review');assert.equal(result.candidates.length,1);assert.equal(result.unoperated,'');assert.match(result.reason,/本次申请付款金额.*OA付款金额.*不一致/);}
 for(const amount of ['',null,'/',-1,'0.001']){const result=match([['OA001',100,'人员甲']],{本次申请付款金额:amount});assert.equal(result.status,'review');assert.match(result.reason,/申请金额异常/);}
});
test('没有有效OA来源时不返回匹配结果',()=>{
 for(const parsed of [null,undefined,{errors:['损坏'],rows:[]}]){const result=C.matchPaymentOA(payment(),parsed);assert.equal(result.status,'review');assert.equal(result.unoperated,'');assert.deepEqual(result.candidates,[]);assert.match(result.reason,/尚未有效读取/);}
});
