const {test}=require('node:test'),assert=require('node:assert/strict');
const C=require('../payment-core.js'),X=require('../vendor/xlsx.full.min.js');
const HEADERS=['流程编号','付款金额','未操作者','收款单位'];
const row=(oa='OA001',amount=100,unoperated='人员甲',supplier='供应商甲')=>[oa,amount,unoperated,supplier];
function book(rows=[row()],headers=HEADERS,name='OA明细'){
 const result=X.utils.book_new();X.utils.book_append_sheet(result,X.utils.aoa_to_sheet([headers,...rows]),name);return result;
}
const parse=(value,selected='')=>C.parsePaymentOA(value,'对公付款.xlsx',X,selected);
const payment=(extra={})=>({合同编号:'OA001',本次申请付款金额:100,供应商全称:'供应商甲',付款主体:'甲公司',...extra});
const match=(rows,extra={})=>C.matchPaymentOA(payment(extra),parse(book(rows)));

test('OA读取收款单位及分金额，保留所有来源且不修改源表',()=>{
 const b=book([row('OA001','￥1,234.56'),row('OA001','1234.56'),row('000123',0,'')]),original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.errors,[]);assert.deepEqual(result.sheetNames,['OA明细']);assert.equal(result.sheetName,'OA明细');assert.equal(result.rows.length,3);
 assert.deepEqual(result.rows[0],{oa:'OA001',supplier:'供应商甲',amount:123456n,amountText:'1,234.56',unoperated:'人员甲',source:{filename:'对公付款.xlsx',sheet:'OA明细',row:2}});
 assert.equal(result.rows[1].source.row,3);assert.equal(result.rows[2].oa,'000123');assert.equal(result.rows[2].amount,0n);assert.equal(result.rows[2].unoperated,'');assert.equal(JSON.stringify(b),original);
});
test('动态列位置、标题行、表头空格与格式空列兼容，付款主体不作为匹配条件',()=>{
 const b=book([]),s=X.utils.aoa_to_sheet([['导出说明'],[],['未操作\n者','收款单位',' 流程编码 ','付款金额','业务所属公司'],['人员甲','供应商甲','OA001',100,'不同主体'],[],['','供应商乙','OA002',0,'其他主体']]);
 s.XFD1048576={t:'z',s:{fill:{fgColor:{rgb:'FFFF00'}}}};s['!ref']='A1:XFD1048576';s.E4={t:'e',v:23,f:'#REF!'};b.Sheets.OA明细=s;
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>[r.oa,r.supplier,r.amount,r.source.row]),[['OA001','供应商甲',10000n,4],['OA002','供应商乙',0n,6]]);
 const matched=C.matchPaymentOA(payment({付款主体:''}),result);assert.equal(matched.status,'matched');assert.equal(matched.unoperated,'人员甲');
});
test('流程编码优先流程编号，空单元格不回退，未用旧编号的公式错误不影响新编号',()=>{
 const b=book([['旧001','OA001',100,'人员甲','供应商甲'],['旧002','',200,'人员乙','供应商乙']],['流程编号','流程编码','付款金额','未操作者','收款单位']);b.Sheets.OA明细.A2={t:'e',f:'#REF!',v:23};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>r.oa),['OA001']);assert.equal(result.warnings.length,1);assert.match(result.warnings[0],/第 3 行.*流程号为空/);
 assert.equal(C.matchPaymentOA(payment(),result).status,'matched');assert.equal(C.matchPaymentOA(payment({合同编号:'旧002',本次申请付款金额:200}),result).status,'review');
 b.Sheets.OA明细.B3={t:'n',f:'A3'};assert.match(parse(b).errors.join(' '),/!B3.*公式错误或缺少有效计算结果/);
});
test('收款单位优先所有供应商别名，存在空值不回退，兼容仅有各别名的文件',()=>{
 for(const alias of ['供应商名称','供应商全称','供应商']){
  const b=book([row()],['流程编号','付款金额','未操作者',alias]);const result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows[0].supplier,'供应商甲');assert.equal(C.matchPaymentOA(payment(),result).status,'matched');
 }
 const b=book([['OA001',100,'人员甲','','供应商甲','供应商乙','供应商丙']],['流程编号','付款金额','未操作者','收款单位','供应商名称','供应商全称','供应商']);b.Sheets.OA明细.E2={t:'e',v:23,f:'#REF!'};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows[0].supplier,'');assert.equal(C.matchPaymentOA(payment(),result).status,'review');assert.deepEqual(C.matchPaymentOA(payment(),result).candidates,[]);
 const alternate=parse(book([['OA001',100,'人员甲','供应商乙','供应商甲','供应商丙']],['流程编号','付款金额','未操作者','供应商全称','供应商名称','供应商']));assert.equal(alternate.rows[0].supplier,'供应商甲');
});
test('多个有效OA表必须选择，没有有效表时明确必需收款单位及其他字段',()=>{
 const b=book();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([HEADERS,row('OA002',200,'人员乙')]),'OA第二表');X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['流程编号'],['OA003']]),'说明');
 const result=parse(b);assert.deepEqual(result.sheetNames,['OA明细','OA第二表']);assert.equal(result.sheetName,'');assert.deepEqual(result.errors,['请选择对公付款OA工作表']);assert.deepEqual(result.rows,[]);
 const second=parse(b,'OA第二表');assert.deepEqual(second.errors,[]);assert.equal(second.rows[0].oa,'OA002');assert.equal(second.rows[0].source.sheet,'OA第二表');
 assert.match(parse(b,'说明').errors[0],/对公付款.xlsx \/ 说明.*收款单位.*付款金额.*未操作者/);assert.match(parse(b,'不存在').errors[0],/不存在/);
 for(const omitted of HEADERS){const headers=HEADERS.filter(h=>h!==omitted);assert.match(parse(book([],headers)).errors[0],/未找到OA表头.*收款单位/);}
 assert.match(parse(X.utils.book_new()).errors[0],/未找到OA表头/);assert.match(parse(book([])).errors[0],/没有OA明细/);
});
test('重复的有效表头报错，不任取一列；低优先级别名重复可忽略',()=>{
 for(const header of HEADERS){const result=parse(book([], [...HEADERS,header]));assert.equal(result.rows.length,0);assert.match(result.errors[0],/第 1 行.*多个列/);}
 const prioritized=parse(book([['旧1','旧2','OA001',100,'人员甲','供应商甲']],['流程编号','流程编号','流程编码','付款金额','未操作者','收款单位']));assert.deepEqual(prioritized.errors,[]);assert.equal(prioritized.rows[0].oa,'OA001');
});
test('四项必要字段公式错误或缺缓存阻断全部记录，有效零和空未操作者缓存保留',()=>{
 for(const address of ['A3','B3','C3','D3'])for(const cell of [{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'},{t:'s',f:'A99'},{t:'z',f:'A99',v:''}]){
  const b=book([row(),row('OA002',200,'人员乙')]);b.Sheets.OA明细[address]=cell;const result=parse(b);assert.equal(result.rows.length,0);assert(result.errors.some(e=>e.includes('对公付款.xlsx / OA明细!'+address)&&/公式错误或缺少有效计算结果/.test(e)));
 }
 const b=book();b.Sheets.OA明细.B2={t:'n',f:'1-1',v:0};b.Sheets.OA明细.C2={t:'s',f:'IF(TRUE,"","人员甲")',v:''};
 const result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows[0].amount,0n);assert.equal(result.rows[0].unoperated,'');assert.equal(C.matchPaymentOA(payment({本次申请付款金额:0}),result).status,'matched');
});
test('空白或异常OA付款金额不能当零，任何异常都不返回可用部分',()=>{
 for(const amount of ['',null,'/',-1,'1.001','1,23',true,Infinity,'9999999999999']){const result=parse(book([row(),row('OA002',amount)]));assert.equal(result.rows.length,0,String(amount));assert(result.errors.some(e=>e.includes('!B3')));}
});
test('无流程号草稿只提示来源，跳过其他必要字段异常；全空行不提示',()=>{
 const b=book([row(),row('','','人员乙'),row('未提交OA',0,'人员丙'),row('/',-1,'人员丁'),['','','',''],row('',100,'人员戊')]);
 b.Sheets.OA明细.B3={t:'e',v:23,f:'#REF!'};b.Sheets.OA明细.C3={t:'n',f:'A99'};b.Sheets.OA明细.D3={t:'e',v:23,f:'#REF!'};b.Sheets.OA明细.A7={t:'s',f:'IF(TRUE,"","OA002")',v:''};
 const original=JSON.stringify(b),result=parse(b);assert.deepEqual(result.errors,[]);assert.equal(result.rows.length,1);assert.equal(result.warnings.length,4);
 for(const r of [3,4,5,7])assert(result.warnings.some(w=>w.includes('对公付款.xlsx / OA明细 第 '+r+' 行：流程号为空，未参与OA核对')));
 assert.equal(C.matchPaymentOA(payment(),result).status,'matched');assert.equal(JSON.stringify(b),original);
 const onlyDrafts=parse(book([row('',100)]));assert.equal(onlyDrafts.rows.length,0);assert.equal(onlyDrafts.warnings.length,1);assert.match(onlyDrafts.errors[0],/没有OA明细/);
});
test('无流程号也不能跳过编号公式错误或缓存缺失',()=>{
 for(const cell of [{t:'e',v:'',f:'#REF!'},{t:'s',f:'A99'},{t:'z',f:'A99',v:''}]){
  const b=book([row(),['','','','']]);b.Sheets.OA明细.A3=cell;const result=parse(b);assert.equal(result.rows.length,0);assert.match(result.errors[0],/!A3.*公式错误或缺少有效计算结果/);assert.deepEqual(result.warnings,[]);
 }
});
test('文本及显示前导零保留，数字OA精度异常有源定位并阻断',()=>{
 const b=book([row('00000123'),row(123)]);b.Sheets.OA明细.A3.z='00000000';const original=JSON.stringify(b),result=parse(b);
 assert.deepEqual(result.errors,[]);assert.deepEqual(result.rows.map(r=>r.oa),['00000123','00000123']);assert.equal(JSON.stringify(b),original);
 assert.equal(C.matchPaymentOA(payment({合同编号:'00000123'}),result).status,'matched');const mismatch=C.matchPaymentOA(payment({合同编号:'123'}),result);assert.equal(mismatch.status,'review');assert.equal(mismatch.suggestedOA,'00000123');
 for(const value of [1000000000000000,Number.MAX_SAFE_INTEGER+1,12.3]){const parsed=parse(book([row(value)]));assert.equal(parsed.rows.length,0);assert.match(parsed.errors[0],/!A2.*精度/);const check=match([row()],{合同编号:value});assert.equal(check.suggestedOA,'OA001');assert.match(check.reason,/精度/);}
});
test('供应商仅trim精确匹配，金额按分，OA规范化后核对，不改变原申请',()=>{
 const record=payment({合同编号:' mdst - dgfk20261201001 ',本次申请付款金额:.1,供应商全称:' 供应商甲 ',付款主体:'不同公司'}),before=JSON.stringify(record),result=C.matchPaymentOA(record,parse(book([row('ＭＤＳＴ－ＤＧＦＫ２０２６１２０１００１','0.10','人员甲\n人员乙',' 供应商甲 ')])));
 assert.equal(result.status,'matched');assert.equal(result.suggestedOA,'MDST-DGFK20261201001');assert.equal(result.unoperated,'人员甲\n人员乙');assert.equal(result.candidates.length,1);assert.equal(JSON.stringify(record),before);
 for(const supplier of ['供应商乙','供应 商甲','供應商甲']){const check=match([row()],{供应商全称:supplier});assert.equal(check.status,'review');assert.deepEqual(check.candidates,[]);assert.equal(check.suggestedOA,'');}
 assert.equal(match([row()],{本次申请付款金额:'100.00'}).status,'matched');
});
test('供应商缺失、占位或多家供应商不可匹配，预算与OA均不模糊合并',()=>{
 for(const supplier of ['', ' ', '/', '待核对','无','1234','多家供应商','甲、乙','甲；乙','甲\n乙']){
  const result=match([row('OA001',100,'人员甲',supplier)],{供应商全称:supplier});assert.equal(result.status,'review',supplier);assert.deepEqual(result.candidates,[]);assert.equal(result.suggestedOA,'');assert.match(result.reason,/供应商为空、名称异常或包含多家供应商/);
 }
});
test('英文法人名称中的逗号和&不推测为多家供应商，仍精确匹配完整名称',()=>{
 for(const supplier of ['Example Healthcare Consulting, LLC','Example R&D Company','Example+Partners']){
  const result=match([row('OA001',100,'人员甲',supplier)],{供应商全称:supplier});assert.equal(result.status,'matched');assert.equal(result.suggestedOA,'OA001');
 }
});
test('唯一名称金额匹配后才核对原OA，缺号异常及不同编号均提示建议号',()=>{
 for(const oa of ['', ' ', '/', '未提交OA','未申请OA','未交OA','OA001;OA002','OA001 OA002','MDST-DGFK20261201001\nMDST-DGFK20261201002','OA001/002','待核对','OA999']){
  const record=payment({合同编号:oa}),before=JSON.stringify(record),result=C.matchPaymentOA(record,parse(book()));assert.equal(result.status,'review',oa);assert.equal(result.suggestedOA,'OA001');assert.equal(result.candidates.length,1);assert.equal(result.unoperated,'人员甲');assert.match(result.reason,/原合同编号.*唯一匹配的OA号.*OA001/s);assert.equal(JSON.stringify(record),before);
 }
});
test('没有名称金额匹配时不按原OA回退，无效来源OA不作为候选',()=>{
 for(const candidate of [row('OA001',100,'人员甲','供应商乙'),row('OA001',200),row('OA001;OA002'),row('OA001',100,'人员甲','')]){
  const result=match([candidate]);assert.equal(result.status,'review');assert.deepEqual(result.candidates,[]);assert.equal(result.suggestedOA,'');assert.equal(result.unoperated,'');assert.match(result.reason,/未找到收款单位.*付款金额.*均一致/);
 }
});
test('同号重复信息一致只算一个流程，保留各源行；原申请重复不累计',()=>{
 const parsed=parse(book([row(),row(' oa001 ','100.00',' 人员甲 ',' 供应商甲 '),row()])),before=JSON.stringify(parsed,(_,v)=>typeof v==='bigint'?String(v):v),result=C.matchPaymentOA(payment(),parsed);
 assert.equal(result.status,'matched');assert.equal(result.unoperated,'人员甲');assert.equal(result.suggestedOA,'OA001');assert.equal(result.candidates.length,3);assert.deepEqual(result.candidates.map(r=>r.source.row),[2,3,4]);assert.deepEqual(C.matchPaymentOA(payment(),parsed),result);assert.equal(JSON.stringify(parsed,(_,v)=>typeof v==='bigint'?String(v):v),before);
});
test('同供应商金额对应多个不同流程，不按原号选择，列全候选',()=>{
 const parsed=parse(book([row(),row('OA002'),row(' oa001 '),row('OA003',200),row('OA004',100,'人员甲','供应商乙')]));
 for(const oa of ['OA001','OA002','未提交OA']){const result=C.matchPaymentOA(payment({合同编号:oa}),parsed);assert.equal(result.status,'review');assert.equal(result.suggestedOA,'');assert.equal(result.unoperated,'');assert.deepEqual(result.candidates.map(r=>r.source.row),[2,3,4]);assert.match(result.reason,/2 个OA流程/);}
});
test('唯一流程仍核验该号所有源行，供应商金额或未操作者冲突不能被初筛隐藏',()=>{
 for(const conflicting of [row('OA001',200),row('OA001',100,'人员乙'),row('OA001',100,''),row('OA001',100,'人员甲','供应商乙'),row('OA001',100,'人员甲','')]){
  for(const rows of [[row(),conflicting],[conflicting,row()]]){const result=match(rows);assert.equal(result.status,'review');assert.equal(result.unoperated,'');assert.equal(result.suggestedOA,'');assert.equal(result.candidates.length,2);assert.match(result.reason,/收款单位、付款金额或未操作者不一致/);}
 }
 const reordered=match([row('OA001',100,'甲\n乙'),row('OA001',100,'乙\n甲')]);assert.equal(reordered.status,'review');
});
test('空未操作者不阻止唯一匹配，同金额不同供应商分别定位正确流程',()=>{
 const parsed=parse(book([row('OA001',100,''),row('OA002',100,'人员乙','供应商乙')]));
 const first=C.matchPaymentOA(payment(),parsed),second=C.matchPaymentOA(payment({合同编号:'OA002',供应商全称:'供应商乙'}),parsed);
 assert.equal(first.status,'matched');assert.equal(first.unoperated,'');assert.equal(first.candidates.length,1);assert.equal(second.status,'matched');assert.equal(second.unoperated,'人员乙');assert.equal(second.candidates[0].source.row,3);
});
test('申请金额必须严格相等，不累计分笔，不以小于OA金额通过',()=>{
 for(const amount of [40,60,99.99,100.01]){const result=match([row()],{本次申请付款金额:amount});assert.equal(result.status,'review');assert.deepEqual(result.candidates,[]);assert.equal(result.unoperated,'');assert.match(result.reason,/未找到收款单位.*付款金额.*均一致/);}
 const separate=match([row('OA001',40),row('OA002',60)]);assert.equal(separate.status,'review');assert.deepEqual(separate.candidates,[]);
 for(const amount of ['',null,'/',-1,'0.001']){const result=match([row()],{本次申请付款金额:amount});assert.equal(result.status,'review');assert.match(result.reason,/申请金额异常/);}
});
test('没有有效OA来源时不返回匹配结果',()=>{
 for(const parsed of [null,undefined,{errors:['损坏'],rows:[]}]){const result=C.matchPaymentOA(payment(),parsed);assert.equal(result.status,'review');assert.equal(result.unoperated,'');assert.equal(result.suggestedOA,'');assert.deepEqual(result.candidates,[]);assert.match(result.reason,/尚未有效读取/);}
});
