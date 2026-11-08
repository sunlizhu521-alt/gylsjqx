const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-ledger-core.js');
const LH=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','是否谈判付款'];
const OH=['流程编号','当前节点','未操作者','流程状态'];
const ledgerRow=(o={})=>({'供应商名称':'供应商甲','付款主体':'国源','对公付款0A流程号':'OA001','申请付款金额':100,'摘要':'货款','应付款时间':'2026-11-06','是否谈判付款':'否',...o});
const flow=(o={})=>({'流程编号':'OA001','当前节点':'出纳执行付款','未操作者':'人员甲','流程状态':'归档',...o});
function sheet(headers,rows,title=false){return X.utils.aoa_to_sheet([...(title?[['导出标题']]:[]),headers,...rows.map(row=>headers.map(h=>row[h]??''))]);}
function ledger(rows){const b=X.utils.book_new();['李奇','徐英','邢芳芳','操思敏','孙立柱'].forEach((p,i)=>X.utils.book_append_sheet(b,sheet(LH,i?[]:rows),'对账明细-'+p));return b;}
function oa(rows,headers=OH,name='Sheet0',title=false){const book=X.utils.book_new();X.utils.book_append_sheet(book,sheet(headers,rows,title),name);return {book,filename:'原始OA.xlsx'};}
const build=(rows,source)=>C.build(ledger(rows),'台账.xlsx',{month:'2026-11',...(source?{oa:source}:{})});
const extras=r=>r.sheets[0].rows.map(row=>row.slice(6).map(C.formatCell));
test('原始OA按编号精确匹配，金额/供应商/主体不参与，也不增付款行或改变原6列',()=>{
 const records=[ledgerRow(),ledgerRow({'申请付款金额':20}),ledgerRow({'供应商名称':'供应商乙','付款主体':'迈德斯特'})],source=oa([flow({'付款金额':999,'供应商':'另一供应商'})],[...OH,'付款金额','供应商']),l=ledger(records),before=JSON.stringify([l,source]),base=C.build(l,'台账.xlsx',{month:'2026-11'}),r=C.build(l,'台账.xlsx',{month:'2026-11',oa:source});
 assert.equal(r.sheets[0].rows.length,3);assert(extras(r).every(row=>JSON.stringify(row)===JSON.stringify(['出纳执行付款','人员甲','归档','已付款'])));assert.deepEqual(r.sheets[0].rows.map(row=>row.slice(0,6)),base.sheets[0].rows.map(row=>row.slice(0,6)));assert.deepEqual(r.sheets[0].sources,base.sheets[0].sources);assert.equal(JSON.stringify([l,source]),before);
});
test('归档/已归档才已付款，未归档未付款，其他状态待核对；办理人付款执行不能代替归档',()=>{
 for(const [state,expected]of [['归档','已付款'],['已归档','已付款'],['未归档','未付款'],['流程归档','待核对'],['','待核对'],['审批中','待核对']]){const r=build([ledgerRow()],oa([flow({'流程状态':state})]));assert.equal(extras(r)[0][3],expected);assert.equal(extras(r)[0][0],'出纳执行付款');}
});
test('相同供应商同金额不同OA分别读取流程信息，不混用另一笔节点或付款状态',()=>{
 const r=build([ledgerRow(),ledgerRow({'对公付款0A流程号':'OA002'})],oa([flow(),flow({'流程编号':'OA002','当前节点':'财务审核','未操作者':'人员乙','流程状态':'未归档'})]));
 assert.deepEqual(r.sheets[0].rows.map(row=>row[2].v),['OA001','OA002']);assert.deepEqual(extras(r),[['出纳执行付款','人员甲','归档','已付款'],['财务审核','人员乙','未归档','未付款']]);
});
test('无OA上传四字段全空；上传后缺号、占位、多号及未匹配只待核对且不借号',()=>{
 assert.deepEqual(extras(build([ledgerRow()])),[['','','','']]);
 for(const code of ['', '未提交OA','未申请OA','未匹配','OA999','OA001/OA002','OA001 OA002','OA001\nOA002']){const r=build([ledgerRow({'对公付款0A流程号':code})],oa([flow()]));assert.deepEqual(extras(r),[['','','','待核对']],code);assert.equal(r.sheets[0].rows[0][2].v,code);assert.match(r.warnings.join('\n'),/台账.xlsx.*李奇 第2行/);}
});
test('相同OA一致多行只读一份信息，任一流程字段冲突不任选并列出来源行',()=>{
 const same=build([ledgerRow()],oa([flow(),flow()]));assert.equal(same.sheets[0].rows.length,1);assert.equal(extras(same)[0][3],'已付款');assert.equal(same.warnings.length,0);
 for(const change of [{'当前节点':'经理'},{'未操作者':'另一人'},{'流程状态':'未归档'}]){const r=build([ledgerRow()],oa([flow(),flow(change)]));assert.deepEqual(extras(r),[['','','','待核对']]);assert.match(r.warnings.join('\n'),/OA001.*冲突.*原始OA.xlsx.*第2行.*第3行/);}
});
test('任意表名按4个表头识别唯一原始OA表，列顺序标题行与流程编码别名兼容',()=>{
 const source=oa([{...flow(),'流程编码':'oa001'}],OH.map(h=>h==='流程编号'?'流程编码':h).reverse(),'审批明细',true);X.utils.book_append_sheet(source.book,sheet(['摘要'],[{'摘要':'说明'}]),'说明');const info=C.validateOA(source.book,source.filename);assert.deepEqual(info,{sheetName:'审批明细',count:1,warnings:[]});assert.equal(extras(build([ledgerRow()],source))[0][3],'已付款');
 X.utils.book_append_sheet(source.book,sheet(OH,[flow()]),'第二份OA');assert.throws(()=>C.validateOA(source.book,source.filename),/多个 OA 工作表/);
 const missing=oa([flow()],OH.filter(h=>h!=='未操作者'));assert.throws(()=>C.validateOA(missing.book,missing.filename),/未找到.*流程编号.*未操作者/);
 const previous=oa([{'对公付款OA号':'OA001','当前节点':'经理','未操作者':'人','流程状态':'归档'}],['对公付款OA号','当前节点','未操作者','流程状态'],'付款进度整理');assert.throws(()=>C.validateOA(previous.book,previous.filename),/未找到/);
});
test('原始OA相关重复表头阻断，必要公式缓存取值，错误或缺缓存精确定位',()=>{
 const duplicate=oa([{...flow(),'流程编码':'OA002'}],[...OH,'流程编码']);assert.throws(()=>C.validateOA(duplicate.book,duplicate.filename),/字段.*流程编号.*重复/);
 const cached=oa([flow({'当前节点':{t:'s',f:'"出纳"',v:'出纳'},'未操作者':{t:'s',f:'""',v:''},'流程状态':{t:'s',f:'"归档"',v:'归档'}})]),r=build([ledgerRow()],cached);assert.deepEqual(extras(r),[['出纳','','归档','已付款']]);assert(r.sheets[0].rows[0].every(c=>c.f===undefined));
 for(const h of OH)for(const cell of [{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const source=oa([flow({[h]:cell})]);assert.throws(()=>C.validateOA(source.book,source.filename),/原始OA.xlsx.*![A-D]2/);}
 const unrelated=oa([flow({'办理人':{t:'n',f:'A99'},'付款金额':{t:'e',v:23}})],[...OH,'办理人','付款金额']);assert.equal(extras(build([ledgerRow()],unrelated))[0][3],'已付款');
});
test('数值OA按显示格式保留前导零匹配，新OA或月份重算不遗留旧状态',()=>{
 const l=ledger([ledgerRow({'对公付款0A流程号':{t:'n',v:123,z:'000000'}})]),source=oa([flow({'流程编号':{t:'n',v:123,z:'000000'}})]);const r=C.build(l,'台账.xlsx',{month:'2026-11',oa:source});assert.equal(r.sheets[0].rows[0][2].v,'000123');assert.equal(extras(r)[0][3],'已付款');
 const changed=oa([flow({'流程编号':'000123','流程状态':'未归档'})]);assert.equal(extras(C.build(l,'台账.xlsx',{month:'2026-11',oa:changed}))[0][3],'未付款');assert.deepEqual(extras(C.build(l,'台账.xlsx',{month:'2026-11'})),[['','','','']]);assert.equal(C.build(l,'台账.xlsx',{month:'2026-12',oa:changed}).sheets[0].rows.length,0);
});
test('无效源OA提示来源不用于匹配，真实十列表可回读并保留原生金额和日期',()=>{
 const source=oa([flow(),flow({'流程编号':'未匹配'})]),info=C.validateOA(source.book,source.filename);assert.equal(info.count,2);assert.match(info.warnings.join('\n'),/原始OA.xlsx.*第3行/);const r=build([ledgerRow()],source),out=C.workbook(r),roundtrip=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer'}),s=roundtrip.Sheets['2026年11月正常付款'];assert.deepEqual(roundtrip.SheetNames,['2026年11月正常付款']);assert.equal(s['!ref'],'A1:J2');assert.equal(s.D2.t,'n');assert.equal(s.F2.t,'n');assert.equal(s.J2.v,'已付款');
});
