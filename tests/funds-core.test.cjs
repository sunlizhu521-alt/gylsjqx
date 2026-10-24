const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const LH=['月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付款金额','是否结清','剩余付款金额','是否谈判付款','是否提交付款申请单'];
const NH=['供应商','付款主体','对公付款0A流程号','合同总金额','已付金额','摘要','申请付款金额','单据号【财务应付单】','创建人','是否结清','剩余付款金额','原备注','付款计划'];
const AH=['申请日期','供应商全称','付款主体','合同号','合同金额','已付金额','付款事由','备注','本次申请付款金额'];
const OH=['流程编号','当前节点','未操作者','流程状态','申请人','收款单位','付款金额'];
function book(name,headers,rows){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows.map(r=>headers.map(h=>r[h]??''))]),name);return b;}
function fixture(){
 const main={'月份':'2026-10-01','供应商名称':'测试供应商甲','付款主体':'迈德斯特','对公付款0A流程号':'MDST-DGFK20261001001','开票时间':'2026-10-01','开票金额':100,'摘要':'九月货款','申请付款金额':100,'应付款时间':'2026-10-31','单据号【财务应付单】':'AP000001','OA合同编号':'MDST-HT20260101','合同总金额':1000,'已付款金额':20,'是否结清':'否','剩余付款金额':80,'是否谈判付款':'是','是否提交付款申请单':''};
 const normal={...main,'供应商名称':'测试供应商乙','对公付款0A流程号':'MDST-DGFK20261001002','单据号【财务应付单】':'AP000002','OA合同编号':'MDST-HT20260102','是否谈判付款':'否','已付款金额':100,'是否结清':'是','剩余付款金额':0};
 const ledger=book('对账明细-李奇',LH,[main,normal]);for(const person of C.PEOPLE.slice(1))X.utils.book_append_sheet(ledger,X.utils.aoa_to_sheet([LH]),'对账明细-'+person);
 const n={'供应商':'测试供应商甲','付款主体':'迈德斯特','对公付款0A流程号':main['对公付款0A流程号'],'合同总金额':1000,'已付金额':0,'摘要':'九月货款','申请付款金额':100,'单据号【财务应付单】':'AP000001','创建人':'李奇','是否结清':'否','剩余付款金额':100,'原备注':'保留备注','付款计划':100};
 const negotiated=book('10月谈判付款主表',NH,[n]);negotiated.Sheets[negotiated.SheetNames[0]].M2={t:'n',v:100,f:'G2'};
 const appRows=[main,normal].map(r=>({'申请日期':'2026-10-10','供应商全称':r['供应商名称'],'付款主体':r['付款主体'],'合同号':r['OA合同编号'],'合同金额':r['合同总金额'],'已付金额':0,'付款事由':r['摘要'],'备注':'','本次申请付款金额':r['申请付款金额']}));
 const oaRows=[main,normal].map(r=>({'流程编号':r['对公付款0A流程号'],'当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇','收款单位':r['供应商名称'],'付款金额':r['申请付款金额']}));
 return {ledger:{book:ledger,name:'资金.xlsx'},negotiated:{book:negotiated,name:'谈判.xlsx'},applications:{book:book('付款申请汇总',AH,appRows),name:'申请.xlsx'},oa:{book:book('Sheet0',OH,oaRows),name:'OA.xlsx'}};
}
function set(input,role,header,row,value,sheet){const b=input[role].book,s=b.Sheets[sheet||b.SheetNames[0]],h=X.utils.sheet_to_json(s,{header:1})[0],c=h.indexOf(header);assert(c>=0);s[X.utils.encode_cell({r:row-1,c})]=typeof value==='object'&&value!==null?value:{t:typeof value==='number'?'n':'s',v:value};}
const update=(r,address)=>r.negotiated.updates.find(u=>u.address===address)?.cell;
function resolveKeep(input){let d={};for(let i=0;i<4;i++){const r=C.build(input,d);if(!r.stats.pending)return r;for(const issue of r.issues)if(!issue.resolved)d[issue.id]=['keep'];}return C.build(input,d);}
test('五张台账、精确付款匹配、正常付款字段顺序、OA状态与不修改源表',()=>{
 const input=fixture(),before=JSON.stringify(input);const r=C.build(input);assert.equal(r.stats.pending,0);assert.equal(r.stats.normalRows,1);assert.equal(r.stats.negotiatedRows,1);assert.deepEqual(r.normalHeaders,C.NORMAL_HEADERS);assert.equal(update(r,'E2').v,20);assert.equal(update(r,'K2').v,80);assert.equal(update(r,'K2').f,'G2-E2');assert.equal(r.normalRows[0][14].f,'H2-M2');assert.equal(r.normalRows[0][14].v,0);assert.equal(r.normalRows[0][13].v,'是');assert.equal(r.normalRows[0][20].v,'李奇');assert.equal(r.normalRows[0][0].t,'n');assert.equal(r.normalRows[0][4].z,C.DATE_FORMAT);assert.equal(JSON.stringify(input),before);assert(!r.negotiated.updates.some(u=>['L2','M2'].includes(u.address)));assert.equal(r.negotiated.rows[0][12].f,'G2');
});
test('谈判工作表名称包含匹配：零、一、多张及选择合法性',()=>{
 const i=fixture();assert.deepEqual(C.negotiatedSheets(i.negotiated.book),['10月谈判付款主表']);X.utils.book_append_sheet(i.negotiated.book,X.utils.aoa_to_sheet([NH]),'谈判付款-反馈');assert.throws(()=>C.build(i),/请选择/);i.negotiated.sheetName='10月谈判付款主表';assert.doesNotThrow(()=>C.build(i));i.negotiated.sheetName='不存在';assert.throws(()=>C.build(i),/请选择/);i.negotiated.book=book('普通表',NH,[]);assert.throws(()=>C.build(i),/未找到名称/);
});
test('同一OA多行一致时折叠，不将归档当成付款或申请提交',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇','测试供应商甲','￥100.00']],{origin:-1});set(i,'oa','流程状态',3,'归档');set(i,'ledger','已付款金额',3,0);set(i,'ledger','是否结清',3,'');const r=C.build(i);assert.equal(r.stats.oaFlows,2);assert.equal(r.normalRows[0][12].v,0);assert.equal(r.normalRows[0][13].v,'否');assert.equal(r.normalRows[0][19].v,'归档');assert.equal(r.stats.pending,0);assert.equal(r.stats.oaConsistent,2);assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002');
});
test('空白实付保持未知，零实付正常计算，斜线不转零',()=>{
 const i=fixture();set(i,'ledger','已付款金额',2,'');let r=C.build(i);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'K2'),undefined);assert(r.issues.some(x=>x.kind==='amount'));assert.equal(resolveKeep(i).stats.pending,0);
 set(i,'ledger','已付款金额',2,0);r=C.build(i);assert.equal(update(r,'E2').v,0);assert.equal(update(r,'K2').v,100);set(i,'ledger','已付款金额',2,'/');r=C.build(i);assert.equal(update(r,'E2'),undefined);assert(r.issues.some(x=>x.kind==='amount'));
});
test('明确金额按分计算且不静默截断第三位或越界金额',()=>{
 assert.equal(C.money({v:'￥1,000.20'}).cents,100020n);assert.equal(C.money({v:0.1+0.2}).cents,30n);assert(C.money({v:0.001}).invalid);assert(C.money({v:'90071992547409.92'}).invalid);assert(C.money({v:'待核对'}).invalid);
 const i=fixture();set(i,'ledger','申请付款金额',2,0.3);set(i,'negotiated','申请付款金额',2,0.3);set(i,'ledger','已付款金额',2,0.1);const r=C.build(i);assert.equal(update(r,'K2').v,0.2);
});
test('已结清与金额冲突需要明确选择，负余额不得自动更新',()=>{
 const i=fixture();set(i,'ledger','是否结清',2,'是');let r=C.build(i),q=r.issues.find(x=>x.kind==='settlement');assert(q);assert.equal(update(r,'E2'),undefined);r=C.build(i,{[q.id]:['recalculate']});assert.equal(update(r,'E2').v,20);assert.equal(update(r,'J2').v,'否');set(i,'ledger','已付款金额',2,101);r=C.build(i);assert.equal(update(r,'E2'),undefined);assert(r.issues.some(x=>x.kind==='amount'));
});
test('付款申请多次日期全部保留，不使用申请表已付覆盖资金台账',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;const row=X.utils.sheet_to_json(s,{header:1})[1];row[0]='2026-10-08';row[5]=90;X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const r=C.build(i);assert.equal(r.stats.pending,0);assert.equal(update(r,'E2').v,20);const dateCol=r.negotiated.headers.indexOf('提交付款申请单日期');assert.equal(r.negotiated.rows[0][dateCol].v,'2026-10-08\n2026-10-10');
});
test('部分申请、缺失编号不猜匹配，人工确认可标记已提交',()=>{
 const i=fixture();set(i,'applications','本次申请付款金额',2,50);let r=C.build(i),q=r.issues.find(x=>x.kind==='application'&&x.source.sheet.includes('谈判'));assert(q);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'待核对');const choice=q.choices.find(c=>c.value.startsWith('match:'));r=C.build(i,{[q.id]:[choice.value]});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'是');assert.equal(update(r,'E2').v,20);
});
test('无申请不推断未提交，选择未提交才更新；无效多选保持待核对',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',AH,[]);let r=C.build(i),q=r.issues.find(x=>x.kind==='application');assert(q);r=C.build(i,{[q.id]:['keep','not-submitted']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);r=C.build(i,{[q.id]:['not-submitted']});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'否');assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('提交付款申请单日期')].v,'');
});
test('同一资金记录不能自动分配给两个谈判明细，人工匹配后禁止重复使用',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});let r=C.build(i),qs=r.issues.filter(x=>x.kind==='payment-match');assert.equal(qs.length,2);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'E3'),undefined);const match=qs[0].choices.find(c=>c.value.startsWith('match:')).value;r=C.build(i,{[qs[0].id]:[match],[qs[1].id]:[match]});assert.equal(update(r,'E2').v,20);assert.equal(update(r,'E3'),undefined);assert.equal(r.issues.find(x=>x.id===qs[1].id).resolved,false);
});
test('缺失付款分类需选择，正常付款包含已结清，不静默纳入空分类',()=>{
 const i=fixture();set(i,'ledger','是否谈判付款',3,'');let r=C.build(i),q=r.issues.find(x=>x.kind==='classification');assert.equal(r.normalRows.length,0);r=C.build(i,{[q.id]:['normal']});assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][13].v,'是');assert.equal(r.normalRows[0][15].v,'否');
});
test('OA重复冲突与多个编号保留每个编号对应状态',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');assert(q);const candidate=q.choices.find(c=>c.label.includes('出纳'));r=C.build(i,{[q.id]:[candidate.value]});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('当前节点')].v,'出纳');
 const j=fixture();set(j,'ledger','对公付款0A流程号',3,'MDST-DGFK20261001001、MDST-DGFK20261001002');set(j,'oa','收款单位',2,'测试供应商乙');r=C.build(j);q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);r=C.build(j,{[q.id]:q.choices.filter(c=>c.value.startsWith('match:')).map(c=>c.value)});assert.equal(r.normalRows[0][17].v,'MDST-DGFK20261001001：财务审核\nMDST-DGFK20261001002：财务审核');assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001001\nMDST-DGFK20261001002');
});
test('源公式仅使用有效缓存，缺缓存保留错误并阻止未核对导出',()=>{
 const i=fixture();set(i,'ledger','已付款金额',2,{t:'n',f:'10+10',v:20});let r=C.build(i);assert.equal(update(r,'E2').v,20);assert.equal(update(r,'E2').f,undefined);set(i,'ledger','已付款金额',2,{t:'n',f:'10+10'});r=C.build(i);assert.equal(update(r,'E2'),undefined);assert(r.stats.pending>0);set(i,'ledger','OA合同编号',3,{t:'e',v:23,w:'#REF!'});r=C.build(i);assert.equal(r.normalRows[0][10].v,'#REF!');assert(r.issues.some(x=>x.kind==='formula'));
});
test('已有追加字段不重复新增，源表空行和非授权字段保持完整',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号']],{origin:'N1'});s['!ref']='A1:T4';s.L4={t:'s',v:'尾注'};const r=C.build(i);assert.equal(r.negotiated.headers.length,20);assert.equal(r.negotiated.rows.length,3);assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);assert.equal(r.negotiated.rows[2][11].v,'尾注');assert(!r.negotiated.updates.some(u=>u.address==='L4'));
});
test('1904来源日期转换为1900体系，原表不变',()=>{
 const i=fixture();i.ledger.book.Workbook={WBProps:{date1904:true}};set(i,'ledger','开票时间',3,{t:'n',v:44834,z:'yyyy-mm-dd'});const r=C.build(i);assert.equal(r.normalRows[0][4].v,44834+1462);assert.equal(i.ledger.book.Sheets['对账明细-李奇'].E3.v,44834);
});
test('OA只读取指定字段，其他重复明细列不构成冲突',()=>{
 const i=fixture();i.oa.book=book('Sheet0',[...OH,'备注','备注'],[{'流程编号':'MDST-DGFK20261001001','当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇'}]);assert.doesNotThrow(()=>C.build(i));
});
test('正常付款已占用的资金来源不能再分配给谈判行，伪造选择也不能覆盖',()=>{
 const i=fixture();set(i,'ledger','是否谈判付款',2,'否');let r=C.build(i),q=r.issues.find(x=>x.kind==='payment-match');assert(q);assert.match(q.message,/已在正常付款输出/);assert.equal(q.choices.some(x=>x.value==='match:对账明细-李奇!2'),false);r=C.build(i,{[q.id]:['match:对账明细-李奇!2']});assert.equal(r.normalRows.length,2);assert.equal(update(r,'E2'),undefined);assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);r=C.build(i,{[q.id]:['keep']});assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);
});
test('源结清为空但谈判原状态与明确金额矛盾也需核对',()=>{
 const i=fixture();set(i,'ledger','是否结清',2,'');set(i,'negotiated','是否结清',2,'是');const r=C.build(i);assert(r.issues.some(x=>x.kind==='settlement'));assert.equal(update(r,'E2'),undefined);
});
test('OA或合同相同但双方财务单号不相交，不得自动匹配',()=>{
 const i=fixture();set(i,'negotiated','单据号【财务应付单】',2,'AP999999');let r=C.build(i);assert(r.issues.some(x=>x.kind==='payment-match'));assert.equal(update(r,'E2'),undefined);
 const j=fixture();set(j,'applications','备注',2,'AP999999');r=C.build(j);assert(r.issues.some(x=>x.kind==='application'&&x.source.sheet.includes('谈判')));
});
test('正常付款缺失供应商或主体不得静默当作正常匹配',()=>{
 const i=fixture();set(i,'ledger','供应商名称',3,'');const r=C.build(i);assert.equal(r.normalRows.length,1);assert(r.issues.some(x=>x.kind==='identity'));
});
test('同AP同金额但有效OA不一致必须核对，人工确认后才同步',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261024999');let r=C.build(i),q=r.issues.find(x=>x.kind==='payment-match');assert(q);assert.equal(update(r,'E2'),undefined);const match=q.choices.find(c=>c.value.startsWith('match:'));assert(match);r=C.build(i,{[q.id]:[match.value]});assert.equal(update(r,'E2').v,20);
});
test('财务单号组部分相交不能自动匹配，集合相同顺序不同仍可匹配',()=>{
 const i=fixture();set(i,'ledger','单据号【财务应付单】',2,'AP000001、AP000003');set(i,'negotiated','单据号【财务应付单】',2,'AP000001、AP000004');let r=C.build(i);assert(r.issues.some(x=>x.kind==='payment-match'));assert.equal(update(r,'E2'),undefined);
 set(i,'negotiated','单据号【财务应付单】',2,'AP000003\nAP000001');r=C.build(i);assert.equal(update(r,'E2').v,20);
 const j=fixture();set(j,'applications','备注',2,'AP000001、AP000099');r=C.build(j);assert(r.issues.some(x=>x.kind==='application'&&x.source.sheet.includes('谈判')));
});
test('付款OA组部分相交不能自动分配已付金额',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261001001、MDST-DGFK20261024999');const r=C.build(i);assert(r.issues.some(x=>x.kind==='payment-match'));assert.equal(update(r,'E2'),undefined);
});
test('明确合同号冲突不能由相同AP掩盖',()=>{
 const i=fixture();set(i,'applications','备注',2,'AP000001');set(i,'applications','合同号',2,'MDST-HT20990101');const r=C.build(i);assert(r.issues.some(x=>x.kind==='application'&&x.source.sheet.includes('谈判')));
});
function historyFixture(value){const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单'],[value,'是']],{origin:'N1'});return i;}
test('已有提交日期与本次日期合并去重排序，反复处理不丢历史',()=>{
 const i=historyFixture('2026-10-09\n2026-10-01、2026-10-10');let r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10');
 set(i,'negotiated','提交付款申请单日期',2,update(r,'N2'));set(i,'applications','申请日期',2,'2026-10-12');r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10\n2026-10-12');
});
test('历史提交日期不能解析则单独核对，保留选择不会清空或覆盖',()=>{
 const i=historyFixture('10月初提交过');let r=C.build(i),q=r.issues.find(x=>x.kind==='application-history');assert(q);assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2'),undefined);r=C.build(i,{[q.id]:['keep']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(r.negotiated.rows[0][13].v,'10月初提交过');assert.equal(update(r,'N2'),undefined);
});
test('未匹配申请且保留原值时历史日期与状态原样保留',()=>{
 const i=historyFixture('2026-09-10、2026-09-12');i.applications.book=book('付款申请汇总',AH,[]);let r=C.build(i),q=r.issues.find(x=>x.kind==='application'&&x.source.sheet.includes('谈判'));r=C.build(i,{[q.id]:['keep']});assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2'),undefined);assert.equal(r.negotiated.rows[0][13].v,'2026-09-10、2026-09-12');
});
test('历史提交日期为Excel日期数值时正确转换后合并',()=>{
 const i=historyFixture(46294);const r=C.build(i);assert.equal(update(r,'N2').v,[X.SSF.format('yyyy-mm-dd',46294),'2026-10-10'].sort().join('\n'));
});
test('历史提交日期公式缺少缓存不能当作空白覆盖',()=>{
 const i=historyFixture('2026-09-10');set(i,'negotiated','提交付款申请单日期',2,{t:'n',f:'DATE(2026,9,10)'});let r=C.build(i),q=r.issues.find(x=>x.kind==='application-history');assert(q);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].f,'DATE(2026,9,10)');r=C.build(i,{[q.id]:['keep']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].f,'DATE(2026,9,10)');
});
test('按收款单位和付款金额匹配，原OA号错或空需确认，新流程编号不覆盖原号',()=>{
 for(const old of ['', 'MDST-DGFK20990101001']){
  const i=fixture();set(i,'ledger','对公付款0A流程号',3,old);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.equal(r.normalRows[0][21].v,'');
  const choice=q.choices.find(c=>c.label.includes('MDST-DGFK20261001002'));assert.match(choice.label,/收款单位 测试供应商乙；付款金额 100/);
  r=C.build(i,{[q.id]:[choice.value]});assert.equal(r.normalRows[0][3].v,old);assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][12].v,100);assert.equal(r.normalRows[0][19].v,'未归档');
 }
});
test('相同OA但收款单位或付款金额不同不得回退按号匹配，也不使用已付金额匹配',()=>{
 for(const [field,value] of [['收款单位','其他供应商'],['付款金额',20]]){
  const i=fixture();set(i,'oa',field,2,value);const r=C.build(i),q=r.issues.find(x=>x.id==='oa:10月谈判付款主表!2');assert(q);assert.equal(q.choices.length,1);assert.match(q.message,/OA 来源记录/);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('流程编号')].v,'');
 }
});
test('同收款单位同金额的多个流程必须选择，旧OA恰在候选中也不自动取用',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','测试供应商乙',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/多个流程/);assert.equal(q.choices.length,3);assert.equal(r.normalRows[0][21].v,'');
 const chosen=q.choices.find(c=>c.label.includes('MDST-DGFK20261001003'));r=C.build(i,{[q.id]:[chosen.value]});assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001003');assert.equal(r.normalRows[0][3].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][17].v,'出纳');
});
test('先校验同流程全部展开行，不能过滤掉收款单位或金额冲突后自动匹配',()=>{
 for(const [payee,amount] of [['其他供应商',100],['测试供应商甲',101]]){
  const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇',payee,amount]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:10月谈判付款主表!2');assert(q);assert.match(q.message,/存在冲突/);assert.equal(q.choices.length,2);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('流程编号')].v,'');
  r=C.build(i,{[q.id]:[q.choices[1].value]});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('流程编号')].v,'MDST-DGFK20261001001');
 }
});
test('同一流程信息冲突时不能同时选两个版本，保留原值也不会覆盖已有流程编号',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0,n=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(n,[['流程编号'],['旧流程记录']],{origin:'N1'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');r=C.build(i,{[q.id]:q.choices.filter(c=>c.value.startsWith('match:')).map(c=>c.value)});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(update(r,'N2'),undefined);r=C.build(i,{[q.id]:['keep']});assert.equal(r.negotiated.rows[0][13].v,'旧流程记录');assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);
});
test('OA金额空白、错误、负数或超过分精度不参与自动匹配，明确零可以匹配',()=>{
 for(const amount of ['', '/', 0.001, -100, {t:'n',f:'100+0'}, {t:'e',v:23}]){
  const i=fixture();set(i,'oa','付款金额',3,amount);const r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.equal(q.choices.length,1);assert.equal(r.normalRows[0][21].v,'');
 }
 const i=fixture();set(i,'ledger','申请付款金额',3,0);set(i,'ledger','已付款金额',3,0);set(i,'oa','付款金额',3,0);let r=C.build(i);assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002');assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);set(i,'oa','付款金额',3,'');r=C.build(i);assert(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'));
});
test('OA匹配不累加多个流程金额，缺少匹配字段明确提示不回退原OA',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',3,'MDST-DGFK20261001001、MDST-DGFK20261001002');set(i,'oa','付款金额',2,50);set(i,'oa','付款金额',3,50);set(i,'oa','收款单位',2,'测试供应商乙');let r=C.build(i);assert(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'));assert.equal(r.normalRows[0][21].v,'');
 const j=fixture();j.oa.book=book('Sheet0',OH.slice(0,5),[{'流程编号':'MDST-DGFK20261001002','当前节点':'出纳'}]);r=C.build(j);assert(r.issues.some(x=>x.kind==='missing-header'&&x.message.includes('收款单位、付款金额')));assert.equal(r.normalRows[0][21].v,'');assert.equal(r.normalRows[0][17].v,'');
});
test('OA进度错误或无缓存公式不得转成23或空白覆盖，重复去重不能掩盖错误',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const i=fixture();set(i,'oa','当前节点',3,bad);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/当前节点/);assert.equal(q.choices.length,1);assert.equal(r.normalRows[0][17].v,'');assert.equal(r.normalRows[0][21].v,'');r=C.build(i,{[q.id]:['keep']});assert.equal(r.normalRows[0][17].v,'');
 }
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;set(i,'oa','当前节点',3,{t:'e',v:23,w:'#REF!'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','23','测试人员','未归档','李奇','测试供应商乙',100]],{origin:-1});const r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/#REF!/);assert.equal(q.choices.length,2);assert.equal(r.normalRows[0][17].v,'');
});
