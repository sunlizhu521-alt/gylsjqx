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
 const appRows=[main,normal].map(r=>({'申请日期':'2026-10-10','供应商全称':r['供应商名称'],'付款主体':r['付款主体'],'合同号':r['对公付款0A流程号'],'合同金额':r['合同总金额'],'已付金额':0,'付款事由':r['摘要'],'备注':r['单据号【财务应付单】'],'本次申请付款金额':r['申请付款金额']}));
 const oaRows=[main,normal].map(r=>({'流程编号':r['对公付款0A流程号'],'当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇','收款单位':r['供应商名称'],'付款金额':r['申请付款金额']}));
 return {ledger:{book:ledger,name:'资金.xlsx'},negotiated:{book:negotiated,name:'谈判.xlsx'},applications:{book:book('付款申请汇总',AH,appRows),name:'申请.xlsx'},oa:{book:book('Sheet0',OH,oaRows),name:'OA.xlsx'}};
}
function set(input,role,header,row,value,sheet){const b=input[role].book,s=b.Sheets[sheet||b.SheetNames[0]],h=X.utils.sheet_to_json(s,{header:1})[0],c=h.indexOf(header);assert(c>=0);s[X.utils.encode_cell({r:row-1,c})]=typeof value==='object'&&value!==null?value:{t:typeof value==='number'?'n':'s',v:value};}
const update=(r,address)=>r.negotiated.updates.find(u=>u.address===address)?.cell;
function resolveKeep(input){let d={};for(let i=0;i<4;i++){const r=C.build(input,d);if(!r.stats.pending)return r;for(const issue of r.issues)if(!issue.resolved)d[issue.id]=['keep'];}return C.build(input,d);}
test('正常付款申请状态旁增加时间汇总，精确名称金额匹配并保留每笔来源',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[2];
 const earlier=row.slice();earlier[0]='2026-10-08';earlier[3]='其他合同编号';
 const partial=row.slice();partial[8]=50;
 const other=row.slice();other[1]='其他供应商';
 X.utils.sheet_add_aoa(s,[earlier,earlier,partial,other],{origin:-1});
 const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.normalHeaders.slice(16,19),['是否提交付款申请单','提交时间汇总','当前节点']);
 assert.equal(r.normalRows[0][16].v,'已提交付款申请单');
 assert.equal(r.normalRows[0][17].v,'2026-10-08-100.00\n2026-10-08-100.00\n2026-10-10-100.00');
 assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('正常付款未申请、金额或名称无法匹配时汇总为空，不沿用旧摘要',()=>{
 for(const [header,v] of [['供应商全称','其他供应商'],['供应商全称',''],['本次申请付款金额',99],['本次申请付款金额',''],['本次申请付款金额','异常'],['本次申请付款金额',0.001]]){
  const i=fixture();set(i,'applications',header,3,v);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['提交时间汇总'],[],['旧摘要']],{origin:'R1'});
  const r=C.build(i);assert.equal(r.normalRows[0][16].v,'未申请');assert.equal(r.normalRows[0][17].v,'');assert.equal(r.reviewRows.length,0);
 }
});
test('正常付款汇总保留零元及日期空白或异常提示，1904日期正确转换',()=>{
 for(const [raw,expected] of [['','日期为空'],['日期待定','日期待定'],['2026-02-30','2026-02-30'],[44834,'2026-10-01']]){
  const i=fixture();set(i,'ledger','申请付款金额',3,0);set(i,'ledger','已付款金额',3,0);set(i,'applications','本次申请付款金额',3,0);set(i,'applications','申请日期',3,raw);
  if(typeof raw==='number')i.applications.book.Workbook={WBProps:{date1904:true}};
  const r=C.build(i);assert.equal(r.normalRows[0][16].v,'已提交付款申请单');assert.equal(r.normalRows[0][17].v,expected+'-0.00');
 }
});
test('五张台账、谈判付款以自身为准、正常付款字段顺序、OA状态与不修改源表',()=>{
 const input=fixture(),before=JSON.stringify(input),r=C.build(input);assert.equal(r.stats.pending,0);assert.equal(r.stats.normalRows,1);assert.equal(r.stats.negotiatedRows,1);assert.deepEqual(r.normalHeaders,C.NORMAL_HEADERS);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);assert(!r.negotiated.updates.some(u=>['E2','J2','K2','L2','M2'].includes(u.address)));assert.equal(r.normalRows[0][14].f,'H2-M2');assert.equal(r.normalRows[0][14].v,0);assert.equal(r.normalRows[0][13].v,'是');assert.equal(r.normalRows[0][21].v,'李奇');assert.equal(r.normalRows[0][0].t,'n');assert.equal(r.normalRows[0][4].z,C.DATE_FORMAT);assert.equal(JSON.stringify(input),before);assert.equal(r.negotiated.rows[0][12].f,'G2');
});
test('谈判工作表名称包含匹配：零、一、多张及选择合法性',()=>{
 const i=fixture();assert.deepEqual(C.negotiatedSheets(i.negotiated.book),['10月谈判付款主表']);X.utils.book_append_sheet(i.negotiated.book,X.utils.aoa_to_sheet([NH]),'谈判付款-反馈');assert.throws(()=>C.build(i),/请选择/);i.negotiated.sheetName='10月谈判付款主表';assert.doesNotThrow(()=>C.build(i));i.negotiated.sheetName='不存在';assert.throws(()=>C.build(i),/请选择/);i.negotiated.book=book('普通表',NH,[]);assert.throws(()=>C.build(i),/未找到名称/);
});
test('同一OA多行一致时折叠，不将归档当成付款或申请提交',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇','测试供应商甲','￥100.00']],{origin:-1});set(i,'oa','流程状态',3,'归档');set(i,'ledger','已付款金额',3,0);set(i,'ledger','是否结清',3,'');const r=C.build(i);assert.equal(r.stats.oaFlows,2);assert.equal(r.normalRows[0][12].v,0);assert.equal(r.normalRows[0][13].v,'否');assert.equal(r.normalRows[0][20].v,'归档');assert.equal(r.stats.pending,0);assert.equal(r.stats.oaConsistent,2);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');
});
test('正常付款空白实付直接保留原情况，零实付正常计算，斜线不转零',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,'');let r=C.build(i);assert.equal(r.normalRows[0][12].v,'');assert.equal(r.normalRows[0][13].v,'是');assert.equal(r.normalRows[0][14].v,0);assert.equal(r.issues.some(x=>x.kind==='amount'),false);assert.equal(resolveKeep(i).stats.pending,0);
 set(i,'ledger','已付款金额',3,0);set(i,'ledger','是否结清',3,'');r=C.build(i);assert.equal(r.normalRows[0][12].v,0);assert.equal(r.normalRows[0][14].v,100);set(i,'ledger','已付款金额',3,'/');r=C.build(i);assert.equal(r.normalRows[0][12].v,'/');assert(r.issues.some(x=>x.kind==='amount'));
});
test('明确金额按分计算且不静默截断第三位或越界金额',()=>{
 assert.equal(C.money({v:'￥1,000.20'}).cents,100020n);assert.equal(C.money({v:0.1+0.2}).cents,30n);assert(C.money({v:0.001}).invalid);assert(C.money({v:'90071992547409.92'}).invalid);assert(C.money({v:'待核对'}).invalid);
 const i=fixture();set(i,'ledger','申请付款金额',3,0.3);set(i,'ledger','已付款金额',3,0.1);set(i,'ledger','是否结清',3,'');const r=C.build(i);assert.equal(r.normalRows[0][14].v,0.2);
});
test('正常付款结清与金额冲突需要明确选择，负余额不得自动更新',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,20);let r=C.build(i),q=r.issues.find(x=>x.kind==='settlement');assert(q);assert.equal(r.normalRows[0][14].v,0);r=C.build(i,{[q.id]:['recalculate']});assert.equal(r.normalRows[0][14].v,80);assert.equal(r.normalRows[0][13].v,'否');set(i,'ledger','已付款金额',3,101);r=C.build(i);assert.equal(r.normalRows[0][14].v,0);assert(r.issues.some(x=>x.kind==='amount'));
});
test('付款申请多次日期全部保留，不使用申请表已付覆盖资金台账',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;const row=X.utils.sheet_to_json(s,{header:1})[1];row[0]='2026-10-08';row[5]=90;X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const r=C.build(i);assert.equal(r.stats.pending,0);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);const dateCol=r.negotiated.headers.indexOf('提交付款申请单日期');assert.equal(r.negotiated.rows[0][dateCol].v,'2026-10-08\n2026-10-10');
});
test('无对应付款OA且单次金额不一致时不累加同供应商申请，旧人工选择不覆盖',()=>{
 const i=fixture();set(i,'applications','合同号',2,'MDST-DGFK20990101999');set(i,'applications','本次申请付款金额',2,50);const s=i.applications.book.Sheets.付款申请汇总,row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const r=C.build(i,{'application:10月谈判付款主表!2':['match:付款申请汇总!2']});assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未申请');assert.equal(r.reviewRows.length,0);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);
});
test('无付款申请记录时两类付款均自动未申请，忽略旧申请确认选择',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i,{'application:10月谈判付款主表!2':['keep','not-submitted']});assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未申请');assert.equal(r.normalRows[0][16].v,'未申请');assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('提交付款申请单日期')].v,'');
});
test('谈判重复记录逐条完整保留，不从资金台账重复分配付款金额',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const r=C.build(i,{'payment-match:10月谈判付款主表!2':['match:对账明细-李奇!2']});assert.equal(r.negotiated.rows.length,2);assert.deepEqual(r.negotiated.rows.map(row=>[row[4].v,row[9].v,row[10].v]),[[0,'否',100],[0,'否',100]]);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'E3'),undefined);
});
test('旧分类是或空但名称金额未命中时纳入正常付款，已结清也保留',()=>{
 for(const flag of ['是','', '待定']){const i=fixture();set(i,'ledger','是否谈判付款',3,flag);const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][13].v,'是');assert.equal(r.normalRows[0][15].v,'否');}
});
test('OA重复冲突与多个编号保留每个编号对应状态',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');assert(q);const candidate=q.choices.find(c=>c.label.includes('出纳'));r=C.build(i,{[q.id]:[candidate.value]});assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('当前节点')].v,'出纳');
 const j=fixture();set(j,'ledger','对公付款0A流程号',3,'MDST-DGFK20261001001、MDST-DGFK20261001002');set(j,'oa','收款单位',2,'测试供应商乙');r=C.build(j);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][18].v,'MDST-DGFK20261001001：财务审核\nMDST-DGFK20261001002：财务审核');assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001001\nMDST-DGFK20261001002');
});
test('正常付款公式仅使用有效缓存，缺缓存保留错误并阻止未核对导出',()=>{
 const i=fixture();set(i,'ledger','已付款金额',3,{t:'n',f:'10+10',v:20});set(i,'ledger','是否结清',3,'');let r=C.build(i);assert.equal(r.normalRows[0][12].v,20);assert.equal(r.normalRows[0][12].f,undefined);set(i,'ledger','已付款金额',3,{t:'n',f:'10+10'});r=C.build(i);assert.equal(r.normalRows[0][12].v,'公式结果不可用');assert(r.stats.pending>0);set(i,'ledger','OA合同编号',3,{t:'e',v:23,w:'#REF!'});r=C.build(i);assert.equal(r.normalRows[0][10].v,'#REF!');assert(r.issues.some(x=>x.kind==='formula'));
});
test('已有追加字段不重复新增，源表空行和非授权字段保持完整',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号']],{origin:'N1'});s['!ref']='A1:T4';s.L4={t:'s',v:'尾注'};const r=C.build(i);assert.equal(r.negotiated.headers.length,21);assert.equal(r.negotiated.headers.at(-1),'提交时间汇总');assert.equal(r.negotiated.rows.length,3);assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);assert.equal(r.negotiated.rows[2][11].v,'尾注');assert(!r.negotiated.updates.some(u=>u.address==='L4'));
});
test('1904来源日期转换为1900体系，原表不变',()=>{
 const i=fixture();i.ledger.book.Workbook={WBProps:{date1904:true}};set(i,'ledger','开票时间',3,{t:'n',v:44834,z:'yyyy-mm-dd'});const r=C.build(i);assert.equal(r.normalRows[0][4].v,44834+1462);assert.equal(i.ledger.book.Sheets['对账明细-李奇'].E3.v,44834);
});
test('登记月份和应付款时间无法识别时完整保留原文，不产生日期核对阻断',()=>{
 for(const [month,due] of [['10月登记',' 月底前付款\n以通知为准 '],['2026-13','2026-02-30'],['/','待定'],['','']]){
  const i=fixture();set(i,'ledger','月份',3,month);set(i,'ledger','应付款时间',3,due);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.normalRows[0][0].v,month);assert.equal(r.normalRows[0][8].v,due);assert.equal(r.normalRows[0][0].t,'s');assert.equal(r.normalRows[0][8].t,'s');assert.equal(r.issues.some(x=>x.kind==='date'),false);assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('登记月份和应付款时间有效时仍导出真实Excel日期及相应格式',()=>{
 const i=fixture();set(i,'ledger','月份',3,'2026-10');set(i,'ledger','应付款时间',3,'2026年10月31日');const r=C.build(i),month=r.normalRows[0][0],due=r.normalRows[0][8];assert.equal(month.t,'n');assert.equal(X.SSF.format('yyyy-mm-dd',month.v),'2026-10-01');assert.equal(month.z,'yyyy"年"m"月"');assert.equal(due.t,'n');assert.equal(X.SSF.format('yyyy-mm-dd',due.v),'2026-10-31');assert.equal(due.z,C.DATE_FORMAT);assert.equal(r.stats.pending,0);
});
test('开票时间无法识别仍保留原文并要求日期核对',()=>{
 const i=fixture();set(i,'ledger','月份',3,'10月登记');set(i,'ledger','应付款时间',3,'月末付款');set(i,'ledger','开票时间',3,'开票待确认');const r=C.build(i),issues=r.issues.filter(x=>x.kind==='date');assert.equal(r.normalRows[0][4].v,'开票待确认');assert.equal(issues.length,1);assert.equal(issues[0].id,'date:对账明细-李奇!3:开票时间');assert.equal(issues[0].resolved,false);assert.equal(r.stats.pending,1);
});
test('OA只读取指定字段，其他重复明细列不构成冲突',()=>{
 const i=fixture();i.oa.book=book('Sheet0',[...OH,'备注','备注'],[{'流程编号':'MDST-DGFK20261001001','当前节点':'财务审核','未操作者':'测试人员','流程状态':'未归档','申请人':'李奇'}]);assert.doesNotThrow(()=>C.build(i));
});
test('名称金额未命中时仍保留谈判原付款，旧来源选择不能覆盖原值',()=>{
 const i=fixture();set(i,'ledger','申请付款金额',2,101);const r=C.build(i,{'payment-match:10月谈判付款主表!2':['match:对账明细-李奇!2']});assert.equal(r.normalRows.length,2);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][10].v,100);
});
test('谈判原付款状态与金额即使不一致也原样保留，不受资金台账覆盖',()=>{
 const i=fixture();set(i,'ledger','是否结清',2,'');set(i,'negotiated','是否结清',2,'是');set(i,'negotiated','剩余付款金额',2,{t:'n',v:99,f:'G2-E2-1'});const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='settlement'),false);assert.equal(update(r,'E2'),undefined);assert.equal(update(r,'J2'),undefined);assert.equal(update(r,'K2'),undefined);assert.equal(r.negotiated.rows[0][9].v,'是');assert.deepEqual(r.negotiated.rows[0][10],i.negotiated.book.Sheets['10月谈判付款主表'].K2);
});
test('资金台账与谈判表财务单号冲突不能补空，申请状态不依赖财务单号',()=>{
 const i=fixture();set(i,'negotiated','单据号【财务应付单】',2,'AP999999');set(i,'negotiated','对公付款0A流程号',2,'');const r=C.build(i);assert.equal(update(r,'C2'),undefined);assert.equal(update(r,'E2'),undefined);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);
 const j=fixture();set(j,'applications','备注',2,'AP999999');assert.equal(C.build(j).issues.some(x=>x.kind==='application'),false);assert.equal(C.build(j).negotiated.rows[0][C.build(j).negotiated.headers.indexOf('是否提交付款申请单')].v,'已全部提交');
});
test('供应商或主体为空直接保留原值，不产生身份核对',()=>{
 for(const field of ['供应商名称','付款主体']){const i=fixture();set(i,'ledger',field,3,'');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='classification'||x.kind==='identity'),false);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][field==='供应商名称'?1:2].v,'');}
 const i=fixture();set(i,'negotiated','供应商',2,'');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind==='identity'),false);assert.equal(r.negotiated.rows[0][0].v,'');assert.equal(r.negotiated.rows[0][1].v,'');
});
test('旧分类否或空但名称金额命中谈判表时只保留谈判原表，不混入正常付款',()=>{
 for(const flag of ['否','']){const i=fixture();set(i,'ledger','是否谈判付款',2,flag);const before=JSON.stringify(i);const r=C.build(i);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][1].v,'测试供应商乙');assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(update(r,'E2'),undefined);assert.equal(JSON.stringify(i),before);}
});
test('同名不同金额分别分类，同名同金额跨主体按出现性归类且谈判付款原样保留',()=>{
 const i=fixture();set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);let r=C.build(i);assert.equal(r.normalRows.length,1);assert.equal(r.normalRows[0][7].v,101);
 set(i,'ledger','申请付款金额',3,100);set(i,'ledger','付款主体',3,'国源');r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.negotiatedLedgerRows,2);assert.equal(r.stats.negotiatedRows,1);assert.equal(update(r,'E2'),undefined);
});
test('分类按分精确比较，规范空白、全角名称和金额表示，零值有效而空金额不算零',()=>{
 const i=fixture();set(i,'ledger','供应商名称',2,' 测试供应商Ａ ');set(i,'negotiated','供应商',2,'测试供应商A');set(i,'ledger','申请付款金额',2,'￥1,000.20');set(i,'negotiated','申请付款金额',2,1000.2);let r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.normalRows.length,1);
 set(i,'ledger','申请付款金额',2,0);set(i,'negotiated','申请付款金额',2,0);r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);set(i,'ledger','申请付款金额',2,'');r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,0);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.normalRows.length,2);assert.equal(r.normalRows[0][7].v,'');assert.equal(r.issues.some(x=>x.id==='classification:对账明细-李奇!2'),false);
});
test('空白及异常金额直接纳入正常付款，保留金额并忽略旧分类决策',()=>{
 for(const amount of ['', '/', -1, '0.001', {t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'ledger','申请付款金额',3,amount);const before=JSON.stringify(i),r=C.build(i,{'classification:对账明细-李奇!3':['keep']});assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(r.normalRows.length,1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.normalRows[0][7].v,typeof amount==='object'?amount.w||'公式结果不可用':amount);assert.equal(r.normalRows[0][15].v,'否');assert.equal(r.negotiated.rows.length,1);assert.equal(JSON.stringify(i),before);
 }
});
test('供应商为空白或错误时自动正常付款，不删除谈判原记录',()=>{
 for(const supplier of ['', '　\n', {t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'ledger','供应商名称',2,supplier);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.stats.negotiatedRows,1);assert.equal(r.stats.negotiatedLedgerRows,0);assert.equal(r.normalRows.length,2);assert.equal(r.normalRows[0][7].v,100);assert.equal(r.normalRows[0][15].v,'否');assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.issues.some(x=>x.kind==='classification'),false);assert.equal(JSON.stringify(i),before);
 }
});
test('正常付款仅保留原OA有值的记录，匹配流程及旧对账值不能替代原OA',()=>{
 const cases=[['',false],['　\n',false],[{t:'s',f:'A99',v:''},false],[{t:'n',f:'A99'},false],[{t:'e',v:23,w:'#REF!'},false],['MDST-DGFK20261001002',true],[{t:'s',f:'A99',v:'MDST-DGFK20261001002'},true],['MDST-DGFK20261001002、MDST-DGFK20261001003',true]];
 for(const [oa,included] of cases){
  const i=fixture();set(i,'ledger','对公付款0A流程号',3,oa);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['是否对账'],['已对账'],[included?'未对账':'已对账']],{origin:'R1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.normalHeaders.at(-1),'是否对账');assert.equal(r.normalRows.length,included?1:0);assert.equal(r.stats.excludedNormalRows,included?0:1);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,0);if(included){assert.equal(r.normalRows[0].at(-1).v,'已对账');assert.equal(r.normalRows[0][3].v,typeof oa==='object'?oa.v:oa);}assert.equal(r.normalHeaders.filter(h=>h==='是否对账').length,1);assert.equal(r.negotiated.headers.includes('是否对账'),false);assert.equal(JSON.stringify(i),before);
 }
 const i=fixture();delete i.ledger.book.Sheets['对账明细-李奇'].D3;assert.equal(C.build(i).normalRows.length,0);delete i.ledger.book.Sheets['对账明细-李奇'].D1;const r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.negotiatedRows,1);assert(r.issues.some(q=>q.kind==='missing-header'&&q.message.includes('对公付款0A流程号')));
});
test('无OA正常记录不进入金额日期及申请核对，保留后续顺序并重排公式行号',()=>{
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[2];X.utils.sheet_add_aoa(s,[row,row],{origin:-1});set(i,'ledger','供应商名称',4,'保留供应商丙');set(i,'ledger','供应商名称',5,'保留供应商丁');set(i,'ledger','对公付款0A流程号',3,'');set(i,'ledger','申请付款金额',3,'异常金额');set(i,'ledger','开票时间',3,'异常日期');i.applications.book=book('付款申请汇总',AH,[]);const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.normalRows.map(row=>row[1].v),['保留供应商丙','保留供应商丁']);assert.deepEqual(r.normalRows.map(row=>row[14].f),['H2-M2','H3-M3']);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.filter(row=>row[0].v==='正常付款').length,0);assert.equal(r.reviewRows.some(row=>row[1].v==='测试供应商乙'),false);assert.equal(JSON.stringify(i),before);
});
test('未带原OA的记录不触发分类待核对，谈判原表和谈判匹配记录仍完整保留',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];row[6]='';X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);for(const r of [2,3])set(i,'ledger','对公付款0A流程号',r,'');set(i,'negotiated','对公付款0A流程号',2,'');const before=JSON.stringify(i),r=C.build(i);
 assert.equal(r.normalRows.length,0);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.stats.negotiatedRows,2);assert.equal(r.negotiated.rows.length,2);assert.equal(r.stats.unclassifiedLedgerRows,0);assert.equal(r.issues.some(q=>q.kind==='classification'),false);assert.equal(JSON.stringify(i),before);
});
test('谈判表同供应商金额缺失时未命中的款项需核对，已命中的款项仍可确定分类',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],row=X.utils.sheet_to_json(s,{header:1})[1];row[6]='';X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'ledger','供应商名称',3,'测试供应商甲');set(i,'ledger','申请付款金额',3,101);let r=C.build(i);assert.equal(r.stats.negotiatedLedgerRows,1);assert.equal(r.normalRows.length,0);const q=r.issues.find(x=>x.id==='classification:对账明细-李奇!3');assert.match(q.message,/谈判来源/);assert.match(q.message,/第 3 行/);r=C.build(i,{[q.id]:['normal']});assert.equal(r.normalRows.length,1);
});
test('重复明细保持原顺序，未选中的谈判反馈表不参与分类，缺失旧分类列不提示',()=>{
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[2];X.utils.sheet_add_aoa(s,[row],{origin:-1});delete s.P1;
 const feedback=book('谈判付款-反馈',NH,[{'供应商':'测试供应商乙','申请付款金额':100}]);X.utils.book_append_sheet(i.negotiated.book,feedback.Sheets['谈判付款-反馈'],'谈判付款-反馈');i.negotiated.sheetName='10月谈判付款主表';const r=C.build(i);assert.equal(r.normalRows.length,2);assert.deepEqual(r.normalRows.map(row=>row[1].v),['测试供应商乙','测试供应商乙']);assert.equal(r.issues.some(x=>x.kind==='missing-header'&&x.message.includes('是否谈判付款')),false);assert.equal(r.stats.negotiatedRows,1);
});
test('同AP同金额但已有OA冲突时不补空，不产生付款来源核对',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261024999');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(r.negotiated.rows[0][2].v,'MDST-DGFK20261001001');assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);assert.equal(update(r,'E2'),undefined);
});
test('财务单号部分相交不能补空，集合相同顺序不同可以补空',()=>{
 const i=fixture();set(i,'ledger','单据号【财务应付单】',2,'AP000001、AP000003');set(i,'negotiated','单据号【财务应付单】',2,'AP000001、AP000004');set(i,'negotiated','付款主体',2,'');let r=C.build(i);assert.equal(update(r,'B2'),undefined);
 set(i,'negotiated','单据号【财务应付单】',2,'AP000003\nAP000001');r=C.build(i);assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'E2'),undefined);
 const j=fixture();set(j,'applications','备注',2,'AP000001、AP000099');r=C.build(j);assert.equal(r.issues.some(x=>x.kind==='application'),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'已全部提交');
});
test('付款OA组部分相交不能补空且不能改变已付金额',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',2,'MDST-DGFK20261001001、MDST-DGFK20261024999');set(i,'negotiated','付款主体',2,'');const r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'E2'),undefined);
});
test('谈判仅供应商金额一致不能绕过OA号，正常付款仍按供应商金额匹配',()=>{
 const i=fixture();X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['OA合同编号'],['MDST-HT20260101']],{origin:'N1'});set(i,'applications','合同号',2,'MDST-HT20990101');set(i,'applications','付款主体',2,'国源');set(i,'applications','付款事由',2,'其他事由');set(i,'applications','备注',2,'AP999999');const r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'未申请');
 const j=fixture();set(j,'ledger','OA合同编号',3,'MDST-HT20990101');set(j,'ledger','付款主体',3,'国源');assert.equal(C.build(j).normalRows[0][16].v,'已提交付款申请单');
});
function historyFixture(value){const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(s,[['提交付款申请单日期','是否提交付款申请单'],[value,'是']],{origin:'N1'});return i;}
test('已有提交日期与本次日期合并去重排序，反复处理不丢历史',()=>{
 const i=historyFixture('2026-10-09\n2026-10-01、2026-10-10');let r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10');
 set(i,'negotiated','提交付款申请单日期',2,update(r,'N2'));set(i,'applications','申请日期',2,'2026-10-12');r=C.build(i);assert.equal(update(r,'N2').v,'2026-10-01\n2026-10-09\n2026-10-10\n2026-10-12');
});
test('历史提交日期不能解析时保留原日期并列核对，不阻止已提交状态',()=>{
 const i=historyFixture('10月初提交过'),r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].v,'10月初提交过');assert.equal(update(r,'O2').v,'已全部提交');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.reviewRows[0][4].v,'已提交付款申请单');assert.equal(r.stats.pending,0);
});
test('未匹配申请时历史日期保留，提交状态重新计算为未申请',()=>{
 const i=historyFixture('2026-09-10、2026-09-12');i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i);assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2').v,'未申请');assert.equal(r.negotiated.rows[0][13].v,'2026-09-10、2026-09-12');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);
});
test('历史提交日期为Excel日期数值时正确转换后合并',()=>{
 const i=historyFixture(46294);const r=C.build(i);assert.equal(update(r,'N2').v,[X.SSF.format('yyyy-mm-dd',46294),'2026-10-10'].sort().join('\n'));
});
test('历史提交日期公式缺缓存时原样保留并列核对，匹配仍为已提交',()=>{
 const i=historyFixture('2026-09-10');set(i,'negotiated','提交付款申请单日期',2,{t:'n',f:'DATE(2026,9,10)'});const r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].f,'DATE(2026,9,10)');assert.equal(update(r,'O2').v,'已全部提交');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);
});
test('正常付款原OA有值但填写错误时按供应商金额自动匹配，保留原OA和是否对账',()=>{
 for(const old of ['待核对', 'MDST-DGFK20990101001']){
  const i=fixture();set(i,'ledger','对公付款0A流程号',3,old);const before=JSON.stringify(i),r=C.build(i,{'oa:对账明细-李奇!3':['keep']});assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);
  assert.equal(r.normalRows[0][3].v,old);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][12].v,100);assert.equal(r.normalRows[0][18].v,'财务审核');assert.equal(r.normalRows[0][19].v,'测试人员');assert.equal(r.normalRows[0][20].v,'未归档');assert.equal(r.normalRows[0][21].v,'李奇');assert.equal(r.normalRows[0].at(-1).v,old?'已对账':'未对账');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判原OA空白错误或未找到时保留原流程信息，不按供应商金额回退匹配',()=>{
 for(const old of ['', '　\n', 'MDST-DGFK20990101001', {t:'e',v:23,w:'#REF!'}, {t:'s',f:'A99'}]){
  const i=fixture();set(i,'negotiated','对公付款0A流程号',2,old);set(i,'ledger','对公付款0A流程号',2,old);const headers=['流程编号','当前节点','未操作者','流程状态','申请人'],values=['原流程','原节点','原人员','原状态','原申请人'];X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),values);assert.deepEqual([r.negotiated.rows[0][4].v,r.negotiated.rows[0][9].v,r.negotiated.rows[0][10].v],[0,'否',100]);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判按付款OA号匹配，不以收款单位或付款金额作为条件',()=>{
 for(const [field,value] of [['收款单位','其他供应商'],['付款金额',20],['付款金额',''],['付款金额',{t:'e',v:23,w:'#REF!'}],['收款单位',{t:'s',f:'A99'}]]){
  const i=fixture();set(i,'oa',field,2,value);const r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'财务审核');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'已付金额'),0);
 }
});
test('无精确OA匹配时两类付款原流程编号及全部进度信息自动保留，旧决定不影响',()=>{
 const headers=['当前节点','未操作者','流程状态','申请人','流程编号'];
 for(const values of [[' 原节点\n复核中 ','原未操作者','原流程状态','原申请人','原流程编号'],['','','','','']]){
  const i=fixture();set(i,'oa','流程编号',2,'MDST-DGFK20990101999');set(i,'oa','付款金额',3,999);X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[headers,[],values],{origin:'R1'});X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});const before=JSON.stringify(i);
  for(const answer of [[],['keep'],['match:Sheet0!2','match:Sheet0!3']]){
   const r=C.build(i,{'oa:对账明细-李奇!3':answer,'oa:10月谈判付款主表!2':answer});assert.deepEqual(r.normalRows[0].slice(18,23).map(c=>c.v),values);assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),values);assert.equal(r.issues.some(q=>q.kind==='oa'),false);assert.equal(r.stats.pending,0);assert.equal(r.stats.oaMatched,0);assert.equal(r.normalRows[0][3].v,'MDST-DGFK20261001002');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20261001001');assert.equal(r.normalRows[0][12].v,100);assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(JSON.stringify(i),before);
  }
 }
});
test('名称金额不可用时正常付款保留原流程，谈判付款仍按原OA匹配',()=>{
 for(const [role,field,value] of [['ledger','供应商名称',''],['ledger','供应商名称',{t:'e',v:23,w:'#REF!'}],['ledger','申请付款金额',''],['ledger','申请付款金额','金额待确认'],['negotiated','供应商',''],['negotiated','申请付款金额','']]){
  const i=fixture(),headers=['当前节点','未操作者','流程状态','申请人','流程编号'],values=['原节点','原未操作者','原状态','原申请人','原流程'];X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[headers,[],values],{origin:'R1'});X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[headers,values],{origin:'N1'});set(i,role,field,role==='ledger'?3:2,value);const before=JSON.stringify(i),r=C.build(i),targetId=role==='ledger'?'对账明细-李奇!3':'10月谈判付款主表!2';assert.equal(r.issues.some(q=>q.id==='oa:'+targetId),false);if(role==='ledger')assert.deepEqual(r.normalRows.find(row=>row[3].v==='MDST-DGFK20261001002').slice(18,23).map(c=>c.v),values);else assert.deepEqual(headers.map(h=>negotiatedValue(r,h)),['财务审核','测试人员','未归档','李奇','MDST-DGFK20261001001']);assert.equal(JSON.stringify(i),before);
 }
});
test('同供应商同金额不同流程自动全部写入，重复展开行不重复写',()=>{
 for(const old of ['MDST-DGFK20990101999', 'MDST-DGFK20261001002']){
  const i=fixture(),s=i.oa.book.Sheets.Sheet0,row=['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','测试供应商乙',100];set(i,'ledger','对公付款0A流程号',3,old);X.utils.sheet_add_aoa(s,[row,row],{origin:-1});const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][3].v,old);assert.equal(r.normalRows[0][18].v,'MDST-DGFK20261001002：财务审核\nMDST-DGFK20261001003：出纳');assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001002：测试人员\nMDST-DGFK20261001003：人员乙');assert.equal(r.normalRows[0][20].v,'MDST-DGFK20261001002：未归档\nMDST-DGFK20261001003：归档');assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002：李奇\nMDST-DGFK20261001003：徐英');assert.equal(r.normalRows[0][12].v,100);assert.equal(r.normalRows[0].at(-1).v,old?'已对账':'未对账');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 }
});
test('正常付款同流程号的收款单位或付款金额冲突仍须核对',()=>{
 for(const [payee,amount] of [['其他供应商',100],['测试供应商乙',101]]){
  const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','财务审核','测试人员','未归档','李奇',payee,amount]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/存在冲突/);assert.equal(r.normalRows[0][22].v,'');r=C.build(i,{[q.id]:[q.choices.find(c=>c.value.startsWith('match:')).value]});assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');
 }
});
test('谈判同OA的明细收款单位金额不同但流程信息一致时只读一次',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','财务审核','测试人员','未归档','李奇','其他单位',20]],{origin:-1});const r=C.build(i);assert.equal(r.stats.pending,0);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'财务审核');assert.equal(negotiatedValue(r,'已付金额'),0);
});
test('谈判多个原OA逐号匹配，忽略同名同金额的其他流程并保留全部申请关联',()=>{
 const i=installmentFixture([{'本次申请付款金额':20},{'合同号':'MDST-DGFK20261001004','本次申请付款金额':30}]),ids='MDST-DGFK20261001003、MDST-DGFK20261001001、MDST-DGFK20261001004';set(i,'negotiated','对公付款0A流程号',2,ids);X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','其他单位',20],['MDST-DGFK20261001999','不应读取','其他人','未归档','其他人','测试供应商甲',100]],{origin:-1});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001003\nMDST-DGFK20261001001');assert.equal(negotiatedValue(r,'当前节点'),'MDST-DGFK20261001003：出纳\nMDST-DGFK20261001001：财务审核');assert.equal(negotiatedValue(r,'对公付款0A流程号'),ids);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-20.00\n2026-10-10-30.00');assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('谈判匹配OA的流程字段异常保留原值，不能用同名同金额的其他编号替代',()=>{
 const i=fixture();set(i,'oa','当前节点',2,{t:'e',v:23,w:'#REF!'});X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001999','出纳','其他人','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:10月谈判付款主表!2');assert(q);assert.match(q.message,/当前节点/);assert.equal(q.choices.length,1);assert.equal(negotiatedValue(r,'流程编号'),'');r=C.build(i,{[q.id]:['keep']});assert.equal(r.stats.pending,0);assert.equal(negotiatedValue(r,'流程编号'),'');
});
test('同一流程进度冲突不能同时选两版本，选择保留时原流程编号及进度都不覆盖',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0,n=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];X.utils.sheet_add_aoa(n,[['流程编号','当前节点'],['旧流程记录','原节点']],{origin:'N1'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001001','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.kind==='oa');r=C.build(i,{[q.id]:q.choices.filter(c=>c.value.startsWith('match:')).map(c=>c.value)});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(update(r,'N2'),undefined);assert.equal(update(r,'O2'),undefined);r=C.build(i,{[q.id]:['keep']});assert.equal(r.negotiated.rows[0][13].v,'旧流程记录');assert.equal(r.negotiated.rows[0][14].v,'原节点');assert.equal(r.negotiated.rows[0][2].v,'MDST-DGFK20261001001');assert.equal(r.negotiated.headers.filter(h=>h==='流程编号').length,1);
});
test('选择冲突流程版本后仍完整写入其他唯一匹配流程，不能选唯一流程绕过冲突',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001003','出纳','人员乙','归档','徐英','测试供应商乙',100],['MDST-DGFK20261001003','复核','人员丙','未归档','徐英','测试供应商乙',100]],{origin:-1});const before=JSON.stringify(i);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);
 r=C.build(i,{[q.id]:['match:Sheet0!3']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(r.normalRows[0][22].v,'');
 const selected=q.choices.find(c=>c.label.includes('MDST-DGFK20261001003')&&c.label.includes('出纳'));assert(selected);r=C.build(i,{[q.id]:[selected.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][18].v,'MDST-DGFK20261001002：财务审核\nMDST-DGFK20261001003：出纳');assert.equal(r.normalRows[0][19].v,'MDST-DGFK20261001002：测试人员\nMDST-DGFK20261001003：人员乙');assert.equal(r.normalRows[0][20].v,'MDST-DGFK20261001002：未归档\nMDST-DGFK20261001003：归档');assert.equal(r.normalRows[0][21].v,'MDST-DGFK20261001002：李奇\nMDST-DGFK20261001003：徐英');assert.equal(r.normalRows[0][3].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][12].v,100);assert.equal(JSON.stringify(i),before);
});
test('有多个冲突流程时每个流程均需选择一个版本，部分选择不能完成核对',()=>{
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','出纳','人员乙','归档','李奇','测试供应商乙',100],['MDST-DGFK20261001003','复核','人员丙','未归档','徐英','测试供应商乙',100],['MDST-DGFK20261001003','审批','人员丁','未归档','徐英','测试供应商乙',100]],{origin:-1});let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);const chosen2=q.choices.find(c=>c.label.includes('MDST-DGFK20261001002')&&c.label.includes('出纳')),chosen3=q.choices.find(c=>c.label.includes('MDST-DGFK20261001003')&&c.label.includes('复核'));assert(chosen2&&chosen3);
 r=C.build(i,{[q.id]:[chosen2.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,false);assert.equal(r.normalRows[0][22].v,'');r=C.build(i,{[q.id]:[chosen2.value,chosen3.value]});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002\nMDST-DGFK20261001003');assert.equal(r.normalRows[0][18].v,'MDST-DGFK20261001002：出纳\nMDST-DGFK20261001003：复核');
});
test('金额空白、错误、负数或超过分精度不参与自动匹配，明确零可以匹配',()=>{
 for(const amount of ['', '/', 0.001, -100, {t:'n',f:'100+0'}, {t:'e',v:23}]){
  const i=fixture();set(i,'oa','付款金额',3,amount);const r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][22].v,'');
 }
 const i=fixture();set(i,'ledger','申请付款金额',3,0);set(i,'ledger','已付款金额',3,0);set(i,'oa','付款金额',3,0);let r=C.build(i);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);set(i,'oa','付款金额',3,'');r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][22].v,'');
});
test('OA匹配不累加多个流程金额，缺少匹配字段明确提示且不回退原OA',()=>{
 const i=fixture();set(i,'oa','付款金额',2,50);set(i,'oa','付款金额',3,50);set(i,'oa','收款单位',2,'测试供应商乙');let r=C.build(i);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);assert.equal(r.normalRows[0][22].v,'');
 const j=fixture();j.oa.book=book('Sheet0',OH.slice(0,5),[{'流程编号':'MDST-DGFK20261001002','当前节点':'出纳'}]);r=C.build(j);assert.equal(r.normalRows[0][22].v,'');assert.equal(r.normalRows[0][18].v,'');assert(r.issues.some(x=>x.kind==='missing-header'&&x.message.includes('收款单位、付款金额')));
});
test('原OA中的多余编号不能混入名称金额精确匹配结果',()=>{
 const i=fixture(),old='MDST-DGFK20261001002、MDST-DGFK20990101999';set(i,'ledger','对公付款0A流程号',3,old);const r=C.build(i);assert.equal(r.normalRows[0][3].v,old);assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][18].v,'财务审核');assert.equal(r.normalRows[0][20].v,'未归档');assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);
});
test('旧OA不在源表时仍按名称金额匹配并更新流程信息，原OA不覆盖',()=>{
 const i=fixture();set(i,'ledger','对公付款0A流程号',3,'MDST-DGFK20260723032');X.utils.sheet_add_aoa(i.ledger.book.Sheets['对账明细-李奇'],[['当前节点','流程状态'],[],['原节点','原状态']],{origin:'R1'});const r=C.build(i);assert.equal(r.normalRows[0][3].v,'MDST-DGFK20260723032');assert.equal(r.normalRows[0][22].v,'MDST-DGFK20261001002');assert.equal(r.normalRows[0][18].v,'财务审核');assert.equal(r.normalRows[0][20].v,'未归档');
});
test('原OA为Excel错误或公式无缓存时排除正常付款，已匹配流程不能使其入选',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){const i=fixture();set(i,'ledger','对公付款0A流程号',3,bad);const r=C.build(i);assert.equal(r.normalRows.length,0);assert.equal(r.stats.excludedNormalRows,1);assert.equal(r.stats.oaMatched,1);assert.equal(r.issues.some(x=>x.id==='oa:对账明细-李奇!3'),false);}
});
test('OA进度错误或无缓存公式不得转成23或空白覆盖，重复去重不能掩盖错误',()=>{
 for(const bad of [{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const i=fixture();set(i,'oa','当前节点',3,bad);let r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/当前节点/);assert.equal(q.choices.length,1);assert.equal(r.normalRows[0][18].v,'');assert.equal(r.normalRows[0][22].v,'');r=C.build(i,{[q.id]:['keep']});assert.equal(r.normalRows[0][18].v,'');
 }
 const i=fixture(),s=i.oa.book.Sheets.Sheet0;set(i,'oa','当前节点',3,{t:'e',v:23,w:'#REF!'});X.utils.sheet_add_aoa(s,[['MDST-DGFK20261001002','23','测试人员','未归档','李奇','测试供应商乙',100]],{origin:-1});const r=C.build(i),q=r.issues.find(x=>x.id==='oa:对账明细-李奇!3');assert(q);assert.match(q.message,/#REF!/);assert.equal(q.choices.length,2);assert.equal(r.normalRows[0][18].v,'');
});

test('谈判唯一同供应商同金额来源只补空OA主体财务单号，不导入资金付款金额或其他字段',()=>{
 const i=fixture();for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,'');set(i,'ledger','摘要',2,'不应覆盖的台账事由');set(i,'ledger','合同总金额',2,99999);const before=JSON.stringify(i),r=C.build(i);
 assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'H2').v,'AP000001');assert.equal(r.negotiated.rows[0][3].v,1000);assert.equal(r.negotiated.rows[0][4].v,0);assert.equal(r.negotiated.rows[0][5].v,'九月货款');assert.equal(r.negotiated.rows[0][10].v,100);assert.equal(r.issues.some(x=>['payment-match','identity'].includes(x.kind)),false);assert.equal(r.issues.some(x=>x.kind==='application'&&x.source.sheet.includes('谈判')),false);assert.equal(r.issues.some(x=>x.id==='oa:10月谈判付款主表!2'),false);assert.equal(JSON.stringify(i),before);
});
test('谈判补空候选不唯一时保持空白，主体已有值可排除另一主体',()=>{
 const i=fixture(),s=i.ledger.book.Sheets['对账明细-李奇'],row=X.utils.sheet_to_json(s,{header:1})[1];for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,'');X.utils.sheet_add_aoa(s,[row],{origin:-1});let r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'C2'),undefined);assert.equal(update(r,'H2'),undefined);assert.equal(r.issues.some(x=>x.kind==='payment-match'),false);
 set(i,'ledger','付款主体',4,'国源');set(i,'negotiated','付款主体',2,'迈德斯特');r=C.build(i);assert.equal(update(r,'B2'),undefined);assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'H2').v,'AP000001');
});
test('谈判已有身份字段含文字、空值公式或错误时不能被覆盖',()=>{
 for(const old of ['手工保留编号',{t:'s',v:'',f:'A99'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();for(const h of ['付款主体','对公付款0A流程号','单据号【财务应付单】'])set(i,'negotiated',h,2,old);const before=JSON.stringify(i),r=C.build(i);for(const a of ['B2','C2','H2'])assert.equal(update(r,a),undefined);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判原已付、结清及剩余空白错误与公式全部保留，不产生付款情况核对',()=>{
 for(const original of ['', '/', {t:'n',v:5,f:'G2-95'}, {t:'n',f:'A99'}, {t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'negotiated','已付金额',2,original);set(i,'negotiated','剩余付款金额',2,original);set(i,'negotiated','是否结清',2,'人工确认');const source=i.negotiated.book.Sheets['10月谈判付款主表'],r=C.build(i);for(const [a,c] of [['E2',4],['J2',9],['K2',10]]){assert.equal(update(r,a),undefined);assert.deepEqual(r.negotiated.rows[0][c],source[a]);}assert.equal(r.issues.some(x=>['amount','settlement','payment-match'].includes(x.kind)&&x.source.sheet.includes('谈判')),false);
 }
});
test('谈判缺少补充字段列时在表尾添加，原列和原付款保持不变',()=>{
 const i=fixture(),removed=['付款主体','对公付款0A流程号','单据号【财务应付单】'],oldSheet=i.negotiated.book.Sheets['10月谈判付款主表'],values=X.utils.sheet_to_json(oldSheet,{header:1})[1],headers=NH.filter(h=>!removed.includes(h));i.negotiated.book=book('10月谈判付款主表',headers,[Object.fromEntries(NH.map((h,c)=>[h,values[c]]))]);const r=C.build(i);assert.deepEqual(r.negotiated.headers.slice(0,headers.length),headers);for(const [h,v] of [['付款主体','迈德斯特'],['对公付款0A流程号','MDST-DGFK20261001001'],['单据号【财务应付单】','AP000001']]){assert.equal(r.negotiated.headers.filter(x=>x===h).length,1);assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf(h)].v,v);}assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('已付金额')].v,0);
});
test('谈判补空不得忽略明确合同号冲突，也不得使用异常来源单元格',()=>{
 const i=fixture();set(i,'negotiated','付款主体',2,'');X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['OA合同编号'],['MDST-HT20990101']],{origin:'N1'});assert.equal(update(C.build(i),'B2'),undefined);
 for(const bad of [{t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const j=fixture();set(j,'negotiated','对公付款0A流程号',2,'');set(j,'ledger','对公付款0A流程号',2,bad);assert.equal(update(C.build(j),'C2'),undefined);}
});

test('未匹配付款申请仅标记未申请，不列入核对sheet，不改变两类付款金额',()=>{
 const i=fixture();set(i,'applications','合同号',2,'MDST-DGFK20990101999');for(const row of [2,3])set(i,'applications','本次申请付款金额',row,50);const before=JSON.stringify(i),r=C.build(i);assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,0);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未申请');assert.equal(r.normalRows[0][16].v,'未申请');assert.deepEqual([negotiatedValue(r,'已付金额'),negotiatedValue(r,'剩余付款金额')],[0,100]);assert.deepEqual([r.normalRows[0][12].v,r.normalRows[0][14].v],[100,0]);assert.equal(JSON.stringify(i),before);
 set(i,'applications','合同号',2,'MDST-DGFK20261001001');for(const row of [2,3])set(i,'applications','本次申请付款金额',row,100);const confirmed=C.build(i);assert.equal(confirmed.reviewRows.length,0);assert.equal(negotiatedValue(confirmed,'是否提交付款申请单'),'已全部提交');assert.equal(confirmed.normalRows[0][16].v,'已提交付款申请单');
});
test('无付款申请候选时两类付款重新标记未申请且不列核对，谈判原提交日期保留',()=>{
 const i=historyFixture('2026-09-10');i.applications.book=book('付款申请汇总',AH,[]);const r=C.build(i);assert.equal(r.reviewRows.length,0);assert.equal(r.normalRows[0][16].v,'未申请');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未申请');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-09-10');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,0);
});
test('多个谈判明细可接收同一申请时不重复分配，保留明细并列出归属核对',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});const a=i.applications.book.Sheets.付款申请汇总,arow=X.utils.sheet_to_json(a,{header:1})[1];X.utils.sheet_add_aoa(a,[arow],{origin:-1});const r=C.build(i);assert.equal(r.reviewRows.length,2);const col=r.negotiated.headers.indexOf('是否提交付款申请单');assert.deepEqual(r.negotiated.rows.map(row=>row[col].v),['待核对','待核对']);assert(r.reviewRows.every(row=>row[5].v.includes('未重复分配')));assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.stats.pending,0);assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('谈判原合并区域只补锚点，不在合并区域其他单元格写值',()=>{
 const i=fixture(),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];X.utils.sheet_add_aoa(s,[row],{origin:-1});set(i,'negotiated','付款主体',2,'');set(i,'negotiated','付款主体',3,'');s['!merges']=[{s:{r:1,c:1},e:{r:2,c:1}}];const before=JSON.stringify(i),r=C.build(i);assert.equal(update(r,'B2').v,'迈德斯特');assert.equal(update(r,'B3'),undefined);assert.equal(JSON.stringify(i),before);
});
test('无法识别的既有OA或财务编号不被当作空白，以有效公式缓存补空时不复制源公式',()=>{
 for(const field of ['对公付款0A流程号','单据号【财务应付单】']){const i=fixture();set(i,'negotiated','付款主体',2,'');set(i,'negotiated',field,2,'待核对');assert.equal(update(C.build(i),'B2'),undefined);}
 const i=fixture();set(i,'negotiated','对公付款0A流程号',2,'');set(i,'ledger','对公付款0A流程号',2,{t:'s',v:'MDST-DGFK20261001001',f:'A99'});const r=C.build(i);assert.equal(update(r,'C2').v,'MDST-DGFK20261001001');assert.equal(update(r,'C2').f,undefined);assert.equal(i.ledger.book.Sheets['对账明细-李奇'].D2.f,'A99');
});

test('正常付款申请仍按名称及金额精确匹配，零有效',()=>{
 const i=fixture();set(i,'applications','供应商全称',3,'其他供应商');assert.equal(C.build(i).normalRows[0][16].v,'未申请');set(i,'applications','供应商全称',3,'　测试供应商乙 ');set(i,'applications','本次申请付款金额',3,'￥100.00');assert.equal(C.build(i).normalRows[0][16].v,'已提交付款申请单');set(i,'applications','本次申请付款金额',3,100.01);assert.equal(C.build(i).normalRows[0][16].v,'未申请');set(i,'ledger','申请付款金额',3,0);set(i,'applications','本次申请付款金额',3,0);assert.equal(C.build(i).normalRows[0][16].v,'已提交付款申请单');
});
test('正常付款申请名称或金额为空错误缺缓存时不命中',()=>{
 for(const amount of ['', '/', -100, 0.001, {t:'n',f:'A99'}, {t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','本次申请付款金额',3,amount);assert.equal(C.build(i).normalRows[0][16].v,'未申请');}
 for(const name of ['', {t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','供应商全称',3,name);assert.equal(C.build(i).normalRows[0][16].v,'未申请');}
});
test('谈判按OA号归集而不校验供应商或单次金额，空错误及不同OA不回退名称金额',()=>{
 for(const name of ['其他供应商','',{t:'e',v:23,w:'#REF!'}]){const i=installmentFixture([{'供应商全称':name,'本次申请付款金额':40},{'供应商全称':'不同名称','本次申请付款金额':60}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已全部提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-40.00\n2026-10-10-60.00');assert.equal(r.reviewRows.length,0);}
 for(const id of ['', 'MDST-DGFK20990101999',{t:'s',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const i=fixture();set(i,'applications','合同号',2,id);const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未申请');assert.equal(negotiatedValue(r,'提交时间汇总'),'');assert.equal(r.reviewRows.length,0);}
});
test('提交日期与汇总相邻，已有字段只重排列顺序，不重复新增或改源表',()=>{
 for(const existing of [[],['提交时间汇总'],['提交付款申请单日期'],['提交付款申请单日期','人工备注','提交时间汇总']]){
  const i=fixture();if(existing.length)X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[existing,existing.map(h=>h==='人工备注'?'必须保留':'')],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i),h=r.negotiated.columnOrder.map(c=>r.negotiated.headers[c]);assert.equal(Math.abs(h.indexOf('提交付款申请单日期')-h.indexOf('提交时间汇总')),1);assert.equal(h.filter(x=>x==='提交时间汇总').length,1);assert.equal(h.filter(x=>x==='提交付款申请单日期').length,1);assert.equal(JSON.stringify(i),before);if(existing.includes('人工备注'))assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('人工备注')].v,'必须保留');
 }
});
test('仅有付款OA号的申请也保留，空金额标记待核对而非静默跳过',()=>{
 const i=installmentFixture([{'供应商全称':'','本次申请付款金额':''}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'待核对');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-金额为空');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/未按零计算/);
});
test('匹配申请日期空白错误或缺缓存时保留原日期，已提交状态有效并进入核对表',()=>{
 for(const d of ['', '近期', {t:'n',f:'DATE(2026,10,10)'},{t:'e',v:23,w:'#REF!'}]){
  const i=historyFixture('2026-09-10');set(i,'applications','申请日期',2,d);const r=C.build(i);assert.equal(update(r,'N2'),undefined);assert.equal(r.negotiated.rows[0][13].v,'2026-09-10');assert.equal(update(r,'O2').v,'已全部提交');assert.equal(r.issues.some(x=>x.kind.startsWith('application')),false);assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.stats.pending,0);
 }
});
test('申请日期列缺失仍按OA号标记谈判已提交，日期保留并列入核对表',()=>{
 const i=fixture();i.applications.book=book('付款申请汇总',['供应商全称','本次申请付款金额','合同号'],[{'供应商全称':'测试供应商甲','本次申请付款金额':100,'合同号':'MDST-DGFK20261001001'},{'供应商全称':'测试供应商乙','本次申请付款金额':100,'合同号':'MDST-DGFK20261001002'}]);const r=C.build(i);assert.equal(r.normalRows[0][16].v,'已提交付款申请单');assert.equal(r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'已全部提交');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(r.stats.pending,0);
});

const negotiatedValue=(r,key,row=0)=>r.negotiated.rows[row][r.negotiated.headers.indexOf(key)].v;
function installmentFixture(entries){
 const i=fixture(),existing=X.utils.sheet_to_json(i.applications.book.Sheets.付款申请汇总);i.applications.book=book('付款申请汇总',AH,[...entries.map(e=>({...existing[0],'合同号':'MDST-DGFK20261001001',...e})),existing[1]]);return i;
}
test('付款申请流程编码替代旧合同号，分批状态和时间汇总一致且原表不变',()=>{
 for(const amounts of [[30,20],[40,60],[60,60],[0]]){
  const i=installmentFixture(amounts.map((amount,n)=>({'申请日期':n?'2026-10-12':'2026-10-10','本次申请付款金额':amount}))),legacy=C.build(i);
  i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:' 流程编码 '};const before=JSON.stringify(i),r=C.build(i);
  assert.deepEqual(r.normalRows,legacy.normalRows);assert.deepEqual(r.negotiated,legacy.negotiated);assert.deepEqual(r.stats,legacy.stats);
  if(r.reviewRows.length)assert.match(r.reviewRows[0][6].v,/流程编码 MDST-DGFK20261001001/);
  assert.equal(JSON.stringify(i),before);
 }
});
test('新旧字段共存仅用流程编码，空白错误不同编号不回退旧合同号',()=>{
 for(const code of ['MDST-DGFK20261001001','MDST-DGFK20990101999','',{t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]),s=i.applications.book.Sheets.付款申请汇总;
  X.utils.sheet_add_aoa(s,[['流程编码'],[''],['MDST-DGFK20261001002']],{origin:'J1'});set(i,'applications','流程编码',2,code);
  const r=C.build(i),matched=code==='MDST-DGFK20261001001';
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),matched?'已部分提交':'未申请');
  assert.equal(negotiatedValue(r,'提交时间汇总'),matched?'2026-10-10-30.00':'');
  assert.equal(r.normalRows[0][16].v,'已提交付款申请单');assert.equal(r.normalRows[0][17].v,'2026-10-10-100.00');
  assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.length,0);
 }
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;
 set(i,'applications','合同号',2,'不匹配的旧合同');X.utils.sheet_add_aoa(s,[['流程编码'],['MDST-DGFK20261001001']],{origin:'J1'});
 assert.equal(negotiatedValue(C.build(i),'是否提交付款申请单'),'已全部提交');
});
test('仅有流程编码的申请不能跳过，金额为空保留并列出核对来源',()=>{
 const i=fixture(),s=i.applications.book.Sheets.付款申请汇总;s.D1={t:'s',v:'流程编码'};
 for(const field of ['供应商全称','本次申请付款金额'])set(i,'applications',field,2,'');
 const r=C.build(i);assert.equal(r.stats.applications,2);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'待核对');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][6].v,/流程编码 MDST-DGFK20261001001/);assert.equal(negotiatedValue(r,'已付金额'),0);
});
test('核对表流程编码对应即为已提交，名称金额异常和超额不抹去提交事实',()=>{
 for(const amount of ['',0.001,-1,120,{t:'e',v:23,w:'#REF!'}]){
  const i=installmentFixture([{'供应商全称':'不同名称','本次申请付款金额':amount}]);i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
  const before=JSON.stringify(i),r=C.build(i);assert.equal(r.reviewRows.length,1);assert.equal(r.reviewRows[0][4].v,'已提交付款申请单');assert(r.reviewRows[0][5].v);assert.match(r.reviewRows[0][6].v,/流程编码 MDST-DGFK20261001001/);
  assert.equal(negotiatedValue(r,'是否提交付款申请单'),'待核对');assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
 }
});
test('同流程归属多条谈判记录时核对表标记已提交，仍保留原记录和归属说明',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),s=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(s,{header:1})[1];
 X.utils.sheet_add_aoa(s,[row],{origin:-1});i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
 const r=C.build(i);assert.equal(r.negotiated.rows.length,2);assert.equal(r.reviewRows.length,2);assert(r.reviewRows.every(row=>row[4].v==='已提交付款申请单'&&row[5].v.includes('未重复分配')));assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('付款申请核对已提交不等于全部提交，日期异常仍保留说明及累计进度',()=>{
 const i=installmentFixture([{'本次申请付款金额':40,'申请日期':'日期待确认'}]);i.applications.book.Sheets.付款申请汇总.D1={t:'s',v:'流程编码'};
 const r=C.build(i);assert.equal(r.reviewRows[0][4].v,'已提交付款申请单');assert.match(r.reviewRows[0][5].v,/日期/);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'日期待确认-40.00');
 i.applications.book=book('付款申请汇总',AH,[]);const empty=C.build(i);assert.equal(empty.reviewRows.length,0);assert.equal(negotiatedValue(empty,'是否提交付款申请单'),'未申请');assert.equal(negotiatedValue(empty,'流程编号'),'MDST-DGFK20261001001','OA流程存在不代表提交过申请');
});
test('谈判分批按供应商及付款OA归集，累计小于等于超过分别为部分全部及待核对',()=>{
 for(const [amounts,status] of [[[30,20],'已部分提交'],[[40,60],'已全部提交'],[[60,60],'待核对'],[[0],'已部分提交']]){
  const i=installmentFixture(amounts.map((amount,n)=>({'申请日期':n?'2026-10-12':'2026-10-10','本次申请付款金额':amount}))),before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),status);assert.equal(negotiatedValue(r,'提交时间汇总'),amounts.map((v,n)=>(n?'2026-10-12':'2026-10-10')+'-'+v.toFixed(2)).join('\n'));assert.equal(r.normalRows[0][16].v,'已提交付款申请单');assert.equal(r.stats.pending,0);assert.deepEqual(r.negotiated.rows[0].slice(4,5).map(c=>c.v),[0]);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(negotiatedValue(r,'是否结清'),'否');assert.equal(r.reviewRows.length,status==='待核对'?1:0);if(status==='待核对')assert.match(r.reviewRows[0][5].v,/累计申请金额 120.00.*超过/);assert.equal(JSON.stringify(i),before);
 }
});
test('谈判同时按单次名称金额和付款OA命中的申请只计一次，分精度累计和零额有效',()=>{
 let i=installmentFixture([{'本次申请付款金额':100}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已全部提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');assert.equal(r.reviewRows.length,0);
 for(const [total,amounts] of [[0,[0]],[0.3,[0.1,0.2]]]){i=installmentFixture(amounts.map(amount=>({'本次申请付款金额':amount})));set(i,'negotiated','申请付款金额',2,total);set(i,'ledger','申请付款金额',2,total);set(i,'oa','付款金额',2,total);r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已全部提交');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);}
});
test('提交时间汇总按日期排序且同日同额源行全部保留，已有汇总列不重复添加',()=>{
 const i=installmentFixture([{'申请日期':'2026-10-12','本次申请付款金额':30},{'申请日期':'2026-10-10','本次申请付款金额':20},{'申请日期':'2026-10-10','本次申请付款金额':20}]);X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['提交时间汇总'],['旧汇总']],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-10-10\n2026-10-12');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-20.00\n2026-10-10-20.00\n2026-10-12-30.00');assert.equal(r.negotiated.headers.filter(h=>h==='提交时间汇总').length,1);assert.equal(JSON.stringify(i),before);
});
test('同名同OA申请金额为空异常负数或精度错误时待核对，不当零不改原付款',()=>{
 for(const amount of ['', '/', -1, 0.001,{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]);set(i,'applications','本次申请付款金额',2,amount);const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'待核对');assert.equal(r.reviewRows.length,1);assert.match(r.reviewRows[0][5].v,/未按零计算/);assert.equal(r.stats.pending,0);assert.equal(negotiatedValue(r,'已付金额'),0);assert.equal(negotiatedValue(r,'剩余付款金额'),100);assert.equal(JSON.stringify(i),before);
 }
});
test('同名OA明确关联优先于另一谈判记录的单次同额，不重复分配',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),n=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(n,{header:1})[1];row[2]='MDST-DGFK20261001003';row[6]=40;X.utils.sheet_add_aoa(n,[row],{origin:-1});X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','李奇','测试供应商甲',40]],{origin:-1});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单',0),'已部分提交');assert.equal(negotiatedValue(r,'提交时间汇总',0),'2026-10-10-40.00');assert.equal(negotiatedValue(r,'是否提交付款申请单',1),'未申请');assert.equal(negotiatedValue(r,'提交时间汇总',1),'');assert.equal(r.reviewRows.length,0);assert.deepEqual(r.negotiated.rows.map(row=>row[4].v),[0,0]);
});
test('申请归属多条谈判记录时保留原提交状态日期汇总，空状态待核对且不阻挡下载',()=>{
 const i=installmentFixture([{'本次申请付款金额':40}]),n=i.negotiated.book.Sheets['10月谈判付款主表'],row=X.utils.sheet_to_json(n,{header:1})[1];X.utils.sheet_add_aoa(n,[row],{origin:-1});X.utils.sheet_add_aoa(n,[['是否提交付款申请单','提交付款申请单日期','提交时间汇总'],['人工确认','2026-09-01','原汇总甲'],['','2026-09-02','原汇总乙']],{origin:'N1'});const before=JSON.stringify(i),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单',0),'人工确认');assert.equal(negotiatedValue(r,'是否提交付款申请单',1),'待核对');assert.equal(negotiatedValue(r,'提交付款申请单日期',0),'2026-09-01');assert.equal(negotiatedValue(r,'提交付款申请单日期',1),'2026-09-02');assert.equal(negotiatedValue(r,'提交时间汇总',0),'原汇总甲');assert.equal(negotiatedValue(r,'提交时间汇总',1),'原汇总乙');assert.equal(r.reviewRows.length,2);assert(r.reviewRows.every(row=>/未重复分配/.test(row[5].v)));assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
});
test('谈判分批申请始终按原付款OA关联，OA表缺记录或总金额不同不影响关联',()=>{
 const i=installmentFixture([{'合同号':'MDST-DGFK20990101999','本次申请付款金额':30}]);set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20990101999');let r=C.build(i);assert.equal(negotiatedValue(r,'流程编号'),'');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');set(i,'applications','合同号',2,'MDST-DGFK20261001001');r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'未申请');assert.equal(negotiatedValue(r,'对公付款0A流程号'),'MDST-DGFK20990101999');
 const j=installmentFixture([{'本次申请付款金额':290000}]);set(j,'negotiated','申请付款金额',2,1202409);r=C.build(j);assert.equal(negotiatedValue(r,'流程编号'),'MDST-DGFK20261001001');assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-290000.00');assert.equal(r.issues.some(q=>q.id==='oa:10月谈判付款主表!2'),false);assert.equal(r.issues.some(q=>q.kind.startsWith('application')),false);
});
test('申请OA对应但供应商为空时仍更新提交汇总，保留历史提交日期',()=>{
 for(const supplier of ['',{t:'e',v:23,w:'#REF!'}]){
  const i=installmentFixture([{'本次申请付款金额':30}]);set(i,'applications','供应商全称',2,supplier);X.utils.sheet_add_aoa(i.negotiated.book.Sheets['10月谈判付款主表'],[['是否提交付款申请单','提交付款申请单日期','提交时间汇总'],['已部分提交','2026-09-01','原汇总']],{origin:'N1'});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已部分提交');assert.equal(negotiatedValue(r,'提交付款申请单日期'),'2026-09-01\n2026-10-10');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-30.00');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);
 }
});
test('单次同额申请含未知多OA也不能直接归集，全部属于目标的多OA申请只计一次',()=>{
 for(const ids of ['MDST-DGFK20990101998、MDST-DGFK20990101999','MDST-DGFK20261001001、MDST-DGFK20990101999']){
  const i=installmentFixture([{'合同号':ids,'本次申请付款金额':100}]),r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),ids.includes('MDST-DGFK20261001001')?'待核对':'未申请');assert.equal(negotiatedValue(r,'提交时间汇总'),'');assert.equal(r.reviewRows.length,ids.includes('MDST-DGFK20261001001')?1:0);if(r.reviewRows.length)assert.match(r.reviewRows[0][5].v,/多个付款 OA 号/);assert.equal(r.stats.pending,0);
 }
 const i=installmentFixture([{'合同号':'MDST-DGFK20261001001、MDST-DGFK20261001003','本次申请付款金额':100}]);set(i,'negotiated','对公付款0A流程号',2,'MDST-DGFK20261001001、MDST-DGFK20261001003');X.utils.sheet_add_aoa(i.oa.book.Sheets.Sheet0,[['MDST-DGFK20261001003','出纳','人员乙','归档','李奇','测试供应商甲',100]],{origin:-1});const r=C.build(i);assert.equal(negotiatedValue(r,'是否提交付款申请单'),'已全部提交');assert.equal(negotiatedValue(r,'提交时间汇总'),'2026-10-10-100.00');assert.equal(r.reviewRows.length,0);assert.equal(r.stats.pending,0);
});
