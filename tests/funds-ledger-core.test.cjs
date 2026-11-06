const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-ledger-core.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const H=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否提交付款申请单','是否结清','是否谈判付款'];
const row=(overrides={})=>({...{'供应商名称':'供应商甲','付款主体':'采购公司','对公付款0A流程号':'OA001','申请付款金额':100,'摘要':'九月货款','应付款时间':'2026-11-06','预估付款时间':'2026-12','是否提交付款申请单':'否','是否结清':'否','是否谈判付款':'否'},...overrides});
function sheet(rows,headers=H,title=false){return X.utils.aoa_to_sheet([...(title?[['台账标题']]:[]),headers,...rows.map(record=>headers.map(h=>record[h]??''))]);}
function book(rows=[]){const b=X.utils.book_new();PEOPLE.forEach((p,i)=>X.utils.book_append_sheet(b,sheet(i?[]:rows),'对账明细-'+p));return b;}
function set(b,header,cell,r=2,name='对账明细-李奇'){b.Sheets[name][X.utils.encode_cell({r:r-1,c:H.indexOf(header)})]=cell;}
const values=result=>result.sheets.map(s=>s.rows.map(r=>r.map(c=>c.v)));
test('仅明确否否进入两类，OA空白/未提交OA未申请，其他非空已申请，提交状态不参与筛选',()=>{
 const b=book([
  row({'对公付款0A流程号':'000001','是否提交付款申请单':'是'}),row({'对公付款0A流程号':'  未提交OA  '}),row({'对公付款0A流程号':'  '}),
  row({'对公付款0A流程号':'待补OA','是否结清':' 否 ','是否谈判付款':'否 '}),
  row({'是否结清':''}),row({'是否谈判付款':''}),row({'是否结清':'是'}),row({'是否谈判付款':'是'}),row({'是否结清':'未结清'}),row({'是否谈判付款':'正常付款'})
 ]),before=JSON.stringify(b),r=C.build(b,'源.xlsx');
 assert.deepEqual(r.sheets.map(s=>s.rows.length),[2,2]);assert.deepEqual(r.counts[0],{sheet:'对账明细-李奇',applied:2,unapplied:2,excluded:6});
 assert.deepEqual(r.sheets[0].rows.map(row=>row[2].v),['000001','待补OA']);assert.equal(r.sheets[0].rows[0][6].v,'是');
 assert.equal(r.sheets[1].rows[0][2].v,'  未提交OA  ');assert.equal(r.sheets[1].rows[1][2].v,'  ');assert.equal(JSON.stringify(b),before);
});
test('按工作簿及行顺序保留重复记录，不读取历史汇总和其他人的表',()=>{
 const b=book([row(),row()]);b.Sheets['对账明细-孙立柱']=sheet([row({'供应商名称':'孙表'})]);
 b.SheetNames.unshift(b.SheetNames.pop());X.utils.book_append_sheet(b,sheet([row({'申请付款金额':'不是金额'})]),'采购付款台账_汇总1017');
 X.utils.book_append_sheet(b,sheet([row()]),'对账明细-李奇历史');X.utils.book_append_sheet(b,sheet([row()]),'对账明细-其他人');
 const r=C.build(b);assert.equal(r.sheets[0].rows.length,3);assert.equal(r.sheets[0].rows[0][0].v,'孙表');
 assert.deepEqual(r.sheets[0].sources,[{sheet:'对账明细-孙立柱',row:2},{sheet:'对账明细-李奇',row:2},{sheet:'对账明细-李奇',row:3}]);
 assert.deepEqual(r.counts.map(c=>c.sheet),b.SheetNames.slice(0,5));
});
test('OA输出文本及前导零保留，未提交 OA的空白兼容且原文不变',()=>{
 const b=book([row(),row({'对公付款0A流程号':' 未提交 OA '})]);set(b,'对公付款0A流程号',{t:'n',v:123,z:'000000'});
 const r=C.build(b),oa=r.sheets[0].rows[0][2];assert.deepEqual(oa,{t:'s',v:'000123',z:'@'});assert.equal(r.sheets[1].rows[0][2].v,' 未提交 OA ');
 const out=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{type:'buffer'});assert.equal(out.Sheets[C.SHEET_NAMES[0]].C2.t,'s');assert.equal(out.Sheets[C.SHEET_NAMES[0]].C2.v,'000123');
});
test('五人表头位置和别名独立识别，允许标题行及日期后缀',()=>{
 const b=book(),headers=[...H].reverse().map(h=>({'供应商名称':'供应商','对公付款0A流程号':'对公付款OA流程号','申请付款金额':'本次申请付款金额','应付款时间':'应付时间','预估付款时间':'预估付款月份','是否提交付款申请单':'是否提交付款审单'}[h]||h));
 const record=Object.fromEntries(H.map((h,i)=>[headers[H.length-i-1],row()[h]]));
 b.Sheets['对账明细-徐英1102']=sheet([record],headers,true);delete b.Sheets['对账明细-徐英'];b.SheetNames[1]='对账明细-徐英1102';
 const r=C.build(b);assert.equal(r.sheets[0].rows[0][0].v,'供应商甲');assert.equal(r.sheets[0].rows[0][3].v,100);assert.deepEqual(r.sheets[0].sources,[{sheet:'对账明细-徐英1102',row:3}]);
});
test('缺人员表、多版本及必要表头/相关重复表头阻断，不任意选列',()=>{
 const missing=book();missing.SheetNames.pop();assert.throws(()=>C.build(missing,'源.xlsx'),/源.xlsx.*孙立柱.*缺失/);
 const duplicate=book();X.utils.book_append_sheet(duplicate,sheet([]),'对账明细-李奇1102');assert.throws(()=>C.build(duplicate),/李奇.*多个版本/);
 for(const required of ['供应商名称','申请付款金额','对公付款0A流程号','是否结清','是否谈判付款']){
  const b=book();b.Sheets['对账明细-李奇']=sheet([],H.filter(h=>h!==required));assert.throws(()=>C.build(b),/缺少必要表头/);
 }
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'供应商':'冲突供应商'})],[...H,'供应商']);assert.throws(()=>C.build(b),/字段“供应商名称”重复/);
});
test('金额0、空白及合法币种千分位保留，负数不取绝对值，无效金额不能变为0',()=>{
 const amounts=[0,'',null,'￥1,234.56','人民币 2,345.67元','(123.45)',-1,0.001],r=C.build(book(amounts.map(value=>row({'申请付款金额':value}))));
 assert.deepEqual(r.sheets[0].rows.map(row=>row[3].v),[0,'','',1234.56,2345.67,-123.45,-1,0.001]);assert.equal(r.sheets[0].rows[1][3].t,'s');
 for(const bad of ['/', '1,23.45','12元待核对',true,Infinity,Number.MAX_SAFE_INTEGER,'123456789.1234567'])assert.throws(()=>C.build(book([row({'申请付款金额':bad})]),'源.xlsx'),/源.xlsx.*对账明细-李奇!D2/);
});
test('应付款与预估日期分别取源，不用另一日期或开票时间补空',()=>{
 const b=book([row({'应付款时间':''}),row({'对公付款0A流程号':'','预估付款时间':''})]);
 b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','开票时间':'2026-09-01','月份':'2026-09'}),row({'对公付款0A流程号':'','预估付款时间':'','开票时间':'2026-09-01','月份':'2026-09'})],[...H,'开票时间','月份']);
 const r=C.build(b);assert.equal(r.sheets[0].rows[0][5].v,'');assert.equal(r.sheets[1].rows[0][5].v,'');assert.deepEqual(r.warnings,[]);
 const missing=book();missing.Sheets['对账明细-李奇']=sheet([row(),row({'对公付款0A流程号':''})],H.filter(h=>!['应付款时间','预估付款时间'].includes(h)));
 const m=C.build(missing);assert.deepEqual(m.sheets.map(s=>s.rows[0][5].v),['','']);assert.equal(m.warnings.length,1);assert.match(m.warnings[0],/所属月份.*未推算/);
});
test('仅无预估列的未申请行按所属月末加账期天数再加7天，月份不猜基准日',()=>{
 const b=book(),headers=[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期'];
 const record=(month,terms,overrides={})=>row({'对公付款0A流程号':'','所属月份':month,'账期':terms,...overrides});
 b.Sheets['对账明细-李奇']=sheet([record('2026-10',60),record('2026年10月','5个月账期'),record('2024-02',0),record('2026-10','票到60天，50%承兑'),record('2026-10','50%承兑'),record('2026-10','票到30天或60天'),record('2026-10',''),record('','60天'),record('2026-10','30-60天'),record('2026-10','30天+30天'),record('2026-10','60天',{'对公付款0A流程号':'OA002','应付款时间':''})],headers);
 const r=C.build(b,'源.xlsx');assert.deepEqual(r.sheets[1].rows.map(row=>C.formatCell(row[5])),['2027-01-06','2027-04-06','2024-03-07','2027-01-06','','','','','','']);
 assert.equal(r.sheets[0].rows[0][5].v,'');assert.equal(r.warnings.length,6);assert(r.warnings.every(w=>/源.xlsx.*对账明细-李奇!/.test(w)));
});
test('已有预估列即使空白也不推算；月份别名、1904缓存日期仅在缺列推算时读取',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'对公付款0A流程号':'','预估付款时间':'','月份':'2026-10','账期':'60天'})],[...H,'月份','账期']);
 assert.equal(C.build(b).sheets[1].rows[0][5].v,'');
 for(const header of ['月份','登记月份']){
  const n=book();n.Workbook={WBProps:{date1904:true}};n.Sheets['对账明细-李奇']=sheet([row({'对公付款0A流程号':'',[header]:{t:'n',v:44840,f:'DATE(2026,10,7)',z:'yyyy-mm'},'账期':'60天'})],[...H.filter(h=>h!=='预估付款时间'),header,'账期']);
  assert.equal(C.formatCell(C.build(n).sheets[1].rows[0][5]),'2027-01-06');
 }
 const invalid=book();invalid.Sheets['对账明细-李奇']=sheet([row({'对公付款0A流程号':'','所属月份':{t:'n',f:'A99'},'账期':'60天'})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);assert.throws(()=>C.build(invalid),/!J2.*公式缺少/);
});
test('账期负数、范围和不确定天数不误算为正的固定天数',()=>{
 for(const terms of ['-30天','－30天','30—60天','30/60天','30／60天','30或60天','30或者60天','约60天','60天以上']){
  const b=book();b.Sheets['对账明细-李奇']=sheet([row({'对公付款0A流程号':'','所属月份':'2026-10','账期':terms})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);
  const r=C.build(b);assert.equal(r.sheets[1].rows[0][5].v,'',terms);assert.equal(r.warnings.length,1,terms);
 }
});
test('年月保留月精度，普通未解析日期保留原文并定位警告',()=>{
 const r=C.build(book([row({'对公付款0A流程号':''}),row({'应付款时间':'预计11月底'}),row({'应付款时间':'2026-02-30'})]),'源.xlsx');
 assert.equal(C.formatCell(r.sheets[1].rows[0][5]),'2026-12');assert.equal(r.sheets[0].rows[0][5].v,'预计11月底');assert.equal(r.sheets[0].rows[1][5].v,'2026-02-30');
 assert.equal(r.warnings.length,2);assert.match(r.warnings[0],/源.xlsx.*对账明细-李奇!F3.*保留原值/);
 for(const value of [{t:'n',v:46332,z:'yyyy-mm-dd'},'2026-11-06',{t:'d',v:new Date('2026-11-06T00:00:00Z')}]){
  const b=book();b.Sheets['对账明细-李奇']=sheet([row({'对公付款0A流程号':'','预估付款月份':value})],H.map(h=>h==='预估付款时间'?'预估付款月份':h));
  assert.equal(C.formatCell(C.build(b).sheets[1].rows[0][5]),'2026-11');
 }
});
test('1904日期和typed Date输出为1900纪元，日期格式及年月精度保留',()=>{
 const b=book([row(),row(),row()]);b.Workbook={WBProps:{date1904:true}};
 set(b,'应付款时间',{t:'n',v:44870,z:'yyyy-mm-dd'});set(b,'应付款时间',{t:'d',v:new Date('2026-11-06T00:00:00Z'),z:'yyyy-mm-dd'},3);set(b,'应付款时间',{t:'n',v:44870,z:'yyyy"年"m"月"'},4);
 const r=C.build(b),out=C.workbook(r);assert.equal(r.sheets[0].rows[0][5].v,46332);assert.equal(C.formatCell(r.sheets[0].rows[0][5]),'2026-11-06');
 assert.equal(C.formatCell(r.sheets[0].rows[1][5]),'2026-11-06');assert.equal(C.formatCell(r.sheets[0].rows[2][5]),'2026年11月');assert.equal(out.Workbook?.WBProps?.date1904,undefined);
});
test('有效公式缓存包括零和空字符串均取值，输出不携带公式且不修改源表',()=>{
 const b=book([row()]);set(b,'申请付款金额',{t:'n',f:'1-1',v:0});set(b,'摘要',{t:'s',f:'""',v:''});set(b,'是否结清',{t:'s',f:'"否"',v:'否'});
 const before=JSON.stringify(b),r=C.build(b);assert.equal(r.sheets[0].rows[0][3].v,0);assert.equal(r.sheets[0].rows[0][4].v,'');
 assert(r.sheets.every(s=>s.rows.every(row=>row.every(c=>c.f===undefined))));assert.equal(JSON.stringify(b),before);
});
test('所需字段公式错误或缺缓存准确定位阻断，排除行非筛选字段及不用的日期不阻断',()=>{
 for(const h of H){
  if(h==='预估付款时间')continue;
  for(const cell of [{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const b=book([row()]);set(b,h,cell);assert.throws(()=>C.build(b,'源.xlsx'),/源.xlsx.*对账明细-李奇![A-J]2/);}
 }
 const b=book([row(),row({'是否结清':'是'})]);set(b,'预估付款时间',{t:'e',v:23});set(b,'申请付款金额',{t:'n',f:'A99'},3);
 b.Sheets['对账明细-李奇'].K2={t:'e',v:23};b.Sheets['对账明细-李奇']['!ref']='A1:K3';assert.equal(C.build(b).sheets[0].rows.length,1);
 const unapplied=book([row({'对公付款0A流程号':''})]);set(unapplied,'预估付款时间',{t:'n',f:'A99'});assert.throws(()=>C.build(unapplied),/!G2/);
});
test('可选正文列缺失只提示，源空白与零不互相转换，格式空尾不扩大扫描范围',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'供应商名称':'','申请付款金额':0})],H.filter(h=>!['付款主体','摘要','是否提交付款申请单'].includes(h)));
 const s=b.Sheets['对账明细-李奇'];s.XFD1048576={t:'s',v:'',s:1};s['!ref']='A1:XFD1048576';
 const r=C.build(b);assert.deepEqual(r.sheets[0].rows[0].filter((_,i)=>[0,1,4,6].includes(i)).map(c=>c.v),['','','','']);assert.equal(r.sheets[0].rows[0][3].v,0);assert.equal(r.warnings.length,1);
});
test('导出仅两张七列表，空类别保留表头，原生数值日期可回读且带筛选列宽',()=>{
 const r=C.build(book([row()])),before=JSON.stringify(r),out=C.workbook(r),roundtrip=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});
 assert.deepEqual(out.SheetNames,C.SHEET_NAMES);assert.deepEqual(roundtrip.SheetNames,C.SHEET_NAMES);
 C.SHEET_NAMES.forEach((name,i)=>{assert.deepEqual(X.utils.sheet_to_json(roundtrip.Sheets[name],{header:1})[0],C.HEADERS[i]);assert.equal(out.Sheets[name]['!cols'].length,7);assert.match(out.Sheets[name]['!autofilter'].ref,/^A1:G[12]$/);});
 assert.equal(roundtrip.Sheets[C.SHEET_NAMES[0]].D2.t,'n');assert.equal(roundtrip.Sheets[C.SHEET_NAMES[0]].F2.t,'n');assert.equal(JSON.stringify(r),before);
 assert.equal(C.exportFilename('/path/采购付款台账.v2.xlsx'),'采购付款台账.v2－资金台账整理.xlsx');assert.equal(C.formatCell({t:'n',v:0}),'0');
});
test('列宽按内容及表头留筛选余量，超长摘要限宽并换行',()=>{
 const r=C.build(book([row({'供应商名称':'这是用于核验列宽的较长供应商名称','摘要':'摘要'.repeat(100)})])),out=C.workbook(r).Sheets[C.SHEET_NAMES[0]],empty=C.workbook(C.build(book())).Sheets[C.SHEET_NAMES[0]];
 assert(out['!cols'][0].wch>empty['!cols'][0].wch);assert.equal(out['!cols'][4].wch,72);assert.equal(out.E2.s.alignment.wrapText,true);
 assert(out['!cols'][2].wch>=20);assert(out['!cols'][6].wch>=20);assert(empty['!cols'].every(c=>c.wch>=12));
});
