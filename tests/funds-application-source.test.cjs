const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,set}=require('./funds-progress-fixtures.cjs');
const progress=result=>result.progressRows.map(row=>Object.fromEntries(result.progressHeaders.map((h,c)=>[h,row[c]])));
const display=c=>c.t==='n'&&c.z?X.SSF.format(c.z,c.v):c.v;
function newApplications(){const i=fixture();for(const row of [2,3,4,5])set(i,'applications','申请日期',row,'2026-11-19');return i;}
const submission=result=>result.negotiated.rows[0][result.negotiated.headers.indexOf('是否提交付款申请单')].v;

test('11月19日含当天起按申请表取应付时间，之前按分类取谈判表或台账',()=>{
 for(const [day,current] of [['2026-11-18 23:59:59',false],['2026-11-19',true],['2026年11月19日',true],['2026-11-20',true],['2027-01-01',true]]){
  const i=fixture();for(const row of [2,3]){set(i,'applications','申请日期',row,day);set(i,'applications','应付时间',row,'2026-12-31');}
  const before=JSON.stringify(i),p=progress(C.build(i));
  assert.equal(display(p[0].应付时间),current?'2026年12月31日':'2026年11月6日',day+'谈判来源');
  assert.equal(display(p[1].应付时间),current?'2026年12月31日':'2026年11月9日',day+'台账来源');
  assert.equal(JSON.stringify(i),before);
 }
});

test('旧申请缺少申请表应付时间不误报，新申请缺少应付时间不回退',()=>{
 for(const type of ['谈判付款','谈判付款-谈判付款表','正常付款-谈判付款表','正常付款','正常付款-正常对账表']){
  const i=fixture();set(i,'applications','付款分类',2,type);set(i,'applications','应付时间',2,'');
  set(i,'applications','申请日期',2,'2026-11-18');let r=C.build(i);
  assert.equal(display(progress(r)[0].应付时间),C.applicationClass(type).source==='谈判付款表'?'2026年11月6日':'2026年11月3日');
  assert(!r.reviewRows.some(row=>row[6].v===i.applications.name&&row[8].v===2&&/应付时间为空或无法识别/.test(row[5].v)));
  set(i,'applications','申请日期',2,'2026-11-19');r=C.build(i);assert.equal(progress(r)[0].应付时间.v,'');
  assert(r.reviewRows.some(row=>row[6].v===i.applications.name&&row[8].v===2&&/应付时间为空或无法识别/.test(row[5].v)));
 }
});

test('申请日期缺失、无年份或公式异常时不猜日期来源，只提示申请日期',()=>{
 for(const raw of ['',0,'11月19日','2026年11月','日期未知',{t:'e',v:23,w:'#REF!'},{t:'n',f:'A99'},{t:'z',f:'A99',v:46345}]){
  const i=fixture();set(i,'applications','申请日期',2,raw);set(i,'applications','应付时间',2,'2027-01-26');
  const r=C.build(i);assert.equal(progress(r)[0].应付时间.v,'');
  const note=r.reviewRows.filter(row=>row[6].v===i.applications.name&&row[8].v===2&&row[5].v.startsWith('付款进度核对')).map(row=>row[5].v).join('\n');
  assert.match(note,/申请日期为空或无法识别/);assert.doesNotMatch(note,/应付时间为空或无法识别/);
 }
});

test('数值及1904申请日期按真实日历判断切换日，同OA跨界申请分别取日期',()=>{
 for(const date1904 of [false,true]){
  const i=fixture();i.applications.book.Workbook={WBProps:{date1904}};
  const serial=day=>(Date.UTC(2026,10,day)-Date.UTC(1899,11,30))/86400000-(date1904?1462:0);
  set(i,'applications','申请日期',2,{t:'n',v:serial(18)+0.9,z:C.DATE_FORMAT});
  set(i,'applications','申请日期',4,{t:'n',v:serial(19),z:C.DATE_FORMAT});
  set(i,'applications','应付时间',2,'2027-01-26');set(i,'applications','应付时间',4,'2027-02-26');
  const p=progress(C.build(i));assert.equal(display(p[0].应付时间),'2026年11月6日');assert.equal(display(p[2].应付时间),'2027年2月26日');
 }
});

test('申请汇总五种分类均识别，保留业务分类及来源，异常不猜测',()=>{
 for(const [label,kind,source] of [
  ['谈判付款','谈判付款','谈判付款表'],['正常付款','正常付款','资金台账'],
  ['谈判付款-谈判付款表','谈判付款','谈判付款表'],['正常付款-谈判付款表','正常付款','谈判付款表'],
  ['正常付款-正常对账表','正常付款','资金台账']
 ])assert.deepEqual(C.applicationClass(label),{label,kind,source});
 assert.deepEqual(C.applicationClass(' 正常付款 － 谈判付款表\n'),C.applicationClass('正常付款-谈判付款表'));
 for(const v of [undefined,null,'','其他','是','正常付款-未知表','constructor','__proto__'])assert.equal(C.applicationClass(v),null);
});

test('新增谈判来源分类参与谈判申请匹配，不被原表是否谈判分批覆盖',()=>{
 for(const type of ['谈判付款-谈判付款表','正常付款-谈判付款表']){
  const i=fixture();for(const row of [2,4])set(i,'applications','付款分类',row,type);
  const ns=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];
  X.utils.sheet_add_aoa(ns,[['是否谈判分批'],[type.startsWith('正常')?'是':'否']],{origin:'O1'});
  X.utils.sheet_add_aoa(i.applications.book.Sheets.付款申请汇总,[['是否已付款'],['否'],[''],['否']],{origin:'G1'});
  const before=JSON.stringify(i),r=C.build(i),p=progress(r);
  assert(!r.issues.some(x=>x.kind==='application-classification'));assert.equal(submission(r),'已经提交部分');
  assert.deepEqual(r.highlights.negotiatedCells,['M2']);
  assert.equal(p[0].付款来源.v,type.startsWith('正常')?'谈判付款表-正常付款':'谈判付款表-谈判付款');
  assert.equal(p[0].是否已付款.v,'未付款','谈判来源仍取申请汇总实付信息');
  assert.equal(JSON.stringify(i),before);
 }
});

test('正常对账表分类仅匹配正常付款，不能作为谈判申请',()=>{
 const i=fixture();for(const row of [2,3,4])set(i,'applications','付款分类',row,'正常付款-正常对账表');
 const r=C.build(i);assert.equal(submission(r),'未提交');assert.deepEqual(r.highlights.negotiatedCells,[]);
 assert.equal(r.normalRows[0][r.normalHeaders.indexOf('是否提交付款申请单')].v,'已提交');
 assert.deepEqual(progress(r).map(row=>row.付款来源.v),['资金台账','资金台账','资金台账','资金台账']);
 assert(!r.issues.some(x=>x.kind==='application-classification'));
});

test('不同申请分类不合为重复申请，同分类跨日重复仍只累计一次',()=>{
 const i=fixture();set(i,'applications','付款分类',2,'谈判付款-谈判付款表');set(i,'applications','付款分类',4,'正常付款-谈判付款表');
 // Both 40 payments must count: 80 exceeds the 70 target, while incorrectly deduplicated 40 does not.
 set(i,'negotiated','申请付款金额',2,70);set(i,'negotiated','剩余付款金额',2,70);
 let r=C.build(i);assert.equal(submission(r),'已经提交全部');assert(!r.reviewRows.some(row=>row[5].v.startsWith('重复提交')));
 set(i,'applications','付款分类',4,'谈判付款-谈判付款表');r=C.build(i);
 assert.equal(submission(r),'已经提交部分');assert.equal(r.reviewRows.filter(row=>row[5].v.startsWith('重复提交')).length,1);
});

test('付款进度逐申请行取应付时间，重新上传更新，不改变正常及谈判原日期和公式',()=>{
 const i=newApplications(),before=C.build(i);
 set(i,'applications','应付时间',2,'2026-12-31');set(i,'applications','应付时间',3,'2027-01-08');set(i,'applications','应付时间',4,'2027-02-06');
 const source=JSON.stringify(i),r=C.build(i),p=progress(r);
 assert.deepEqual(p.slice(0,3).map(row=>display(row.应付时间)),['2026年12月31日','2027年1月8日','2027年2月6日']);
 assert.deepEqual(r.normalRows,before.normalRows);assert.deepEqual(r.negotiated,before.negotiated);assert.deepEqual(r.highlights,before.highlights);
 assert.equal(p[0].应付时间.t,'n');assert.equal(JSON.stringify(i),source);
 set(i,'applications','应付时间',2,'2027-03-02');assert.equal(display(progress(C.build(i))[0].应付时间),'2027年3月2日');
});

test('申请应付时间缺失或异常不借其他表填充，提示精确申请来源',()=>{
 for(const v of ['', '日期未知', {t:'e',v:23,w:'#REF!'},{t:'s',f:'A99'},{t:'z',f:'A99',v:0}]){
  const i=newApplications();set(i,'applications','应付时间',2,v);i.applications.name='付款单申请汇总-最新.xlsx';
  const r=C.build(i),due=progress(r)[0].应付时间,notes=r.reviewRows.filter(row=>row[6].v===i.applications.name&&row[8].v===2).map(row=>row[5].v).join('\n');
  assert.notEqual(due.t,'n');assert.match(notes,/应付时间为空或无法识别/);assert.match(notes,/付款单申请汇总-最新.xlsx \/ 付款申请汇总 第 2 行/);
  assert.doesNotMatch(notes,/negotiated.xlsx.*应付时间|ledger.xlsx.*应付时间/);assert.equal(r.stats.pending,0);
 }
 const i=newApplications();delete i.applications.book.Sheets.付款申请汇总.Z1;
 assert(progress(C.build(i)).every(row=>row.应付时间.v===''),'缺整列也不回退其他表');
});

test('申请应付时间兼容表头、真实日期及1904日期，零值不当空白',()=>{
 const i=newApplications(),s=i.applications.book.Sheets.付款申请汇总;s.Z1.v='应付款时间';
 s.Z2={t:'n',v:46328-1462,z:C.DATE_FORMAT};i.applications.book.Workbook={WBProps:{date1904:true}};
 assert.equal(display(progress(C.build(i))[0].应付时间),'2026年11月2日');
 s.Z2={t:'d',v:new Date('2027-01-26T00:00:00Z')};assert.equal(display(progress(C.build(i))[0].应付时间),'2027年1月26日');
 s.Z2={t:'n',v:0,z:C.DATE_FORMAT};assert.notEqual(progress(C.build(i))[0].应付时间.v,'');
});

test('分类公式没有结果时不能用手动决定绕过本次上传分类',()=>{
 for(const value of [{t:'s',f:'A99'},{t:'z',f:'A99',v:'谈判付款'},{t:'e',v:23,w:'#REF!'}]){
  const i=fixture();set(i,'applications','付款分类',2,value);
  const r=C.build(i),issue=r.issues.find(x=>x.kind==='application-classification');assert(issue);assert.equal(progress(r)[0].付款来源.v,'');
  assert.equal(progress(r)[0].是否已付款.v,'待核对');assert.deepEqual(issue.choices.map(x=>x.value),['keep']);
  assert.equal(C.build(i,{[issue.id]:['negotiated']}).stats.pending,1);
 }
});
