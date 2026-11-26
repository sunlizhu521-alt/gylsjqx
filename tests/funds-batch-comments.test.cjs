const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,set,ID,AH,book}=require('./funds-progress-fixtures.cjs');
const P=['批次（付款节奏跟进）：','申请日期（付款节奏跟进）：','OA备注时间（付款节奏跟进）：','付款时间（付款节奏跟进）：'];
function sample(amounts=[40,60],apps=[['2026-11-01',40,'','是','是']]){
 const i=fixture(),s=i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]];
 for(const a of Object.keys(s))if(/^[A-Z]+\d+$/.test(a)&&X.utils.decode_cell(a).c>=12)delete s[a];
 X.utils.sheet_add_aoa(s,[amounts.map((_,n)=>'2026-11-'+String(n+2).padStart(2,'0')),amounts],{origin:'M1'});
 i.applications.book=book('付款申请汇总',[[...AH,'核对说明','是否OA备注','是否已付款'],...apps.map(([date,amount,batch,oa,paid])=>[date,'合成供应商1','迈德斯特',ID(1),amount,'谈判付款',batch,oa,paid])]);
 return i;
}
const source=i=>i.negotiated.book.Sheets[i.negotiated.book.SheetNames[0]],states=r=>r.progressRows.map(row=>row[r.progressHeaders.indexOf('是否已付款')].v);
test('四项批注与红黄集合一致，只读取申请专用字段，不修改源公式和人工内容',()=>{
 const i=sample(),s=source(i);s.M2={t:'n',v:40,f:'D2-60',c:[{a:'人工',t:'保留人工说明\n11月5日核实-已支付'}]};
 X.utils.sheet_add_aoa(s,[['是否OA备注','是否已付款'],['保留甲','保留乙']],{origin:'O1'});
 const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.highlights.negotiatedCells,['M2']);assert.deepEqual(r.highlights.negotiatedPaidCells,['M2']);
 assert.deepEqual(r.batchComments.M2,[P[0]+'共2批，第1批',P[1]+'2026年11月1日',P[2]+'已备注，时间未填',P[3]+'已付款，时间未填']);
 assert.equal(r.applicationComments.M2,P[1]+'2026年11月1日');assert.deepEqual(r.applicationBatchMatches['付款申请汇总!2'],{address:'M2',unique:true});
 assert(!r.negotiated.managedColumns.includes(14));assert(!r.negotiated.managedColumns.includes(15));assert.equal(JSON.stringify(i),before);
});
test('完整日期计划列可移动，月度支付不计批次，按表头日期排序而非源列顺序',()=>{
 const i=sample([40,60,20]),s=source(i);s.M1.v='2027-01-02';s.N1.v='2026-11-03';s.O1.v='2026-12-04';
 X.utils.sheet_add_aoa(s,[['2027年1月支付','人工备注','2027-02-01'],[40,'保留',10]],{origin:'P1'});s.R2.c=[{a:'人工',t:'人工说明'}];
 const r=C.build(i);assert.equal(r.negotiated.batchPlans[0].count,4);assert.deepEqual(r.negotiated.batchPlans[0].cells.map(b=>b.address),['N2','O2','M2','R2']);assert.equal(r.batchComments.M2[0],P[0]+'共4批，第3批');assert.equal(r.batchComments.R2[0],P[0]+'共4批，第4批');assert(!r.batchComments.P2);
});
test('有效非零计划格决定单笔或分批，空OA与未提交OA不影响，零空异常不计负数照计',()=>{
 for(const oa of [ID(1),'未提交OA',''])for(const [amounts,count,kind]of [[[40,0,'',null],1,'单笔'],[[40,-5,'待定',{t:'n',f:'A99'}],2,'分批'],[[0,'',null],0,'待核对']]){
  const i=sample(amounts,[]);set(i,'negotiated','对公付款0A流程号',2,oa);const before=JSON.stringify(i),r=C.build(i);
  assert.equal(r.negotiated.batchPlans[0].count,count);assert.equal(r.negotiated.batchPlans[0].kind,kind);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(JSON.stringify(i),before);
 }
});
test('明确不同批号的同OA同额申请独立匹配，同一批重复申请合并日期并保留进度行',()=>{
 const i=sample([50,50],[['2026-11-01',50,'第1批','否','否'],['2026-11-02',50,'第2批','2026-11-03','2026-11-04'],['2026-11-03',50,'第2批','2026-11-03','2026-11-04']]),r=C.build(i);
 assert.deepEqual(r.highlights.negotiatedCells,['M2','N2']);assert.deepEqual(r.highlights.negotiatedPaidCells,['N2']);assert.deepEqual(states(r),['未付款','已付款','已付款']);
 assert.equal(r.batchComments.N2[1],P[1]+'2026年11月2日；2026年11月3日');assert.equal(r.batchComments.N2[2],P[2]+'2026年11月3日');assert.equal(r.batchComments.N2[3],P[3]+'2026年11月4日');assert.equal(r.reviewRows.filter(row=>row[5].v.startsWith('重复提交')).length,1);
});
test('原申请日期唯一定位同额批次，同额不同日期记录不先合并再错误消耗一批',()=>{
 const i=sample([50,50],[['2026-11-01',50,'','是','否'],['2026-11-02',50,'','是','2026-11-05']]),s=source(i);
 s.M2.c=[{a:'人工',t:P[1]+'2026年11月2日'}];s.N2.c=[{a:'人工',t:P[1]+'2026年11月1日'}];
 const r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,['N2','M2']);assert.deepEqual(r.highlights.negotiatedPaidCells,['M2']);assert.equal(r.batchComments.M2[3],P[3]+'2026年11月5日');assert.equal(r.batchComments.N2[3],P[3]+'未付款');assert(!r.reviewRows.some(row=>row[5].v.startsWith('重复提交')));
});
test('同额无唯一依据继续申请黄，不把付款时间挂到首批，进度仍按本条申请显示已付款',()=>{
 const i=sample([50,50],[['2026-11-01',50,'','2026-11-02','2026-11-03']]),r=C.build(i);
 assert.deepEqual(r.highlights.negotiatedCells,['M2']);assert.deepEqual(r.highlights.negotiatedPaidCells,[]);assert.equal(r.batchComments.M2[3],P[3]+'未记录');assert.equal(r.batchComments.M2[2],P[2]+'未记录');assert.deepEqual(states(r),['已付款']);assert.equal(r.applicationBatchMatches['付款申请汇总!2'].unique,false);assert(r.reviewRows.some(row=>row[5].v.includes('同额批次缺少唯一付款归属')));
});
test('再次上传不把上次推测批次写入的申请日期当成已付款归属证据',()=>{
 const i=sample([50,50],[['2026-11-01',50,'','是','是']]),first=C.build(i);
 assert.equal(first.applicationBatchMatches['付款申请汇总!2'].unique,false);
 assert.match(first.batchComments.M2[1],/批次归属待核对/);
 for(const [address,lines] of Object.entries(first.batchComments))source(i)[address].c=[{a:'付款节奏跟进',t:lines.join('\n')}];
 const second=C.build(i);assert.deepEqual(second.highlights.negotiatedPaidCells,[]);
 assert.equal(second.applicationBatchMatches['付款申请汇总!2'].unique,false);
});
test('批号与原日期证据冲突或原日期指向两批时不标红',()=>{
 for(const explicit of ['', '第1批']){
  const i=sample([50,50],[['2026-11-01',50,explicit,'是','是']]),s=source(i);s.N2.c=[{a:'人工',t:P[1]+'2026年11月1日'}];if(!explicit)s.M2.c=[{a:'人工',t:P[1]+'2026年11月1日'}];
  const r=C.build(i);assert.deepEqual(r.highlights.negotiatedPaidCells,[]);assert(r.reviewRows.some(row=>/唯一付款归属|批号与原申请日期批注冲突/.test(row[5].v)));
 }
 const i=sample([50,50],[['2026-11-01',50,'第2批','是','是']]);for(const address of ['M2','N2'])source(i)[address].c=[{a:'人工',t:P[1]+'2026年11月1日'}];assert.deepEqual(C.build(i).highlights.negotiatedPaidCells,['N2'],'明确批号可区分共用申请日期的两批');
});
test('同OA同额但供应商或主体不同不借用申请付款字段染红，申请进度仍按自身来源显示',()=>{
 for(const [header,value]of [['供应商全称','其他供应商'],['付款主体','国源']]){
  const i=sample();set(i,'applications',header,2,value);const r=C.build(i);assert.deepEqual(r.highlights.negotiatedPaidCells,[]);assert.equal(states(r)[0],'已付款');assert.equal(r.batchComments.M2?.[3],P[3]+'未记录');assert(r.reviewRows.some(row=>row[5].v.includes('付款申请归属不明')));
 }
});
test('无当次申请的已有批注保留三项历史工具信息，仅重新编号，未批注格不新增批注',()=>{
 const i=sample([40,60],[]),s=source(i),historical=[P[0]+'共9批，第8批',P[1]+'2026年10月20日',P[2]+'2026年10月21日',P[3]+'2026年10月22日'];s.N2.c=[{a:'人工',t:'人工原文\n'+historical.join('\n')}];
 const before=JSON.stringify(i),r=C.build(i);assert.deepEqual(r.batchComments.N2,[P[0]+'共2批，第2批',...historical.slice(1)]);assert(!r.batchComments.M2);assert.deepEqual(r.applicationComments,{});assert.equal(JSON.stringify(i),before);
});
test('是或付款日期与同批否异常冲突时不染红，各申请进度仍独立读取',()=>{
 for(const negative of ['否','待确认']){
  const i=sample([40,60],[['2026-11-01',40,'第1批','是','是'],['2026-11-02',40,'第1批','是',negative]]),r=C.build(i);
  assert.deepEqual(r.highlights.negotiatedPaidCells,[]);assert.equal(states(r)[0],'已付款');assert.equal(states(r)[1],negative==='否'?'未付款':'待核对');assert(r.reviewRows.some(row=>row[5].v.includes('相反或异常内容')));
 }
});
test('事件日期支持时分秒和1904数值，异常时刻不冒充付款时间',()=>{
 const i=sample([40,60],[['2026-11-01',40,'第1批','2026-11-16 09:05','2026-11-17 10:06:07']]);let r=C.build(i);
 assert.equal(r.batchComments.M2[2],P[2]+'2026年11月16日 09:05');assert.equal(r.batchComments.M2[3],P[3]+'2026年11月17日 10:06:07');
 i.applications.book.Workbook={WBProps:{date1904:true}};set(i,'applications','是否已付款',2,46343-1462+10/24+6/1440);r=C.build(i);assert.equal(r.batchComments.M2[3],P[3]+'2026年11月17日 10:06');
 set(i,'applications','是否已付款',2,'2026-11-17 25:00');assert.deepEqual(states(C.build(i)),['待核对']);
});
test('无缓存公式的占位零不作为付款日期，1904日期系统同样待核对',()=>{
 const i=sample();i.applications.book.Workbook={WBProps:{date1904:true}};
 const sheet=i.applications.book.Sheets['付款申请汇总'];
 sheet.H2={t:'z',f:'DATE(2026,11,15)',v:0};sheet.I2={t:'z',f:'DATE(2026,11,16)',v:0};
 const r=C.build(i);assert.deepEqual(states(r),['待核对']);assert.deepEqual(r.highlights.negotiatedPaidCells,[]);
 assert.match(r.batchComments.M2[2],/待核对/);assert.match(r.batchComments.M2[3],/待核对/);
});
test('年份冲突必须核对，确认仅改变排序年月，原表头与五组合计仍读源数据',()=>{
 const i=sample([40,60],[['2026-11-01',40,'第2批','是','是']]),s=source(i);s.M1.v='2026-04-12';s.N1.v='2027-03-15';X.utils.sheet_add_aoa(s,[['2027年4月支付'],[40]],{origin:'P1'});
 const before=JSON.stringify(i),initial=C.build(i),q=initial.issues.find(q=>q.kind==='batch-year');assert(q);assert.equal(q.resolved,false);assert(q.choices.some(c=>c.value==='year:2027'));
 const r=C.build(i,{[q.id]:['year:2027']});assert.equal(r.issues.find(x=>x.id===q.id).resolved,true);assert.deepEqual(r.negotiated.batchPlans[0].cells.map(b=>b.address),['N2','M2']);assert.equal(r.batchComments.M2[0],P[0]+'共2批，第2批');assert.deepEqual(r.negotiated.batchYearAdjustments,[{address:'M1',originalYear:2026,year:2027,date:'2027-04-12'}]);assert.equal(JSON.stringify(i),before);
 const keep=C.build(i,{[q.id]:['keep']});assert.deepEqual(keep.negotiated.batchPlans[0].cells.map(b=>b.address),['M2','N2']);
});
test('付款来源列紧跟申请日期，旧文件缺新列待核对，工具历史付款行不补判',()=>{
 const i=fixture(),s=source(i);s.M2.c=[{a:'人工',t:P[3]+'2026年11月3日'}];const r=C.build(i);
 assert.equal(r.progressHeaders.length,13);assert.deepEqual(r.progressHeaders.slice(0,2),['申请日期','付款来源']);assert.deepEqual(r.progressRows.map(row=>row[1].v),['谈判付款表','资金台账','谈判付款表','资金台账']);assert.equal(states(r)[0],'待核对');assert.equal(r.totalValidation.headers.length,C.TOTAL_VALIDATION_HEADERS.length);assert.equal(C.OUTPUT_SHEETS.includes('总额校验核对'),false);
});
