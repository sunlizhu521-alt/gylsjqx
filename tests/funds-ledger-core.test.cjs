const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-ledger-core.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const H=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否谈判付款'];
const row=(o={})=>({'供应商名称':'供应商甲','付款主体':'国源','对公付款0A流程号':'OA001','申请付款金额':100,'摘要':'货款','应付款时间':'2026-11-06','预估付款时间':'2026-12','是否谈判付款':'否',...o});
function sheet(rows,headers=H,title=false){return X.utils.aoa_to_sheet([...(title?[['台账标题']]:[]),headers,...rows.map(r=>headers.map(h=>r[h]??''))]);}
function book(rows=[]){const b=X.utils.book_new();PEOPLE.forEach((p,i)=>X.utils.book_append_sheet(b,sheet(i?[]:rows),'对账明细-'+p));return b;}
const build=(b,month='2026-11')=>C.build(b,'台账.xlsx',{month,today:'2026-11-09'});
const rows=r=>r.sheets[0].rows;
function set(b,h,cell,r=2){b.Sheets['对账明细-李奇'][X.utils.encode_cell({r:r-1,c:H.indexOf(h)})]=cell;}
test('单月单表12列，仅谈判否筛选，已结清仍保留且结清列无需存在或可含错误',()=>{
 const records=[row({'是否结清':'是'}),row({'是否结清':'否'}),row({'是否结清':{t:'e',v:23}}),row({'是否谈判付款':' 否 '}),row({'是否谈判付款':'是'}),row({'是否谈判付款':''}),row({'是否谈判付款':'正常付款'})],b=book();b.Sheets['对账明细-李奇']=sheet(records,[...H,'是否结清']);
 const before=JSON.stringify(b),r=build(b);assert.equal(r.sheets.length,1);assert.equal(r.sheets[0].name,'2026年11月正常付款');assert.equal(rows(r).length,4);assert(rows(r).every(row=>row.length===12));assert.deepEqual(r.counts[0],{sheet:'对账明细-李奇',included:4,excluded:3,undated:0});assert.equal(JSON.stringify(b),before);assert.equal(rows(build(book([row()]))).length,1);
 assert.deepEqual(C.HEADERS,['供应商','付款主体','对公付款OA流程号','申请付款金额','摘要','应付款时间','需要提交付款审单的日期','当前节点','未操作者','流程状态','是否已付款','是否提交纸质付款审单']);assert(rows(r).every(row=>row.slice(7).every(c=>c.v==='')));
});
test('应付款时间优先，空才用预估日期；不依OA分表，无OA与未提交OA仍可入当月',()=>{
 const r=build(book([row(),row({'对公付款0A流程号':'','应付款时间':'','预估付款时间':'2026-11-02'}),row({'对公付款0A流程号':'未提交OA','应付款时间':'2026-11-03','预估付款时间':'2026-12-01'}),row({'应付款时间':'2026-12-01','预估付款时间':'2026-11-01'})]));
 assert.deepEqual(rows(r).map(row=>C.formatCell(row[5])),['2026年11月02日','2026年11月03日','2026年11月06日']);assert.deepEqual(rows(r).map(row=>row[2].v),['','未提交OA','OA001']);assert.equal(r.counts[0].excluded,1);
});
test('跨月按最终日期分配，升序且同日按工作簿和来源行顺序，保留重复记录',()=>{
 const b=book([row({'摘要':'李奇晚','应付款时间':'2026-11-20'}),row({'摘要':'李奇早','应付款时间':'2026-11-02'}),row({'摘要':'李奇重复','应付款时间':'2026-11-02'}),row({'应付款时间':'2026-10-31'}),row({'应付款时间':'2026-12-01'})]);
 b.Sheets['对账明细-孙立柱']=sheet([row({'摘要':'孙先','应付款时间':{t:'n',v:46328.75,z:'yyyy-mm-dd'}})]);b.SheetNames.unshift(b.SheetNames.pop());
 const r=build(b);assert.deepEqual(rows(r).map(row=>row[4].v),['孙先','李奇早','李奇重复','李奇晚']);assert.deepEqual(r.sheets[0].sources.map(s=>[s.sheet,s.row]),[['对账明细-孙立柱',2],['对账明细-李奇',3],['对账明细-李奇',4],['对账明细-李奇',2]]);
 assert.equal(rows(build(b,'2026-10')).length,1);assert.equal(rows(build(b,'2026-12')).length,1);assert.equal(rows(build(b,'2027-01')).length,0);
});
test('年月精度保留，不转换成虚构日期；月精度在当月具体日期前按来源序',()=>{
 const b=book([row({'应付款时间':'2026年11月','摘要':'月精度'}),row({'应付款时间':'2026-11-01','摘要':'日精度'})]);const r=build(b);assert.equal(C.formatCell(rows(r)[0][5]),'2026年11月');assert.equal(rows(r)[0][4].v,'月精度');
 for(const value of [{t:'n',v:46332,z:'yyyy-mm-dd'},'2026-11-06',{t:'d',v:new Date('2026-11-06T00:00:00Z')}]){
  const n=book();n.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款月份':value})],H.map(h=>h==='预估付款时间'?'预估付款月份':h));assert.equal(C.formatCell(rows(build(n))[0][5]),'2026年11月');
 }
});
test('应付与预估均空时按月末加账期加7归月，但仅显示未对账待定；有无预估列规则一致',()=>{
 const headers=[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期'],b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':'60天'}),row({'应付款时间':'','所属月份':'2026-10','账期':'5个月账期'}),row({'应付款时间':'','所属月份':'2024-02','账期':0})],headers);
 for(const [month,sourceRow]of [['2027-01',2],['2027-04',3],['2024-03',4]]){const r=build(b,month);assert.equal(rows(r).length,1);assert.equal(C.formatCell(rows(r)[0][5]),'未对账待定');assert.equal(r.sheets[0].sources[0].row,sourceRow);}
 const existing=book();existing.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款时间':'','所属月份':'2026-10','账期':'60天'})],[...H,'所属月份','账期']);const r=build(existing,'2027-01');assert.equal(rows(r).length,1);assert.equal(rows(r)[0][5].v,'未对账待定');assert.equal(r.counts[0].undated,0);assert.match(r.warnings.join('\n'),/台账.xlsx.*李奇 第2行.*未对账待定.*留空.*补充具体付款日期/);
 const out=C.workbook(r),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});assert.equal(read.Sheets[out.SheetNames[0]].F2.t,'s');assert.equal(read.Sheets[out.SheetNames[0]].F2.v,'未对账待定');assert.equal(rows(build(existing,'2026-12')).length,0);
});
test('推算日期与真实日期一起升序，推算同日仍保留来源顺序，源数据不变',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([
  row({'摘要':'晚','应付款时间':'2026-11-07'}),
  row({'摘要':'推算一','应付款时间':'','预估付款时间':'','所属月份':'2026-09','账期':'票到30天'}),
  row({'摘要':'明确预估','应付款时间':'','预估付款时间':'2026-11-06'}),
  row({'摘要':'推算二','应付款时间':'','预估付款时间':'','所属月份':'2026-09','账期':'票到30天'}),
  row({'摘要':'早','应付款时间':'2026-11-05'})
 ],[...H,'所属月份','账期']);const before=JSON.stringify(b),r=build(b);
 assert.deepEqual(rows(r).map(row=>row[4].v),['早','推算一','明确预估','推算二','晚']);
 assert.deepEqual(rows(r).map(row=>C.formatCell(row[5])),['2026年11月05日','未对账待定','2026年11月06日','未对账待定','2026年11月07日']);assert.equal(JSON.stringify(b),before);
 for(const bad of [{t:'n',f:'DATE(2026,11,6)'},{t:'e',v:23}]){set(b,'预估付款时间',bad,3);assert.throws(()=>build(b),/李奇!G3/);}
 set(b,'预估付款时间',{t:'s',f:'IF(TRUE,"",1)',v:''},3);assert.equal(rows(build(b))[1][5].v,'未对账待定');
});
test('账期范围/百分数/负数不误推日期，缺年月等异常按来源提示且不进任何月',()=>{
 for(const terms of ['50%承兑','-30天','30/60天','30或60天','约60天','30天+30天']){const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':terms})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);const r=build(b,'2027-01');assert.equal(rows(r).length,0,terms);assert.equal(r.counts[0].undated,1);assert.match(r.warnings.join('\n'),/台账.xlsx.*李奇.*账期/);}
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':'60天，50%承兑'})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);assert.equal(rows(build(b,'2027-01')).length,1);
});
test('非空非法应付或预估不推算；无法定位月份记录单独undated并给来源',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'预计月底','预估付款时间':'2026-11-01'}),row({'应付款时间':'2026-02-30'}),row({'应付款时间':'','预估付款时间':''}),row({'应付款时间':'','预估付款时间':'待定','所属月份':'2026-09','账期':'30天'})],[...H,'所属月份','账期']);
 const r=build(b);assert.equal(rows(r).length,0);assert.equal(r.counts[0].undated,4);assert.equal(r.warnings.length,4);assert(r.warnings.every(w=>/台账.xlsx.*李奇/.test(w)));assert.match(r.warnings[0],/预计月底.*未归入任何月份/);assert.match(r.warnings[3],/待定.*未归入任何月份/);
});
test('1904与typed Date正确归月和回读，金额0/空白/合法币种保留，不把非法金额变0',()=>{
 const b=book([row({'应付款时间':{t:'n',v:44870,z:'yyyy-mm-dd'},'申请付款金额':0}),row({'应付款时间':{t:'d',v:new Date('2026-11-06T00:00:00Z')},'申请付款金额':''}),row({'应付款时间':'2026-11-07','申请付款金额':'￥1,234.56'})]);b.Workbook={WBProps:{date1904:true}};const r=build(b);assert.deepEqual(rows(r).map(row=>row[3].v),[0,'',1234.56]);assert.equal(rows(r)[0][5].v,46332);assert.equal(C.formatCell(rows(r)[1][5]),'2026年11月06日');
 for(const bad of ['/',true,'1,23.45',Infinity,Number.MAX_SAFE_INTEGER])assert.throws(()=>build(book([row({'申请付款金额':bad})])),/台账.xlsx.*李奇!D2/);
});
test('有效公式缓存取值包括零和空字符串；必要错误阻断，不读结清/未选日期/排除款无关字段',()=>{
 const b=book([row()]);set(b,'申请付款金额',{t:'n',f:'1-1',v:0});set(b,'应付款时间',{t:'s',f:'""',v:''});set(b,'预估付款时间',{t:'n',f:'DATE(2026,11,6)',v:46332,z:'yyyy-mm-dd'});const before=JSON.stringify(b),r=build(b);assert.equal(rows(r)[0][3].v,0);assert(rows(r)[0].every(c=>c.f===undefined));assert.equal(JSON.stringify(b),before);
 for(const h of ['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','是否谈判付款'])for(const cell of [{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const n=book([row()]);set(n,h,cell);assert.throws(()=>build(n),/台账.xlsx.*李奇!/);}
 const n=book([row({'预估付款时间':{t:'n',f:'A99'}}),row({'是否谈判付款':'是','应付款时间':{t:'n',f:'A99'}}),row({'应付款时间':'2026-12-01','申请付款金额':{t:'n',f:'A99'}})]);assert.equal(rows(build(n)).length,1);
});
test('五人唯一表、日期后缀、标题行及字段别名，历史表不读取，冲突表头不任取',()=>{
 const b=book();const aliases={'供应商名称':'供应商','对公付款0A流程号':'对公付款OA流程号','申请付款金额':'本次申请付款金额','应付款时间':'应付时间'},headers=H.map(h=>aliases[h]||h).reverse(),record=Object.fromEntries(Object.entries(row()).map(([h,v])=>[aliases[h]||h,v]));
 b.Sheets['对账明细-徐英1102']=sheet([record],headers,true);delete b.Sheets['对账明细-徐英'];b.SheetNames[1]='对账明细-徐英1102';X.utils.book_append_sheet(b,sheet([row({'申请付款金额':'无效'})]),'采购付款台账_历史汇总');assert.deepEqual(build(b).sheets[0].sources,[{sheet:'对账明细-徐英1102',row:3}]);
 const missing=book();missing.SheetNames.pop();assert.throws(()=>build(missing),/孙立柱.*缺失/);const duplicate=book();X.utils.book_append_sheet(duplicate,sheet([]),'对账明细-李奇1102');assert.throws(()=>build(duplicate),/李奇.*多个版本/);
 for(const h of ['供应商名称','申请付款金额','对公付款0A流程号','是否谈判付款']){const n=book();n.Sheets['对账明细-李奇']=sheet([],H.filter(v=>v!==h));assert.throws(()=>build(n),/缺少必要表头/);}
 const n=book();n.Sheets['对账明细-李奇']=sheet([row({'供应商':'另一家'})],[...H,'供应商']);assert.throws(()=>build(n),/字段.*重复/);
});
test('OA按文本保留前导零，原始金额与重复行保留，空类别导出仍为单表12列',()=>{
 const b=book([row(),row()]);set(b,'对公付款0A流程号',{t:'n',v:123,z:'000000'});const r=build(b),out=C.workbook(r),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});assert.deepEqual(out.SheetNames,['2026年11月正常付款']);assert.equal(read.Sheets[out.SheetNames[0]].C2.v,'000123');assert.equal(read.Sheets[out.SheetNames[0]].C2.t,'s');assert.equal(read.Sheets[out.SheetNames[0]].D2.t,'n');assert.equal(read.Sheets[out.SheetNames[0]].F2.t,'n');assert.equal(read.Sheets[out.SheetNames[0]].F2.z,'yyyy"年"mm"月"dd"日"');assert.equal(read.Sheets[out.SheetNames[0]].F2.w,'2026年11月06日');assert.equal(out.Sheets[out.SheetNames[0]]['!autofilter'].ref,'A1:L3');assert.equal(out.Sheets[out.SheetNames[0]]['!cols'].length,12);
 const empty=C.workbook(build(book(),'2027-01'));assert.deepEqual(empty.SheetNames,['2027年1月正常付款']);assert.equal(empty.Sheets[empty.SheetNames[0]]['!ref'],'A1:L1');assert.equal(C.exportFilename('2026-11-08',3),'资金台账整理-2026-11-08-V3.xlsx');
});
test('北京时间当前月份默认，空值及不合法年月阻断，表头列宽可见且长摘要换行',()=>{
 const now=new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit'}).format(new Date());assert.equal(C.build(book(),'台账.xlsx').month,now);
 for(const month of ['',null,'2026-00','2026-13','1899-12','10000-01','2026/11'])assert.throws(()=>C.build(book(),'台账.xlsx',{month}),/有效月份/);
 const b=book([row({'摘要':'摘要'.repeat(100)})]);b.Sheets['对账明细-李奇'].XFD1048576={t:'s',v:'',s:1};b.Sheets['对账明细-李奇']['!ref']='A1:XFD1048576';const s=C.workbook(build(b)).Sheets['2026年11月正常付款'];assert.equal(s['!cols'][4].wch,72);assert(s['!cols'].every(c=>c.wch>=12));assert.equal(s.E2.s.alignment.wrapText,true);
});
test('审单日期：本周及逾期取本周一，未来取前一周周一，周日和周一边界正确',()=>{
 const cases=[['01','09'],['06','09'],['09','09'],['12','09'],['15','09'],['16','09'],['18','09'],['22','09'],['23','16'],['25','16'],['30','23']];
 const b=book(cases.map(([due])=>row({'应付款时间':'2026-11-'+due}))),before=JSON.stringify(b),result=build(b);
 assert.deepEqual(rows(result).map(r=>C.formatCell(r[6])),cases.map(([,day])=>'2026年11月'+day+'日'));assert.equal(JSON.stringify(b),before);
 assert(rows(result).every(r=>r[6].t==='n'&&Number.isInteger(r[6].v)&&r[6].z==='yyyy"年"mm"月"dd"日"'&&!r[6].f));
 const out=C.workbook(result),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}),s=read.Sheets[read.SheetNames[0]];
 assert.equal(s.G1.v,'需要提交付款审单的日期');assert.equal(s.G2.t,'n');assert.equal(s.G2.w,'2026年11月09日');assert.equal(s.G2.z,'yyyy"年"mm"月"dd"日"');assert.equal(s['!autofilter'].ref,'A1:L12');assert(out.Sheets[out.SheetNames[0]]['!cols'][6].wch>=26);
});
test('审单日期跨月跨年，不以选中月份代替当前周；同一文件基准日期变化后重算',()=>{
 for(const [today,due,expected]of [['2026-12-31','2027-01-01','2026年12月28日'],['2026-12-31','2027-01-11','2027年01月04日'],['2026-03-01','2026-03-02','2026年02月23日'],['2026-03-01','2026-03-09','2026年03月02日'],['2026-11-09','2026-10-01','2026年11月09日']]){
  const result=C.build(book([row({'应付款时间':due})]),'跨月台账.xlsx',{month:due.slice(0,7),today});assert.equal(C.formatCell(rows(result)[0][6]),expected);
 }
 const b=book([row({'应付款时间':'2026-11-06'})]);for(const [today,expected]of [['2026-11-08','2026年11月02日'],['2026-11-09','2026年11月09日'],['2026-11-16','2026年11月16日']])assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{month:'2026-11',today}))[0][6]),expected);
 for(const today of ['',null,'2026-11','2026-11-9','2026-02-30','2026-13-01','2026-00-01','0000-01-01'])assert.throws(()=>C.build(b,'台账.xlsx',{month:'2026-11',today}),/基准.*有效日期/);
});
test('缺日或内部推算不生成审单日期，明确预估可计算，提示来源但仍可导出',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([
  row({'应付款时间':'2026年11月'}),
  row({'应付款时间':{t:'n',v:46332,z:'yyyy-mm'}}),
  row({'应付款时间':'','预估付款时间':'2026-11'}),
  row({'应付款时间':'','预估付款时间':'','所属月份':'2026-08','账期':'60天'}),
  row({'应付款时间':'','预估付款时间':'2026-11-25'})
 ],[...H,'所属月份','账期']);const result=build(b);
 assert.deepEqual(rows(result).map(r=>C.formatCell(r[6])),['','','','','2026年11月16日']);assert.equal(result.warnings.length,4);
 for(const sourceRow of [2,3,4,5])assert(result.warnings.some(w=>w.includes('台账.xlsx / 对账明细-李奇 第'+sourceRow+'行')&&w.includes('补充具体付款日期')));
 const out=C.workbook(result);assert.equal(out.Sheets[out.SheetNames[0]].G2.v,'');assert.equal(out.Sheets[out.SheetNames[0]].G6.t,'n');
 const formatted=book();formatted.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款月份':'2026-11-25'})],H.map(h=>h==='预估付款时间'?'预估付款月份':h));assert.equal(rows(build(formatted))[0][6].v,'');
});
test('审单日期用有效公式及1904日期的日历日期，不带时分秒',()=>{
 const b=book([row({'应付款时间':{t:'n',f:'DATE(2026,11,25)',v:44889.75,z:'yyyy-mm-dd hh:mm'}})]);b.Workbook={WBProps:{date1904:true}};
 const result=build(b);assert.equal(C.formatCell(rows(result)[0][5]),'2026年11月25日');assert.equal(C.formatCell(rows(result)[0][6]),'2026年11月16日');assert(!rows(result)[0][6].f);
});
test('默认基准按北京时间，跨周午夜立即变更，不依机器时区',t=>{
 t.mock.timers.enable({apis:['Date'],now:new Date('2026-11-08T15:59:59Z')});const b=book([row()]);
 assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{month:'2026-11'}))[0][6]),'2026年11月02日');
 t.mock.timers.setTime(Date.parse('2026-11-08T16:00:00Z'));
 assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{month:'2026-11'}))[0][6]),'2026年11月09日');
 assert.equal(C.build(book(),'台账.xlsx',{today:'2027-01-01'}).month,'2027-01');
});
