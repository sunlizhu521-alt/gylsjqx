const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-ledger-core.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const H=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否谈判付款'];
const row=(o={})=>({'供应商名称':'供应商甲','付款主体':'国源','对公付款0A流程号':'OA001','申请付款金额':100,'摘要':'货款','应付款时间':'2026-11-06','预估付款时间':'2026-12','是否谈判付款':'否',...o});
function sheet(rows,headers=H,title=false){return X.utils.aoa_to_sheet([...(title?[['台账标题']]:[]),headers,...rows.map(r=>headers.map(h=>r[h]??''))]);}
function book(rows=[]){const b=X.utils.book_new();PEOPLE.forEach((p,i)=>X.utils.book_append_sheet(b,sheet(i?[]:rows),'对账明细-'+p));return b;}
const build=b=>C.build(b,'台账.xlsx',{today:'2026-11-09'});
// Existing payment-field assertions exclude the newly prepended source owner.
const rows=r=>r.sheets[0].rows.map(row=>row.slice(1));
for(const date1904 of [false,true])test(`实际导出日期为文本，${date1904?'1904':'1900'}来源、金额零值及内部日期不变`,async()=>{
 const Export=require('../funds-ledger-export.js');
 const b=book([row({'应付款时间':{t:'n',v:date1904?44870:46332,z:'yyyy-mm-dd',f:'DATE(2026,11,6)'},'申请付款金额':0}),row({'应付款时间':'2026年12月'}),row({'应付款时间':'待定','申请付款金额':''})]);
 b.Workbook={WBProps:{date1904}};
 const result=build(b),before=JSON.stringify(result),bytes=await Export(result),out=X.read(bytes,{type:'array',cellNF:true}),s=out.Sheets.正常付款;
 assert.deepEqual(out.SheetNames,['正常付款']);
 assert.deepEqual([s.G2.v,s.L2.v,s.M2.v],['2026年11月06日','2026年11月','2026年11月09日']);
 for(const address of ['G2','L2','M2','G3']){assert.equal(s[address].t,'s');assert.equal(s[address].z,'@');assert.equal(s[address].f,undefined);}
 assert.equal(s.G3.v,'2026年12月');assert.equal(s.L3.v,'');assert.equal(s.M3.v,'');assert.equal(s.G4.v,'待定');assert.equal(s.E4.v,'');
 assert.equal(s.E2.t,'n');assert.equal(s.E2.v,0);assert.equal(s.E3.t,'n');assert.equal(s.E3.v,100);
 assert.equal([s.G2,s.L2,s.M2,s.E2].filter(c=>c.t==='n').reduce((sum,c)=>sum+c.v,0),0);
 assert.equal(JSON.stringify(result),before);assert.equal(result.sheets[0].rows[0][6].t,'n');assert.equal(C.workbook(result).Sheets.正常付款.G2.t,'n');
 const zip=await require('../vendor/jszip.min.js').loadAsync(bytes),styles=await zip.file('xl/styles.xml').async('string'),xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
 const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g),borders=styles.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)[1].match(/<border>[\s\S]*?<\/border>/g);
 for(const c of xml.match(/<c\b[^>]+>/g)){
  const address=c.match(/\br="([A-Z]+\d+)"/)[1],row=Number(address.match(/\d+$/)[0]),style=xfs[Number(c.match(/\bs="(\d+)"/)[1])],border=borders[Number(style.match(/\bborderId="(\d+)"/)[1])];
  for(const side of ['left','right','top','bottom'])assert.match(border,new RegExp(`<${side} style="${row===1?'thin':'hair'}"><color rgb="FF404040"/></${side}>`),address);
 }
});
test('有内容行列交叉处含空白格都画边框，全空行列不画，零和false有效且红色不变',async()=>{
 const Export=require('../funds-ledger-export.js'),result=build(book([row({'申请付款金额':0})]));
 result.sheets[0].rows[0][5]={t:'s',v:' \n '};result.sheets[0].rows[0][7]={t:'b',v:false};result.sheets[0].rows[0][10]={t:'s',v:'已付款'};
 const emptyRow=()=>Array.from({length:14},()=>({t:'s',v:' \n '})),zeroRow=emptyRow(),falseRow=emptyRow();zeroRow[4]={t:'n',v:0};falseRow[7]={t:'b',v:false};
 result.sheets[0].rows.push(emptyRow(),zeroRow,falseRow);result.sheets[0].headers=[...result.sheets[0].headers];result.sheets[0].headers[13]=' ';
 const before=JSON.stringify(result),bytes=await Export(result),out=X.read(bytes,{type:'array',cellNF:true,cellStyles:true}),s=out.Sheets.正常付款;
 const zip=await require('../vendor/jszip.min.js').loadAsync(bytes),styles=await zip.file('xl/styles.xml').async('string'),xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
 const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
 for(const [address,borderId] of [['A1',1],['N1',1],['E2',2],['F2',2],['H2',2],['I2',2],['N2',0],['A3',0],['E3',0],['A4',2],['E4',2],['N4',0],['A5',2],['H5',2],['N5',0]]){
  const tag=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0],style=xfs[Number(tag.match(/\bs="(\d+)"/)[1])];assert.match(style,new RegExp(`borderId="${borderId}"`),address);
  if(address.endsWith('2'))assert.equal(s[address].s.fgColor.rgb,'FF9999',address);
 }
 assert.equal(s['!ref'],'A1:N5');assert.equal(s.O2,undefined);assert.equal(s.A6,undefined);
 assert.equal(s.E2.v,0);assert.equal(s.E2.t,'n');assert.equal(s.E2.z,result.sheets[0].rows[0][4].z);assert.equal(s.F2.v,' \n ');assert.equal(s.H2.v,false);assert.equal(s.H2.t,'b');assert.equal(s.I2.v,'');assert.equal(JSON.stringify(result),before);
});
function set(b,h,cell,r=2){b.Sheets['对账明细-李奇'][X.utils.encode_cell({r:r-1,c:H.indexOf(h)})]=cell;}
test('全部月份单表14列，仅谈判否筛选，已结清仍保留且结清列无需存在或可含错误',()=>{
 const records=[row({'是否结清':'是'}),row({'是否结清':'否'}),row({'是否结清':{t:'e',v:23}}),row({'是否谈判付款':' 否 '}),row({'是否谈判付款':'是'}),row({'是否谈判付款':''}),row({'是否谈判付款':'正常付款'})],b=book();b.Sheets['对账明细-李奇']=sheet(records,[...H,'是否结清']);
 const before=JSON.stringify(b),r=build(b);assert.equal(r.sheets.length,1);assert.equal(r.sheets[0].name,'正常付款');assert.equal(rows(r).length,4);assert(r.sheets[0].rows.every(row=>row.length===14));assert.deepEqual(r.counts[0],{sheet:'对账明细-李奇',included:4,excluded:3,undated:0});assert.equal(JSON.stringify(b),before);assert.equal(rows(build(book([row()]))).length,1);
 assert.deepEqual(C.HEADERS,['采购跟单员','供应商','付款主体','对公付款OA流程号','申请付款金额','摘要','应付款时间','当前节点','未操作者','流程状态','是否已付款','申请付款月份','需要提交付款审单的日期','是否提交纸质付款审单']);assert(rows(r).every(row=>row.filter((_,c)=>c>=6&&c!==10&&c!==11).every(c=>c.v==='')));
});
test('应付款时间优先，空才用预估日期；不依OA分表，无OA与未提交OA均保留',()=>{
 const r=build(book([row(),row({'对公付款0A流程号':'','应付款时间':'','预估付款时间':'2026-11-02'}),row({'对公付款0A流程号':'未提交OA','应付款时间':'2026-11-03','预估付款时间':'2026-12-01'}),row({'应付款时间':'2026-12-01','预估付款时间':'2026-11-01'})]));
 assert.deepEqual(rows(r).map(row=>C.formatCell(row[5])),['2026年11月02日','2026年11月03日','2026年11月06日','2026年12月01日']);assert.deepEqual(rows(r).map(row=>row[2].v),['','未提交OA','OA001','OA001']);assert.equal(r.counts[0].excluded,0);
});
test('跨月记录全部保留，升序且同日按工作簿和来源行顺序，保留重复记录',()=>{
 const b=book([row({'摘要':'李奇晚','应付款时间':'2026-11-20'}),row({'摘要':'李奇早','应付款时间':'2026-11-02'}),row({'摘要':'李奇重复','应付款时间':'2026-11-02'}),row({'应付款时间':'2026-10-31'}),row({'应付款时间':'2026-12-01'})]);
 b.Sheets['对账明细-孙立柱']=sheet([row({'摘要':'孙先','应付款时间':{t:'n',v:46328.75,z:'yyyy-mm-dd'}})]);b.SheetNames.unshift(b.SheetNames.pop());
 const r=build(b);assert.deepEqual(rows(r).map(row=>row[4].v),['货款','孙先','李奇早','李奇重复','李奇晚','货款']);assert.deepEqual(r.sheets[0].sources.map(s=>[s.sheet,s.row]),[['对账明细-李奇',5],['对账明细-孙立柱',2],['对账明细-李奇',3],['对账明细-李奇',4],['对账明细-李奇',2],['对账明细-李奇',6]]);
 assert.equal(rows(build(b)).length,6);
});
test('年月精度保留，不转换成虚构日期；月精度在当月具体日期前按来源序',()=>{
 const b=book([row({'应付款时间':'2026年11月','摘要':'月精度'}),row({'应付款时间':'2026-11-01','摘要':'日精度'})]);const r=build(b);assert.equal(C.formatCell(rows(r)[0][5]),'2026年11月');assert.equal(rows(r)[0][4].v,'月精度');
 for(const value of [{t:'n',v:46332,z:'yyyy-mm-dd'},'2026-11-06',{t:'d',v:new Date('2026-11-06T00:00:00Z')}]){
  const n=book();n.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款月份':value})],H.map(h=>h==='预估付款时间'?'预估付款月份':h));assert.equal(C.formatCell(rows(build(n))[0][5]),'2026年11月');
 }
});
test('应付与预估均空时按月末加账期加7归月，但仅显示未对账待定；有无预估列规则一致',()=>{
 const headers=[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期'],b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':'60天'}),row({'应付款时间':'','所属月份':'2026-10','账期':'5个月账期'}),row({'应付款时间':'','所属月份':'2024-02','账期':0})],headers);
 const inferred=build(b);assert.equal(rows(inferred).length,3);assert(rows(inferred).every(r=>r[5].v==='未对账待定'));assert.deepEqual(inferred.sheets[0].sources.map(s=>s.row),[4,2,3]);
 const existing=book();existing.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款时间':'','所属月份':'2026-10','账期':'60天'})],[...H,'所属月份','账期']);const r=build(existing);assert.equal(rows(r).length,1);assert.equal(rows(r)[0][5].v,'未对账待定');assert.equal(r.counts[0].undated,0);assert.match(r.warnings.join('\n'),/台账.xlsx.*李奇 第2行.*未对账待定.*留空.*补充具体付款日期/);
 const out=C.workbook(r),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});assert.equal(read.Sheets[out.SheetNames[0]].G2.t,'s');assert.equal(read.Sheets[out.SheetNames[0]].G2.v,'未对账待定');
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
test('账期范围/百分数/负数不误推日期，缺年月等异常按来源提示且保留原记录',()=>{
 for(const terms of ['50%承兑','-30天','30/60天','30或60天','约60天','30天+30天']){const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':terms})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);const r=build(b);assert.equal(rows(r).length,1,terms);assert.equal(rows(r)[0][10].v,'');assert.equal(rows(r)[0][11].v,'');assert.equal(r.counts[0].undated,1);assert.match(r.warnings.join('\n'),/台账.xlsx.*李奇.*账期/);}
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','所属月份':'2026-10','账期':'60天，50%承兑'})],[...H.filter(h=>h!=='预估付款时间'),'所属月份','账期']);assert.equal(rows(build(b)).length,1);
});
test('非空非法应付或预估不推算；无法确定日期记录保留并给来源',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'预计月底','预估付款时间':'2026-11-01'}),row({'应付款时间':'2026-02-30'}),row({'应付款时间':'','预估付款时间':''}),row({'应付款时间':'','预估付款时间':'待定','所属月份':'2026-09','账期':'30天'})],[...H,'所属月份','账期']);
 const r=build(b);assert.equal(rows(r).length,4);assert.equal(r.counts[0].undated,4);assert(rows(r).every(row=>row[10].v===''&&row[11].v===''));assert.deepEqual(rows(r).map(row=>row[5].v),['预计月底','2026-02-30','未对账待定','待定']);assert(r.warnings.every(w=>/台账.xlsx.*李奇/.test(w)));assert.match(r.warnings.join('\n'),/预计月底.*保留原文/);assert.match(r.warnings.join('\n'),/待定.*保留原文/);for(const sourceRow of [2,3,4,5])assert(r.warnings.some(w=>w.includes('第'+sourceRow+'行')&&w.includes('留空')));
});
test('1904与typed Date正确归月和回读，金额0/空白/合法币种保留，不把非法金额变0',()=>{
 const b=book([row({'应付款时间':{t:'n',v:44870,z:'yyyy-mm-dd'},'申请付款金额':0}),row({'应付款时间':{t:'d',v:new Date('2026-11-06T00:00:00Z')},'申请付款金额':''}),row({'应付款时间':'2026-11-07','申请付款金额':'￥1,234.56'})]);b.Workbook={WBProps:{date1904:true}};const r=build(b);assert.deepEqual(rows(r).map(row=>row[3].v),[0,'',1234.56]);assert.equal(rows(r)[0][5].v,46332);assert.equal(C.formatCell(rows(r)[1][5]),'2026年11月06日');
 for(const bad of ['/',true,'1,23.45',Infinity,Number.MAX_SAFE_INTEGER])assert.throws(()=>build(book([row({'申请付款金额':bad})])),/台账.xlsx.*李奇!D2/);
});
test('有效公式缓存取值包括零和空字符串；必要错误阻断，不读结清/未选日期/排除款无关字段',()=>{
 const b=book([row()]);set(b,'申请付款金额',{t:'n',f:'1-1',v:0});set(b,'应付款时间',{t:'s',f:'""',v:''});set(b,'预估付款时间',{t:'n',f:'DATE(2026,11,6)',v:46332,z:'yyyy-mm-dd'});const before=JSON.stringify(b),r=build(b);assert.equal(rows(r)[0][3].v,0);assert(rows(r)[0].every(c=>c.f===undefined));assert.equal(JSON.stringify(b),before);
 for(const h of ['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','是否谈判付款'])for(const cell of [{t:'n',f:'A99'},{t:'e',v:23,w:'#REF!'}]){const n=book([row()]);set(n,h,cell);assert.throws(()=>build(n),/台账.xlsx.*李奇!/);}
 const n=book([row({'预估付款时间':{t:'n',f:'A99'}}),row({'是否谈判付款':'是','应付款时间':{t:'n',f:'A99'}})]);assert.equal(rows(build(n)).length,1);const otherMonth=book([row({'应付款时间':'2026-12-01','申请付款金额':{t:'n',f:'A99'}})]);assert.throws(()=>build(otherMonth),/李奇!D2.*公式/);
});
test('五人唯一表、日期后缀、标题行及字段别名，历史表不读取，冲突表头不任取',()=>{
 const b=book();const aliases={'供应商名称':'供应商','对公付款0A流程号':'对公付款OA流程号','申请付款金额':'本次申请付款金额','应付款时间':'应付时间'},headers=H.map(h=>aliases[h]||h).reverse(),record=Object.fromEntries(Object.entries(row()).map(([h,v])=>[aliases[h]||h,v]));
 b.Sheets['对账明细-徐英1102']=sheet([record],headers,true);delete b.Sheets['对账明细-徐英'];b.SheetNames[1]='对账明细-徐英1102';X.utils.book_append_sheet(b,sheet([row({'申请付款金额':'无效'})]),'采购付款台账_历史汇总');assert.deepEqual(build(b).sheets[0].sources,[{sheet:'对账明细-徐英1102',row:3}]);
 const missing=book();missing.SheetNames.pop();assert.throws(()=>build(missing),/孙立柱.*缺失/);const duplicate=book();X.utils.book_append_sheet(duplicate,sheet([]),'对账明细-李奇1102');assert.throws(()=>build(duplicate),/李奇.*多个版本/);
 for(const h of ['供应商名称','申请付款金额','对公付款0A流程号','是否谈判付款']){const n=book();n.Sheets['对账明细-李奇']=sheet([],H.filter(v=>v!==h));assert.throws(()=>build(n),/缺少必要表头/);}
 const n=book();n.Sheets['对账明细-李奇']=sheet([row({'供应商':'另一家'})],[...H,'供应商']);assert.throws(()=>build(n),/字段.*重复/);
});
test('OA按文本保留前导零，原始金额与重复行保留，空类别导出仍为单表14列',()=>{
 const b=book([row(),row()]);set(b,'对公付款0A流程号',{t:'n',v:123,z:'000000'});const r=build(b),out=C.workbook(r),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true});assert.deepEqual(out.SheetNames,['正常付款']);assert.equal(read.Sheets[out.SheetNames[0]].D2.v,'000123');assert.equal(read.Sheets[out.SheetNames[0]].D2.t,'s');assert.equal(read.Sheets[out.SheetNames[0]].E2.t,'n');assert.equal(read.Sheets[out.SheetNames[0]].G2.t,'n');assert.equal(read.Sheets[out.SheetNames[0]].G2.z,'yyyy"年"mm"月"dd"日"');assert.equal(read.Sheets[out.SheetNames[0]].G2.w,'2026年11月06日');assert.equal(out.Sheets[out.SheetNames[0]]['!autofilter'].ref,'A1:N3');assert.equal(out.Sheets[out.SheetNames[0]]['!cols'].length,14);
 const empty=C.workbook(build(book()));assert.deepEqual(empty.SheetNames,['正常付款']);assert.equal(empty.Sheets[empty.SheetNames[0]]['!ref'],'A1:N1');assert.equal(C.exportFilename('2026-11-08',3),'资金台账整理-2026-11-08-V3.xlsx');
});
test('不再按月份筛选，表头列宽可见且长摘要换行',()=>{
 assert(!Object.hasOwn(C.build(book(),'台账.xlsx'),'month'));const all=book([row({'应付款时间':'2026-10-01'}),row({'应付款时间':'2027-01-01'})]);assert.equal(rows(build(all)).length,2);
 const b=book([row({'摘要':'摘要'.repeat(100)})]);b.Sheets['对账明细-李奇'].XFD1048576={t:'s',v:'',s:1};b.Sheets['对账明细-李奇']['!ref']='A1:XFD1048576';const s=C.workbook(build(b)).Sheets['正常付款'];assert.equal(s['!cols'][5].wch,72);assert(s['!cols'].every(c=>c.wch>=12));assert.equal(s.F2.s.alignment.wrapText,true);
});
test('审单日期：本周及逾期取本周一，未来取前一周周一，周日和周一边界正确',()=>{
 const cases=[['01','09'],['06','09'],['09','09'],['12','09'],['15','09'],['16','09'],['18','09'],['22','09'],['23','16'],['25','16'],['30','23']];
 const b=book(cases.map(([due])=>row({'应付款时间':'2026-11-'+due}))),before=JSON.stringify(b),result=build(b);
 assert.deepEqual(rows(result).map(r=>C.formatCell(r[11])),cases.map(([,day])=>'2026年11月'+day+'日'));assert.equal(JSON.stringify(b),before);
 assert(rows(result).every(r=>r[11].t==='n'&&Number.isInteger(r[11].v)&&r[11].z==='yyyy"年"mm"月"dd"日"'&&!r[11].f));
 const out=C.workbook(result),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}),s=read.Sheets[read.SheetNames[0]];
 assert.equal(s.M1.v,'需要提交付款审单的日期');assert.equal(s.M2.t,'n');assert.equal(s.M2.w,'2026年11月09日');assert.equal(s.M2.z,'yyyy"年"mm"月"dd"日"');assert.equal(s['!autofilter'].ref,'A1:N12');assert(out.Sheets[out.SheetNames[0]]['!cols'][12].wch>=26);
});
test('审单日期跨月跨年，使用当前周边界；同一文件基准日期变化后重算',()=>{
 for(const [today,due,expected]of [['2026-12-31','2027-01-01','2026年12月28日'],['2026-12-31','2027-01-11','2027年01月04日'],['2026-03-01','2026-03-02','2026年02月23日'],['2026-03-01','2026-03-09','2026年03月02日'],['2026-11-09','2026-10-01','2026年11月09日']]){
  const result=C.build(book([row({'应付款时间':due})]),'跨月台账.xlsx',{today});assert.equal(C.formatCell(rows(result)[0][11]),expected);
 }
 const b=book([row({'应付款时间':'2026-11-06'})]);for(const [today,expected]of [['2026-11-08','2026年11月02日'],['2026-11-09','2026年11月09日'],['2026-11-16','2026年11月16日']])assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{today}))[0][11]),expected);
 for(const today of ['',null,'2026-11','2026-11-9','2026-02-30','2026-13-01','2026-00-01','0000-01-01'])assert.throws(()=>C.build(b,'台账.xlsx',{today}),/基准.*有效日期/);
});
test('缺日或内部推算不生成审单日期，明确预估可计算，提示来源但仍可导出',()=>{
 const b=book();b.Sheets['对账明细-李奇']=sheet([
  row({'应付款时间':'2026年11月'}),
  row({'应付款时间':{t:'n',v:46332,z:'yyyy-mm'}}),
  row({'应付款时间':'','预估付款时间':'2026-11'}),
  row({'应付款时间':'','预估付款时间':'','所属月份':'2026-08','账期':'60天'}),
  row({'应付款时间':'','预估付款时间':'2026-11-25'})
 ],[...H,'所属月份','账期']);const result=build(b);
 assert.deepEqual(rows(result).map(r=>C.formatCell(r[11])),['','','','','2026年11月16日']);assert.equal(result.warnings.length,4);
 for(const sourceRow of [2,3,4,5])assert(result.warnings.some(w=>w.includes('台账.xlsx / 对账明细-李奇 第'+sourceRow+'行')&&w.includes('补充具体付款日期')));
 const out=C.workbook(result);assert.equal(out.Sheets[out.SheetNames[0]].M2.v,'');assert.equal(out.Sheets[out.SheetNames[0]].M6.t,'n');
 const formatted=book();formatted.Sheets['对账明细-李奇']=sheet([row({'应付款时间':'','预估付款月份':'2026-11-25'})],H.map(h=>h==='预估付款时间'?'预估付款月份':h));assert.equal(rows(build(formatted))[0][11].v,'');
});
test('审单日期用有效公式及1904日期的日历日期，不带时分秒',()=>{
 const b=book([row({'应付款时间':{t:'n',f:'DATE(2026,11,25)',v:44889.75,z:'yyyy-mm-dd hh:mm'}})]);b.Workbook={WBProps:{date1904:true}};
 const result=build(b);assert.equal(C.formatCell(rows(result)[0][5]),'2026年11月25日');assert.equal(C.formatCell(rows(result)[0][11]),'2026年11月16日');assert(!rows(result)[0][11].f);
});
test('默认基准按北京时间，跨周午夜立即变更，不依机器时区',t=>{
 t.mock.timers.enable({apis:['Date'],now:new Date('2026-11-08T15:59:59Z')});const b=book([row()]);
 assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{}))[0][11]),'2026年11月02日');
 t.mock.timers.setTime(Date.parse('2026-11-08T16:00:00Z'));
 assert.equal(C.formatCell(rows(C.build(b,'台账.xlsx',{}))[0][11]),'2026年11月09日');
 assert(!Object.hasOwn(C.build(book(),'台账.xlsx',{today:'2027-01-01'}),'month'));
});

test('申请付款月份取提交审单日期所在月，是真实月精度日期，跨月跨年及逾期不取应付月份',()=>{
 const cases=[['2026-11-09','2026-12-01','2026年11月','2026年11月23日'],['2026-11-09','2026-12-08','2026年11月','2026年11月30日'],['2026-12-28','2027-01-04','2026年12月','2026年12月28日'],['2027-01-04','2026-12-01','2027年01月','2027年01月04日']];
 for(const [today,due,month,submission]of cases){
  const result=C.build(book([row({'应付款时间':due})]),'跨月.xlsx',{today}),record=rows(result)[0];assert.equal(C.formatCell(record[10]),month);assert.equal(C.formatCell(record[11]),submission);assert.equal(record[10].t,'n');assert.equal(record[10].z,'yyyy"年"mm"月"');assert.equal(X.SSF.parse_date_code(record[10].v).d,1);
  const out=C.workbook(result),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}),sheet=read.Sheets.正常付款;assert.equal(sheet.L1.v,'申请付款月份');assert.equal(sheet.L2.w,month);assert.equal(sheet.M2.w,submission);assert.equal(sheet.N1.v,'是否提交纸质付款审单');assert.equal(sheet['!autofilter'].ref,'A1:N2');
 }
});
test('跨月排序保持年月精度次序，无有效日期记录放末尾，不丢失零值空金额和重复记录',()=>{
 const b=book([row({'应付款时间':'2026-12','申请付款金额':0}),row({'应付款时间':'待定','申请付款金额':''}),row({'应付款时间':'2026-11-30'}),row({'应付款时间':'2026-12-02'}),row({'应付款时间':'待定','申请付款金额':''})]);const before=JSON.stringify(b),r=build(b);
 assert.deepEqual(r.sheets[0].sources.map(s=>s.row),[4,2,5,3,6]);assert.deepEqual(rows(r).map(r=>r[3].v),[100,0,100,'','']);assert.equal(JSON.stringify(b),before);assert(rows(r).filter(r=>r[5].v==='待定').every(r=>r[10].v===''&&r[11].v===''));
 for(const cell of [{t:'n',f:'1+1'},{t:'e',v:23}])assert.throws(()=>build(book([row({'应付款时间':'待定','申请付款金额':cell})])),/李奇!D2/);
});

test('采购跟单员为供应商前首列，按真实来源人员填写，兼容日期后缀和重排，预览导出一致',()=>{
 const b=book(),names=['对账明细-李奇','对账明细徐英1109','对账明细－邢芳芳（20261109）','对账明细-操思敏-2026-11-09','对账明细-孙立柱'];
 b.Sheets={};b.SheetNames=[names[4],names[0],names[2],names[1],names[3]];
 for(let i=0;i<names.length;i++)b.Sheets[names[i]]=sheet([row({'供应商名称':'供应商'+i,'对公付款0A流程号':'OA00'+i,'应付款时间':'2026-11-'+String(20-i).padStart(2,'0')})],i%2?[...H].reverse():H);
 const applications=X.utils.book_new();X.utils.book_append_sheet(applications,X.utils.aoa_to_sheet([['流程编码'],['OA001'],['OA003']]),'付款申请汇总');
 const before=JSON.stringify(b),r=C.build(b,'原台账.xlsx',{today:'2026-11-09',applications:{book:applications,filename:'申请.xlsx'}}),result=r.sheets[0];
 assert.deepEqual(result.headers.slice(0,2),['采购跟单员','供应商']);assert.equal(result.headers.length,14);assert.deepEqual(result.rows.map(r=>r[0].v),['操思敏','徐英','孙立柱','邢芳芳','李奇']);
 const out=C.workbook(r),read=X.read(X.write(out,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true}),exported=read.Sheets.正常付款;
 for(let i=0;i<result.rows.length;i++){const record=result.rows[i],person=PEOPLE[names.indexOf(result.sources[i].sheet)];assert.equal(record[0].v,person);assert.equal(record[0].t,'s');assert.equal(exported['A'+(i+2)].v,person);assert.equal(exported['B'+(i+2)].v,record[1].v);assert.equal(exported['N'+(i+2)].v,record[13].v);}
 assert.equal(exported['!autofilter'].ref,'A1:N6');assert.equal(JSON.stringify(b),before);
});
