const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const NEG=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','已付金额','剩余付款金额','是否结清','是否提交付款申请单','谈判明细','分批总金额','核对'];
const APP=['申请日期','供应商名称','付款主体','流程编码','本次申请付款金额','付款分类'];
const OA=['流程编号','收款单位','付款金额',...C.FLOW_FIELDS];
const negCode='MDST-DGFK20261001001',normalCode='MDST-DGFK20261001002';
function book(name,headers,rows){const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([headers,...rows.map(row=>headers.map(h=>row[h]??''))]),name);return b;}
function fixture(plans=[100],amounts=[100]){
 const base={'供应商名称':'谈判供应商','付款主体':'迈德斯特','对公付款0A流程号':negCode,'申请付款金额':100,'已付金额':0,'剩余付款金额':100,'是否结清':'否','是否提交付款申请单':'未提交','是否谈判付款':'是','应付款时间':'2026-10-31','开票时间':'2026-10-01','登记月份':'2026-10','谈判明细':'原文'};
 const normal={...base,'供应商名称':'正常供应商','对公付款0A流程号':normalCode,'申请付款金额':80,'剩余付款金额':80,'是否谈判付款':'否','是否提交付款申请单':'已提交'};
 const ledger=book('对账明细-李奇',C.NORMAL_HEADERS,[base,normal]);for(const p of C.PEOPLE.slice(1))X.utils.book_append_sheet(ledger,X.utils.aoa_to_sheet([C.NORMAL_HEADERS]),'对账明细-'+p);
 const dates=plans.map((_,i)=>'2026-11-'+String(i+2).padStart(2,'0')),negotiated=book('最新版谈判付款',NEG.concat(dates),[{...base,...Object.fromEntries(dates.map((d,i)=>[d,plans[i]]))}]);
 const applications=book('申请汇总',APP,amounts.map((a,i)=>({'申请日期':'2026-10-'+(20+i),'供应商名称':'谈判供应商','付款主体':'迈德斯特','流程编码':negCode,'本次申请付款金额':a,'付款分类':'谈判付款'})).concat([{'申请日期':'2026-10-27','供应商名称':'正常供应商','付款主体':'迈德斯特','流程编码':normalCode,'本次申请付款金额':80,'付款分类':'正常付款'}]));
 const oa=book('Sheet0',OA,[base,normal].map(r=>({'流程编号':r['对公付款0A流程号'],'收款单位':r['供应商名称'],'付款金额':r['申请付款金额'],当前节点:'审批',未操作者:'人',流程状态:'处理中',申请人:'人'})));
 return {ledger:{book:ledger,name:'资金.xlsx'},negotiated:{book:negotiated,name:'谈判.xlsx'},applications:{book:applications,name:'申请.xlsx'},oa:{book:oa,name:'OA.xlsx'}};
}
const sheet=(i,role)=>i[role].book.Sheets[i[role].book.SheetNames[0]];
const reviews=r=>r.reviewRows.filter(row=>row[5].v.startsWith('申请标红核对：'));
const state=r=>r.negotiated.rows[0][r.negotiated.headers.indexOf('是否提交付款申请单')].v;
test('逐笔精确匹配只标一个分批单元格，正常付款只标本次匹配行，保留原公式和源对象',()=>{
 const i=fixture([60,40],[60]);sheet(i,'negotiated').L2={t:'n',v:60,f:'D2-40'};const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.highlights,{negotiatedCells:['L2'],normalRows:[2]});assert.equal(state(r),'已经提交部分');assert.equal(r.negotiated.rows[0][11].f,'D2-40');assert.equal(r.stats.pending,0);assert.equal(reviews(r).length,0);assert.equal(JSON.stringify(i),before);
});
test('按分比较文本金额与公式缓存，逐笔不累计凑批次，跨日重复只计一次',()=>{
 const i=fixture(['￥60.00',40],[60,40,60]),r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,['L2','M2']);assert.equal(state(r),'已经提交全部');assert.equal(r.reviewRows.filter(row=>row[5].v.startsWith('重复提交')).length,1);assert.equal(reviews(r).length,0);
 const j=fixture([100],[60,40]),result=C.build(j);assert.deepEqual(result.highlights.negotiatedCells,[]);assert.equal(state(result),'已经提交全部');assert.equal(reviews(result).length,2);assert.match(reviews(result)[0][5].v,/60.00.*未找到金额相等/s);assert.match(reviews(result)[1][5].v,/40.00.*未找到金额相等/s);assert.equal(result.stats.pending,0);
});
test('同额多批按实际日期先后标红，重复申请不消耗后续批次',()=>{
 const i=fixture([90000,90000],[90000,90000]),s=sheet(i,'negotiated');s.L1={t:'s',v:'2026-11-23'};s.M1={t:'s',v:'2026-10-26'};
 const before=JSON.stringify(i),r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,['M2']);assert.equal(reviews(r).length,0);assert.equal(r.stats.pending,0);assert.equal(r.reviewRows.filter(row=>row[5].v.startsWith('重复提交')).length,1);assert.equal(JSON.stringify(i),before);
});
test('没有同额批次只提醒不阻止，包含金额及付款申请来源行',()=>{
 const r=C.build(fixture([],[50]));assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(reviews(r).length,1);assert.equal(r.stats.pending,0);const message=reviews(r)[0][5].v;assert.match(message,/50.00/);assert.match(message,/申请.xlsx \/ 申请汇总 第 2 行/);assert.match(message,/无金额相等的批次/);
 assert.match(message,/需要处理/);assert.match(message,/谈判.xlsx \/ 最新版谈判付款 第 2 行/);
});
test('同流程1183.85与11183.85分别判断，已标红批次不掩盖另一笔金额核对和处理指引',()=>{
 const i=fixture([11183.85],[1183.85,11183.85]),applications=sheet(i,'applications'),negotiated=sheet(i,'negotiated');
 const applicationRows=X.utils.sheet_to_json(applications,{header:1});
 for(let r=1;r<=3;r++)for(let c=0;c<APP.length;c++)delete applications[X.utils.encode_cell({r,c})];
 X.utils.sheet_add_aoa(applications,[applicationRows[1]],{origin:'A12'});X.utils.sheet_add_aoa(applications,[applicationRows[2]],{origin:'A67'});X.utils.sheet_add_aoa(applications,[applicationRows[3]],{origin:'A68'});
 i.applications.book.SheetNames=['付款申请汇总'];i.applications.book.Sheets={'付款申请汇总':applications};i.applications.name='付款申请汇总_1024.xlsx';
 negotiated.D2={t:'n',v:11183.85};negotiated.F2={t:'n',v:11183.85};
 const negotiatedRow=X.utils.sheet_to_json(negotiated,{header:1})[1];for(let c=0;c<NEG.length+1;c++)delete negotiated[X.utils.encode_cell({r:1,c})];X.utils.sheet_add_aoa(negotiated,[negotiatedRow],{origin:'A49'});
 const ledger=sheet(i,'ledger');for(const h of ['申请付款金额','剩余金额'])ledger[X.utils.encode_cell({r:1,c:C.NORMAL_HEADERS.indexOf(h)})]={t:'n',v:11183.85};sheet(i,'oa').C2={t:'n',v:11183.85};
 const before=JSON.stringify(i),r=C.build(i),lines=reviews(r);assert.deepEqual(r.highlights.negotiatedCells,['L49']);assert.equal(lines.length,1);assert.equal(r.stats.pending,0);
 assert.equal(r.negotiated.rows[47][r.negotiated.headers.indexOf('是否提交付款申请单')].v,'已经提交全部');assert.equal(r.negotiated.rows[47][11].v,11183.85);assert.equal(JSON.stringify(i),before);
 const message=lines[0][5].v;assert.match(message,/申请金额 1183\.85/);assert.match(message,/未找到金额相等的分批/);assert.match(message,/付款申请汇总_1024.xlsx \/ 付款申请汇总 第 12 行/);assert.match(message,/谈判.xlsx \/ 最新版谈判付款 第 49 行/);assert.doesNotMatch(message,/第 67 行/);
 assert.match(message,/需要处理/);assert.match(message,/核实[^。\n]*本次申请付款金额/);assert.match(message,/(?:填错|错误|不正确)[^。\n]*(?:修改|更正|改)[^。\n]*申请/);assert.match(message,/金额正确[^。\n]*谈判[^。\n]*(?:日期|金额)/);assert.match(message,/重新上传/);
});
test('独立申请依次使用未标红批次，批次用完后仅提醒不重复使用',()=>{
 const i=fixture([100,100],[100,100,100]),s=sheet(i,'applications');for(const row of [2,3,4])s['D'+row]={t:'s',v:''};
 const r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,['L2','M2']);assert.equal(reviews(r).length,1);assert.match(reviews(r)[0][5].v,/同额分批已全部对应其他申请/);assert.match(reviews(r)[0][5].v,/第 4 行/);assert.equal(r.stats.pending,0);
 assert.match(reviews(r)[0][5].v,/需要处理/);
});
test('最早日期相同仍需核对，不凭列顺序选择批次',()=>{
 const i=fixture([50,50],[50]),s=sheet(i,'negotiated');s.M1={t:'s',v:'2026-11-02'};const r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.match(reviews(r)[0][5].v,/最早日期存在多个同额批次/);assert.equal(r.stats.pending,0);
 assert.match(reviews(r)[0][5].v,/需要处理/);
});
test('同流程对应多笔款项归属不明不标红，独立核对绕过谈判付款过滤且不增加pending',()=>{
 const i=fixture([60,40],[60]),s=sheet(i,'negotiated');X.utils.sheet_add_aoa(s,[X.utils.sheet_to_json(s,{header:1})[1]],{origin:-1});const before=JSON.stringify(i),r=C.build(i);
 assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(reviews(r).length,2);assert(reviews(r).every(row=>row[5].v.includes('归属不明')&&row[5].v.includes('60.00')));assert.equal(r.stats.pending,0);assert.equal(JSON.stringify(i),before);
 for(const row of reviews(r))assert.match(row[5].v,/需要处理/);
});
test('分类隔离且原已提交状态不能触发红色，未申请不新增标红核对',()=>{
 const i=fixture([100],[100]);sheet(i,'applications').F2={t:'s',v:'正常付款'};sheet(i,'applications').E3={t:'n',v:79};sheet(i,'negotiated').H2={t:'s',v:'已经提交全部'};const r=C.build(i);
 assert.deepEqual(r.highlights,{negotiatedCells:[],normalRows:[]});assert.equal(state(r),'未提交');assert.equal(r.normalRows[0][r.normalHeaders.indexOf('是否提交付款申请单')].v,'未申请');assert.equal(reviews(r).length,0);
 const noApplications=fixture([100],[]);sheet(noApplications,'applications').E2={t:'n',v:79};assert.deepEqual(C.build(noApplications).highlights,{negotiatedCells:[],normalRows:[]});
});
test('零、空、负、异常分批和无缓存公式不标红，不按申请日期选择批次',()=>{
 for(const value of [0,'',-1,'待定',{t:'e',v:23},{t:'n',f:'D2'},{t:'n',v:100.001}]){const i=fixture([value],[100]),r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(reviews(r).length,1);assert.equal(r.stats.pending,0);assert.match(reviews(r)[0][5].v,/需要处理/);}
 const same=fixture([50,50],[50]);sheet(same,'applications').A2={t:'s',v:'2026-11-03'};assert.deepEqual(C.build(same).highlights.negotiatedCells,['L2']);
 for(const value of [0,-1,'',{t:'e',v:23}]){const i=fixture([100],[value]);sheet(i,'negotiated').H2={t:'s',v:'未提交'};const r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,[]);for(const row of reviews(r))assert.match(row[5].v,/需要处理/);}
});
test('日期表头支持文本和1904数值，金额或表头合并不猜归属',()=>{
 const i=fixture([100],[100]),s=sheet(i,'negotiated');i.negotiated.book.Workbook={WBProps:{date1904:true}};s.L1={t:'n',v:44866,z:C.DATE_FORMAT};assert.deepEqual(C.build(i).highlights.negotiatedCells,['L2']);
 for(const range of ['L2:M2','L1:M1','L2:L3']){const j=fixture([100,40],[100]),s=sheet(j,'negotiated');s['!merges']=[X.utils.decode_range(range)];const r=C.build(j);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(reviews(r).length,1);assert.match(reviews(r)[0][5].v,/合并|同额/);assert.match(reviews(r)[0][5].v,/需要处理/);}
 const invalid=fixture([100],[100]);sheet(invalid,'negotiated').L1={t:'s',v:'3月25'};const r=C.build(invalid);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.match(reviews(r)[0][5].v,/日期无法识别/);assert.match(reviews(r)[0][5].v,/需要处理/);
});
test('重复申请的无批提醒只保留一次，正常付款按真实输出行号标红',()=>{
 const i=fixture([],[100,100,100]),s=sheet(i,'ledger'),normal=X.utils.sheet_to_json(s,{header:1})[2].slice();normal[C.NORMAL_HEADERS.indexOf('供应商名称')]='第二正常供应商';normal[C.NORMAL_HEADERS.indexOf('对公付款0A流程号')]='MDST-DGFK20261001003';X.utils.sheet_add_aoa(s,[normal],{origin:-1});
 const r=C.build(i);assert.equal(reviews(r).length,1);assert.equal((reviews(r)[0][5].v.match(/申请金额 100.00/g)||[]).length,1);assert.deepEqual(r.highlights.normalRows,[2]);assert.equal(r.normalRows.length,2);
});
test('几十笔不匹配申请每笔单独核对，保持短提示且不罗列无关日期',()=>{
 const i=fixture(Array(20).fill(9999),Array.from({length:40},(_,n)=>n+1)),r=C.build(i),lines=reviews(r);assert.equal(lines.length,40);assert.equal(r.stats.pending,0);assert.deepEqual(r.highlights.negotiatedCells,[]);
 for(const line of lines){assert.match(line[5].v,/候选日期：无金额相等的批次/);assert.doesNotMatch(line[5].v,/2026年11月/);assert(line[5].v.length<350);assert.equal((line[5].v.match(/申请来源：/g)||[]).length,1);assert.match(line[5].v,/需要处理/);}
 const firstHighlight=r.reviewRows.findIndex(row=>row[5].v.startsWith('申请标红核对：'));assert(r.reviewRows.slice(firstHighlight).every(row=>row[5].v.startsWith('申请标红核对：')));
});
test('保留的历史已提交状态不能冒充本次匹配，源数据前置标题不改变原单元格坐标',()=>{
 const i=fixture();sheet(i,'oa').B3={t:'s',v:'另一个供应商'};const r=C.build(i);assert.equal(r.normalRows[0][r.normalHeaders.indexOf('是否提交付款申请单')].v,'已提交');assert.deepEqual(r.highlights.normalRows,[]);
 const j=fixture(),s=sheet(j,'negotiated'),moved={};for(const [address,value]of Object.entries(s))if(/^[A-Z]+\d+$/.test(address)){const p=X.utils.decode_cell(address);p.r+=2;moved[X.utils.encode_cell(p)]=value;}moved.A1={t:'s',v:'原表标题'};moved['!ref']='A1:L4';j.negotiated.book.Sheets[j.negotiated.book.SheetNames[0]]=moved;assert.deepEqual(C.build(j).highlights.negotiatedCells,['L4']);
});
test('付款归属不明的100个同额日期候选分条完整列出，不产生超高单行',()=>{
 const i=fixture(Array(100).fill(100),[100]),s=sheet(i,'negotiated'),expected=[];
 for(let c=0;c<100;c++){const d=new Date(Date.UTC(2026,0,1+c*7));s[X.utils.encode_cell({r:0,c:c+11})]={t:'s',v:d.toISOString().slice(0,10)};expected.push({date:d.getUTCFullYear()+'年'+(d.getUTCMonth()+1)+'月'+d.getUTCDate()+'日',address:X.utils.encode_cell({r:1,c:c+11})});}
 X.utils.sheet_add_aoa(s,[X.utils.sheet_to_json(s,{header:1})[1]],{origin:-1});
 const before=JSON.stringify(i),r=C.build(i),lines=reviews(r).filter(row=>row[8].v===2);assert.equal(lines.length,10);assert.equal(reviews(r).length,20);assert.equal(r.stats.pending,0);assert.deepEqual(r.highlights.negotiatedCells,[]);assert.equal(JSON.stringify(i),before);
 const found=[];for(const [index,line]of lines.entries()){
  const message=line[5].v;assert(message.length<500);assert.match(message,/申请金额 100.00/);assert.match(message,/申请.xlsx \/ 申请汇总 第 2 行/);assert(message.includes('候选日期（'+(index*10+1)+'-'+(index*10+10)+'/100）'));
  for(const m of message.matchAll(/(\d{4}年\d{1,2}月\d{1,2}日)（([A-Z]+2)，100.00）/g))found.push({date:m[1],address:m[2]});
  const formatted=C.formatReviewText(message);let heightLines=0;for(const text of formatted.split('\n')){let used=0;heightLines++;for(const ch of text){const width=(ch.charCodeAt(0)>255?2:1)*12;if(used&&used+width>(255-3)*7){heightLines++;used=0;}used+=width;}}assert(heightLines*20+10<=409.5,'导出最大列宽下完整展示');
 }
 assert.deepEqual(found,expected,'所有日期与地址按原顺序保留，不丢失、不重复');
});
test('申请日期只随本次匹配输出，重复申请日期去重排序并附于同一批次',()=>{
 const i=fixture([100,100],[100,100,100]),apps=sheet(i,'applications');apps.A2.v='2026-10-23';apps.A3.v='2026-10-20';apps.A4.v='2026-10-23';
 const before=JSON.stringify(i),r=C.build(i);assert.deepEqual(r.highlights.negotiatedCells,['L2']);assert.deepEqual(r.applicationComments,{L2:'申请日期（资金整理）：2026年10月20日；2026年10月23日'});
 const c=r.normalRows[0][r.normalHeaders.indexOf('申请日期')];assert.equal(c.t,'n');assert.equal(c.z,C.DATE_FORMAT);assert.equal(X.SSF.format(c.z,c.v),'2026年10月27日');assert.equal(JSON.stringify(i),before);
 apps.E5.v=79;sheet(i,'ledger')[X.utils.encode_cell({r:2,c:C.NORMAL_HEADERS.indexOf('申请日期')})]={t:'s',v:'历史日期不沿用'};const next=C.build(i);assert.equal(next.normalRows[0][next.normalHeaders.indexOf('申请日期')].v,'');
});
test('多次正常申请日期合并，异常和空日期明确显示，日期采用申请文件纪元',()=>{
 const i=fixture(),apps=sheet(i,'applications'),normalRow=X.utils.sheet_to_json(apps,{header:1})[2];
 X.utils.sheet_add_aoa(apps,[[...normalRow],[...normalRow]],{origin:-1});apps.A3={t:'n',v:44860};apps.A4={t:'s',v:'2026-10-26'};apps.A5={t:'s',v:'2026-10-27'};
 i.applications.book.Workbook={WBProps:{date1904:true}};let r=C.build(i);assert.equal(r.normalRows[0][r.normalHeaders.indexOf('申请日期')].v,'2026年10月26日；2026年10月27日');
 apps.A2.v='';apps.A3={t:'e',v:23,w:'#REF!'};apps.A4.v='';apps.A5.v='日期待确认';r=C.build(i);
 assert.equal(r.applicationComments.L2,'申请日期（资金整理）：日期未填写');assert.equal(r.normalRows[0][r.normalHeaders.indexOf('申请日期')].v,'日期待核对：#REF!；日期未填写；日期待核对：日期待确认');
});
test('独立同额申请依次分配的批注只包含各自日期，未标红批次无批注',()=>{
 const i=fixture([100,100,100],[100,100]),apps=sheet(i,'applications');apps.D2.v=apps.D3.v='';const r=C.build(i);
 assert.deepEqual(r.applicationComments,{L2:'申请日期（资金整理）：2026年10月20日',M2:'申请日期（资金整理）：2026年10月21日'});
 apps.F2.v=apps.F3.v='正常付款';const next=C.build(i);assert.deepEqual(next.applicationComments,{});
});
test('正常付款按申请付款的日期升序，同周稳定、无效日期置后，标红与公式跟随记录',()=>{
 for(const date1904 of [false,true]){
  const i=fixture(),ledger=sheet(i,'ledger'),headers=C.NORMAL_HEADERS,base=X.utils.sheet_to_json(ledger)[1];
  const setDue=(s,row,value)=>{s[X.utils.encode_cell({r:row-1,c:headers.indexOf('应付款时间')})]=typeof value==='object'?value:{t:typeof value==='number'?'n':'s',v:value};};
  setDue(ledger,3,'2027-01-04');
  const cases=[['空白',''],['同周先','2026-11-08'],['跨年','2026-01-01'],['同周后','2026-11-02'],['文本异常','未确定'],['日期零',0],['错误公式',{t:'e',v:23,w:'#REF!',f:'A99'}],['无缓存公式',{t:'n',f:'A99'}],['数值日期',46328-(date1904?1462:0)],['年末','2026-12-31']];
  cases.forEach(([name,due],j)=>{
   const record={...base,'供应商名称':name,'对公付款0A流程号':'MDST-DGFK202611'+String(j+1).padStart(5,'0'),'申请付款金额':81+j,'已付金额':j,'是否结清':'否'};
   X.utils.sheet_add_aoa(ledger,[headers.map(h=>record[h]??'')],{origin:-1});setDue(ledger,j+4,due);
  });
  i.ledger.book.Workbook={WBProps:{date1904}};
  const before=JSON.stringify(i),r=C.build(i),at=h=>headers.indexOf(h),names=r.normalRows.map(row=>row[at('供应商名称')].v);
  assert.deepEqual(names,[...(date1904?['日期零']:[]),'跨年','同周先','同周后','数值日期','年末','正常供应商','空白','文本异常',...(date1904?[]:['日期零']),'错误公式','无缓存公式']);
  const validCount=date1904?7:6;assert.deepEqual(r.normalRows.slice(0,validCount).map(row=>row[at('申请付款的日期')].v),[...(date1904?['1903/12/21周']:[]),'2025/12/22周','2026/10/26周','2026/10/26周','2026/10/26周','2026/12/21周','2026/12/28周']);
  assert(r.normalRows.slice(validCount).every(row=>row[at('申请付款的日期')].v==='日期待确认'));
  const submitted=names.indexOf('正常供应商');assert.deepEqual(r.highlights.normalRows,[submitted+2]);assert.equal(r.normalRows[submitted][at('申请日期')].v,46322);assert.equal(r.normalRows[submitted][at('是否提交付款申请单')].v,'已提交');
  for(const [j,row]of r.normalRows.entries()){assert.equal(row[at('剩余金额')].f,'I'+(j+2)+'-N'+(j+2));assert.equal(row[at('剩余金额')].v,row[at('申请付款金额')].v-row[at('已付金额')].v);}
  assert.equal(JSON.stringify(i),before);assert.deepEqual(C.build(i).normalRows,r.normalRows);assert.deepEqual(r.applicationComments,{L2:'申请日期（资金整理）：2026年10月20日'});
 }
});
