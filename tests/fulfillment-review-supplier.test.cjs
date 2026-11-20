const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js'),P=require('../fulfillment-purchase.js');
const filename='订单履约明细跟进表截止2026.11.13-11.19更新版.xlsx';
const channel='海外事业一部*欧洲部*亚马逊欧盟组';
const cell=v=>({t:typeof v==='number'?'n':'s',v:v??''});
function snapshot(items=[{}]){
 const defaults={'采购下单人':'徐英','事业部':'海外事业一部','销售分渠道':channel,'供应商简称':'迈锐','系列':'AX25','物料编码':'001','采购订单号':'CGDD015712','未交付数量':10,'下单月份':'2026年11月'};
 return {rows:items.map(item=>C.COLUMNS.map(name=>cell({...defaults,...item}[name]))),sources:items.map((_,i)=>({sheet:'徐英',row:i+2})),counts:[{sheet:'徐英',count:items.length}],warnings:[]};
}
function directory(rows){
 const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['供应商','供应商简称'],...rows]),'供应商名录');return {filename:'供应商分工.xlsx',book};
}
function source(kind,rows){
 const receipt=kind==='receipt',headers=receipt?['入库日期','单据编号','订单单号','供应商','物料编码','实收数量']:['采购日期','单据编号','供应商','物料编码','采购数量','事业部'];
 const data=rows.map((row,i)=>{const r={supplier:'迈锐',code:'NEW',order:'CGDD999999',qty:12,date:'2026-11-13',...row};return receipt?[r.date,'CGRK'+(i+1),r.order,r.supplier,r.code,r.qty]:[r.date,r.order,r.supplier,r.code,r.qty,channel];});
 const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...data]),'明细');return P.parse({filename:receipt?'采购入库单.xlsx':'采购订单.xlsx',book},kind);
}
function run(kind,{current=snapshot(),previous=null,records=[{}],catalog=null,comparison,override={}}={}){
 const args={kind,filename,current,previous,directory:catalog,purchase:source(kind,records),comparison:comparison??C.seriesComparison(previous,current),...override};
 const before=structuredClone(args),result=P.enrich(args);assert.deepEqual(args,before,'核对范围筛选不改写源数据、比较行或固定历史');
 assert(Array.isArray(result.reviewIssues),'返回单独的核对展示范围');
 for(const issue of result.reviewIssues)assert(result.issues.includes(issue),'展示核对是原始问题子集，不能重写或补造问题');
 return result;
}

for(const kind of ['purchase','receipt']){
 test(kind+'：只过滤无关供应商核对，相关供应商的新物料或单号仍保留',()=>{
  const result=run(kind,{records:[{supplier:'无关供应商有限公司',qty:99},{supplier:' 迈 锐 ',qty:12},{supplier:'迈锐',code:'001',order:'CGDD015712',qty:20}]});
  assert.equal(result.issues.length,2);assert.equal(result.reviewIssues.length,1);assert.equal(result.reviewIssues[0].supplier,'迈 锐');
  assert.equal(result.issues[0].supplier,'无关供应商有限公司');assert.equal(result.issues[0].qty,99);
  assert.equal(result.rows[0][8].v,20,'只改核对展示，确定采购或入库数量不受影响');assert.equal(result.stats.unmatchedRows,2);
  const baseline=run(kind,{records:[{supplier:'迈锐',code:'001',order:'CGDD015712',qty:20}]});assert.deepEqual(result.rows,baseline.rows);
 });

 test(kind+'：全称简称依据目录对应，目录内但履约未涉及的公司仍排除',()=>{
  const catalog=directory([['嘉兴迈锐医疗器械有限公司','迈锐'],['仅在目录有限公司','目录供应商']]);
  const result=run(kind,{catalog,records:[{supplier:'嘉兴迈锐医疗器械有限公司'},{supplier:'仅在目录有限公司'}]});
  assert.equal(result.issues.length,2);assert.deepEqual(result.reviewIssues.map(i=>i.supplier),['嘉兴迈锐医疗器械有限公司']);
  const reverse=run(kind,{catalog,current:snapshot([{'供应商简称':'嘉兴迈锐医疗器械有限公司'}]),records:[{supplier:'迈锐'}]});assert.equal(reverse.reviewIssues.length,1);
 });

 test(kind+'：供应商范围取全部本期履约行，零差异组也保留相关核对，上期独有公司不纳入',()=>{
  const current=snapshot(),previous=snapshot([{}, {'供应商简称':'仅上期供应商','采购订单号':'CGDD888888','物料编码':'999'}]);
  const result=run(kind,{current,previous,records:[{supplier:'迈锐'},{supplier:'仅上期供应商',code:'999',order:'CGDD888888',qty:''}]});
  assert(result.rows.every(row=>row[3].v!=='迈锐'),'同量组已从系列对比过滤');
  assert.equal(result.issues.length,2);assert.deepEqual(result.reviewIssues.map(i=>i.supplier),['迈锐']);assert.equal(result.rows[0][8].v,'待核对','历史候选的数量状态不因隐藏核对改变');
  const noDifference=run(kind,{current,previous:snapshot(),records:[{supplier:'迈锐'}]});assert.deepEqual(noDifference.rows,[]);assert.equal(noDifference.reviewIssues.length,1);
 });

 test(kind+'：同订单物料关联到本期的供应商冲突必须保留，不能隐藏相关记录',()=>{
  const result=run(kind,{records:[{supplier:'不同供应商有限公司',code:'001',order:'CGDD015712',qty:30}]});
  assert.equal(result.reviewIssues.length,1);assert.match(result.reviewIssues[0].reason,/冲突/);assert.equal(result.rows[0][8].v,'待核对');
  if(kind==='receipt'){assert.match(result.rows[0][10].v,/待核对候选数量：30/);assert.match(result.rows[0][10].v,/不能计作本组入库/);}
 });

 test(kind+'：仅匹配历史订单物料的冲突不因与本期同五维组而扩大核对范围',()=>{
  const current=snapshot(),previous=snapshot([{'物料编码':'999','采购订单号':'CGDD888888','未交付数量':5}]);
  const result=run(kind,{current,previous,records:[{supplier:'本期未涉及的冲突供应商',code:'999',order:'CGDD888888',qty:30}]});
  assert.equal(result.issues.length,1);assert.match(result.issues[0].reason,/冲突/);assert.deepEqual(result.reviewIssues,[]);
  assert.equal(result.rows[0][8].v,'待核对','只筛核对展示，不改变历史候选造成的数量状态');
  if(kind==='receipt')assert.match(result.rows[0][10].v,/待核对候选数量：30/);
 });

 test(kind+'：共享简称只保留潜在相关核对，不合并两家公司或扩大已明确全称的范围',()=>{
  const catalog=directory([['供应商甲有限公司','共同简称'],['供应商乙有限公司','共同简称']]);
  const ambiguous=run(kind,{catalog,current:snapshot([{'供应商简称':'共同简称'}]),records:[{supplier:'供应商甲有限公司'},{supplier:'供应商乙有限公司'}]});
  assert.equal(ambiguous.reviewIssues.length,2);assert.equal(ambiguous.rows[0][8].v,0,'只是保留核对，不按同简称分配数量');
  const explicit=run(kind,{catalog,current:snapshot([{'供应商简称':'供应商甲有限公司'}]),records:[{supplier:'供应商甲有限公司'},{supplier:'供应商乙有限公司'}]});
  assert.deepEqual(explicit.reviewIssues.map(i=>i.supplier),['供应商甲有限公司']);assert.equal(explicit.issues.length,2);
 });

 test(kind+'：供应商空白及全局文件或日期异常保留，不能当作无关供应商隐藏',()=>{
  const blank=run(kind,{records:[{supplier:'',qty:0}]});assert.equal(blank.reviewIssues.length,1);assert.match(blank.reviewIssues[0].reason,/供应商为空/);assert.equal(blank.reviewIssues[0].qty,0);
  const broken=run(kind,{override:{purchase:{filename:'损坏文件.xlsx',error:'读取失败',records:[]}}});assert.equal(broken.reviewIssues.length,1);assert.match(broken.reviewIssues[0].reason,/读取失败/);
  const badDate=run(kind,{override:{filename:'没有日期的履约表.xlsx'}});assert.equal(badDate.reviewIssues.length,1);assert.match(badDate.reviewIssues[0].reason,/日期/);
  const absent=run(kind,{override:{purchase:null}});assert.deepEqual(absent.reviewIssues,[]);
  const legacy=run(kind,{override:{filename:'旧履约表.xlsx',purchase:null}});assert.deepEqual(legacy.reviewIssues,[]);
 });
}
