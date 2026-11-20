const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js'),P=require('../fulfillment-purchase.js');
const FILENAME='订单履约明细跟进表截止2026.11.13-11.19更新版.xlsx';
const PH=['采购日期','创建人','事业部','单据编号','供应商','物料编码','采购数量','剩余入库数量','关闭状态','采购下单人','销售分渠道','供应商简称','系列'];
const cell=value=>value&&typeof value==='object'?structuredClone(value):{t:typeof value==='number'?'n':'s',v:value??''};
function current(items=[{}]){
 const defaults={'采购下单人':'徐英','事业部':'海外事业一部','销售分渠道':'海外事业一部*欧洲部*亚马逊欧盟组','供应商简称':'迈锐','系列':'AX25','物料编码':'001','SKU':'AX25-EU','采购订单号':'CGDD015712','下单月份':'2026年11月','未交付数量':10};
 return {rows:items.map(item=>C.COLUMNS.map(key=>cell({...defaults,...item}[key]))),sources:items.map((_,i)=>({sheet:'徐英',row:i+2})),counts:[{sheet:'徐英',count:items.length}],warnings:[]};
}
function order(extra={}){return {'采购日期':'2026-11-13','创建人':'孙立柱','事业部':'海外事业一部*欧洲部*亚马逊欧盟组','单据编号':'CGDD015712','供应商':'迈锐','物料编码':'001','采购数量':150,'剩余入库数量':1,'关闭状态':'未关闭',...extra};}
function input(records,headers=PH){const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(header=>record[header]??''))]),'采购明细');return {filename:'采购订单.xlsx',book};}
function directory(records){const headers=['采购下单人','正确物料编码','供应商','供应商简称'],book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records.map(row=>headers.map(header=>row[header]??''))]),'产品线明细');return {filename:'供应商分工.xlsx',book};}
function enrich({items=[{}],records=[order()],filename=FILENAME,comparison,previous=null,catalog=null}={}){
 const now=current(items),purchase=records===null?null:P.parse(input(records)),args={filename,current:now,previous,purchase,directory:catalog,comparison:comparison??C.seriesComparison(previous,now)};
 const before=structuredClone(args),result=P.enrich(args);assert.deepEqual(args,before,'补充列不得改写采购、履约、比较行或分工资料');return result;
}
const rowFor=(result,labels)=>result.rows.find(row=>labels.every((value,index)=>value===undefined||row[index].v===value));
const note=result=>JSON.stringify(result.issues);

test('有日期和数量但身份字段全空的采购行仍进入核对，只有明确合计或空行跳过',()=>{
 const source=input([order(),{'采购日期':'2026-11-14','采购数量':99},{'供应商':'合计','采购数量':249},{}]);
 const parsed=P.parse(source);assert.equal(parsed.records.length,2);
 assert.match(parsed.records[1].issues.join('；'),/物料编码为空/);
 const now=current(),result=P.enrich({filename:FILENAME,current:now,comparison:C.seriesComparison(null,now),purchase:parsed});
 assert.equal(result.rows[0][8].v,150);assert.match(note(result),/采购明细!G3/);assert.match(note(result),/供应商为空/);
});

test('标题取本次文件日期范围，含起止日，支持跨月跨年并拒绝无效或多范围',()=>{
 const normal=P.dateRange(FILENAME);assert.equal(normal.start,'2026-11-13');assert.equal(normal.end,'2026-11-19');assert.equal(normal.label,'11.13-11.19新增采购订单');assert.equal(normal.endSerial-normal.startSerial,6);
 const month=P.dateRange('履约2026.10.30-11.5汇总版_V9.xlsx');assert.equal(month.start,'2026-10-30');assert.equal(month.end,'2026-11-05');assert.equal(month.label,'10.30-11.5新增采购订单');
 for(const filename of ['履约2026.12.26-1.1.xlsx','履约2026.12.26-2027.1.1.xlsx']){const range=P.dateRange(filename);assert.equal(range.start,'2026-12-26');assert.equal(range.end,'2027-01-01');assert.match(range.label,/2026/);assert.match(range.label,/2027/);}
 for(const filename of ['履约更新版.xlsx','履约2026.2.29-3.5.xlsx','履约2026.11.19-11.13.xlsx','履约2026.13.1-13.7.xlsx','履约2026.11.13-11.19对比2026.11.6-11.12.xlsx'])assert.equal(P.dateRange(filename),null,filename);
});

test('原采购数量按日期首尾均计入，已关闭不排除，不用剩余入库数量或订单数',()=>{
 const result=enrich({records:[order({'采购日期':'2026-11-12','采购数量':999}),order({'采购日期':'2026-11-13','采购数量':150,'剩余入库数量':3}),order({'采购日期':'2026-11-19 23:59:59','采购数量':200,'剩余入库数量':0,'关闭状态':'已关闭'}),order({'采购日期':'2026-11-20','采购数量':888})]});
 assert.equal(result.rows[0][8].t,'n');assert.equal(result.rows[0][8].v,350);assert.equal(result.columns[7],'差异数量');assert.equal(result.columns[8],'11.13-11.19新增采购订单');assert.deepEqual(result.columns.slice(9),['备注','判断']);
});

test('同单据同物料的等额多行逐行合计，小数精确，不按单号或数量去重',()=>{
 const result=enrich({records:[order({'采购数量':0.1}),order({'采购数量':0.2}),order({'采购数量':0.1}),order({'采购数量':0})]});
 assert.equal(result.rows[0][8].v,0.4);assert.equal(result.rows[0][8].t,'n');assert.equal(result.rows[0][8].f,undefined);
});

test('超出精度的采购数量文本和合计不能四舍五入后静默计数',()=>{
 for(const records of [[order({'采购数量':'0.123456789012345678901'})],[order({'采购数量':Number.MAX_SAFE_INTEGER}),order({'采购数量':1})]]){
  const result=enrich({records});assert.equal(result.rows[0][8].v,'待核对');assert.match(note(result),/精确|精度/);
 }
 assert.equal(enrich({records:[order({'采购数量':'0001,234.00'})]}).rows[0][8].v,1234);
});

test('完整渠道路径与五维组对应，不把同系列其他供应商或渠道采购混入',()=>{
 const items=[{}, {'供应商简称':'其他供应商','采购订单号':'CGDD000001'}, {'销售分渠道':'海外事业一部*美洲部','采购订单号':'CGDD000002'}];
 const result=enrich({items,records:[order({'单据编号':'CGDD099999','采购数量':150}),order({'单据编号':'CGDD088888','供应商':'其他供应商','采购数量':40}),order({'单据编号':'CGDD077777','事业部':'海外事业一部*美洲部','采购数量':60})]});
 assert.equal(rowFor(result,['徐英','海外事业一部','海外事业一部*欧洲部*亚马逊欧盟组','迈锐'])[8].v,150);
 assert.equal(rowFor(result,['徐英','海外事业一部','海外事业一部*欧洲部*亚马逊欧盟组','其他供应商'])[8].v,40);
 assert.equal(rowFor(result,['徐英','海外事业一部','海外事业一部*美洲部','迈锐'])[8].v,60);
});

test('分工目录全称简称加正确物料定位下单人，创建人不替代采购下单人',()=>{
 const catalog=directory([{'采购下单人':'徐英','正确物料编码':'001','供应商':'嘉兴迈锐医疗器械有限公司','供应商简称':'迈锐'}]);
 const result=enrich({items:[{}, {'采购下单人':'李奇','采购订单号':'CGDD000002'}],catalog,records:[order({'单据编号':'CGDD099999','供应商':'嘉兴迈锐医疗器械有限公司','创建人':'孙立柱','采购数量':150})]});
 assert.equal(rowFor(result,['徐英'])[8].v,150);assert.equal(rowFor(result,['李奇'])[8].v,0);
});

test('同单据和物料能唯一归组时使用确切来源，不因宽事业部产生跨渠道重复分配',()=>{
 const result=enrich({items:[{}, {'销售分渠道':'海外事业一部*美洲部','采购订单号':'CGDD000002'}],records:[order({'事业部':'海外事业一部','采购数量':80})]});
 assert.equal(rowFor(result,['徐英','海外事业一部','海外事业一部*欧洲部*亚马逊欧盟组'])[8].v,80);
 assert.equal(rowFor(result,['徐英','海外事业一部','海外事业一部*美洲部'])[8].v,0);
});

test('同单据物料不能覆盖采购下单人、渠道、系列或目录供应商的明确冲突',()=>{
 for(const extra of [{'采购下单人':'李奇'},{'销售分渠道':'海外事业一部*美洲部'},{'系列':'AX35'},{'供应商':'完全不同的供应商有限公司'}]){
  const result=enrich({records:[order(extra)]});assert.match(String(result.rows[0][8].v),/待核对/);assert(result.issues.length,JSON.stringify(extra));
 }
 const ownerCatalog=directory([{'采购下单人':'李奇','正确物料编码':'001','供应商':'嘉兴迈锐医疗器械有限公司','供应商简称':'迈锐'}]);
 const ownerConflict=enrich({catalog:ownerCatalog,records:[order({'供应商':'嘉兴迈锐医疗器械有限公司'})]});assert.match(String(ownerConflict.rows[0][8].v),/待核对/);assert(ownerConflict.issues.length);
 const supplierCatalog=directory([{'采购下单人':'徐英','正确物料编码':'001','供应商':'嘉兴迈锐医疗器械有限公司','供应商简称':'迈锐'},{'采购下单人':'徐英','正确物料编码':'001','供应商':'其他供应商有限公司','供应商简称':'其他供应商'}]);
 const supplierConflict=enrich({catalog:supplierCatalog,records:[order({'供应商':'其他供应商有限公司'})]});assert.match(String(supplierConflict.rows[0][8].v),/待核对/);assert(supplierConflict.issues.length);
});

test('同订单物料本期已明确移入新组时只取本期归属，不重复分给上期组',()=>{
 const previous=current([{'采购下单人':'李奇'}]);
 const result=enrich({previous,records:[order({'采购数量':60})]});
 assert.equal(rowFor(result,['徐英'])[8].v,60);assert.equal(rowFor(result,['李奇'])[8].v,0);
});

test('同供应商物料在宽事业部跨组且无确切订单时保留待核对，不随意选一组或多算',()=>{
 const result=enrich({items:[{}, {'销售分渠道':'海外事业一部*美洲部','采购订单号':'CGDD000002'}],records:[order({'单据编号':'CGDD099999','事业部':'海外事业一部','采购数量':80})]});
 assert(result.rows.every(row=>String(row[8].v).includes('待核对')));assert(result.issues.length);assert.match(note(result),/采购明细/);assert.match(note(result),/2/);
});

test('无对应来源不从SKU猜系列，异供应商不混入，核对保留来源',()=>{
 const result=enrich({records:[order({'单据编号':'CGDD099999','供应商':'不同公司','物料编码':'新编码','SKU':'AX25-NEW','采购数量':777})]});
 assert.equal(result.rows[0][8].v,0);assert(result.issues.length);assert.match(note(result),/采购明细/);assert.match(note(result),/CGDD099999|新编码|2/);
});

test('空数量、错误或无缓存公式不可作为零；有效零仍是数值',()=>{
 for(const qty of ['',{t:'e',v:23,w:'#REF!'},{t:'n',f:'SUM(A99:A100)'},{t:'z',f:'SUM(A99:A100)',v:0}]){
  const result=enrich({records:[order({'采购数量':qty})]});assert(result.issues.length,JSON.stringify(qty));assert.match(note(result),/数量|公式|错误|缓存/);assert.notEqual(result.rows[0][8].v,0,'不完整数量不能显示确定零');
 }
 const zero=enrich({records:[order({'采购数量':0})]});assert.equal(zero.rows[0][8].v,0);assert.equal(zero.rows[0][8].t,'n');
});

test('缺失无效或错误日期需核对，不静默漏算成已确认零',()=>{
 for(const date of ['', '2026-02-30',{t:'e',v:15,w:'#VALUE!'}]){
  const result=enrich({records:[order({'采购日期':date})]});assert(result.issues.length,JSON.stringify(date));assert.match(note(result),/日期|错误/);assert.notEqual(result.rows[0][8].v,0);
 }
});

test('1904日期系统及原生Excel日期按实际日期筛选',()=>{
 const data=input([order({'采购日期':C.dateCell({v:'2026-11-13'}).v-1462,'采购数量':12})]);data.book.Workbook={WBProps:{date1904:true}};
 const now=current(),args={filename:FILENAME,current:now,comparison:C.seriesComparison(null,now),purchase:P.parse(data)};
 assert.equal(P.enrich(args).rows[0][8].v,12);
});

test('缺失或重复必需表头明确拒绝，不接受空数量默认值',()=>{
 for(const field of ['采购日期','单据编号','供应商','物料编码','采购数量']){
  assert.throws(()=>P.parse(input([order()],PH.filter(header=>header!==field))),/缺少|缺失|未找到|表头/);
  assert.throws(()=>P.parse(input([order()],[...PH,field])),/重复|表头/);
 }
});

test('未上传或日期无法识别时明确展示缺项；插列保留备注判断且导出同步表头',()=>{
 const now=current(),comparison=C.seriesComparison(null,now);comparison[0][8]={t:'s',v:'人工备注保留',z:'@'};comparison[0][9]={t:'s',v:'原判断内容',z:'@'};
 const result=enrich({records:null,comparison});assert.equal(result.rows[0][8].v,'未上传采购订单');assert.deepEqual(result.rows[0].slice(9),comparison[0].slice(8));
 const dated=enrich({filename:'无日期履约表.xlsx'});assert.match(String(dated.rows[0][8].v),/日期待确认/);
 const legacy=enrich({filename:'无日期履约表.xlsx',records:null,comparison});assert.deepEqual(legacy.columns,C.SERIES_COLUMNS);assert.deepEqual(legacy.rows,comparison);
 const book=C.workbook({...now,seriesComparison:result.rows,seriesColumns:result.columns}),saved=X.read(X.write(book,{type:'buffer',bookType:'xlsx'})),sheet=saved.Sheets.系列对比;
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1,defval:''})[0],result.columns);assert.equal(sheet.I2.v,'未上传采购订单');assert.equal(sheet.J2.v,'人工备注保留');assert.equal(sheet.K2.v,'原判断内容');assert.equal(sheet['!autofilter'].ref,'A1:K2');
});
