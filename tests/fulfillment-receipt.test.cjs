const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js'),P=require('../fulfillment-purchase.js');
const FILENAME='订单履约明细跟进表截止2026.11.13-11.19更新版.xlsx';
const HEADERS=['入库日期','单据类型','单据编号','订单单号','供应商','物料编码','实收数量','应收数量','计价数量','创建人'];
const channel='海外事业一部*欧洲部*亚马逊欧盟组';
const cell=value=>value&&typeof value==='object'?structuredClone(value):{t:typeof value==='number'?'n':'s',v:value??''};
function current(items=[{}]){
 const defaults={'采购下单人':'徐英','事业部':'海外事业一部','销售分渠道':channel,'供应商简称':'迈锐','系列':'AX25','物料编码':'001','采购订单号':'CGDD015712','下单月份':'2026年11月','未交付数量':10};
 return {rows:items.map(item=>C.COLUMNS.map(key=>cell({...defaults,...item}[key]))),sources:items.map((_,i)=>({sheet:'徐英',row:i+2})),counts:[{sheet:'徐英',count:items.length}],warnings:[]};
}
function receipt(extra={}){return {'入库日期':'2026-11-13','单据类型':'标准采购入库','单据编号':'CGRK001','订单单号':'CGDD015712','供应商':'迈锐','物料编码':'001','实收数量':12,'应收数量':99,'计价数量':88,'创建人':'孙立柱',...extra};}
function input(records,headers=HEADERS){const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records.map(record=>headers.map(header=>record[header]??''))]),'入库明细');return {filename:'采购入库单.xlsx',book};}
function enrich({items=[{}],records=[receipt()],filename=FILENAME,comparison,previous=null,directory=null,source}={}){
 const now=current(items),purchase=records===null?null:P.parse(source||input(records),'receipt'),args={kind:'receipt',filename,current:now,previous,purchase,directory,comparison:comparison??C.seriesComparison(previous,now)};
 const before=structuredClone(args),result=P.enrich(args);assert.deepEqual(args,before,'不改写来源、历史分组或人工备注');return result;
}
const issueText=result=>JSON.stringify(result.issues);

test('入库单按入库日期和实收数量汇总，源订单匹配，无部门渠道亦可准确归组',()=>{
 const parsed=P.parse(input([receipt()]),'receipt');assert.equal(parsed.records[0].order,'CGDD015712');assert.equal(parsed.records[0].document,'CGRK001');assert.equal(parsed.records[0].qty,12);
 const result=enrich({records:[receipt({'入库日期':'2026-11-12','实收数量':999}),receipt({'实收数量':12}),receipt({'入库日期':'2026-11-19 23:59:59','实收数量':8}),receipt({'入库日期':'2026-11-20','实收数量':777})]});
 assert.equal(result.rows[0][8].v,20);assert.equal(result.rows[0][8].t,'n');assert.equal(result.rows[0][8].f,undefined);assert.equal(result.columns[8],'11.13-11.19采购入库单数量');assert.equal(result.stats.periodRows,2);assert.equal(result.stats.matchedRows,2);assert.deepEqual(result.issues,[]);
});

test('入库单据编号不当作采购单号；无对应源订单且有多个分组时不猜归属',()=>{
 const result=enrich({items:[{}, {'采购订单号':'CGDD888888','系列':'AX35'}],records:[receipt({'单据编号':'CGDD015712','订单单号':'CGDD999999'})]});
 assert(result.rows.every(r=>r[8].v==='待核对'));assert.equal(result.stats.uncertainRows,1);assert.match(issueText(result),/CGDD999999/);assert.equal(result.issues[0].document,'CGDD015712');assert.match(result.issues[0].source,/入库明细!G2/);
});

test('入库没有部门时仅同供应商和物料唯一组可回退，不混入异供应商或其他物料',()=>{
 const result=enrich({items:[{}, {'供应商简称':'其他供应商','采购订单号':'CGDD888888'}, {'物料编码':'002','采购订单号':'CGDD777777','系列':'AX35'}],records:[receipt({'订单单号':'CGDD999999'})]});
 assert.equal(result.rows.find(r=>r[3].v==='迈锐'&&r[4].v==='AX25')[8].v,12);assert.equal(result.stats.matchedRows,1);assert(result.rows.filter(r=>r[3].v!=='迈锐'||r[4].v!=='AX25').every(r=>r[8].v===0));
 const channels=enrich({items:[{}, {'销售分渠道':'海外事业一部*美洲部','采购订单号':'CGDD888888'}],records:[receipt({'订单单号':'CGDD999999'})]});assert(channels.rows.every(r=>r[8].v==='待核对'));assert.match(issueText(channels),/多个系列/);
 const missing=enrich({records:[receipt({'订单单号':'CGDD999999','供应商':'完全不同供应商'})]});assert.equal(missing.rows[0][8].v,0);assert.equal(missing.stats.unmatchedRows,1);
});

test('当前同订单物料唯一归属优先于历史，已消失记录仍可按历史源订单归组',()=>{
 const previous=current([{'采购下单人':'李奇'}]),result=enrich({previous});
 assert.equal(result.rows.find(r=>r[0].v==='徐英')[8].v,12);assert.equal(result.rows.find(r=>r[0].v==='李奇')[8].v,0);
 const gone=enrich({items:[],previous:current()});assert.equal(gone.rows[0][8].v,12);
});

test('同源订单物料跨两个系列时标待核对，不均分也不任选',()=>{
 const result=enrich({items:[{}, {'系列':'AX35'}]});assert.equal(result.stats.uncertainRows,1);assert(result.rows.every(r=>r[8].v==='待核对'));assert.match(issueText(result),/多个系列/);
});

test('来源订单不能掩盖供应商及显式分组冲突',()=>{
 for(const overrides of [{'供应商':'其他供应商'},{'事业部':'国内事业部'},{'采购下单人':'李奇'},{'销售分渠道':'另一渠道'},{'系列':'AX35'}]){
  const source=input([receipt(overrides)],[...HEADERS,'事业部','采购下单人','销售分渠道','系列']);const result=enrich({source});assert.equal(result.rows[0][8].v,'待核对');assert.match(issueText(result),/冲突/);
 }
});

test('供应商目录用于全称与简称对应，不从创建人推断采购人',()=>{
 const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['采购下单人','正确物料编码','供应商','供应商简称'],['徐英','001','嘉兴迈锐医疗器械有限公司','迈锐']]),'产品线明细');
 const result=enrich({records:[receipt({'供应商':'嘉兴迈锐医疗器械有限公司'})],directory:{filename:'分工.xlsx',book}});assert.equal(result.rows[0][8].v,12);assert.deepEqual(result.issues,[]);
});

test('重复入库明细逐行累加，零与小数保留数值格式，不按单据编号去重',()=>{
 const result=enrich({records:[receipt({'实收数量':0.1}),receipt({'实收数量':0.2}),receipt({'实收数量':0.1}),receipt({'实收数量':0})]});assert.equal(result.rows[0][8].v,0.4);assert.equal(result.rows[0][8].z,'0.0');assert.equal(result.stats.matchedRows,4);
});

test('缺失错误负数及精度异常的实收数量须核对，不能取应收或计价数量替代',()=>{
 for(const qty of ['',-1,{t:'e',v:23}, {t:'n',f:'SUM(G100:G101)'},'0.123456789012345678901']){
  const result=enrich({records:[receipt({'实收数量':qty})]});assert.equal(result.rows[0][8].v,'待核对');assert(result.issues.length);assert.match(issueText(result),/数量|公式|错误|精确/);
 }
 const result=enrich({records:[receipt({'实收数量':Number.MAX_SAFE_INTEGER}),receipt({'实收数量':1})]});assert.equal(result.rows[0][8].v,'待核对');assert.match(issueText(result),/精确/);
});

test('空或无效入库日期须核对，Excel 1904日期系统仍按实际日期筛选',()=>{
 for(const date of ['', '2026-02-30',{t:'e',v:15}]){const result=enrich({records:[receipt({'入库日期':date})]});assert.equal(result.rows[0][8].v,'待核对');assert.match(issueText(result),/入库日期|错误/);}
 const source=input([receipt({'入库日期':C.dateCell({v:'2026-11-19'}).v-1462})]);source.book.Workbook={WBProps:{date1904:true}};assert.equal(enrich({source}).rows[0][8].v,12);
});

test('合并数量仅首行计数，文本编号格式和明确合计行按原有规则处理',()=>{
 const source=input([receipt({'物料编码':{t:'n',v:1,z:'000'}}),receipt({'物料编码':'001','实收数量':''}),{'入库日期':'合计','实收数量':12}]);
 source.book.Sheets.入库明细['!merges']=[X.utils.decode_range('G2:G3')];const parsed=P.parse(source,'receipt');assert.equal(parsed.records.length,2);assert.equal(parsed.records[0].code,'001');assert.equal(parsed.records[1].qty,0);assert.equal(enrich({source}).rows[0][8].v,12);
});

test('缺失或重复必要入库字段拒绝，不能把采购日期或收料源号当成入库采购关联',()=>{
 for(const header of ['入库日期','订单单号','供应商','物料编码','实收数量']){
  assert.throws(()=>P.parse(input([receipt()],HEADERS.filter(h=>h!==header)),'receipt'),/缺少|表头/);
  assert.throws(()=>P.parse(input([receipt()],[...HEADERS,header]),'receipt'),/重复/);
 }
 assert.throws(()=>P.parse(input([receipt({'源单编号':'CGSL001'})],HEADERS.map(h=>h==='订单单号'?'源单编号':h)),'receipt'),/订单单号/);
 assert.throws(()=>P.parse(input([receipt({'采购日期':'2026-11-13'})],HEADERS.map(h=>h==='入库日期'?'采购日期':h)),'receipt'),/入库日期/);
});

test('未上传无期间或读取失败均有独立提示，列名跨月跨年更新且保留备注判断',()=>{
 const now=current(),comparison=C.seriesComparison(null,now);comparison[0][8]=cell('人工备注');comparison[0][9]=cell('原判断');
 const empty=enrich({records:null,comparison});assert.equal(empty.rows[0][8].v,'未上传采购入库单');assert.deepEqual(empty.rows[0].slice(9),comparison[0].slice(8));
 const missing=enrich({filename:'履约.xlsx'});assert.equal(missing.columns[8],'采购入库单数量（日期待确认）');assert.equal(missing.rows[0][8].v,'日期待确认');
 const legacy=enrich({filename:'履约.xlsx',records:null});assert.deepEqual(legacy.columns,C.SERIES_COLUMNS);
 const failed=P.enrich({kind:'receipt',filename:FILENAME,current:now,comparison,purchase:{filename:'坏文件.xlsx',error:'字段缺失'}});assert.equal(failed.rows[0][8].v,'入库文件待核对');
 for(const [filename,label]of [['履约2026.10.30-11.5.xlsx','10.30-11.5采购入库单数量'],['履约2026.12.26-1.1.xlsx','2026.12.26-2027.1.1采购入库单数量']])assert.equal(enrich({filename}).columns[8],label);
 const result=enrich({comparison}),book=C.workbook({...now,seriesComparison:result.rows,seriesColumns:result.columns}),saved=X.read(X.write(book,{type:'buffer',bookType:'xlsx'})),sheet=saved.Sheets.系列对比;
 assert.equal(sheet.I1.v,'11.13-11.19采购入库单数量');assert.equal(sheet.I2.t,'n');assert.equal(sheet.I2.v,12);assert.equal(sheet.J2.v,'人工备注');assert.equal(sheet.K2.v,'原判断');assert.equal(sheet['!autofilter'].ref,'A1:K2');
});
