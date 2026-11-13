const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const HEADERS=['采购下单人','事业部','销售分渠道','系列','上期未交付数据','本期未交付数据','差异数量','备注','判断'];
const fields=['采购下单人','事业部','销售分渠道','系列','未交付数量'];
function snapshot(items,sources){
 const result={rows:items.map(values=>C.COLUMNS.map(key=>{
  const i=fields.indexOf(key),value=i<0?'':values[i]??'';
  return value&&typeof value==='object'?{...value}:{t:typeof value==='number'?'n':'s',v:value};
 })),counts:[],warnings:[]};
 if(sources)result.sources=sources.map((sheet,i)=>({sheet,row:i+2}));
 return result;
}
const values=rows=>rows.map(row=>row.map(cell=>cell.v));
function orders(items){
 const defaults={'采购下单人':'甲','事业部':'部','销售分渠道':'渠','系列':'S','未交付数量':1,'下单月份':'2026-11'};
 return {rows:items.map(item=>C.COLUMNS.map(key=>{const value=({...defaults,...item})[key]??'';return typeof value==='object'?{...value}:{t:typeof value==='number'?'n':'s',v:value};})),sources:items.map(item=>item.source),counts:[],warnings:[]};
}

test('系列对比按四维联合键精确汇总，两期并集按本期优先排序，不改输入',()=>{
 assert.deepEqual(C.SERIES_COLUMNS,HEADERS);
 const before=snapshot([
  ['甲','国内事业部','店铺','S',0.1],['甲','国内事业部','店铺','S',0.2],
  ['甲','国内事业部','店铺','旧系列',7],['乙','国内事业部','店铺','S',5]
 ]);
 const current=snapshot([
  [' 甲 ','国内事业部 ',' 店铺','S ',0.4],['甲','国内事业部','店铺','S',0.1],
  ['甲','海外事业一部','店铺','S',2],['甲','国内事业部','新渠道','S',3],
  ['甲','国内事业部','店铺','新系列',4],['丙','国内事业部','店铺','S',5],
  ['','','','',0]
 ]);
 const original=structuredClone({before,current}),rows=C.seriesComparison(before,current);
 assert.deepEqual(values(rows),[
  ['甲','国内事业部','店铺','S',0.3,0.5,0.2,'',''],
  ['甲','海外事业一部','店铺','S',0,2,2,'',''],['甲','国内事业部','新渠道','S',0,3,3,'',''],
  ['甲','国内事业部','店铺','新系列',0,4,4,'',''],['丙','国内事业部','店铺','S',0,5,5,'',''],
  ['甲','国内事业部','店铺','旧系列',7,0,-7,'',''],['乙','国内事业部','店铺','S',5,0,-5,'','']
 ]);
 assert.deepEqual({before,current},original);
 for(const row of rows){assert(row.slice(0,4).every(c=>c.t==='s'));assert(row.slice(4,7).every(c=>c.t==='n'));assert.equal(row[7].t,'s');assert.equal(row[7].v,'');assert.equal(row[8].t,'s');assert.equal(row[8].v,'');}
});

test('首次没有上期保留零和空数量，有上期时零差异整行不输出',()=>{
 const current=snapshot([['甲','部门','渠道','零',0],['甲','部门','渠道','空','']]);
 assert.deepEqual(values(C.seriesComparison(null,current)),[['甲','部门','渠道','零','',0,'','',''],['甲','部门','渠道','空','',0,'','','']]);
 assert.deepEqual(values(C.seriesComparison(snapshot([]),current)),[]);
 assert.deepEqual(C.seriesComparison(null,snapshot([])),[]);
});

test('只使用四个人员来源，兼容日期后缀，采购下单人取字段而非sheet名称',()=>{
 const sourceNames=['徐英','操思敏1113','邢芳芳20261113','李奇','孙立柱','汇总表','徐英备份'];
 const current=snapshot(sourceNames.map((_,i)=>['真实采购员','部门','渠道','S',i<4?i+1:999]),sourceNames);
 assert.deepEqual(values(C.seriesComparison(null,current)),[['真实采购员','部门','渠道','S','',10,'','','']]);
 const previous=snapshot([['真实采购员','部门','渠道','S',2],['真实采购员','部门','渠道','S',999]],['徐英','孙立柱']);
 assert.deepEqual(values(C.seriesComparison(previous,current)),[['真实采购员','部门','渠道','S',2,10,8,'','']]);
});

test('微小数量、精确小数和负数保持精度，联合键不以简单分隔符拼接',()=>{
 const current=snapshot([
  ['甲','部','渠','S',1e-7],['甲','部','渠','S',2e-7],
  ['甲','部','渠','小数',0.1],['甲','部','渠','小数',0.2],['甲','部','渠','小数',-0.1],
  ['A|B','C','D','E',1],['A','B|C','D','E',2]
 ]);
 assert.deepEqual(values(C.seriesComparison(null,current)),[
  ['甲','部','渠','S','',3e-7,'','',''],['甲','部','渠','小数','',0.2,'','',''],
  ['A|B','C','D','E','',1,'','',''],['A','B|C','D','E','',2,'','','']
 ]);
});

test('差异按本期减上期精确计算，零差异过滤且备注始终留空',()=>{
 const previous=snapshot([['甲','部','渠','增加',0.1],['甲','部','渠','减少',0.3],['甲','部','渠','不变',0.3],['甲','部','渠','微小',1e-7]]);
 const current=snapshot([['甲','部','渠','增加',0.3],['甲','部','渠','减少',0.1],['甲','部','渠','不变',0.3],['甲','部','渠','微小',3e-7]]);
 const rows=C.seriesComparison(previous,current);
 assert.deepEqual(rows.map(row=>[row[3].v,row[6].v,row[7].v]),[['增加',0.2,''],['减少',-0.2,''],['微小',2e-7,'']]);
 assert(rows.every(row=>row[6].t==='n'&&row[7].t==='s'&&!row[6].f&&!row[7].f));
});

test('数量错误、非数字、无结果公式及超出安全范围不静默当零',()=>{
 for(const quantity of ['待确认',{t:'e',v:23},{t:'n',f:'1+2'},NaN,Infinity,Number.MAX_SAFE_INTEGER+1]){
  const current=snapshot([['甲','部门','渠道','S',quantity]]);
  assert.throws(()=>C.seriesComparison(null,current),/数量|数字|范围|公式|错误/);
  assert.throws(()=>C.seriesComparison(current,snapshot([])),/数量|数字|范围|公式|错误/);
 }
 assert.throws(()=>C.seriesComparison(null,snapshot([
  ['甲','部','渠','S',Number.MAX_SAFE_INTEGER],['甲','部','渠','S',1]
 ])),/范围|精确|数量/);
});

test('单物料数量有效但系列总量溢出时，历史推进失败且原快照不变',()=>{
 const previous=snapshot([['甲','部门','渠道','S',1]]),current=snapshot([
  ['甲','部门','渠道','S',6000000000000000],['甲','部门','渠道','S',6000000000000000]
 ]);
 previous.rows[0][7]={t:'s',v:'MAT-A'};
 current.rows.forEach((row,i)=>{row[7]={t:'s',v:'MAT-'+(i?'B':'A')};});
 const history={version:1,previous:null,current:{filename:'有效历史.xlsx',result:previous}},original=structuredClone(history);
 assert.throws(()=>C.advance(history,{filename:'系列溢出.xlsx',result:current}),/范围|精确|数量/);
 assert.deepEqual(history,original);
});

test('两期数量各自有效但差值溢出时明确拦截，历史保持不变',()=>{
 const previous=snapshot([['甲','部门','渠道','S',-6000000000000000]]),current=snapshot([['甲','部门','渠道','S',6000000000000000]]);
 previous.rows[0][7]={t:'s',v:'MAT-OLD'};current.rows[0][7]={t:'s',v:'MAT-NEW'};
 assert.doesNotThrow(()=>C.seriesComparison(null,previous));assert.doesNotThrow(()=>C.seriesComparison(null,current));
 assert.throws(()=>C.seriesComparison(previous,current),/范围|精确|数量/);
 const history={version:1,previous:null,current:{filename:'有效上期.xlsx',result:previous}},original=structuredClone(history);
 assert.throws(()=>C.advance(history,{filename:'差值溢出.xlsx',result:current}),/范围|精确|数量/);
 assert.deepEqual(history,original);
});

test('正差异只比较本期最新月份与上期同月，唯一采购单物料定位新增行并压缩连续行号',()=>{
 const old=[1,2,3].map(i=>({'采购订单号':'PO'+i,'物料编码':'M'+i,'SKU':'SKU'+i,source:{sheet:'徐英1112',row:i+3}}));
 const before=orders([...old,{'采购订单号':'OLD','物料编码':'OLD','下单月份':'2026-10','未交付数量':5}]);
 const current=orders([
  {...old[2],source:{sheet:'李奇',row:40}},
  {...old[0],'未交付数量':2,'OA备货单号':'变更OA','合同约定交期':'2027-01-01','是否正常履约':'延期',source:{sheet:'徐英1113',row:8}},
  {...old[1],'下单月份':'2026-11-29',source:{sheet:'操思敏',row:90}},
  {'采购订单号':'PO4','物料编码':'M4','SKU':'SKU4',source:{sheet:'徐英',row:12}},
  {'采购订单号':'PO5','物料编码':'M5','SKU':'SKU5',source:{sheet:'徐英',row:13}},
  {'采购订单号':'OLD','物料编码':'OLD','下单月份':'2026-10','未交付数量':5}
 ]);
 const input=structuredClone({before,current}),rows=C.seriesComparison(before,current);
 assert.equal(rows[0][6].v,3);
 assert.equal(rows[0][8].v,'2026年11月净增2行（上期3行，本期5行）；涉及行：徐英第12-13行');
 assert.equal(rows[0][7].v,'');assert.equal(rows[0][8].t,'s');assert.equal(rows[0][8].z,'@');
 assert.deepEqual({before,current},input);
 const saved=X.read(X.write(C.workbook({...current,seriesComparison:rows}),{type:'buffer',bookType:'xlsx'}));
 assert.equal(saved.Sheets.系列对比.I2.v,rows[0][8].v);assert.equal(saved.Sheets.系列对比.H2?.v??'','');
});

test('选择本期最新月份而非上期最新月份，且月份比较不混入其他四维分组',()=>{
 const before=orders([
  {'采购订单号':'A','物料编码':'M1'},
  {'采购订单号':'FUTURE','物料编码':'M9','下单月份':'2026-12'},
  {'系列':'其他','采购订单号':'OTHER','物料编码':'M8','下单月份':'2027-01'}
 ]);
 const current=orders([
  {'采购订单号':'A','物料编码':'M1',source:{sheet:'李奇',row:2}},
  {'采购订单号':'B','物料编码':'M2',source:{sheet:'李奇',row:12}},
  {'采购订单号':'C','物料编码':'M3',source:{sheet:'李奇',row:15}},
  {'系列':'其他','采购订单号':'OTHER','物料编码':'M8','下单月份':'2027-01'}
 ]);
 const rows=C.seriesComparison(before,current);
 assert.equal(rows.length,1);assert.equal(rows[0][6].v,1);
 assert.equal(rows[0][8].v,'2026年11月净增2行（上期1行，本期3行）；涉及行：李奇第12、15行');
});

test('有有效OA号时回退匹配，采购单物料联合键可区分同单不同物料与SKU',()=>{
 const before=orders([
  {'OA备货单号':'OA1','物料编码':'M1','SKU':'S1'},
  {'采购订单号':'PO1','物料编码':'M2','SKU':'S2'}
 ]);
 const current=orders([
  {'OA备货单号':'OA1','物料编码':'M1','SKU':'S1','未交付数量':2,source:{sheet:'李奇',row:6}},
  {'采购订单号':'PO1','OA备货单号':'后来补充','物料编码':'M2','SKU':'S2',source:{sheet:'操思敏',row:7}},
  {'采购订单号':'PO1','物料编码':'M3','SKU':'S3',source:{sheet:'操思敏',row:12}},
  {'OA备货单号':'OA2','物料编码':'M4','SKU':'S4',source:{sheet:'徐英',row:19}}
 ]);
 assert.equal(C.seriesComparison(before,current)[0][8].v,'2026年11月净增2行（上期2行，本期4行）；涉及行：操思敏第12行；徐英第19行');
});

test('仅数量增加或最新月份行数未增加时不推断新增单，旧月份新增行也不计入',()=>{
 const before=orders([{'采购订单号':'A','物料编码':'M','未交付数量':1}]);
 const current=orders([
  {'采购订单号':'A','物料编码':'M','未交付数量':2},
  {'采购订单号':'OLD','物料编码':'OLD','下单月份':'2026-10','未交付数量':10}
 ]);
 const rows=C.seriesComparison(before,current);assert.equal(rows[0][6].v,11);assert.equal(rows[0][8].v,'');
 const replacement=orders([{'采购订单号':'B','物料编码':'N','未交付数量':20}]);
 assert.equal(C.seriesComparison(before,replacement)[0][8].v,'','行数不变时更换订单编码也不能报新增');
});

test('缺编号、重复标识或同时新增删除时仅给本期相关行，不能认定唯一新增行',()=>{
 const scenarios=[
  [[{'物料编码':'M'}],[{'物料编码':'M'},{'物料编码':'N'}]],
  [[{'OA备货单号':'无','物料编码':'M'}],[{'OA备货单号':'无','物料编码':'M'},{'OA备货单号':'/','物料编码':'N'}]],
  [[{'采购订单号':'A','物料编码':'M'}],[{'采购订单号':'A','物料编码':'M'},{'采购订单号':'A','物料编码':'M'}]],
  [[{'采购订单号':'OLD','物料编码':'OLD'}],[{'采购订单号':'A','物料编码':'M'},{'采购订单号':'B','物料编码':'N'}]]
 ];
 for(const [old,now]of scenarios){
  const before=orders(old),current=orders(now.map((row,i)=>({...row,source:{sheet:'徐英',row:i+12}})));
  assert.equal(C.seriesComparison(before,current)[0][8].v,'2026年11月净增1行（上期1行，本期2行）；本期相关行（新增行待确认）：徐英第12-13行');
 }
 const before=orders([{'采购订单号':'A','物料编码':'M'},{'采购订单号':'A','物料编码':'M'}]);
 const current=orders([{'采购订单号':'A','物料编码':'M'},{'采购订单号':'B','物料编码':'N'},{'采购订单号':'C','物料编码':'P'}]);
 assert.equal(C.seriesComparison(before,current)[0][8].v,'2026年11月净增1行（上期2行，本期3行）；本期相关行（新增行待确认）：汇总表第2-4行');
 const shiftedBefore=orders([{'采购订单号':'A','物料编码':'M','下单月份':'2026-10'},{'采购订单号':'B','物料编码':'N'}]);
 const shiftedNow=orders([{'采购订单号':'A','物料编码':'M','未交付数量':2},{'采购订单号':'B','物料编码':'N'}]);
 assert.equal(C.seriesComparison(shiftedBefore,shiftedNow)[0][8].v,'2026年11月净增1行（上期1行，本期2行）；本期相关行（新增行待确认）：汇总表第2-3行','旧明细跨月调整不能标成明确新增行');
});

test('月份为空或无效不借下单日期推断，斜杠和历史订单沿用已有月份口径',()=>{
 for(const month of ['', '-', ' - ', '月份待定','2026-13']){
  const current=orders([{'采购订单号':'A','物料编码':'M','下单月份':month,'下单日期':'2026-11-01'}]);
  assert.equal(C.seriesComparison(orders([]),current)[0][8].v,'');
 }
 for(const [month,label]of [['/','2026年4月'],['历史订单','2025年1月']]){
  const current=orders([{'采购订单号':'A','物料编码':'M','下单月份':month}]);
  assert.equal(C.seriesComparison(orders([]),current)[0][8].v,`${label}净增1行（上期0行，本期1行）；涉及行：汇总表第2行`);
 }
});

test('负差异和首期判断留空，零差异整行删除仍先验证两期系列总量',()=>{
 const before=orders([{'采购订单号':'A','物料编码':'M','未交付数量':10}]);
 const current=orders([{'采购订单号':'A','物料编码':'M'},{'采购订单号':'B','物料编码':'N'}]);
 assert.equal(C.seriesComparison(before,current)[0][8].v,'');
 assert.equal(C.seriesComparison(null,current)[0][8].v,'');
 assert.deepEqual(C.seriesComparison(current,structuredClone(current)),[]);
 const unsafe=orders([{'物料编码':'M','未交付数量':6000000000000000},{'物料编码':'N','未交付数量':6000000000000000}]);
 assert.throws(()=>C.seriesComparison(unsafe,structuredClone(unsafe)),/范围|精确|数量/);
});

test('系列对比按指定九列导出，保留数字、空白上期与差异、H备注和I判断及普通筛选',()=>{
 const current=snapshot([['甲','国内事业部','渠道','系列',0.3],['乙','','','',0]]);
 const result={...current,review:[],differences:[],seriesComparison:C.seriesComparison(null,current)};
 const bytes=X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),saved=X.read(bytes,{cellNF:true,sheetStubs:true});
 assert.deepEqual(saved.SheetNames,['汇总表','合并核对','系列对比']);
 const sheet=saved.Sheets.系列对比;
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1,defval:''}),[HEADERS,['甲','国内事业部','渠道','系列','',0.3,'','',''],['乙','','','','',0,'','','']]);
 assert.equal(sheet['!autofilter'].ref,'A1:I3');
 assert.equal(sheet.F2.t,'n');assert.equal(sheet.F3.t,'n');assert.equal(sheet.E2?.v??'','');assert.equal(sheet.G2?.v??'','');assert.equal(sheet.H2?.v??'','');assert.equal(sheet.I2?.v??'','');
 assert(Object.values(sheet).filter(c=>c&&typeof c==='object').every(c=>!c.f));
 assert.deepEqual(C.workbook(current).SheetNames,['汇总表'],'没有系列对比属性的事业部分导出不额外加表');
});
