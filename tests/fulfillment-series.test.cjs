const test=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const HEADERS=['采购下单人','事业部','销售分渠道','系列','上期未交付数据','本期未交付数据','差异数量','备注'];
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
  ['甲','国内事业部','店铺','S',0.3,0.5,0.2,''],
  ['甲','海外事业一部','店铺','S',0,2,2,''],['甲','国内事业部','新渠道','S',0,3,3,''],
  ['甲','国内事业部','店铺','新系列',0,4,4,''],['丙','国内事业部','店铺','S',0,5,5,''],
  ['','','','',0,0,0,''],['甲','国内事业部','店铺','旧系列',7,0,-7,''],['乙','国内事业部','店铺','S',5,0,-5,'']
 ]);
 assert.deepEqual({before,current},original);
 for(const row of rows){assert(row.slice(0,4).every(c=>c.t==='s'));assert(row.slice(4,7).every(c=>c.t==='n'));assert.equal(row[7].t,'s');assert.equal(row[7].v,'');}
});

test('首次没有上期留空，本期零和空数量保留，存在空上期时才填零',()=>{
 const current=snapshot([['甲','部门','渠道','零',0],['甲','部门','渠道','空','']]);
 assert.deepEqual(values(C.seriesComparison(null,current)),[['甲','部门','渠道','零','',0,'',''],['甲','部门','渠道','空','',0,'','']]);
 assert.deepEqual(values(C.seriesComparison(snapshot([]),current)),[['甲','部门','渠道','零',0,0,0,''],['甲','部门','渠道','空',0,0,0,'']]);
 assert.deepEqual(C.seriesComparison(null,snapshot([])),[]);
});

test('只使用四个人员来源，兼容日期后缀，采购下单人取字段而非sheet名称',()=>{
 const sourceNames=['徐英','操思敏1113','邢芳芳20261113','李奇','孙立柱','汇总表','徐英备份'];
 const current=snapshot(sourceNames.map((_,i)=>['真实采购员','部门','渠道','S',i<4?i+1:999]),sourceNames);
 assert.deepEqual(values(C.seriesComparison(null,current)),[['真实采购员','部门','渠道','S','',10,'','']]);
 const previous=snapshot([['真实采购员','部门','渠道','S',2],['真实采购员','部门','渠道','S',999]],['徐英','孙立柱']);
 assert.deepEqual(values(C.seriesComparison(previous,current)),[['真实采购员','部门','渠道','S',2,10,8,'']]);
});

test('微小数量、精确小数和负数保持精度，联合键不以简单分隔符拼接',()=>{
 const current=snapshot([
  ['甲','部','渠','S',1e-7],['甲','部','渠','S',2e-7],
  ['甲','部','渠','小数',0.1],['甲','部','渠','小数',0.2],['甲','部','渠','小数',-0.1],
  ['A|B','C','D','E',1],['A','B|C','D','E',2]
 ]);
 assert.deepEqual(values(C.seriesComparison(null,current)),[
  ['甲','部','渠','S','',3e-7,'',''],['甲','部','渠','小数','',0.2,'',''],
  ['A|B','C','D','E','',1,'',''],['A','B|C','D','E','',2,'','']
 ]);
});

test('差异按本期减上期精确计算正负零和微小小数，备注始终留空',()=>{
 const previous=snapshot([['甲','部','渠','增加',0.1],['甲','部','渠','减少',0.3],['甲','部','渠','不变',0.3],['甲','部','渠','微小',1e-7]]);
 const current=snapshot([['甲','部','渠','增加',0.3],['甲','部','渠','减少',0.1],['甲','部','渠','不变',0.3],['甲','部','渠','微小',3e-7]]);
 const rows=C.seriesComparison(previous,current);
 assert.deepEqual(rows.map(row=>[row[3].v,row[6].v,row[7].v]),[['增加',0.2,''],['减少',-0.2,''],['不变',0,''],['微小',2e-7,'']]);
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

test('系列对比按指定八列导出，保留数字、空白上期与差异、空备注及普通筛选',()=>{
 const current=snapshot([['甲','国内事业部','渠道','系列',0.3],['乙','','','',0]]);
 const result={...current,review:[],differences:[],seriesComparison:C.seriesComparison(null,current)};
 const bytes=X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),saved=X.read(bytes,{cellNF:true,sheetStubs:true});
 assert.deepEqual(saved.SheetNames,['汇总表','合并核对','系列对比']);
 const sheet=saved.Sheets.系列对比;
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1,defval:''}),[HEADERS,['甲','国内事业部','渠道','系列','',0.3,'',''],['乙','','','','',0,'','']]);
 assert.equal(sheet['!autofilter'].ref,'A1:H3');
 assert.equal(sheet.F2.t,'n');assert.equal(sheet.F3.t,'n');assert.equal(sheet.E2?.v??'','');assert.equal(sheet.G2?.v??'','');assert.equal(sheet.H2?.v??'','');
 assert(Object.values(sheet).filter(c=>c&&typeof c==='object').every(c=>!c.f));
 assert.deepEqual(C.workbook(current).SheetNames,['汇总表'],'没有系列对比属性的事业部分导出不额外加表');
});
