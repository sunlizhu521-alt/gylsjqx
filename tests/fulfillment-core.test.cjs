const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
function fixture(){const w=X.utils.book_new();for(const name of C.SHEETS){const h=[...C.COLUMNS].reverse();const r=h.map(k=>k==='采购下单人'?'原表采购员':k==='物料编码'?'00123':k==='SKU'?'SKU-001':k==='合同约定交期'?'2026年8月14日':k==='生产中交付时间'?'/':k==='未生产预计交付时间'?'':k==='未交付数量'?0:'样例');X.utils.book_append_sheet(w,X.utils.aoa_to_sheet([['标题'],h,r,r]),name);}return w;}
test('五表按姓名顺序匹配不同列位，保留重复、零和文本编码',()=>{const r=C.merge(fixture());assert.equal(r.rows.length,10);assert.deepEqual(r.counts.map(x=>x.sheet),C.SHEETS);assert.equal(r.rows[0][0].v,'原表采购员');assert.equal(C.COLUMNS[1],'事业部');assert.equal(r.rows[0][7].v,'00123');assert.equal(r.rows[0][10].v,0);assert.equal(r.rows[0][16].t,'n');assert.equal(r.rows[0][17].v,'/');assert.equal(r.rows[0][18].v,'');const w=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(w.Sheets.汇总表.Q2.z,C.DATE_FORMAT);assert.equal(w.Sheets.汇总表['!autofilter'].ref,'A1:T11');});
test('日期真实类型、1904体系、闰日、无效日期与空白',()=>{const c=C.dateCell({v:'2026-08-14 23:59:59'},false);assert.equal(X.SSF.format(c.z,c.v),'2026年8月14日');assert.equal(C.dateCell({v:c.v-1462},true).v,c.v);assert.ok(C.dateCell({v:'2026-02-29'}).invalid);assert.equal(C.dateCell({v:'2024-02-29'}).t,'n');assert.equal(C.dateCell({v:''}).v,'');assert.ok(C.dateCell({v:'10月中旬'}).invalid);});
test('缺表、缺列、重复列明确阻止导出；日期异常保留且溯源',()=>{let w=fixture();delete w.Sheets.徐英;assert.throws(()=>C.merge(w),/缺少工作表/);w=fixture();w.Sheets.徐英.A2.v='缺失';assert.throws(()=>C.merge(w),/缺少列/);w=fixture();w.Sheets.徐英.U2={t:'s',v:'SKU'};w.Sheets.徐英['!ref']='A1:U4';assert.throws(()=>C.merge(w),/重复列/);w=fixture();w.Sheets.徐英.D3={t:'s',v:'待定'};let r=C.merge(w);assert.equal(r.warnings.length,1);assert.match(r.warnings[0],/徐英!D3/);});
test('导出名称使用上传名称和日期，更新版替换为汇总版',()=>{
 assert.equal(C.exportFilename('订单履约明细跟进表截止2026.10.8-10.10更新版.xlsx'),'订单履约明细跟进表截止2026.10.8-10.10汇总版.xlsx');
 assert.equal(C.exportFilename('订单2026.11.1更新版.XLS'),'订单2026.11.1汇总版.xlsx');
 assert.equal(C.exportFilename('订单2026.10.8.xlsx'),'订单2026.10.8汇总版.xlsx');
 assert.equal(C.exportFilename('订单汇总版.xlsx'),'订单汇总版.xlsx');
});
function snapshot(items){return {rows:items.map(([code,qty,owner='采购员',supplier='供应商'])=>C.COLUMNS.map(k=>({t:k==='未交付数量'&&typeof qty==='number'?'n':'s',v:({'物料编码':code,'未交付数量':qty,'采购下单人':owner,'供应商简称':supplier,'SKU':'SKU-'+code,'物料名称':'物料'+code})[k]??''}))),counts:[],sources:[],warnings:[]};}
test('差异按物料汇总，保留多采购员及供应商，精确小数、新增、删除和零',()=>{
 const old=snapshot([['001',0.1],['001',0.2],['002',5],['003',0],['004',10]]);
 const now=snapshot([['001',0.3],['004',2,'甲','厂甲'],['004',3,'乙','厂乙'],['005',4],['006','']]);
 const diff=C.difference(old,now);assert.deepEqual(diff.map(r=>[r[2].v,r[5].v]),[['004',-5],['005',4],['002',-5]]);
 assert.equal(diff[0][0].v,'甲 / 乙');assert.equal(diff[0][1].v,'厂甲 / 厂乙');
 assert.deepEqual(C.difference(null,now),[]);
 const book=C.workbook({...now,differences:diff});assert.deepEqual(book.SheetNames,['汇总表','差异']);assert.equal(book.Sheets.差异.F2.v,-5);assert.equal(book.Sheets.差异['!autofilter'].ref,'A1:F4');
});
test('首次基准、重复上传、历史推进及异常不修改旧快照',()=>{
 const a={filename:'a.xlsx',result:snapshot([['001',5]])},b={filename:'b.xlsx',result:snapshot([['001',3]])};
 const first=C.advance(null,a);assert.equal(first.previous,null);
 assert.equal(C.advance(first,{...a,filename:'副本.xlsx'}),first);
 const second=C.advance(first,b);assert.equal(second.previous,a);assert.equal(C.advance(second,b),second);
 assert.throws(()=>C.advance(second,{result:snapshot([['001','待定']])}),/不是有效数字/);
 assert.throws(()=>C.advance(second,{result:snapshot([['',2]])}),/物料编码为空/);
 assert.equal(second.current,b);assert.equal(second.previous,a);
 assert.deepEqual(C.difference(snapshot([['001',1]]),snapshot([['001','']])).map(r=>r[5].v),[-1]);
});
