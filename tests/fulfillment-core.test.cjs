const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
function fixture(){const w=X.utils.book_new();for(const name of C.SHEETS){const h=[...C.COLUMNS].reverse();const r=h.map(k=>k==='采购下单人'?'原表采购员':k==='物料编码'?'00123':k==='SKU'?'SKU-001':k==='合同约定交期'?'2026年8月14日':k==='生产中交付时间'?'/':k==='未生产预计交付时间'?'':k==='未交付数量'?0:'样例');X.utils.book_append_sheet(w,X.utils.aoa_to_sheet([['标题'],h,r,r]),name);}return w;}
test('四表按姓名顺序匹配不同列位，保留重复、零和文本编码',()=>{const r=C.merge(fixture());assert.equal(r.rows.length,8);assert.deepEqual(r.counts.map(x=>x.sheet),C.SHEETS);assert.equal(r.rows[0][0].v,'原表采购员');assert.equal(C.COLUMNS[1],'事业部');assert.equal(r.rows[0][7].v,'00123');assert.equal(r.rows[0][10].v,0);assert.equal(r.rows[0][16].t,'n');assert.equal(r.rows[0][17].v,'/');assert.equal(r.rows[0][18].v,'');const w=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(w.Sheets.汇总表.Q2.z,C.DATE_FORMAT);assert.equal(w.Sheets.汇总表['!autofilter'].ref,'A1:W9');});
test('日期真实类型、1904体系、闰日、无效日期与空白',()=>{const c=C.dateCell({v:'2026-08-14 23:59:59'},false);assert.equal(X.SSF.format(c.z,c.v),'2026年8月14日');assert.equal(C.dateCell({v:c.v-1462},true).v,c.v);assert.ok(C.dateCell({v:'2026-02-29'}).invalid);assert.equal(C.dateCell({v:'2024-02-29'}).t,'n');assert.equal(C.dateCell({v:''}).v,'');assert.ok(C.dateCell({v:'10月中旬'}).invalid);});
test('缺表、缺列、重复列明确阻止导出；日期异常保留且溯源',()=>{let w=fixture();delete w.Sheets.徐英;assert.throws(()=>C.merge(w),/缺少工作表/);w=fixture();w.Sheets.徐英.D2.v='缺失';assert.throws(()=>C.merge(w),/缺少列/);w=fixture();w.Sheets.徐英.X2={t:'s',v:'SKU'};w.Sheets.徐英['!ref']='A1:X4';assert.throws(()=>C.merge(w),/重复列/);w=fixture();w.Sheets.徐英.G3={t:'s',v:'待定'};let r=C.merge(w);assert.equal(r.warnings.length,1);assert.match(r.warnings[0],/徐英!G3/);});
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

test('姓名日期后缀、三列映射、缺列留空及旧历史兼容',()=>{
 const w=fixture();w.Sheets['邢芳芳1017']=w.Sheets.邢芳芳;delete w.Sheets.邢芳芳;
 const s=w.Sheets['邢芳芳1017'];s.B2.v='OA备货流程号';s.A3={t:'s',v:'CGDD001'};s.B3={t:'s',v:'OA001'};s.C3={t:'n',v:46275};
 const r=C.merge(w),row=r.rows[6];assert.equal(row[21].v,'OA001');assert.equal(row[22].v,'CGDD001');assert.equal(X.SSF.format(row[20].z,row[20].v),'2026年9月');
 assert.equal(r.sources[6].sheet,'邢芳芳1017');
 for(const col of ['A','B','C'])delete s[col+'2'];const missing=C.merge(w);assert.deepEqual(missing.rows[6].slice(20).map(c=>c.v),['','','']);assert.equal(missing.warnings.length,3);
 const old={...r,rows:r.rows.map(row=>row.slice(0,20))};assert.deepEqual(C.difference(old,r),[]);assert.doesNotThrow(()=>C.workbook(old));
 w.Sheets.邢芳芳=s;assert.throws(()=>C.merge(w),/多个工作表/);delete w.Sheets.邢芳芳;w.Sheets['邢芳芳备份']=s;delete w.Sheets['邢芳芳1017'];assert.throws(()=>C.merge(w),/缺少工作表/);
});

test('陆续交货取自然月末，文件名年份优先、显式年份优先并保留真实日期',()=>{
 const formatted=(value,year)=>{const c=C.dateCell({v:value},false,year);return X.SSF.format(c.z,c.v);};
 assert.equal(formatted('10月陆续交货',2026),'2026年10月31日');
 assert.equal(formatted(' 4月 陆续交货 ',2026),'2026年4月30日');
 assert.equal(formatted('2月陆续交货',2026),'2026年2月28日');
 assert.equal(formatted('2028年2月陆续交货',2026),'2028年2月29日');
 assert.equal(formatted('12月陆续交货',2026),'2026年12月31日');
 for(const v of ['0月陆续交货','13月陆续交货','10月中旬','10月陆续交货待确认'])assert.ok(C.dateCell({v},false,2026).invalid);
 const w=fixture();w.Sheets.徐英.G3={t:'s',v:'2月陆续交货'};
 const r=C.merge(w,'订单履约明细跟进表截止2028.2.1更新版.xlsx');
 assert.equal(r.warnings.length,0);assert.equal(w.Sheets.徐英.G3.v,'2月陆续交货');
 const saved=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{cellNF:true});
 const cell=saved.Sheets.汇总表.Q2;
 assert.equal(cell.t,'n');assert.equal(X.SSF.format(cell.z,cell.v),'2028年2月29日');
});

test('相同汇总补存或更新原文件时不推进对比基准',()=>{
 const result=C.merge(fixture()),previous={filename:'上次',result};
 const current={filename:'本次',result},history={version:1,previous,current};
 const next={...current,sourceBytes:new Uint8Array([1,2]).buffer};
 const upgraded=C.advance(history,next);assert.equal(upgraded.previous,previous);assert.equal(upgraded.current,next);
 assert.equal(C.advance(upgraded,{...next,sourceBytes:new Uint8Array([1,2]).buffer}),upgraded);
 const changed={...next,sourceBytes:new Uint8Array([1,3]).buffer};
 assert.equal(C.advance(upgraded,changed).current,changed);
 assert.equal(C.advance(upgraded,changed).previous,previous);
});

test('孙立柱及日期后缀表不汇总，旧历史按同样口径排除以免产生虚假差异',()=>{
 const book=fixture();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['无需校验表头'],['原始内容']]),'孙立柱1023');
 assert.equal(C.merge(book).rows.length,8);assert.equal(book.Sheets.孙立柱1023.A2.v,'原始内容');
 const result=snapshot([['001',5],['001',99],['002',9]]);
 result.sources=[{sheet:'徐英',row:2},{sheet:'孙立柱',row:2},{sheet:'孙立柱1023',row:3}];
 result.counts=[{sheet:'徐英',count:1},{sheet:'孙立柱',count:1},{sheet:'孙立柱1023',count:1}];
 const filtered=C.currentScope(result);
 assert.equal(filtered.rows.length,1);assert.equal(result.rows.length,3);
 assert.deepEqual(filtered.counts,[{sheet:'徐英',count:1}]);
 assert.deepEqual(C.difference(filtered,snapshot([['001',5]])),[]);
});
