const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
function fixture(){const w=X.utils.book_new();for(const name of C.SHEETS){const h=[...C.COLUMNS.slice(0,23)].reverse().concat('下单日期');const r=h.map(k=>k==='下单日期'?'':k==='下单月份'?'2026-10':k==='采购下单人'?'原表采购员':k==='物料编码'?'00123':k==='SKU'?'SKU-001':k==='合同约定交期'?'2026年8月14日':k==='生产中交付时间'?'/':k==='未生产预计交付时间'?'':k==='未交付数量'?0:'样例');X.utils.book_append_sheet(w,X.utils.aoa_to_sheet([['标题'],h,r,r]),name);}return w;}
test('四表按姓名顺序匹配不同列位，保留重复、零和文本编码',()=>{const r=C.merge(fixture());assert.equal(r.rows.length,8);assert.deepEqual(r.counts.map(x=>x.sheet),C.SHEETS);assert.equal(r.rows[0][0].v,'原表采购员');assert.equal(C.COLUMNS[1],'事业部');assert.equal(r.rows[0][7].v,'00123');assert.equal(r.rows[0][10].v,0);assert.equal(r.rows[0][16].t,'n');assert.equal(r.rows[0][17].v,'/');assert.equal(r.rows[0][18].v,'');const w=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(w.Sheets.汇总表.R2.z,C.DATE_FORMAT);assert.equal(w.Sheets.汇总表['!autofilter'].ref,'A1:U9');});
test('日期真实类型、1904体系、闰日、无效日期与空白',()=>{const c=C.dateCell({v:'2026-08-14 23:59:59'},false);assert.equal(X.SSF.format(c.z,c.v),'2026年8月14日');assert.equal(C.dateCell({v:c.v-1462},true).v,c.v);assert.ok(C.dateCell({v:'2026-02-29'}).invalid);assert.equal(C.dateCell({v:'2024-02-29'}).t,'n');assert.equal(C.dateCell({v:''}).v,'');assert.ok(C.dateCell({v:'10月中旬'}).invalid);});
test('历史订单按2025年1月保存真实日期，忽略空格且不猜测其他说明',()=>{
 for(const value of ['历史订单',' 历史\n订单 '])for(const date1904 of [false,true]){const cell=C.dateCell({v:value},date1904,2030);assert.equal(cell.t,'n');assert.equal(cell.v,45658);assert.equal(X.SSF.format(cell.z,cell.v),'2025年1月');}
 assert.ok(C.dateCell({v:'历史订单待确认'}).invalid);assert.equal(C.dateCell({v:''}).v,'');assert.equal(C.dateCell({v:'/'}).v,'/');
 const source=fixture();source.Sheets.徐英.X3={t:'s',v:'历史订单'};source.Sheets.徐英.C3={t:'s',v:'历史订单'};const result=C.merge(source),sheet=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true}).Sheets.汇总表;
 assert.equal(sheet.D2.t,'n');assert.equal(X.SSF.format(sheet.D2.z,sheet.D2.v),'2025年1月');assert.equal(result.rows[0][C.COLUMNS.indexOf('下单月份')].v,45658);assert.ok(!result.warnings.some(w=>w.includes('历史订单')));assert.equal(source.Sheets.徐英.X3.v,'历史订单');
});
test('旧缓存仅迁移日期字段的历史订单，不改数量与其他文字',()=>{
 const original=snapshot([['MAT',123]]);original.sources=[{sheet:'徐英',row:2}];for(const key of ['下单日期','下单月份','合同约定交期','生产中交付时间','未生产预计交付时间','物料名称'])original.rows[0][C.COLUMNS.indexOf(key)]={t:'s',v:'历史订单'};
 original.warnings=['徐英!D2：日期“历史订单”无法识别，保留原文','徐英!E2：日期“待确认”无法识别，保留原文'];const migrated=C.currentScope(original);
 assert.equal(migrated.rows[0][C.COLUMNS.indexOf('下单日期')].v,45658);assert.equal(migrated.rows[0][C.COLUMNS.indexOf('下单月份')].v,45658);assert.equal(migrated.rows[0][C.COLUMNS.indexOf('未交付数量')].v,123);assert.equal(migrated.rows[0][C.COLUMNS.indexOf('物料名称')].v,'历史订单');assert.deepEqual(migrated.warnings,[original.warnings[1]]);assert.equal(original.rows[0][C.COLUMNS.indexOf('下单日期')].v,'历史订单');
});
test('下单月份斜杠默认2026年4月，仅处理指定字段并保留源值',()=>{
 for(const date1904 of [false,true]){const source=fixture();source.Workbook={WBProps:{date1904}};source.Sheets.徐英.C3={t:'s',v:' / '};source.Sheets.徐英.C4={t:'s',v:''};source.Sheets.徐英.X3={t:'s',v:'/'};
 const result=C.merge(source),cell=result.rows[0][C.COLUMNS.indexOf('下单月份')];assert.equal(cell.t,'n');assert.equal(cell.v,46113);assert.equal(X.SSF.format(cell.z,cell.v),'2026年4月');assert.equal(result.rows[1][20].v,'');assert.equal(result.rows[0][23].v,'/');assert.equal(result.rows[0][17].v,'/');assert.equal(source.Sheets.徐英.C3.v,' / ');
 const cached=structuredClone(result);cached.rows[0][20]={t:'s',v:' / '};assert.deepEqual(C.currentScope(cached).rows[0][20],cell);assert.equal(cached.rows[0][20].v,' / ');}
});
test('年月文本保存真实日期并显示中文月份，无效月份保留原文',()=>{
 for(const value of ['2025-11','2025-1','2025年11月',' 2025 / 11 ']){const cell=C.dateCell({v:value});assert.equal(cell.t,'n');assert.equal(X.SSF.format(cell.z,cell.v),value==='2025-1'?'2025年1月':'2025年11月');}
 for(const value of ['2025-0','2025-13','2025-11待定'])assert.ok(C.dateCell({v:value}).invalid);
 const source=fixture();source.Sheets.徐英.C3={t:'s',v:'2025-11'};source.Sheets.徐英.X3={t:'s',v:'2025-11'};source.Sheets.徐英.N3={t:'s',v:'2025-11'};const result=C.merge(source);assert.equal(X.SSF.format(result.rows[0][20].z,result.rows[0][20].v),'2025年11月');assert.equal(source.Sheets.徐英.C3.v,'2025-11');assert.equal(result.rows[0][9].v,'2025-11');
 const saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(saved.Sheets.汇总表.D2.t,'n');assert.equal(X.SSF.format(saved.Sheets.汇总表.D2.z,saved.Sheets.汇总表.D2.v),'2025年11月');
 const cached=structuredClone(result);cached.rows[0][20]={t:'s',v:'2025-11'};cached.rows[0][23]={t:'s',v:'2025-11'};cached.warnings=['徐英!X3：日期“2025-11”无法识别，保留原文','徐英!X4：日期“2025-13”无法识别，保留原文'];const migrated=C.currentScope(cached);assert.equal(migrated.rows[0][20].t,'n');assert.equal(migrated.rows[0][23].t,'n');assert.deepEqual(migrated.warnings,[cached.warnings[1]]);
});
test('OA备货单号缺失或斜杠填无，已有编号和其他单号不变，预览导出缓存一致',()=>{
 for(const value of [undefined,'',' \n ','/',' / ','OA001','0','无']){const source=fixture();source.Sheets.徐英.B3={t:'s',v:value};source.Sheets.徐英.A3={t:'s',v:''};const result=C.merge(source),expected=['OA001','0','无'].includes(value)?value:'无';assert.equal(result.rows[0][21].v,expected);assert.equal(result.rows[0][22].v,'');assert.equal(source.Sheets.徐英.B3.v,value);
 const saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}));assert.equal(saved.Sheets.汇总表.E2.v,expected);
 result.rows[0][21]={t:'s',v:value};assert.equal(C.currentScope(result).rows[0][21].v,expected);assert.equal(C.summaryRows(result.rows)[0][4].v,expected);}
 const source=fixture();delete source.Sheets.徐英.B2;const result=C.merge(source);assert.equal(result.rows[0][21].v,'无');assert.ok(result.warnings.includes('徐英：缺少列“OA备货单号”，已填“无”'));
 const old=snapshot([['MAT',123]]);old.rows[0]=old.rows[0].slice(0,20);assert.equal(C.currentScope(old).rows[0][21].v,'无');assert.equal(C.summaryRows(old.rows)[0][4].v,'无');assert.equal(old.rows[0].length,20);
});
test('缺表、缺列、重复列明确阻止导出；日期异常保留且溯源',()=>{let w=fixture();delete w.Sheets.徐英;assert.throws(()=>C.merge(w),/缺少工作表/);w=fixture();w.Sheets.徐英.D2.v='缺失';assert.throws(()=>C.merge(w),/缺少列/);w=fixture();w.Sheets.徐英.Y2={t:'s',v:'SKU'};w.Sheets.徐英['!ref']='A1:Y4';assert.throws(()=>C.merge(w),/重复列/);w=fixture();w.Sheets.徐英.G3={t:'s',v:'待定'};let r=C.merge(w);assert.equal(r.warnings.length,1);assert.match(r.warnings[0],/徐英!G3/);});
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
 const r=C.merge(w),row=r.rows[4];assert.equal(row[21].v,'OA001');assert.equal(row[22].v,'CGDD001');assert.equal(X.SSF.format(row[20].z,row[20].v),'2026年9月');
 assert.equal(r.sources[4].sheet,'邢芳芳1017');
 for(const col of ['A','B','C'])delete s[col+'2'];const missing=C.merge(w);assert.deepEqual(missing.rows[4].slice(20).map(c=>c.v),['','无','','']);assert.equal(missing.warnings.length,3);
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
 const cell=saved.Sheets.汇总表.R2;
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

test('汇总21列按指定顺序，月份及OA字段对应，旧缓存缺月份不借用下单日期',()=>{
 const expected=['事业部','销售分渠道','运营','下单月份','OA流程号','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','为生产预计交付时间','是否正常履约'];
 const w=fixture();w.Sheets.徐英.X3={t:'s',v:'2026年10月23日'};w.Sheets.徐英.B2.v='OA流程号';w.Sheets.徐英.B3.v='OA123';
 const result=C.merge(w),book=C.workbook(result),sheet=book.Sheets.汇总表;
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],expected);
 assert.equal(X.SSF.format(sheet.D2.z,sheet.D2.v),'2026年10月');
 assert.equal(sheet.E2.v,'OA123');assert.equal(sheet.I2.v,'00123');assert.equal(sheet.L2.v,0);
 const oldRow=structuredClone(result.rows[0]);oldRow[20]={t:'s',v:''};assert.equal(C.summaryRows([oldRow])[0][3].v,'');assert.equal(C.summaryRows([result.rows[0].slice(0,23)])[0][3].v,sheet.D2.v);
});

test('四人月份按各自表头读取，月份显示年月而交期保留具体日期',()=>{
 const book=X.utils.book_new(),months=['2025-11',46275,'历史订单','/'];
 for(const [i,name] of C.SHEETS.entries()){
  const columns=C.COLUMNS.filter(k=>k!=='下单日期'),headers=columns.slice(i*3).concat(columns.slice(0,i*3));
  const row=headers.map(k=>k==='下单月份'?months[i]:k==='物料编码'?'000123':k==='SKU'?'000SKU':k==='OA备货单号'?'0000123':C.QUANTITY_FIELDS.includes(k)?1:['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(k)?'2026-10-23':'样例');
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,row]),name);
 }
 const original=structuredClone(book),result=C.merge(book),preview=C.summaryRows(result.rows),saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true}).Sheets.汇总表;
 assert.deepEqual(preview.map(row=>X.SSF.format(row[3].z,row[3].v)),['2025年11月','2026年9月','2025年1月','2026年4月']);
 assert.deepEqual(result.warnings,[]);assert.deepEqual(book,original);
 for(let i=0;i<4;i++){
  const month=saved['D'+(i+2)];assert.equal(month.t,'n');assert.equal(month.z,'yyyy"年"m"月"');assert.equal(month.v,preview[i][3].v);
  for(const col of ['R','S','T'])assert.equal(X.SSF.format(saved[col+(i+2)].z,saved[col+(i+2)].v),'2026年10月23日');
  for(const [col,value] of [['E','0000123'],['I','000123'],['J','000SKU']]){assert.equal(saved[col+(i+2)].t,'s');assert.equal(saved[col+(i+2)].v,value);}
 }
 assert.ok(C.workbook(result).Sheets.汇总表['!cols'][3].wch>=14);
});

test('汇总数量用真实数值格式显示整数、负数和实际小数位，空白仍为空白',()=>{
 const result=snapshot([['MAT',0],['EMPTY','']]),values=[0,-7,1.25,1e-7,-0.0000012,10];
 values.forEach((v,i)=>result.rows[0][10+i]={t:typeof v==='number'?'n':'s',v,z:'General'});
 const original=structuredClone(result),preview=C.summaryRows(result.rows)[0],sheet=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true,sheetStubs:true}).Sheets.汇总表;
 const formats=['0','0','0.00','0.0000000','0.0000000','0'],display=['0','-7','1.25','0.0000001','-0.0000012','10'];
 for(let i=0;i<6;i++){const cell=sheet[X.utils.encode_cell({r:1,c:11+i})];assert.equal(cell.t,'n');assert.equal(cell.v,values[i]);assert.equal(cell.z,formats[i]);assert.equal(C.formatCell(cell),display[i]);assert.equal(preview[11+i].z,cell.z);assert.equal(C.formatCell(preview[11+i]),display[i]);}
 const empty=C.summaryRows(result.rows)[1][11];assert.equal(empty.v,'');assert.notEqual(empty.t,'n');assert.ok(sheet.L3==null||sheet.L3.v==='');assert.deepEqual(result,original);
});

test('旧缓存月份在显示时规范化，不改记录、数量、差异或历史基准',()=>{
 const current={filename:'本次.xlsx',result:snapshot([['MAT',5]])},previous={filename:'上次.xlsx',result:snapshot([['MAT',7]])},history={version:1,current,previous};
 current.result.differences=C.difference(previous.result,current.result);const diffBefore=structuredClone(C.workbook(current.result).Sheets.差异);
 for(const [value,expected] of [[46275,'2026年9月'],['2026-10-23','2026年10月'],['2025-11','2025年11月'],['历史订单','2025年1月'],['/','2026年4月'],['',''],['2026-13','2026-13']]){
  current.result.rows[0][20]={t:typeof value==='number'?'n':'s',v:value,z:'General'};const before=structuredClone(history),preview=C.summaryRows(current.result.rows),cell=preview[0][3],book=C.workbook(current.result);
  assert.equal(cell.z?X.SSF.format(cell.z,cell.v):cell.v,expected);assert.deepEqual(history,before);assert.equal(C.advance(history,{...current}),history);assert.deepEqual(book.Sheets.差异,diffBefore);
 }
 const source=fixture();source.Sheets.徐英.C3={t:'s',v:'2026-13'};source.Sheets.徐英.C4={t:'s',v:''};const result=C.merge(source);
 assert.equal(C.summaryRows(result.rows)[0][3].v,'2026-13');assert.equal(C.summaryRows(result.rows)[1][3].v,'');assert.ok(result.warnings.some(w=>w.includes('徐英!C3')&&w.includes('2026-13')));
});
