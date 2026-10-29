const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
function fixture(){const w=X.utils.book_new();for(const name of C.SHEETS){const h=[...C.COLUMNS.slice(0,23)].reverse().concat('下单日期');const r=h.map(k=>k==='下单日期'?'':k==='下单月份'?'2026-10':k==='采购下单人'?'原表采购员':k==='物料编码'?'00123':k==='SKU'?'SKU-001':k==='合同约定交期'?'2026年8月14日':k==='生产中交付时间'?'/':k==='未生产预计交付时间'?'':k==='未交付数量'?0:'样例');X.utils.book_append_sheet(w,X.utils.aoa_to_sheet([['标题'],h,r,r]),name);}return w;}
test('四表按姓名顺序匹配不同列位，保留重复、零和文本编码',()=>{const r=C.merge(fixture());assert.equal(r.rows.length,8);assert.deepEqual(r.counts.map(x=>x.sheet),C.SHEETS);assert.equal(r.rows[0][0].v,'原表采购员');assert.equal(C.COLUMNS[1],'事业部');assert.equal(r.rows[0][7].v,'00123');assert.equal(r.rows[0][10].v,0);assert.equal(r.rows[0][16].t,'n');assert.equal(r.rows[0][17].v,'/');assert.equal(r.rows[0][18].v,'');const w=X.read(X.write(C.workbook(r),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(w.Sheets.汇总表.S2.t,'s');assert.equal(w.Sheets.汇总表.S2.z,'@');assert.equal(w.Sheets.汇总表.S2.v,'2026年8月14日');assert.equal(w.Sheets.汇总表['!autofilter'].ref,'A1:V9');});
test('日期真实类型、1904体系、闰日、无效日期与空白',()=>{const c=C.dateCell({v:'2026-08-14 23:59:59'},false);assert.equal(X.SSF.format(c.z,c.v),'2026年8月14日');assert.equal(C.dateCell({v:c.v-1462},true).v,c.v);assert.ok(C.dateCell({v:'2026-02-29'}).invalid);assert.equal(C.dateCell({v:'2024-02-29'}).t,'n');assert.equal(C.dateCell({v:''}).v,'');assert.ok(C.dateCell({v:'10月中旬'}).invalid);});
test('历史订单按2025年1月保存真实日期，忽略空格且不猜测其他说明',()=>{
 for(const value of ['历史订单',' 历史\n订单 '])for(const date1904 of [false,true]){const cell=C.dateCell({v:value},date1904,2030);assert.equal(cell.t,'n');assert.equal(cell.v,45658);assert.equal(X.SSF.format(cell.z,cell.v),'2025年1月');}
 assert.ok(C.dateCell({v:'历史订单待确认'}).invalid);assert.equal(C.dateCell({v:''}).v,'');assert.equal(C.dateCell({v:'/'}).v,'/');
 const source=fixture();source.Sheets.徐英.X3={t:'s',v:'历史订单'};source.Sheets.徐英.C3={t:'s',v:'历史订单'};const result=C.merge(source),sheet=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true}).Sheets.汇总表;
 assert.equal(sheet.E2.t,'n');assert.equal(X.SSF.format(sheet.E2.z,sheet.E2.v),'2025年1月');assert.equal(result.rows[0][C.COLUMNS.indexOf('下单月份')].v,45658);assert.ok(!result.warnings.some(w=>w.includes('历史订单')));assert.equal(source.Sheets.徐英.X3.v,'历史订单');
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
 const saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.equal(saved.Sheets.汇总表.E2.t,'n');assert.equal(X.SSF.format(saved.Sheets.汇总表.E2.z,saved.Sheets.汇总表.E2.v),'2025年11月');
 const cached=structuredClone(result);cached.rows[0][20]={t:'s',v:'2025-11'};cached.rows[0][23]={t:'s',v:'2025-11'};cached.warnings=['徐英!X3：日期“2025-11”无法识别，保留原文','徐英!X4：日期“2025-13”无法识别，保留原文'];const migrated=C.currentScope(cached);assert.equal(migrated.rows[0][20].t,'n');assert.equal(migrated.rows[0][23].t,'n');assert.deepEqual(migrated.warnings,[cached.warnings[1]]);
});
test('OA备货单号缺失或斜杠填无，已有编号和其他单号不变，预览导出缓存一致',()=>{
 for(const value of [undefined,'',' \n ','/',' / ','OA001','0','无']){const source=fixture();source.Sheets.徐英.B3={t:'s',v:value};source.Sheets.徐英.A3={t:'s',v:''};const result=C.merge(source),expected=['OA001','0','无'].includes(value)?value:'无';assert.equal(result.rows[0][21].v,expected);assert.equal(result.rows[0][22].v,'');assert.equal(source.Sheets.徐英.B3.v,value);
 const saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}));assert.equal(saved.Sheets.汇总表.F2.v,expected);
 result.rows[0][21]={t:'s',v:value};assert.equal(C.currentScope(result).rows[0][21].v,expected);assert.equal(C.summaryRows(result.rows,result.sources)[0][5].v,expected);}
 const source=fixture();delete source.Sheets.徐英.B2;const result=C.merge(source);assert.equal(result.rows[0][21].v,'无');assert.ok(result.warnings.includes('徐英：缺少列“OA备货单号”，已填“无”'));
 const old=snapshot([['MAT',123]]);old.rows[0]=old.rows[0].slice(0,20);assert.equal(C.currentScope(old).rows[0][21].v,'无');assert.equal(C.summaryRows(old.rows)[0][5].v,'无');assert.equal(old.rows[0].length,20);
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

test('陆续交货取自然月末，内部真实日期导出为完整年月日文本',()=>{
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
 const cell=saved.Sheets.汇总表.S2;
 assert.equal(cell.t,'s');assert.equal(cell.z,'@');assert.equal(cell.v,'2028年2月29日');
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

test('汇总22列按指定顺序，月份及OA字段对应，旧缓存缺月份不借用下单日期',()=>{
 const expected=['采购跟单员','事业部','销售分渠道','运营','下单月份','OA流程号','供应商简称','产品线','系列','物料编码','SKU','物料名称','未交付数量','已下单未备料未生产','已备料未生产','生产中产品','完工未发产品','已发货数量','合同约定交期','生产中交付时间','为生产预计交付时间','是否正常履约'];
 const w=fixture();w.Sheets.徐英.X3={t:'s',v:'2026年10月23日'};w.Sheets.徐英.B2.v='OA流程号';w.Sheets.徐英.B3.v='OA123';
 const result=C.merge(w),book=C.workbook(result),sheet=book.Sheets.汇总表;
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],expected);
 assert.equal(X.SSF.format(sheet.E2.z,sheet.E2.v),'2026年10月');
 assert.equal(sheet.F2.v,'OA123');assert.equal(sheet.J2.v,'00123');assert.equal(sheet.M2.v,0);
 const oldRow=structuredClone(result.rows[0]);oldRow[20]={t:'s',v:''};assert.equal(C.summaryRows([oldRow])[0][4].v,'');assert.equal(C.summaryRows([result.rows[0].slice(0,23)])[0][4].v,sheet.E2.v);
});

test('四人月份按各自表头读取，月份为真实日期，其余非数量列为文本且交期显示年月日',()=>{
 const book=X.utils.book_new(),months=['2025-11',46275,'历史订单','/'];
 for(const [i,name] of C.SHEETS.entries()){
  const columns=C.COLUMNS.filter(k=>k!=='下单日期'),headers=columns.slice(i*3).concat(columns.slice(0,i*3));
  const row=headers.map(k=>k==='下单月份'?months[i]:k==='运营'?101:k==='物料编码'?'000123':k==='SKU'?'000SKU':k==='OA备货单号'?'0000123':C.QUANTITY_FIELDS.includes(k)?1:['合同约定交期','生产中交付时间','未生产预计交付时间'].includes(k)?'2026-10-23':'样例');
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,row]),name);
 }
 const original=structuredClone(book),result=C.merge(book),preview=C.summaryRows(result.rows,result.sources),saved=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true}).Sheets.汇总表;
 assert.deepEqual(preview.map(row=>X.SSF.format(row[4].z,row[4].v)),['2025年11月','2026年9月','2025年1月','2026年4月']);
 assert.deepEqual(result.warnings,[]);assert.deepEqual(book,original);
 for(let i=0;i<4;i++){
  const month=saved['E'+(i+2)];assert.equal(month.t,'n');assert.equal(month.z,'yyyy"年"m"月"');assert.equal(month.v,preview[i][4].v);
  for(const col of ['S','T','U'])assert.equal(saved[col+(i+2)].v,'2026年10月23日');
  for(let col=0;col<22;col++){if(col===4||(col>=12&&col<=17))continue;const cell=saved[X.utils.encode_cell({r:i+1,c:col})];assert.equal(cell.t,'s');assert.equal(cell.z,'@');assert.equal(typeof cell.v,'string');assert.equal(preview[i][col].t,'s');assert.equal(preview[i][col].z,'@');assert.equal(preview[i][col].v,cell.v);}
  assert.equal(saved['D'+(i+2)].v,'101');
  for(const [col,value] of [['F','0000123'],['J','000123'],['K','000SKU']]){assert.equal(saved[col+(i+2)].t,'s');assert.equal(saved[col+(i+2)].v,value);}
 }
 assert.ok(C.workbook(result).Sheets.汇总表['!cols'][4].wch>=14);
});

test('汇总数量用真实数值格式显示整数、负数和实际小数位，空白仍为空白',()=>{
 const result=snapshot([['MAT',0],['EMPTY','']]),values=[0,-7,1.25,1e-7,-0.0000012,10];
 values.forEach((v,i)=>result.rows[0][10+i]={t:typeof v==='number'?'n':'s',v,z:'General'});
 const original=structuredClone(result),preview=C.summaryRows(result.rows,result.sources)[0],sheet=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true,sheetStubs:true}).Sheets.汇总表;
 const formats=['0','0','0.00','0.0000000','0.0000000','0'],display=['0','-7','1.25','0.0000001','-0.0000012','10'];
 for(let i=0;i<6;i++){const cell=sheet[X.utils.encode_cell({r:1,c:12+i})];assert.equal(cell.t,'n');assert.equal(cell.v,values[i]);assert.equal(cell.z,formats[i]);assert.equal(C.formatCell(cell),display[i]);assert.equal(preview[12+i].z,cell.z);assert.equal(C.formatCell(preview[12+i]),display[i]);}
 const empty=C.summaryRows(result.rows,result.sources)[1][12];assert.equal(empty.v,'');assert.notEqual(empty.t,'n');assert.ok(sheet.M3==null||sheet.M3.v==='');assert.deepEqual(result,original);
});

test('旧缓存月份在显示时规范化，不改记录、数量、差异或历史基准',()=>{
 const current={filename:'本次.xlsx',result:snapshot([['MAT',5]])},previous={filename:'上次.xlsx',result:snapshot([['MAT',7]])},history={version:1,current,previous};
 current.result.differences=C.difference(previous.result,current.result);const diffBefore=structuredClone(C.workbook(current.result).Sheets.差异);
 for(const [value,expected] of [[46275,'2026年9月'],['2026-10-23','2026年10月'],['2025-11','2025年11月'],['历史订单','2025年1月'],['/','2026年4月'],['',''],['2026-13','2026-13']]){
  current.result.rows[0][20]={t:typeof value==='number'?'n':'s',v:value,z:'General'};const before=structuredClone(history),preview=C.summaryRows(current.result.rows),cell=preview[0][4],book=C.workbook(current.result);
  assert.equal(cell.z?X.SSF.format(cell.z,cell.v):cell.v,expected);assert.deepEqual(history,before);assert.equal(C.advance(history,{...current}),history);assert.deepEqual(book.Sheets.差异,diffBefore);
 }
 const source=fixture();source.Sheets.徐英.C3={t:'s',v:'2026-13'};source.Sheets.徐英.C4={t:'s',v:''};const result=C.merge(source);
 assert.equal(C.summaryRows(result.rows,result.sources)[0][4].v,'2026-13');assert.equal(C.summaryRows(result.rows,result.sources)[1][4].v,'');assert.ok(result.warnings.some(w=>w.includes('徐英!C3')&&w.includes('2026-13')));
});

function reviewFixture(records){
 const book=X.utils.book_new();
 for(const [i,person] of C.SHEETS.entries()){
  const columns=C.COLUMNS.slice(i*5).concat(C.COLUMNS.slice(0,i*5));
  const rows=records.filter((_,index)=>index%4===i).map((value,index)=>columns.map(field=>({'采购下单人':'表内其他人','事业部':'海外事业一部','销售分渠道':'海外事业一部','下单日期':'2026-10-01','物料编码':'000'+i+index,'未交付数量':1.25,...value})[field]??''));
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([columns.map(field=>field==='销售分渠道'?'销售分渠道部':field),...rows]),i===2?'邢芳芳1017':person);
 }
 return book;
}
const reviewObjects=rows=>rows.map(row=>Object.fromEntries(C.REVIEW_COLUMNS.map((key,i)=>[key,row[i].v])));
test('合并核对覆盖全部运营对应关系，不改原值，22列跟单员来自完整工作表名',()=>{
 const groups=[['倪芷涵 王程','海外事业一部*欧洲部*亚马逊英国组'],['雷维明 王青萍 胡鑫敏 富晓霞 范俊 赵雅杰 张梦露','海外事业一部*欧洲部*亚马逊欧盟组'],['金何强 王鑫','海外事业一部*多渠道部'],['薛文乐 赵恬幸','海外事业一部*美洲部'],['王志豪','国内事业部*TEMU业务部']];
 const records=groups.flatMap(([names,channel])=>names.split(' ').map(name=>({'运营':name,expected:channel}))),book=reviewFixture(records),before=structuredClone(book),result=C.merge(book),original=structuredClone(result);
 const audit=reviewObjects(C.reviewRows(result,{checks:[],issues:[]}));
 for(const record of records){const found=audit.find(row=>row.运营===record.运营&&row.核对字段==='销售分渠道');assert.equal(found.建议值,record.expected);assert.match(found.需要处理,/核实.*修改源表/);assert.equal(found.原值,'海外事业一部');}
 const domestic=audit.find(row=>row.运营==='王志豪'&&row.核对字段==='事业部');assert.equal(domestic.建议值,'国内事业部');
 assert.deepEqual(book,before);assert.deepEqual(result,original);
 const summary=C.summaryRows(result.rows,result.sources);assert.equal(C.SUMMARY_COLUMNS.length,22);assert.equal(C.SUMMARY_COLUMNS[0],'采购跟单员');
 assert.deepEqual(summary.map(row=>row[0].v),result.sources.map(source=>source.sheet));assert(summary.some(row=>row[0].v==='邢芳芳1017'));assert(!summary.some(row=>row[0].v==='表内其他人'));
 for(const record of records){record.销售分渠道=record.expected;record.事业部=record.expected.split('*')[0];}
 assert.deepEqual(C.reviewRows(C.merge(reviewFixture(records)),{checks:[],issues:[]}),[]);
});
test('所有日期检查空白，10月边界起检查渠道；未知运营不猜测，不匹配部分姓名',()=>{
 const inputs=[
  {'物料编码':'before','下单日期':'2026-09-30','运营':'倪芷涵'},
  {'物料编码':'cutoff','下单日期':'2026-10-01','运营':' 倪 芷 涵 '},
  {'物料编码':'after','下单日期':'2026-11-01','运营':'薛文乐'},
  {'物料编码':'blank','下单日期':'2025-11-01','事业部':' ','销售分渠道':'\n'},
  {'物料编码':'unknown','运营':'倪芷涵、王程'},
  {'物料编码':'validUnknown','运营':'未配置','销售分渠道':'海外事业一部*美洲部'},
  {'物料编码':'constructor','运营':'constructor'},
  {'物料编码':'other','事业部':'海外事业二部','销售分渠道':'海外事业二部','运营':'未配置'}
 ];
 const rows=reviewObjects(C.reviewRows(C.merge(reviewFixture(inputs)),{checks:[],issues:[]}));
 assert(!rows.some(row=>['before','validUnknown','other'].includes(row.物料编码)));
 assert.equal(rows.find(row=>row.物料编码==='cutoff').建议值,'海外事业一部*欧洲部*亚马逊英国组');
 assert.equal(rows.find(row=>row.物料编码==='after').建议值,'海外事业一部*美洲部');
 assert.equal(rows.filter(row=>row.物料编码==='blank').length,2);
 for(const code of ['unknown','constructor'])assert.equal(rows.find(row=>row.物料编码===code).建议值,'');
});
test('日期优先、月份回退、无效日期和范围冲突都明确报告，保留原日期',()=>{
 const inputs=[
  {'物料编码':'month','下单日期':'','下单月份':'2026-10','运营':'倪芷涵'},
  {'物料编码':'date','下单日期':'2026-10-31','下单月份':'2026-11','运营':'王程'},
  {'物料编码':'conflict','下单日期':'2026-09-30','下单月份':'2026-10','运营':'王程'},
  {'物料编码':'badDate','下单日期':'待定','下单月份':'2026-10','运营':'王程'},
  {'物料编码':'badMonth','下单日期':'2026-10-01','下单月份':'2026-13','运营':'王程'},
  {'物料编码':'empty','下单日期':'','下单月份':'','运营':'王程'},
  {'物料编码':'old','下单日期':'','下单月份':'历史订单','运营':'王程'},
  {'物料编码':'slash','下单日期':'','下单月份':'/','运营':'王程'}
 ];
 const result=C.merge(reviewFixture(inputs)),before=structuredClone(result),rows=reviewObjects(C.reviewRows(result,{checks:[],issues:[]}));
 for(const code of ['month','date'])assert(rows.some(row=>row.物料编码===code&&row.建议值==='海外事业一部*欧洲部*亚马逊英国组'));
 for(const code of ['conflict','badDate','badMonth','empty']){const issues=rows.filter(row=>row.物料编码===code);assert.equal(issues.length,1);assert.equal(issues[0].核对类型,'日期范围');assert.equal(issues[0].建议值,'');}
 for(const code of ['old','slash'])assert(!rows.some(row=>row.物料编码===code));assert.deepEqual(result,before);
});
test('数量合计和单元格异常写入同一核对表，导出重读数字格式、筛选和顺序',()=>{
 const V=require('../fulfillment-validation.js'),book=reviewFixture([{}, {}, {}, {}]),result=C.merge(book);result.rows[0][C.COLUMNS.indexOf('未交付数量')].v=9;
 const check=V.validate(book,result,C);assert.equal(check.passed,false);result.review=C.reviewRows(result,check);result.differences=[];
 const objects=reviewObjects(result.review),totals=objects.filter(row=>row.核对类型==='数量合计');assert.equal(totals.length,30);assert.equal(totals.filter(row=>row.来源工作表==='合计').length,6);
 assert(totals.some(row=>row.原表合计===1.25&&row.汇总合计===9&&row.差额===7.75&&row.结果==='不一致'));
 const issue=objects.find(row=>row.核对类型==='数量／转换异常');assert.equal(issue.来源工作表,'徐英');assert.equal(issue.来源位置,'K2');assert.match(issue.需要处理,/修正后重新上传/);
 const output=X.read(X.write(C.workbook(result),{type:'buffer',bookType:'xlsx'}),{cellNF:true});assert.deepEqual(output.SheetNames,['汇总表','合并核对','差异']);
 const sheet=output.Sheets.合并核对;assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],C.REVIEW_COLUMNS);assert.equal(sheet['!autofilter'].ref,'A1:N'+(result.review.length+1));
 const index=objects.findIndex(row=>row.原表合计===1.25&&row.差额===7.75)+2;assert.equal(sheet['J'+index].t,'n');assert.equal(sheet['J'+index].z,'0.00');assert.equal(sheet['L'+index].v,7.75);
 assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1,defval:''}).slice(1),result.review.map(row=>row.map(cell=>cell.v)));
});
test('销售分渠道部别名不能将失效公式当作无关辅助列丢弃，重名表头仍拦截',()=>{
 const V=require('../fulfillment-validation.js'),book=reviewFixture([{}]);book.Sheets.徐英.C2={t:'n',v:0,f:'#REF!'};
 const materialized=V.materialize(book);assert(materialized.issues.some(issue=>issue.sheet==='徐英'&&issue.cell==='C2'));assert.equal(materialized.book.Sheets.徐英.C2.f,'#REF!');
 book.Sheets.徐英.Y1={t:'s',v:'销售分渠道'};book.Sheets.徐英['!ref']='A1:Y2';assert.throws(()=>C.merge(book),/重复列“销售分渠道”/);
});
