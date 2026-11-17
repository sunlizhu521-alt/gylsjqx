const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const X=require('../vendor/xlsx.full.min.js');
const Zip=require('../vendor/jszip.min.js');
const Core=require('../funds-core.js');
const Validation=require('../funds-total-validation.js');
const HEADERS=['供应商','对公付款0A流程号','OA-申请付款金额','谈判明细','分批总金额','2026年11月2日','2026年11月16日','2026年11月支付','2026年12月支付'];
const serial=s=>(Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
function table(data,headers=HEADERS){
 const sheet=X.utils.aoa_to_sheet([headers,...data]),map={},canonical={'供应商':'供应商名称','OA-申请付款金额':'申请付款金额'};
 headers.forEach((h,c)=>{if(typeof h==='string')map[canonical[h]||h]=c;});
 const records=data.map((_,i)=>({r:i+1,source:{file:'源表.xlsx',sheet:'谈判付款',row:i+2},values:Object.fromEntries(Object.entries(map).map(([h,c])=>[h,sheet[X.utils.encode_cell({r:i+1,c})]||{t:'s',v:''}]))}));
 return {sheet,map,records,headers,headerRow:0,range:X.utils.decode_range(sheet['!ref']),date1904:false,name:'源表.xlsx',sheetName:'谈判付款'};
}
const row=(amount=100)=>['供应商甲','OA-001',amount,'11月2日－'+amount,amount,amount,null,amount];
const build=t=>Validation.build(t,Core.money);
const values=(result,r=0)=>Object.fromEntries(result.headers.map((header,c)=>[header,result.rows[r][c].v]));

test('逐原行保留重复记录，五组金额按分相等；计数不同不导致金额不一致且不改来源',()=>{
 const source=table([row(),['供应商甲','OA-001',0.3,'11月2日－0.1\n11月16日－0.2',0.3,0.3,0,0.3,0]]),before=JSON.stringify(source),result=build(source),data=values(result,1);
 assert.deepEqual(result.stats,{total:2,excludedNotPay:0,consistent:2,inconsistent:0,unverified:0});
 assert.equal(data['谈判明细金额合计'],0.3);assert.equal(data['谈判明细有效条数'],2);assert.equal(data['日期计划非零金额格数'],1);assert.equal(data['月度有数值月份数'],2);
 assert.equal(data['校验结果'],'一致');assert.equal(result.rows[1][5].z,'#,##0.00');assert.equal(JSON.stringify(source),before);
 assert.equal(result.headers.length,19);assert.ok(result.rows.every(r=>r.length===result.headers.length));assert.equal(result.canExport,undefined);
});

test('仅排除源OA列有效的不支付标记，保留原行号、零金额记录及来源数据',()=>{
 for(const marker of [{t:'s',v:'不支付'},{t:'s',v:'\u3000不 支\n付\u00a0'},{t:'s',f:'"不支付"',v:'不支付'}]){
  const source=table([row(),row(),row(0)]);source.sheet.B3=marker;
  const before=JSON.stringify(source),result=build(source);
  assert.deepEqual(result.stats,{total:2,excludedNotPay:1,consistent:2,inconsistent:0,unverified:0});
  assert.deepEqual(result.rows.map(r=>r[2].v),[2,4]);assert.equal(values(result,1)['OA-申请付款金额'],0);
  assert.equal(JSON.stringify(source),before);
 }
});

test('较长说明、其他列的不支付、空OA及无有效缓存的OA公式均不作为排除标记',()=>{
 const markers=[
  {t:'s',v:'是否不支付'},{t:'s',v:'不支付待确认'},{t:'s',v:'不支付（暂缓）'},
  {t:'s',v:'未提交OA'},{t:'s',v:''},{t:'s',v:'\u3000 '},
  {t:'e',v:'不支付',w:'#VALUE!'},{t:'s',f:'"不支付"',w:'不支付'},
  {t:'z',f:'"不支付"',v:'不支付'},{t:'str',f:'"不支付"',v:'',w:'不支付'}
 ];
 for(const marker of markers){
  const source=table([row()]);source.sheet.B2=marker;source.sheet.A2.v='不支付';
  const result=build(source);
  assert.deepEqual(result.stats,{total:1,excludedNotPay:0,consistent:1,inconsistent:0,unverified:0},JSON.stringify(marker));
 }
 const source=table([row()]);source.sheet.D2.v='不支付';delete source.sheet.B2;
 const result=build(source);assert.equal(result.stats.total,1);assert.equal(result.stats.excludedNotPay,0);assert.equal(result.stats.unverified,1);
});

test('不支付按map所指的原单元格判断，不使用固定列或过期记录值；缺少OA列不排除',()=>{
 const headers=[HEADERS[1],HEADERS[0],...HEADERS.slice(2)],moved=row(),source=table([[moved[1],moved[0],...moved.slice(2)]],headers);
 source.sheet.A2={t:'s',v:'不支付'};assert.equal(build(source).stats.excludedNotPay,1);
 source.sheet.A2={t:'s',v:'OA-001'};source.records[0].values['对公付款0A流程号']={t:'s',v:'不支付'};
 assert.equal(build(source).stats.excludedNotPay,0);
 delete source.map['对公付款0A流程号'];source.sheet.A2={t:'s',v:'不支付'};
 assert.equal(build(source).stats.total,1);assert.equal(build(source).stats.excludedNotPay,0);
});

test('有效不支付即使金额异常也不核对；全部排除时保留19列表头且不修改源表',()=>{
 const source=table([row(),row()]);source.sheet.B2={t:'s',v:'不支付'};source.sheet.B3={t:'s',v:'不支付'};
 source.sheet.C2={t:'e',v:7};source.sheet.D2={t:'s',v:'异常明细'};source.sheet.E3={t:'n',f:'1+1'};source.sheet.F3={t:'s',v:'无效金额'};
 const before=JSON.stringify(source),result=build(source);
 assert.deepEqual(result.stats,{total:0,excludedNotPay:2,consistent:0,inconsistent:0,unverified:0});
 assert.deepEqual(result.rows,[]);assert.deepEqual(result.headers,Validation.HEADERS);assert.equal(result.headers.length,19);assert.equal(JSON.stringify(source),before);
});

test('支付与不支付跨记录合并的金额仍归属不清，排除记录不能掩盖合并冲突',()=>{
 for(const c of [4,5]){
  const source=table([row(),row()]);source.sheet.B2={t:'s',v:'不支付'};source.sheet['!merges']=[{s:{r:1,c},e:{r:2,c}}];
  const before=JSON.stringify(source),result=build(source),data=values(result);
  assert.deepEqual(result.stats,{total:1,excludedNotPay:1,consistent:0,inconsistent:0,unverified:1});
  assert.equal(data['来源行号'],3);assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],new RegExp(c===4?'E2:E3':'F2:F3'));
  assert.equal(JSON.stringify(source),before);
 }
});

test('五组包含独立分批总金额，差额按各组合计减OA列，说明准确指向原单元格',()=>{
 const source=table([row()]);source.sheet.E2.v=90;const result=build(source),data=values(result);
 assert.equal(data['校验结果'],'不一致');assert.equal(data['分批总金额减OA差额'],-10);assert.equal(data['谈判明细减OA差额'],0);
 assert.match(data['需要处理'],/分批总金额E2/);assert.match(data['需要处理'],/日期计划F2、G2/);assert.match(data['需要处理'],/月度支付H2、I2/);assert.equal(result.stats.inconsistent,1);
});

test('有效零值含公式缓存零保留；空白组与缺少OA不是零，记录不丢弃',()=>{
 const source=table([row(0),row(0),row(10)]);source.sheet.E2={t:'n',f:'SUM(F2:G2)',v:0};delete source.sheet.F3;delete source.sheet.C4;
 const result=build(source),zero=values(result),blank=values(result,1),noOA=values(result,2);
 assert.equal(zero['校验结果'],'一致');assert.equal(zero['分批总金额'],0);assert.equal(zero['日期计划非零金额格数'],0);assert.equal(zero['月度有数值月份数'],1);
 assert.equal(blank['日期计划金额合计'],'');assert.equal(blank['校验结果'],'待核对');assert.match(blank['需要处理'],/空白不按零/);
 assert.equal(noOA['OA-申请付款金额'],'');assert.equal(noOA['谈判明细金额合计'],10);assert.equal(noOA['谈判明细减OA差额'],'');assert.equal(result.rows.length,3);
});

test('公式错误、无缓存及SheetJS占位零不当作金额零；保留可解析小计而不冒充完整差额',()=>{
 for(const broken of [{t:'n',f:'1+1'},{t:'z',f:'1+1',v:0},{t:'str',f:'1+1',v:''},{t:'e',v:7,w:'#DIV/0!'}]){
  const source=table([row()]);source.sheet.G2=broken;const data=values(build(source));
  assert.equal(data['日期计划金额合计'],100);assert.equal(data['校验结果'],'待核对');assert.equal(data['日期计划减OA差额'],'');assert.match(data['需要处理'],/G2公式结果错误或缺失/);assert.match(data['需要处理'],/仅保留可解析有效金额小计/);
 }
});

test('明细按有效笔数计数，保留正负金额及可解析小计，异常笔和括号说明定位到源行',()=>{
 const source=table([row(80),row(80)]);source.sheet.D2.v='1、11月2日－100（已确认）\n2、11月16日－－20';source.sheet.D3.v='11月2日－100\n不完整\n11月16日－－20';
 const result=build(source),valid=values(result),partial=values(result,1);
 assert.equal(valid['谈判明细金额合计'],80);assert.equal(valid['谈判明细有效条数'],2);assert.equal(valid['校验结果'],'一致');assert.match(valid['需要处理'],/D2第1笔括号说明/);
 assert.equal(partial['谈判明细金额合计'],80);assert.equal(partial['谈判明细有效条数'],2);assert.equal(partial['校验结果'],'待核对');assert.equal(partial['谈判明细减OA差额'],'');assert.match(partial['需要处理'],/D3第2笔/);
});

test('总额相等仍识别月份错位，按明细减月度提示金额及来源，19列结构和源表不变',()=>{
 const source=table([row()]);source.sheet.H2.v=80;source.sheet.I2={t:'n',v:20};
 const before=JSON.stringify(source),result=build(source),data=values(result);
 assert.deepEqual(result.stats,{total:1,excludedNotPay:0,consistent:0,inconsistent:1,unverified:0});
 assert.equal(data['谈判明细金额合计'],100);assert.equal(data['月度支付金额合计'],100);
 for(const field of ['谈判明细减OA差额','分批总金额减OA差额','日期计划减OA差额','月度支付减OA差额'])assert.equal(data[field],0);
 assert.match(data['需要处理'],/2026年11月[^\n]*谈判明细D2[^\n]*100\.00[^\n]*月度支付H2[^\n]*80\.00[^\n]*差额20\.00/);
 assert.match(data['需要处理'],/2026年12月[^\n]*0\.00[^\n]*I2[^\n]*20\.00[^\n]*差额-20\.00/);
 assert.equal(result.headers.length,19);assert.equal(result.rows[0].length,19);assert.equal(JSON.stringify(source),before);
 source.sheet.B2={t:'s',v:'不支付'};assert.equal(build(source).stats.excludedNotPay,1);assert.deepEqual(build(source).rows,[]);
});

test('同年月多笔按分汇总正负和零金额，接受带年及不带年的仅月份明细',()=>{
 const source=table([row(0.3)]);source.sheet.D2.v='1、2026年11月－0.1\n2、11月16日－0.3\n3、11月－-0.1\n4、12月－0';source.sheet.I2={t:'n',f:'0',v:0};
 const data=values(build(source));
 assert.equal(data['谈判明细金额合计'],0.3);assert.equal(data['谈判明细有效条数'],4);assert.equal(data['月度有数值月份数'],2);assert.equal(data['校验结果'],'一致');
 assert.doesNotMatch(data['需要处理'],/差额|无法完整解析/);
});

test('跨年同月分别核对，缺年份时拒绝跨年歧义及日期和月度表头年份冲突',()=>{
 const source=table([row()]);source.sheet.H1.v='2025年11月支付';source.sheet.I1.v='2026年11月支付';
 source.sheet.D2.v='2025年11月2日－40\n2026年11月－60';source.sheet.H2.v=60;source.sheet.I2={t:'n',v:40};
 let data=values(build(source));assert.equal(data['校验结果'],'不一致');
 assert.match(data['需要处理'],/2025年11月[^\n]*差额-20\.00/);assert.match(data['需要处理'],/2026年11月[^\n]*差额20\.00/);
 source.sheet.H2.v=40;source.sheet.I2.v=60;assert.equal(values(build(source))['校验结果'],'一致');
 source.sheet.D2.v='11月－100';data=values(build(source));assert.equal(data['校验结果'],'待核对');assert.equal(data['谈判明细有效条数'],0);assert.match(data['需要处理'],/年份.*唯一/);
 const conflict=table([row()]);conflict.sheet.H1.v='2027年11月支付';conflict.sheet.D2.v='11月－100';
 data=values(build(conflict));assert.equal(data['校验结果'],'待核对');assert.equal(data['谈判明细有效条数'],0);assert.match(data['需要处理'],/年份.*唯一|年份.*冲突/);
 conflict.sheet.H1.v='备注';conflict.sheet.I2={t:'n',v:0};data=values(build(conflict));
 assert.equal(data['谈判明细有效条数'],1);assert.equal(data['谈判明细金额合计'],100);assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年11月/);
});

test('明细月份有金额而月度列缺失、空白或公式异常时待核对，不能补零',()=>{
 const cases=[undefined,{t:'s',v:' '},{t:'e',v:7,w:'#DIV/0!'},{t:'n',f:'100'},{t:'z',f:'100',v:0}];
 for(const monthly of cases){
  const source=table([row()]);source.sheet.D2.v='2026年11月－100';source.sheet.I2={t:'n',v:0};
  if(monthly)source.sheet.H2=monthly;else delete source.sheet.H2;
  const data=values(build(source));assert.equal(data['校验结果'],'待核对',JSON.stringify(monthly));
  assert.match(data['需要处理'],/2026年11月/);assert.match(data['需要处理'],/H2/);assert.doesNotMatch(data['需要处理'],/2026年11月[^\n]*差额100\.00/);
 }
 const missing=table([row()]);missing.sheet.D2.v='2026年11月－100';missing.sheet.H1.v='备注';missing.sheet.I2={t:'n',v:0};
 const data=values(build(missing));assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年11月/);assert.match(data['需要处理'],/缺少.*月度|未找到.*月度/);
});

test('双方无金额的月份不提示，明确零金额配空白月度仍需核对',()=>{
 const source=table([row()]);let data=values(build(source));assert.equal(data['校验结果'],'一致');assert.doesNotMatch(data['需要处理'],/2026年12月/);
 source.sheet.I2={t:'n',v:0};data=values(build(source));assert.equal(data['校验结果'],'一致');assert.doesNotMatch(data['需要处理'],/2026年12月/);
 source.sheet.D2.v='11月－100\n12月－0';delete source.sheet.I2;data=values(build(source));
 assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年12月/);assert.match(data['需要处理'],/I2/);
});

test('同年月重复月度表头不自动合并掩盖问题，提示两处表头位置',()=>{
 const source=table([row()]);source.sheet.I1.v='2026年11月支付';source.sheet.H2.v=40;source.sheet.I2={t:'n',v:60};
 const data=values(build(source));assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年11月[^\n]*表头重复/);assert.match(data['需要处理'],/H1/);assert.match(data['需要处理'],/I1/);
});

test('不完整明细不以部分小计做月差额，其他组不完整时保留月差异但优先待核对',()=>{
 const source=table([row()]);source.sheet.D2.v='11月2日－100\n不完整';source.sheet.H2.v=80;source.sheet.I2={t:'n',v:20};
 let data=values(build(source));assert.equal(data['校验结果'],'待核对');assert.equal(data['谈判明细金额合计'],100);assert.match(data['需要处理'],/D2第2笔/);
 assert.doesNotMatch(data['需要处理'],/2026年(?:11|12)月[^\n]*差额/);
 source.sheet.D2.v='11月2日－100';delete source.sheet.C2;data=values(build(source));
 assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年11月[^\n]*差额20\.00/);
});

test('月合计及月差额超出安全整数分时待核对，即使跨月抵消后的总额为零',()=>{
 const max='90071992547409.91',source=table([row(0)]);source.sheet.I2={t:'n',v:0};
 source.sheet.D2.v='2026年11月－'+max+'\n11月－0.01\n12月－-'+max+'\n12月－-0.01';
 let data=values(build(source));assert.equal(data['谈判明细金额合计'],0);assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/2026年11月/);assert.match(data['需要处理'],/超出安全整数分范围/);
 source.sheet.D2.v='2026年11月－'+max+'\n12月－-'+max;source.sheet.H2={t:'s',v:'-'+max};source.sheet.I2={t:'s',v:max};
 data=values(build(source));assert.equal(data['谈判明细金额合计'],0);assert.equal(data['月度支付金额合计'],0);assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/差额[^\n]*超出安全整数分范围/);
});

test('计划和月度按表头识别，列插入移动、分散、隐藏均有效，命名日期元数据不计金额',()=>{
 const headers=['应付款时间','2026年12月支付','供应商','摘要','2026年11月16日','OA-申请付款金额','2026年11月支付','谈判明细','分批总金额','对公付款0A流程号','申请日期','2026年11月2日'];
 const source=table([[999,20,'甲','说明',30,100,80,'11月2日－70\n11月16日－30',100,'OA-1',999,70]],headers);
 source.sheet['!cols']=headers.map((_,c)=>({hidden:[1,4,6,11].includes(c)}));source.sheet.A1.z='yyyy"年"m"月"d"日"';source.sheet.K1.z='yyyy-mm-dd';
 assert.deepEqual(Validation.selectColumns(source),{dateColumns:[4,11],monthColumns:[1,6]});const result=build(source),data=values(result);
 assert.equal(data['日期计划金额合计'],100);assert.equal(data['月度支付金额合计'],100);assert.equal(data['校验结果'],'不一致');assert.equal(data['日期计划非零金额格数'],2);assert.match(data['需要处理'],/2026年11月[^\n]*H2[^\n]*G2[^\n]*差额20\.00/);
 const inserted=table([[999,123,20,'甲','说明',30,46328,100,80,'11月2日－70\n11月16日－30',100,'OA-1',999,70]],[headers[0],123,...headers.slice(1,5),46328,...headers.slice(5)]);
 const insertedData=values(build(inserted));assert.deepEqual(Validation.selectColumns(inserted),{dateColumns:[5,13],monthColumns:[2,8]});assert.equal(insertedData['日期计划金额合计'],100);assert.equal(insertedData['校验结果'],'不一致');assert.match(insertedData['需要处理'],/2026年11月[^\n]*J2[^\n]*I2[^\n]*差额20\.00/);
});

test('Excel日期、1904系统及完整文本表头有效；缺年、仅月、非法日期和非日期数值格式不误判',()=>{
 const source=table([],['供应商','OA-申请付款金额','谈判明细',serial('2026-11-02'),'2026年11月支付','11月16日','2026年11月','2026-02-30','应付款时间',serial('2026-12-02'),'2026/12/15']);
 source.sheet.D1={t:'n',v:serial('2026-11-02'),z:'yyyy-mm-dd'};source.sheet.J1={t:'n',v:serial('2026-12-02'),z:'0.00'};
 assert.deepEqual(Validation.selectColumns(source),{dateColumns:[3,10],monthColumns:[4]});
 delete source.sheet.D1.z;assert.ok(!Validation.selectColumns(source).dateColumns.includes(3));source.sheet.D1.w='2026年11月2日';assert.ok(Validation.selectColumns(source).dateColumns.includes(3));
 source.sheet.D1.w='46328';assert.ok(!Validation.selectColumns(source).dateColumns.includes(3));source.sheet.D1={t:'n',f:'DATE(2026,11,2)',v:0,z:'yyyy-mm-dd'};source.date1904=true;assert.ok(Validation.selectColumns(source).dateColumns.includes(3));
 source.sheet.D1={t:'z',f:'DATE(2026,11,2)',v:0,z:'yyyy-mm-dd'};assert.ok(!Validation.selectColumns(source).dateColumns.includes(3));
 source.sheet.D1={t:'d',v:new Date('2026-11-02T00:00:00Z')};assert.ok(Validation.selectColumns(source).dateColumns.includes(3));
});

test('缺年份无唯一来源、非法日期、非法金额均待核对，不推断或吞掉异常',()=>{
 for(const detail of ['2026年2月30日－100','1月2日－100','11月2日－待支付100','11月2日－1.001']){
  const source=table([row()]);source.sheet.D2.v=detail;const data=values(build(source));assert.equal(data['校验结果'],'待核对',detail);assert.equal(data['谈判明细金额合计'],'',detail);assert.equal(data['谈判明细有效条数'],0);assert.match(data['需要处理'],/D2第1笔/);
 }
 const source=table([row()]);delete source.map['分批总金额'];source.sheet.F1={t:'s',v:'备注'};source.sheet.G1={t:'s',v:'申请日期'};
 const data=values(build(source));assert.equal(data['分批总金额'],'');assert.equal(data['日期计划金额合计'],'');assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/缺少“分批总金额”/);
 source.sheet.D2.v='；;';assert.match(values(build(source))['需要处理'],/D2谈判明细没有可解析的日期和金额/);
});

test('跨记录或跨日期合并不重复分配；金额、求和及差额超出安全整数分范围不输出近似值',()=>{
 const merged=table([row(),row()]);merged.sheet['!merges']=[{s:{r:1,c:5},e:{r:2,c:5}}];const result=build(merged);
 assert.ok(result.rows.every((_,i)=>values(result,i)['校验结果']==='待核对'));assert.match(values(result)['需要处理'],/F2:F3/);
 const max='90071992547409.91',overflow=table([row()]);overflow.sheet.F2={t:'s',v:max};overflow.sheet.G2={t:'n',v:0.01};
 let data=values(build(overflow));assert.equal(data['日期计划金额合计'],'');assert.equal(data['校验结果'],'待核对');assert.match(data['需要处理'],/超出安全整数分范围/);
 const diff=table([row()]);diff.sheet.C2={t:'s',v:'-'+max};diff.sheet.E2={t:'s',v:max};data=values(build(diff));assert.equal(data['分批总金额减OA差额'],'');assert.equal(data['校验结果'],'待核对');
});

test('真实XML无缓存公式读取为t:z时不充当0；显式缓存0仍有效',async()=>{
 const source=table([row(0)]),book=X.utils.book_new();X.utils.book_append_sheet(book,source.sheet,'谈判付款');
 for(const cache of ['', '<v/>','<v>0</v>']){
  const zip=await Zip.loadAsync(X.write(book,{type:'buffer',bookType:'xlsx'})),path='xl/worksheets/sheet1.xml',xml=await zip.file(path).async('string');
  zip.file(path,xml.replace(/<c\b[^>]*\br="E2"[^>]*>[\s\S]*?<\/c>/,'<c r="E2"><f>SUM(F2:G2)</f>'+cache+'</c>'));
  source.sheet=X.read(await zip.generateAsync({type:'nodebuffer'}),{type:'buffer',sheetStubs:true,cellFormula:true,cellNF:true}).Sheets['谈判付款'];const data=values(build(source));
  assert.equal(data['校验结果'],cache==='<v>0</v>'?'一致':'待核对');assert.equal(data['分批总金额'],cache==='<v>0</v>'?0:'');
 }
});

test('浏览器UMD和CommonJS提供同样的独立接口；空数据不生成人造总计',()=>{
 const context=vm.createContext({});vm.runInContext(fs.readFileSync(require.resolve('../funds-total-validation.js'),'utf8'),context);
 assert.equal(typeof context.FundsTotalValidation.build,'function');assert.equal(typeof context.FundsTotalValidation.selectColumns,'function');assert.deepEqual(Array.from(context.FundsTotalValidation.HEADERS),Validation.HEADERS);
 const result=build(table([]));assert.deepEqual(result.rows,[]);assert.deepEqual(result.stats,{total:0,excludedNotPay:0,consistent:0,inconsistent:0,unverified:0});
});
