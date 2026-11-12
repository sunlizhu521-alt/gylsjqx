const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Core = require('../negotiated-payment-summary-core.js');

test('两个付款区块分别移除全空日期列，零值、负数和正负抵消均保留',()=>{
  const result={dates:[0,1,2,3,4],rows:[
    {negotiatedCents:[100,null,0,-50,null],normalCents:[null,0,null,null,null]},
    {negotiatedCents:[-100,null,null,null,null],normalCents:[null,null,null,60,null]}
  ]},before=JSON.stringify(result);
  assert.deepEqual(Core.dateColumns(result),{negotiated:[0,2,3],normal:[1,3]});
  assert.equal(JSON.stringify(result),before);
  result.rows=[];
  assert.deepEqual(Core.dateColumns(result),{negotiated:[],normal:[]});
});

const headers = ['供应商','是否谈判分批','谈判明细','2026年11月支付','2027年1月支付','付款主体','OA-申请付款金额'];
const detail = (first,second) => '11月25日－'+first+(second === undefined ? '' : '\n1月10日－'+second);
function source(rows=[], labels=headers) {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([labels,...rows]),'谈判付款');
  return {book,filename:'付款节奏.xlsx'};
}

test('同供应商跨主体按具体日合并，来源顺序及重复申请均保留，两区共用完整日期', () => {
  const input = source([
    ['乙公司','是',detail(100,20),999,999,'迈德斯特'],['甲公司','否','11月15日－30',500,null,'国源'],
    ['乙公司','是',detail(100,20),999,999,'国源'],['乙公司','否',detail(50),111,null,'迈德斯特']
  ]);
  input.book.Sheets.谈判付款.H1={t:'s',v:'是否结清'}; input.book.Sheets.谈判付款.H2={t:'s',v:'是'};
  input.book.Sheets.谈判付款.C2.s={fill:{fgColor:{rgb:'FF0000'}}};
  input.book.Sheets.谈判付款.G2={t:'e',f:'1/0',v:7};
  X.utils.book_append_sheet(input.book,X.utils.aoa_to_sheet([headers,['无关','否',detail(999999)]]),'正常付款');
  const before=JSON.stringify(input),r=Core.build(input);
  assert.equal(r.canExport,true); assert.equal(r.recordCount,4);
  assert.deepEqual(r.counts,{negotiated:2,normal:2,unresolved:0});
  assert.deepEqual(r.rows.map(x=>x.supplier),['乙公司','甲公司']);
  assert.deepEqual(r.dates,[{key:'2026-11-15',label:'2026年11月15日'},{key:'2026-11-25',label:'2026年11月25日'},{key:'2027-01-10',label:'2027年1月10日'}]);
  assert.deepEqual(r.rows[0],{supplier:'乙公司',negotiatedCents:[null,20000,4000],normalCents:[null,5000,null],negotiatedTotalCents:24000,normalTotalCents:5000});
  assert.deepEqual(r.totals,{negotiatedCents:24000,normalCents:8000,unresolvedCents:0});
  assert.equal(JSON.stringify(input),before);
});

test('字段按表头定位，仅谈判明细金额参与，月列和排期金额及其坏公式忽略', () => {
  const input=source([[999,' 是 ',' 甲公司 ','1、2027年1月10日－5\n2、2026年12月25日－3\n3、2026年9月30日－1',888,777]],['2027年1月支付','是否谈判分批','供应商名称','谈判明细','2026年12月支付','2026年11月2日']);
  input.book.Sheets.谈判付款.A2={t:'e',f:'BAD()',v:7};
  const r=Core.build(input);
  assert.equal(r.canExport,true);assert.equal(r.errors.length,0);
  assert.deepEqual(r.dates.map(d=>d.key),['2026-09-30','2026-12-25','2027-01-10']);
  assert.deepEqual(r.rows[0].negotiatedCents,[100,300,500]);
  assert.equal(r.rows[0].supplier,'甲公司');
});

test('同日多笔、同行重复及两类均累计，不把月金额放入某个具体日',()=>{
  const r=Core.build(source([
    ['甲','是','1、11月25日－100\n2、11月25日－100\n3、11月26日－30',9000],
    ['甲','否','11月25日－50',8000]
  ]));
  assert.equal(r.canExport,true);assert.deepEqual(r.dates.map(d=>d.key),['2026-11-25','2026-11-26']);
  assert.deepEqual(r.rows[0].negotiatedCents,[20000,3000]);assert.deepEqual(r.rows[0].normalCents,[5000,null]);
});

test('无年份优先取同月唯一年月表头，忽略排期日期的不同年份；不借相邻月份推年',()=>{
  const input=source([['甲','是','4月25日－100',999,888]],['供应商','是否谈判分批','谈判明细','2027年4月支付','2026年4月13日']);
  assert.deepEqual(Core.build(input).dates,[{key:'2027-04-25',label:'2027年4月25日'}]);
  const conflicting=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细','2026年4月支付','2027年4月支付','2027年4月13日']);
  assert.equal(Core.build(conflicting).canExport,false);assert.match(Core.build(conflicting).errors.join('\n'),/C2.*年份不唯一.*D1.*E1/);
  const missing=source([['甲','是','12月25日－100']],['供应商','是否谈判分批','谈判明细','2026年11月支付','2027年1月支付']);
  assert.equal(Core.build(missing).canExport,false);assert.match(Core.build(missing).errors.join('\n'),/C2.*缺少年份/);
  const duplicated=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细','2027年4月支付','2027年4月支付']);
  assert.equal(Core.build(duplicated).canExport,true,'仅年份依据重复且一致不重复金额');
});

test('缺月度依据才按同月完整日期表头唯一年份读取，支持日期序列和1904工作簿',()=>{
  for(const h of ['2027年4月13日','2027-04-13','2027/4/13','2027.4.13',46490]) {
    const input=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细',h]);
    if(typeof h==='number')input.book.Sheets.谈判付款.D1.z='yyyy-mm-dd';
    const r=Core.build(input);assert.equal(r.canExport,true,String(h));assert.equal(r.dates[0].key,'2027-04-25');
  }
  const input=source([['甲','否','4月25日－100']],['供应商','是否谈判分批','谈判明细',46490-1462]);
  input.book.Sheets.谈判付款.D1.z='yyyy"年"m"月"d"日"';
  input.book.Workbook={WBProps:{date1904:true}};
  assert.equal(Core.build(input).dates[0].key,'2027-04-25');
  const conflict=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细','2026年4月13日','2027年4月13日']);
  assert.match(Core.build(conflict).errors.join('\n'),/年份不唯一/);
});

test('普通数字表头不能提供年份，日期数值须有日期格式或明确完整日期显示',()=>{
  for(const cell of [{t:'n',v:46490},{t:'n',v:46490,z:'General'},{t:'n',v:46490,z:'0" days"'},{t:'n',v:46490,z:'h:mm'},{t:'n',v:46490,w:'4/13'}]) {
    const input=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细','参考']);input.book.Sheets.谈判付款.D1=cell;
    const r=Core.build(input);assert.equal(r.canExport,false,JSON.stringify(cell));assert.match(r.errors.join('\n'),/缺少年份/);
  }
  const displayed=source([['甲','是','4月25日－100']],['供应商','是否谈判分批','谈判明细','参考']);
  displayed.book.Sheets.谈判付款.D1={t:'n',v:46490,w:'2027年4月13日'};
  assert.equal(Core.build(displayed).dates[0].key,'2027-04-25');
  const zero=source([['甲','是','1月15日－100']],['供应商','是否谈判分批','谈判明细','参考']);
  zero.book.Sheets.谈判付款.D1={t:'n',v:0,z:'yyyy-mm-dd'};zero.book.Workbook={WBProps:{date1904:true}};
  assert.equal(Core.build(zero).dates[0].key,'1904-01-15');
});

test('完整年月日直接使用，闰日和日期边界严格校验，不需要年参考表头',()=>{
  const h=['供应商','是否谈判分批','谈判明细'];
  const input=source([['甲','是','2028年2月29日－1;2026-12-31－2；2027/1/1－3\n1900.1.1－4\n9999年12月31日－5']],h);
  const r=Core.build(input);assert.equal(r.canExport,true);assert.deepEqual(r.dates.map(d=>d.key),['1900-01-01','2026-12-31','2027-01-01','2028-02-29','9999-12-31']);
  for(const date of ['2026年2月29日','2026年4月31日','2026年13月1日','2026年0月1日','2026年11月0日','1899年12月31日','0000年11月25日','11月','2026年11月']) {
    const bad=Core.build(source([['甲','是',date+'－100']]));assert.equal(bad.canExport,false,date);assert.match(bad.errors.join('\n'),/C2/);
  }
  const explicit=source([['甲','是','2028年4月25日－100']],['供应商','是否谈判分批','谈判明细','2026年4月支付','2027年4月支付']);
  assert.equal(Core.build(explicit).dates[0].key,'2028-04-25','显式年份不受参考表头冲突影响');
});

test('支持序号及纯文字括号注释；未完整解析的尾注不静默忽略或部分累计',()=>{
  const r=Core.build(source([['甲','是','1、11月25日－300,000.00（沟通延后）\n2、1月10日－337,812.00'] ]));
  assert.equal(r.canExport,true);assert.equal(r.rows[0].negotiatedTotalCents,63781200);assert.match(r.warnings[0],/C2.*沟通延后/);
  const inline=Core.build(source([['甲','是','1、11月25日－1 2、1月10日－2']]));assert.equal(inline.totals.negotiatedCents,300);
  for(const content of ['11月25日－575.00 200+375','11月25日－100（另付20）','11月25日－100；不明说明','11月25日－100\n1月10日－待定','11月25日－','11月25日－100abc','11月25日－100及11月26日－20','11月25日－100\n/']) {
    const bad=Core.build(source([['甲','是',content]]));assert.equal(bad.canExport,false,content);assert.equal(bad.totals.negotiatedCents,0,'整个异常源行不部分累计');assert.match(bad.errors.join('\n'),/C2/);
  }
});

test('供应商只去首尾空格，不模糊归并名称，不修改原数据', () => {
  const r=Core.build(source([['甲 公司','是',detail(1)],['甲公司','是',detail(2)],[' 甲公司 ','否',detail(3)]]));
  assert.equal(r.rows.length,2);assert.equal(r.rows[0].supplier,'甲 公司');assert.equal(r.rows[1].normalTotalCents,300);
});

test('空明细保留供应商并阻止，明确零保留，未发生付款的日期保持空白',()=>{
  const r=Core.build(source([['甲','是','11月25日－0'],['乙','否','1月10日－0'],['丙','否','']]));
  assert.equal(r.canExport,false);assert.equal(r.recordCount,3);assert.equal(r.rows.length,3);
  assert.deepEqual(r.rows[0].negotiatedCents,[0,null]);assert.deepEqual(r.rows[1].normalCents,[null,0]);
  assert.deepEqual(r.rows[2].normalCents,[null,null]);assert.match(r.errors.join('\n'),/C4.*明细为空.*具体付款日期/);
  for(const value of [null,undefined,'  ',false,0]) {const bad=Core.build(source([['甲','是',value]]));assert.equal(bad.canExport,false);assert.equal(bad.rows[0].supplier,'甲');}
  const missing=Core.build(source([['','是',detail(10)]]));assert.equal(missing.canExport,false);assert.match(missing.errors.join('\n'),/A2.*供应商/);
});

test('跨记录的纵向合并明细不复制也不猜分配，其他合并不影响',()=>{
  const input=source([['甲','是',detail(100,20)],['甲','否',null]]);
  input.book.Sheets.谈判付款['!merges']=[{s:{r:1,c:2},e:{r:2,c:2}}];
  const before=JSON.stringify(input),r=Core.build(input);
  assert.equal(r.canExport,false);assert.equal(r.recordCount,2);assert.equal(r.totals.negotiatedCents,0);assert.equal(r.totals.normalCents,0);
  assert(r.errors.some(e=>e.includes('C2')&&e.includes('跨多条记录')));assert(r.errors.some(e=>e.includes('C3')&&e.includes('跨多条记录')));assert.equal(JSON.stringify(input),before);
  const other=source([['甲','是',detail(100)],['甲','否',detail(20)]]);other.book.Sheets.谈判付款['!merges']=[{s:{r:1,c:5},e:{r:2,c:5}}];
  assert.equal(Core.build(other).canExport,true);
  const padded=source([['甲','是',detail(100)]]);padded.book.Sheets.谈判付款['!merges']=[{s:{r:1,c:2},e:{r:2,c:2}}];
  assert.equal(Core.build(padded).totals.negotiatedCents,10000,'合并区只有一条记录不重复累计');
});

test('分类缺失或异常需明确选择，按地址重算且已知分类不覆盖',()=>{
  const input=source([['甲','',detail(100,20)],['甲','待核对',detail(30)],['乙','是',detail(1,2)]]),r=Core.build(input);
  assert.equal(r.canExport,false);assert.equal(r.errors.length,0);assert.deepEqual(r.counts,{negotiated:1,normal:0,unresolved:2});
  assert.deepEqual(r.totals,{negotiatedCents:300,normalCents:0,unresolvedCents:15000});
  assert.deepEqual(r.unresolved.map(x=>[x.id,x.row,x.amountCents]),[['B2',2,12000],['B3',3,3000]]);
  assert.match(r.unresolved[0].source,/付款节奏.xlsx \/ 谈判付款 \/ 第2行/);
  const resolved=Core.build(input,{B2:'是',B3:'否',B4:'否'});assert.equal(resolved.canExport,true);
  assert.deepEqual(resolved.totals,{negotiatedCents:12300,normalCents:3000,unresolvedCents:0});
  assert.equal(Core.build(input,{B2:'yes',B3:'否'}).unresolved.length,1);assert.equal(Core.build(input).unresolved.length,2);
});

test('金额按分四舍五入，币种千分位、科学计数和负数合法',()=>{
  const r=Core.build(source([['甲','是',detail('￥1,234.56元','-1.005')],['甲','是',detail(0.1,'1e-2')],['甲','是',detail(0.2,1.005)]]));
  assert.equal(r.canExport,true);assert.deepEqual(r.rows[0].negotiatedCents,[123486,1]);assert.equal(r.warnings.length,2);
  for(const value of ['false','Infinity','NaN','1,00','待付款','/','90071992547409.92']) {
    const bad=Core.build(source([['甲','是',detail(value)]]));assert.equal(bad.canExport,false,value);assert.match(bad.errors.join('\n'),/C2.*金额/);
  }
});

test('单笔、同日、供应商及全表合计保护安全整数分边界',()=>{
  const max='90071992547409.91';assert.equal(Core.build(source([['甲','是',detail(max)]])).totals.negotiatedCents,Number.MAX_SAFE_INTEGER);
  for(const rows of [
    [['甲','是',detail(max)],['甲','是',detail(0.01)]],
    [['甲','是',detail(max,0.01)]],
    [['甲','是',detail(max)],['乙','是',detail(0.01)]],
    [['甲','',detail(max,0.01)]]
  ]) {const r=Core.build(source(rows));assert.equal(r.canExport,false);assert.match(r.errors.join('\n'),/安全整数分/);}
});

test('必要字段错误或公式缺缓存阻止，缓存明细文本有效，分类选择不绕过公式错误',()=>{
  for(const address of ['A2','B2','C2']) for(const value of [{t:'n',f:'1+1'},{t:'z',f:'1+1',v:0},{t:'str',f:'A1',v:''},{t:'e',v:23,f:'bad'}]) {
    const input=source([['甲','是',detail(10)]]);input.book.Sheets.谈判付款[address]=value;
    const r=Core.build(input,{B2:'是'});assert.equal(r.canExport,false);assert(r.errors.some(e=>e.includes(address)&&e.includes('公式')));
  }
  const input=source([['甲','是',detail(10)]]);input.book.Sheets.谈判付款.C2={t:'str',f:'D2',v:detail(0)};
  assert.equal(Core.build(input).canExport,true);assert.equal(Core.build(input).rows[0].negotiatedCents[0],0);
});

test('真实OOXML必要公式缺缓存阻止导出，文本零金额缓存有效',async()=>{
  const original=X.write(source([['甲','是',detail(10)]]).book,{type:'buffer',bookType:'xlsx'});
  for(const cache of ['', '<v/>', '<v>11月25日－0</v>']) {
    const zip=await Zip.loadAsync(original),file='xl/worksheets/sheet1.xml',xml=await zip.file(file).async('string');
    zip.file(file,xml.replace(/<c\b[^>]*\br="C2"[^>]*>[\s\S]*?<\/c>/,`<c r="C2" t="str"><f>D2</f>${cache}</c>`));
    const input={book:X.read(await zip.generateAsync({type:'nodebuffer'}),{type:'buffer',sheetStubs:true,cellFormula:true}),filename:'公式.xlsx'};
    const r=Core.build(input);assert.equal(r.canExport,cache.includes('11月25日'));if(!r.canExport)assert.match(r.errors.join('\n'),/C2/);
  }
});

test('必要表头重复、缺失不猜测；无关年月列格式不影响显式日期',()=>{
  assert.equal(Core.build({book:{Sheets:{正常付款:{}}}}).canExport,false);
  for(const h of [
    ['供应商','谈判明细'],['供应商','是否谈判分批'],
    ['供应商','供应商名称','是否谈判分批','谈判明细'],
    ['供应商','是否谈判分批','是否谈判分批','谈判明细'],
    ['供应商','是否谈判分批','谈判明细','谈判明细']
  ]) {const r=Core.build(source([],h));assert.equal(r.canExport,false);assert(r.errors.length,h.join(','));}
  const explicit=source([['甲','是','2026年11月25日－1']],['供应商','是否谈判分批','谈判明细','11月支付','2026年13月支付']);
  assert.equal(Core.build(explicit).canExport,true);
});

test('大格式范围不扫描，无关行或其他字段不产生付款记录',()=>{
  const input=source([['甲','是',detail(0)]]),sheet=input.book.Sheets.谈判付款;
  sheet['!ref']='A1:XFD1048576';sheet.XFD1048576={t:'z',s:{fill:{fgColor:{rgb:'FFFFFF'}}}};sheet.Z1000={t:'s',v:'备注说明'};
  const r=Core.build(input);assert.equal(r.canExport,true);assert.equal(r.recordCount,1);assert.equal(Core.build(source()).canExport,false);
});

test('浏览器与CommonJS导出相同构建接口',()=>{
  const context={};vm.runInNewContext(fs.readFileSync(require.resolve('../negotiated-payment-summary-core.js'),'utf8'),context);
  const r=context.NegotiatedPaymentSummaryCore.build(source([['甲','否',detail(12.34)]]));assert.equal(r.canExport,true);assert.equal(r.totals.normalCents,1234);assert.equal(r.dates[0].key,'2026-11-25');
});

function paymentSource(rows, dates=['2026年11月2日','2026年11月16日']) {
  return source(rows,['供应商','是否谈判分批','谈判明细',...dates,'2026年11月支付','2026年12月支付']);
}
function comment(input,address,value) {
  input.book.Sheets.谈判付款[address].c=[{a:'人工核实',t:value}];
}
const paidAt=(r,type,supplier,date)=>r.paid[type][r.rows.findIndex(row=>row.supplier===supplier)][r.dates.findIndex(d=>d.key===date)];

test('金额单元格明确支付批注生成矩阵，精确日期优先，同月唯一排期可定位，金额和源内容不变',()=>{
  const input=paymentSource([
    ['甲','是','11月2日－700000\n11月25日－30',700000,30],
    ['乙','否','11月25日－100',null,100]
  ]),baseline=Core.build(input);
  comment(input,'D2','11月3日-OA已评论\n11月5日核实-已支付\n申请日期（付款节奏跟进）：2026年11月2日\n11月5日核实-已支付');
  comment(input,'E2','11/5 出纳核实-本批已支付30.00元');comment(input,'E3','已付款');
  const before=JSON.stringify(input),r=Core.build(input);
  assert.equal(r.canExport,true);assert.deepEqual(r.paid,{negotiated:[[true,true],[false,false]],normal:[[false,false],[false,true]]});
  assert.deepEqual(r.rows,baseline.rows);assert.deepEqual(r.totals,baseline.totals);assert.equal(r.warnings.length,0);assert.equal(JSON.stringify(input),before);
});

test('只有明确完成的批注肯定支付，申请日期、原填色、已付金额、归档及结清不作证据',()=>{
  for(const note of ['已支付','已付款','2026年11月5日-确认：已付清','11/5 出纳核实-本批已支付100.00元','已支付：人民币100元']) {
    const input=paymentSource([['甲','是','11月2日－100',100]]);comment(input,'D2',note);
    const r=Core.build(input);assert.equal(paidAt(r,'negotiated','甲','2026-11-02'),true,note);assert.equal(r.warnings.length,0);
  }
  const input=paymentSource([['甲','是','11月2日－100',100]]),sheet=input.book.Sheets.谈判付款;
  sheet.D2.s={fill:{fgColor:{rgb:'FFFF0000'}}};sheet.H1={t:'s',v:'已付金额'};sheet.H2={t:'n',v:100};sheet.I1={t:'s',v:'是否结清'};sheet.I2={t:'s',v:'是'};sheet.J1={t:'s',v:'流程状态'};sheet.J2={t:'s',v:'归档'};
  for(const note of ['申请日期（付款节奏跟进）：2026年11月2日','11月3日-OA已评论','']) {
    comment(input,'D2',note);const r=Core.build(input);assert.equal(paidAt(r,'negotiated','甲','2026-11-02'),false,note);assert.equal(r.warnings.length,0,'普通无支付信息批注不产生噪音');
  }
  sheet.F2={t:'n',v:100,c:[{a:'人工',t:'已支付'}]};
  const monthlyOnly=Core.build(input);assert.equal(paidAt(monthlyOnly,'negotiated','甲','2026-11-02'),false);assert.equal(monthlyOnly.warnings.length,0,'月支付列批注不作为付款证据');
});

test('否定、计划、疑问、部分支付、矛盾和批注明确金额不一致均不确认，非阻塞提示来源',()=>{
  for(const note of ['未支付','未付款','未到账','计划已支付','预计已支付100元','是否已支付','已支付？','如果已支付请反馈','已申请付款','已部分支付100元','部分已支付100元','已支付99.99元','已支付0元','已支付-100元','已支付100.001元','已支付\n银行退汇','已支付\n11月6日未支付','已支付\n是否已支付','已支付100元\n已支付99元']) {
    const input=paymentSource([['甲','是','11月2日－100',100]]);comment(input,'D2',note);
    const r=Core.build(input);assert.equal(r.canExport,true,note);assert.equal(paidAt(r,'negotiated','甲','2026-11-02'),false,note);assert.equal(r.warnings.length,1,note);assert.match(r.warnings[0],/D2.*支付批注未用于标红.*请核对/);
  }
});

test('同额两批先按精确日期保留归属，第一笔支付不会带红后面的同额付款',()=>{
  const input=paymentSource([['甲','是','11月2日－1090000\n11月25日－1090000',1090000,1090000]]);
  comment(input,'D2','11月5日核实-已支付');
  let r=Core.build(input);assert.deepEqual(r.paid.negotiated,[[true,false]]);assert.equal(r.warnings.length,0);
  delete input.book.Sheets.谈判付款.D2.c;comment(input,'E2','11月20日核实-已支付');
  r=Core.build(input);assert.deepEqual(r.paid.negotiated,[[false,true]]);assert.equal(r.warnings.length,0,'无批注精确日期批次必须保留自己的归属');
});

test('同月同额无法唯一配对时只有全部源格明确已付且数量相等才确认，不重复使用一格',()=>{
  const input=paymentSource([['甲','是','11月25日－100\n11月26日－100',100,100]]);
  comment(input,'D2','已支付');let r=Core.build(input);
  assert.deepEqual(r.paid.negotiated,[[false,false]]);assert.equal(r.warnings.length,1);assert.match(r.warnings[0],/同月同金额的多批归属不唯一/);
  comment(input,'E2','已支付');r=Core.build(input);assert.deepEqual(r.paid.negotiated,[[true,true]]);assert.equal(r.warnings.length,0);
  delete input.book.Sheets.谈判付款.E2;r=Core.build(input);assert.deepEqual(r.paid.negotiated,[[false,false]]);assert.equal(r.warnings.length,1,'一个付款格不能重复确认两批');
  const duplicate=paymentSource([['甲','是','11月25日－100\n11月25日－100',100,100]],['2026年11月25日','2026年11月25日']);
  comment(duplicate,'D2','已支付');comment(duplicate,'E2','已支付');assert.deepEqual(Core.build(duplicate).paid.negotiated,[[true]]);
});

test('同供应商同日所有组成款项均明确支付才标红，跨行或同日不同金额部分已付不标红',()=>{
  const input=paymentSource([['甲','是','11月2日－100',100],['甲','是','11月2日－200',200]]);comment(input,'D2','已支付');
  let r=Core.build(input);assert.equal(r.rows[0].negotiatedCents[0],30000);assert.deepEqual(r.paid.negotiated,[[false]]);
  comment(input,'D3','已支付200元');assert.deepEqual(Core.build(input).paid.negotiated,[[true]]);
  const sameRow=paymentSource([['甲','是','11月2日－100\n11月2日－200',100,200]],['2026年11月2日','2026年11月2日']);
  comment(sameRow,'D2','已支付');assert.deepEqual(Core.build(sameRow).paid.negotiated,[[false]]);
  comment(sameRow,'E2','已支付');assert.deepEqual(Core.build(sameRow).paid.negotiated,[[true]]);
});

test('不同供应商、分类、月份和年份隔离，未配对支付批注只提醒不影响金额或下载',()=>{
  const input=paymentSource([['甲','是','11月2日－100',100],['甲','否','11月2日－100',100],['乙','是','11月2日－100',100]]);comment(input,'D2','已支付');
  let r=Core.build(input);assert.deepEqual(r.paid.negotiated,[[true],[false]]);assert.deepEqual(r.paid.normal,[[false],[false]]);
  for(const date of ['2026年12月2日','2027年11月2日']) {
    const other=paymentSource([['甲','是','2026年11月25日－100',100]],[date]);comment(other,'D2','已支付');r=Core.build(other);
    assert.equal(r.canExport,true);assert.deepEqual(r.paid.negotiated,[[false]]);assert.match(r.warnings[0],/D2.*不跨月配对/);
  }
  const mismatch=paymentSource([['甲','是','11月2日－200',100]]);comment(mismatch,'D2','已支付100元');r=Core.build(mismatch);
  assert.equal(r.rows[0].negotiatedTotalCents,20000);assert.deepEqual(r.paid.negotiated,[[false]]);assert.match(r.warnings[0],/唯一谈判明细/);
});

test('跨记录合并金额及不可用公式不作为支付证据，不新增阻塞错误',()=>{
  const input=paymentSource([['甲','是','11月2日－100',100],['乙','否','11月2日－100',null]]);comment(input,'D2','已支付');
  input.book.Sheets.谈判付款['!merges']=[{s:{r:1,c:3},e:{r:2,c:3}}];
  let r=Core.build(input);assert.equal(r.canExport,true);assert.deepEqual(r.paid,{negotiated:[[false],[false]],normal:[[false],[false]]});assert.match(r.warnings[0],/D2.*跨多条记录合并/);
  const bad=paymentSource([['甲','是','11月2日－100',100]]);bad.book.Sheets.谈判付款.D2={f:'BAD()',t:'e',v:7,c:[{t:'已支付'}]};
  r=Core.build(bad);assert.equal(r.canExport,true);assert.equal(r.errors.length,0);assert.deepEqual(r.paid.negotiated,[[false]]);assert.match(r.warnings[0],/D2.*公式结果不可用/);
});

test('真实案例的申请日期批注不标红，明确支付批注只确认对应的700000和1090000批次',()=>{
  const input=paymentSource([
    ['科福','是','9月30日－615833\n11月2日－700000\n11月25日－670407',615833,700000,670407],
    ['迈锐','是','9月30日－211820\n11月2日－1090000\n11月25日－1090000',211820,1090000,1090000]
  ],['2026年9月28日','2026年11月2日','2026年11月16日']);
  comment(input,'D2','申请日期（付款节奏跟进）：2026年9月28日');comment(input,'D3','申请日期（付款节奏跟进）：2026年9月28日');
  comment(input,'E2','11月5日核实-已支付');comment(input,'E3','申请日期（付款节奏跟进）：2026年11月2日\n11月5日核实-已支付');
  const r=Core.build(input);assert.equal(r.canExport,true);assert.deepEqual(r.paid.negotiated,[[false,true,false],[false,true,false]]);assert.equal(r.warnings.length,0);
  delete input.book.Sheets.谈判付款.E2.c;assert.equal(paidAt(Core.build(input),'negotiated','科福','2026-11-02'),false,'重新上传重新读取本次批注，不沿用旧标记');
});

const snapshot=process.env.NEGOTIATED_DATE_QA_SNAPSHOT||'/tmp/negotiated-date-inspect.json';
test('真实样表快照119条47家：空明细和额外金额说明阻止下载，四月取2027且不读月金额',{skip:!fs.existsSync(snapshot)},()=>{
  const input=JSON.parse(fs.readFileSync(snapshot,'utf8')),r=Core.build(input);
  assert.equal(r.recordCount,119);assert.equal(r.rows.length,47);assert.deepEqual(r.counts,{negotiated:56,normal:62,unresolved:1});
  assert.deepEqual(r.totals,{negotiatedCents:3428869573,normalCents:1063642003,unresolvedCents:18000000});assert.equal(r.dates.length,16);
  assert.equal(r.canExport,false);assert.equal(r.unresolved[0].id,'G40');assert.equal(r.unresolved[0].amountCents,18000000);
  assert.equal(r.errors.length,4,r.errors.join('\n'));for(const address of ['L25','L44','L73','L78'])assert(r.errors.some(e=>e.includes('/ '+address+' ' )||e.includes('/ '+address+'：')),address);
  assert(r.dates.some(d=>d.key==='2027-04-25'));assert(!r.dates.some(d=>d.key==='2026-04-25'));assert(r.dates.some(d=>d.key==='2026-11-02'));
  assert.equal(r.warnings.filter(w=>w.includes('括号说明')&&w.includes('沟通延后')).length,2); // This old snapshot verifies amounts only, never current payment comments.
  const original=JSON.stringify(input);
  for(const [address,cell] of Object.entries(input.book.Sheets.谈判付款))if(/^(AI|AJ|AK|AL|AM|AN|AO|AP)\d+$/.test(address)&&!address.endsWith('1'))cell.v=99999999;
  const changed=Core.build(input);assert.deepEqual(changed.totals,r.totals,'月度金额变动不参与汇总');
  assert.notEqual(JSON.stringify(input),original);
});
