const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Core = require('../funds-budget-summary-core.js');

const H = ['预估付款月份','供应商名称','分类','申请付款金额','是否已支付','是否剔除'];
const serial = date => (Date.parse(date+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const row = (month='2026年11月',category='谈判付款表-未暂停',amount=100,paid='未支付') => [month,'',category,amount,paid,'否'];
function source(rows,headers=H) {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'付款汇总明细');
  return {book,filename:'合成付款明细.xlsx'};
}

test('固定七列表头；表头定位、空供应商、重复行及剔除列都不丢金额，原工作簿不变', () => {
  const input = source([
    row('2026年11月','谈判付款表-未暂停',100,'已支付'),
    row('2026年11月','谈判付款表-未暂停',100,'已支付'),
    [...row('2026-11-20','谈判付款表-暂停付款',30), '无关'],
    row('2026/11/2','正常付款',20)
  ]);
  input.book.Sheets.付款汇总明细.F2.v='是';
  X.utils.book_append_sheet(input.book,X.utils.aoa_to_sheet([['申请总金额'],[999999]]),'付款预算汇总');
  const before = JSON.stringify(input), result = Core.build(input);
  assert.deepEqual(Core.HEADERS,['付款月份','分类','状态','申请总金额','已付款金额','剩余付款金额','备注']);
  assert.equal(result.recordCount,4); assert.equal(result.canExport,true);
  assert.deepEqual(result.groups.map(g => [g.category,g.state,g.totalCents,g.paidCents]),[
    ['谈判付款表','未暂停',20000,20000],['谈判付款表','暂停付款',3000,0],['正常付款','正常付款',2000,0]
  ]);
  assert.deepEqual(result.totals,{totalCents:25000,paidCents:20000,remainingCents:5000});
  assert.equal(JSON.stringify(input),before);
  const shuffled = source([['已支付',5,'正常付款','2026年11月']],['是否已支付','申请付款金额','分类','预估付款月份']);
  assert.equal(Core.build(shuffled).totals.paidCents,500);
});

test('只有trim后精确已支付计已付，其余状态与空白全部保留总额', () => {
  const statuses = ['已支付',' 已支付 ','已付款','部分支付','待核对','未支付','',null,'已 支付'];
  const result = Core.build(source(statuses.map(state => row('2026年11月','正常付款',10,state))));
  assert.equal(result.canExport,true); assert.equal(result.recordCount,statuses.length);
  assert.deepEqual(result.totals,{totalCents:9000,paidCents:2000,remainingCents:7000});
});

test('实际年月合并并升序，跨年分开，空白及缺年份均计入最后时间待定组', () => {
  const result = Core.build(source([
    row('2027年1月','正常付款',1),row(serial('2026-12-28')+0.5,'正常付款',2),row('2026-12-01T03:04:05Z','正常付款',3),
    row('时间待定','正常付款',4),row('','正常付款',5),row('11月','正常付款',6),row('2026-02-30','正常付款',7),row(11,'正常付款',8),
    row('2026年11月','正常付款',9),row('2026年11月','谈判付款表-暂停付款',10),row('2026年11月','谈判付款表-未暂停',11)
  ]));
  assert.equal(result.canExport,true); assert.equal(result.recordCount,11); assert.equal(result.warnings.length,5);
  assert.deepEqual(result.groups.map(g => [g.monthKey,g.state,g.totalCents]),[
    ['2026-11','未暂停',1100],['2026-11','暂停付款',1000],['2026-11','正常付款',900],
    ['2026-12','正常付款',500],['2027-01','正常付款',100],['','正常付款',3000]
  ]);
  assert.equal(result.totals.totalCents,6600);
});

test('1904日期系统及Date对象使用实际月份；无效Excel闰日保留为时间待定', () => {
  const input = source([row(serial('2026-11-01')-1462,'正常付款',1)]);
  input.book.Workbook={WBProps:{date1904:true}};
  const result=Core.build(input);
  assert.equal(result.date1904,true); assert.equal(result.groups[0].monthKey,'2026-11');
  input.book.Sheets.付款汇总明细.A2={t:'d',v:new Date('2026-12-02T00:00:00Z')};
  assert.equal(Core.build(input).groups[0].monthKey,'2026-12');
  const leap=Core.build(source([row(60,'正常付款',1)]));
  assert.equal(leap.groups[0].monthKey,''); assert.equal(leap.warnings.length,1);
});

test('按分精确四舍五入，零、负数和合法金额文字不被剔除', () => {
  const values = [0,-2.345,'1.005','￥1,234.56元','1e-2','90071992547409.91'];
  const results = values.map(value => Core.build(source([row('2026年11月','正常付款',value,'已支付')])));
  assert.deepEqual(results.map(r => r.totals.totalCents),[0,-235,101,123456,1,Number.MAX_SAFE_INTEGER]);
  assert.ok(results.every(r => r.canExport && r.recordCount===1 && r.totals.remainingCents===0));
  assert.equal(results[1].warnings.length,1); assert.equal(results[2].warnings.length,1);
  assert.equal(Core.build(source([row('2026年11月','正常付款',-10),row('2026年11月','正常付款',0)])).totals.totalCents,-1000);
});

test('空金额、异常金额、异常分类阻止导出并给出来源单元格', () => {
  for (const value of [null,'',false,'待核对','1,00',Infinity,'90071992547409.92']) {
    const result=Core.build(source([row('2026年11月','正常付款',value)]));
    assert.equal(result.canExport,false,String(value)); assert.equal(result.recordCount,1);
    assert.match(result.errors.join('\n'),/合成付款明细.xlsx \/ 付款汇总明细 \/ D2.*金额/);
  }
  for (const category of ['',null,'谈判付款','__proto__']) {
    const result=Core.build(source([row('2026年11月',category,1)]));
    assert.equal(result.canExport,false); assert.match(result.errors.join('\n'),/C2.*分类/);
  }
});

test('四个必要字段的公式无缓存或错误均阻止导出，缓存零值有效', () => {
  for (const address of ['A2','C2','D2','E2']) for (const cell of [{t:'n',f:'1+1'},{t:'z',f:'1+1',v:0},{t:'str',f:'A1',v:''},{t:'e',v:7,w:'#DIV/0!'}]) {
    const input=source([row()]); input.book.Sheets.付款汇总明细[address]=cell;
    const result=Core.build(input);
    assert.equal(result.canExport,false,address); assert.ok(result.errors.some(error => error.includes(address) && error.includes('公式')));
  }
  const input=source([row()]); input.book.Sheets.付款汇总明细.D2={t:'n',f:'1-1',v:0};
  assert.equal(Core.build(input).canExport,true); assert.equal(Core.build(input).totals.totalCents,0);
});

test('真实XML无缓存公式经SheetJS生成的t:z零值阻止导出，显式缓存零值仍有效', async () => {
  const original = X.write(source([row()]).book,{type:'buffer',bookType:'xlsx'});
  async function readFormula(cache, type='') {
    const zip = await Zip.loadAsync(original), file = 'xl/worksheets/sheet1.xml';
    const xml = await zip.file(file).async('string');
    assert.match(xml,/<c\b[^>]*\br="D2"[^>]*>[\s\S]*?<\/c>/);
    zip.file(file,xml.replace(/<c\b[^>]*\br="D2"[^>]*>[\s\S]*?<\/c>/,'<c r="D2"'+type+'><f>1+1</f>'+cache+'</c>'));
    return {book:X.read(await zip.generateAsync({type:'nodebuffer'}),{type:'buffer',cellFormula:true,sheetStubs:true,cellNF:true,cellDates:false}),filename:'真实XML无缓存.xlsx'};
  }
  for (const type of ['', ' t="n"']) for (const cache of ['', '<v/>', '<v></v>']) {
    const input = await readFormula(cache,type), cell = input.book.Sheets.付款汇总明细.D2;
    assert.equal(cell.f,'1+1'); assert.equal(cell.t,'z'); assert.equal(cell.v,type ? undefined : 0);
    const result = Core.build(input);
    assert.equal(result.canExport,false); assert.equal(result.recordCount,1);
    assert.match(result.errors.join('\n'),/真实XML无缓存.xlsx \/ 付款汇总明细 \/ D2.*公式/);
  }
  for (const cache of ['<v>0</v>','<v>2</v>']) {
    const input = await readFormula(cache), result = Core.build(input);
    assert.equal(input.book.Sheets.付款汇总明细.D2.t,'n');
    assert.equal(result.canExport,true); assert.equal(result.totals.totalCents,cache.includes('>0<') ? 0 : 200);
  }
  const explicitEmpty = await readFormula('<v></v>',' t="str"');
  assert.equal(Core.build(explicitEmpty).canExport,false);
});

test('BigInt累加同时保护分组、总额、已付及剩余的安全边界', () => {
  const max='90071992547409.91';
  const overflow=Core.build(source([row('2026年11月','正常付款',max),row('2026年11月','正常付款','0.01')]));
  assert.equal(overflow.canExport,false); assert.equal(overflow.totals.totalCents,null); assert.match(overflow.errors.join('\n'),/安全整数分/);
  const totalOnly=Core.build(source([row('2026年11月','正常付款',max),row('2026年12月','正常付款','0.01')]));
  assert.equal(totalOnly.canExport,false); assert.equal(totalOnly.groups[0].totalCents,Number.MAX_SAFE_INTEGER);
  const paidOnly=Core.build(source([row('2026年11月','正常付款',max,'已支付'),row('2026年11月','正常付款','0.01','已支付'),row('2026年11月','正常付款','-0.01')]));
  assert.equal(paidOnly.canExport,false);
  const remainingOnly=Core.build(source([row('2026年11月','正常付款',max),row('2026年11月','正常付款','0.01'),row('2026年11月','正常付款','-0.01','已支付')]));
  assert.equal(remainingOnly.canExport,false);
});

test('只读指定明细表；缺表、缺字段及重复字段明确报错', () => {
  assert.equal(Core.build({book:{Sheets:{'其他明细':{}}}}).canExport,false);
  const missing=Core.build(source([['2026年11月','正常付款',1]],['预估付款月份','分类','申请付款金额']));
  assert.match(missing.errors.join('\n'),/是否已支付/);
  const duplicate=Core.build(source([],H.concat('申请付款金额')));
  assert.match(duplicate.errors.join('\n'),/表头重复/);
});

test('存储范围中的空白格式不会产生记录，有内容但关键字段为空的行不静默丢弃', () => {
  const input=source([row('2026年11月','正常付款',0)]), sheet=input.book.Sheets.付款汇总明细;
  sheet['!ref']='A1:XFD1048576'; sheet.A1048576={t:'z',s:{fill:{fgColor:{rgb:'FFFFFFFF'}}}};
  assert.equal(Core.build(input).recordCount,1);
  sheet.B5={t:'s',v:'仅供应商，关键字段待填'};
  const result=Core.build(input);
  assert.equal(result.recordCount,2); assert.equal(result.canExport,false); assert.match(result.errors.join('\n'),/C5|D5/);
  const empty=Core.build(source([]));
  assert.equal(empty.canExport,true); assert.equal(empty.recordCount,0); assert.deepEqual(empty.groups,[]);
  assert.deepEqual(empty.totals,{totalCents:0,paidCents:0,remainingCents:0});
});

test('浏览器UMD入口与CommonJS接口一致', () => {
  const context={}; vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../funds-budget-summary-core.js'),'utf8'),context);
  assert.equal(typeof context.FundsBudgetSummaryCore.build,'function');
  assert.deepEqual(JSON.parse(JSON.stringify(context.FundsBudgetSummaryCore.build(source([row()])))),Core.build(source([row()])));
});
