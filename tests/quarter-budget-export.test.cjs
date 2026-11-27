const test = require('node:test');
const assert = require('node:assert/strict');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Export = require('../quarter-budget-export.js');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const record = (changes = {}) => ({id:'r1',month:46327,monthKey:'2026-11',supplier:'供应商甲',subject:'迈德斯特',oa:'000123',summary:'货款',amount:100,cents:10000,status:'未支付',notes:'',category:'未预估付款',owner:'李奇',excluded:'否',related:false,...changes});
const result = (records = [record()], changes = {}) => ({records,summary:{classified:[{monthKey:'2026-11',category:'未预估付款',totalCents:10000,paidCents:0,remainingCents:10000}],monthly:[{monthKey:'2026-11',totalCents:10000,paidCents:0,remainingCents:10000}],totals:{totalCents:10000,paidCents:0,remainingCents:10000},...changes}});
async function read(data) {
  const bytes = await Export(data);
  assert(bytes instanceof Uint8Array);
  const book = X.read(bytes,{type:'array',cellNF:true,cellStyles:true});
  const zip = await Zip.loadAsync(bytes);
  return {bytes,book,zip};
}
test('1904日期系统合并导出为年月文本，内部月份保持原数值', async () => {
  const Core = require('../quarter-budget-core.js');
  const source = {filename:'1904.xlsx',fingerprint:'1904',book:{SheetNames:['付款汇总明细'],Workbook:{WBProps:{date1904:true}},Sheets:{付款汇总明细:X.utils.aoa_to_sheet([Core.HEADERS,[44865,'合成供应商','主体','FP001','货款',100,'未支付','原备注','人工分类','跟单员','否']])}}};
  const empty = role => ({filename:role+'.xlsx',fingerprint:role,book:{SheetNames:['明细'],Sheets:{明细:X.utils.aoa_to_sheet([Core.HEADERS])}}});
  const built = Core.build({negotiated:empty('negotiated'),ledger:empty('ledger'),budget:source},{budgetStartMonth:'2026-11'});
  assert.equal(built.records[0].month,44865);
  const {book} = await read(built);
  assert.equal(book.Sheets.付款汇总明细.A2.v,'2026年11月');
  assert.equal(book.Sheets.付款汇总明细.A2.t,'s');
  assert.equal(book.Sheets.付款汇总明细.A2.z,'@');
  assert.equal(book.Sheets.付款汇总明细.A2.w,'2026年11月');
  assert.match(book.Sheets.付款汇总明细.H2.v,/原备注/);assert.match(book.Sheets.付款汇总明细.H2.v,/1904.xlsx/);
  assert.equal(source.book.Sheets.付款汇总明细.A2.v,44865);
});
const xml = (zip, index) => zip.file(`xl/worksheets/sheet${index}.xml`).async('string');
function stylesFrom(text) {
  return {
    xfs:text.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g),
    fills:text.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/)[1].match(/<fill>[\s\S]*?<\/fill>/g),
    borders:text.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)[1].match(/<border>[\s\S]*?<\/border>/g)
  };
}
function styleAt(sheetXml, styles, address) {
  const attrs = sheetXml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0];
  const xf = styles.xfs[Number(attrs.match(/\bs="(\d+)"/)?.[1] || 0)];
  const fill = styles.fills[Number(xf.match(/\bfillId="(\d+)"/)[1])];
  return {xf,color:fill.match(/\brgb="([^"]+)"/)?.[1],border:styles.borders[Number(xf.match(/\bborderId="(\d+)"/)[1])]};
}

test('固定11列，两张表；日期为文本，金额精度、空值和文本不变且没有公式', async () => {
  const data = result([
    record({amount:123.456,cents:12346,summary:'=1+1',notes:'=HYPERLINK("https://example.invalid","text")'}),
    record({id:'r2',month:'月份待定',monthKey:'',amount:0,cents:0,oa:'=SUM(1,2)'}),
    record({id:'r3',month:null,monthKey:'',amount:null,cents:null,summary:'',excluded:'是',notes:'缺少金额，人工确认剔除'}),
    record({id:'r4',month:'2027年11月/12月',monthKey:'',amount:'待核对',cents:null,excluded:'是',notes:'异常金额，人工确认剔除'})
  ]);
  const before = JSON.stringify(data), {book} = await read(data), sheet = book.Sheets.付款汇总明细;
  assert.deepEqual(book.SheetNames,['付款汇总明细','付款预算汇总']);
  assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],['预估付款月份','供应商名称','付款主体','对公付款0A流程号','摘要','申请付款金额','是否已支付','备注','分类','跟单员','是否剔除']);
  assert.equal(sheet.A2.t,'s'); assert.equal(sheet.A2.v,'2026年11月'); assert.equal(sheet.A2.z,'@'); assert.equal(sheet.A2.w,'2026年11月');
  assert.equal(sheet.F2.v,123.456); assert.equal(sheet.F2.z,'#,##0.00'); assert.equal(sheet.F2.w,'123.46');
  assert.equal(sheet.F3.v,0); assert.equal(sheet.F3.t,'n'); assert.equal(sheet.F4.v,''); assert.equal(sheet.F4.t,'s');
  assert.equal(sheet.A3.v,'月份待定'); assert.equal(sheet.A4.v,''); assert.equal(sheet.A5.v,'2027年11月/12月'); assert.equal(sheet.F5.v,'待核对');
  assert.equal(sheet.D2.v,'000123'); assert.equal(sheet.D2.t,'s'); assert.equal(sheet.D3.v,'=SUM(1,2)');
  assert.equal(sheet.E2.v,'=1+1'); assert.equal(sheet.H2.v,'=HYPERLINK("https://example.invalid","text")');
  for (const name of book.SheetNames) for (const [address,value] of Object.entries(book.Sheets[name])) if(!address.startsWith('!')) assert.equal(value.f,undefined);
  assert.equal(JSON.stringify(data),before);
});

test('同页按月分类和按月汇总，记录数不同也正确，剔除记录仍保留明细但不重算core汇总', async () => {
  const totals={totalCents:12345,paidCents:2345,remainingCents:10000};
  const data=result([record(),record({id:'excluded',amount:999,excluded:'是',related:true})],{
    classified:[{monthKey:'2026-11',category:'正常付款',totalCents:2345,paidCents:2345,remainingCents:0},{monthKey:'2026-11',category:'未预估付款',totalCents:5000,paidCents:0,remainingCents:5000},{monthKey:'',category:'未预估付款',totalCents:5000,paidCents:0,remainingCents:5000}],
    monthly:[{monthKey:'2026-11',totalCents:7345,paidCents:2345,remainingCents:5000},{monthKey:'',totalCents:5000,paidCents:0,remainingCents:5000}],totals
  });
  const {book}=await read(data),s=book.Sheets.付款预算汇总;
  assert.deepEqual(X.utils.sheet_to_json(s,{header:1})[0],['付款月份','分类','申请总金额','已付款金额','剩余付款金额','备注']);
  assert.deepEqual(X.utils.sheet_to_json(s,{header:1})[6],['付款月份','申请总金额','已付款金额','剩余付款金额','备注']);
  assert.equal(s.A2.v,'2026年11月'); assert.equal(s.A2.t,'s'); assert.equal(s.A2.z,'@'); assert.equal(s.A4.v,'月份待核对'); assert.equal(s.A9.v,'月份待核对');
  assert.deepEqual([s.C5.v,s.D5.v,s.E5.v],[123.45,23.45,100]);
  assert.deepEqual([s.B10.v,s.C10.v,s.D10.v],[123.45,23.45,100]);
  assert.equal(s.F5.v,''); assert.equal(s.E10.v,'');
  assert.equal(book.Sheets.付款汇总明细.F3.v,999); assert.equal(book.Sheets.付款汇总明细.K3.v,'是');
});

test('related无论是否剔除都整行浅红，其余已支付仅金额和状态浅红，其他白底', async () => {
  const {zip}=await read(result([
    record({related:true,excluded:'是'}),record({related:true,excluded:'否'}),record({status:'已支付'}),record()
  ]));
  const styles=stylesFrom(await zip.file('xl/styles.xml').async('string')),sheet=await xml(zip,1);
  for(const row of[2,3])for(let col=0;col<11;col++)assert.equal(styleAt(sheet,styles,X.utils.encode_col(col)+row).color,'FFFFC7CE');
  for(let col=0;col<11;col++)assert.equal(styleAt(sheet,styles,X.utils.encode_col(col)+'4').color,[5,6].includes(col)?'FFFFC7CE':'FFFFFFFF');
  for(let col=0;col<11;col++)assert.equal(styleAt(sheet,styles,X.utils.encode_col(col)+'5').color,'FFFFFFFF');
});

test('普通明细筛选可取消，全部可见，两页冻结表头、居中换行边框，无Excel Table', async () => {
  const {zip,book}=await read(result([record(),record({id:'other'})]));
  assert.equal(book.Sheets.付款汇总明细['!autofilter'].ref,'A1:K3');
  assert(!Object.keys(zip.files).some(name=>name.startsWith('xl/tables/')));
  const styles=stylesFrom(await zip.file('xl/styles.xml').async('string'));
  for(const index of[1,2]){
    const sheet=await xml(zip,index);
    assert.match(sheet,/<pane[^>]*ySplit="1"[^>]*topLeftCell="A2"[^>]*state="frozen"/);
    assert.doesNotMatch(sheet,/hidden="1"|<filterColumn\b|<tableParts\b/);
    const style=styleAt(sheet,styles,'A2').xf;
    assert.match(style,/horizontal="center" vertical="center" wrapText="1" shrinkToFit="0"/);
    assert.match(style,/borderId="2"/);
    for(const c of sheet.match(/<c\b[^>]+>/g)){
      const address=c.match(/\br="([A-Z]+\d+)"/)[1],row=Number(address.match(/\d+$/)[0]);
      const border=styleAt(sheet,styles,address).border,expected=row===1||(index===2&&row===5)?'thin':'hair';
      for(const side of ['left','right','top','bottom'])assert.match(border,new RegExp(`<${side} style="${expected}"><color rgb="FF404040"/></${side}>`),address);
    }
  }
  assert.match(await xml(zip,1),/<autoFilter ref="A1:K3"\/>/);
  assert.doesNotMatch(await xml(zip,2),/<autoFilter\b/);
});

test('有内容的行包含空白备注也画边框，全空行和分隔行不画，零及false有效',async()=>{
  const empty={month:null,monthKey:'',supplier:' ',subject:'',oa:'',summary:null,amount:null,cents:0,status:'',notes:' \n ',category:'',owner:'',excluded:'',related:false};
  const {zip,book}=await read(result([record({supplier:false,notes:' \n ',summary:null,amount:0,cents:0,related:true}),record(empty),record({...empty,amount:0,cents:0}),record({...empty,supplier:false})],{classified:[],monthly:[],totals:{totalCents:0,paidCents:0,remainingCents:0}}));
  const styles=stylesFrom(await zip.file('xl/styles.xml').async('string')),detail=await xml(zip,1),summary=await xml(zip,2);
  for(const address of ['E2','H2','B2','F2']){assert.match(styleAt(detail,styles,address).border,/style="hair"/);assert.equal(styleAt(detail,styles,address).color,'FFFFC7CE');}
  for(let col=0;col<11;col++)assert.doesNotMatch(styleAt(detail,styles,X.utils.encode_col(col)+'3').border,/style=/);
  for(const address of ['A4','F4','H4','A5','B5','H5'])assert.match(styleAt(detail,styles,address).border,/style="hair"/);
  assert.equal(book.Sheets.付款汇总明细.B2.v,'false');assert.equal(book.Sheets.付款汇总明细.F2.v,0);assert.equal(book.Sheets.付款汇总明细.F2.z,'#,##0.00');assert.equal(book.Sheets.付款汇总明细.H2.v,' \n ');
  assert.match(styleAt(summary,styles,'B2').border,/style="hair"/);assert.match(styleAt(summary,styles,'C2').border,/style="hair"/);assert.match(styleAt(summary,styles,'B2').xf,/fontId="2"/);
  assert.doesNotMatch(summary,/<c\b[^>]*\br="[A-Z]+3"/);assert.equal(book.Sheets.付款汇总明细['!ref'],'A1:K5');assert.equal(book.Sheets.付款汇总明细.L2,undefined);
});

test('多行备注按内容增加行高，不裁切内容，超过Excel行高上限明确报错', async () => {
  const note='分批付款：保留对应OA与原金额。\n'.repeat(12)+'尾行';
  const {zip,book}=await read(result([record({notes:note})]));
  const row=(await xml(zip,1)).match(/<row\b(?=[^>]*\br="2")[^>]*>/)[0];
  const height=Number(row.match(/\bht="([^"]+)"/)[1]);
  assert(height>=13*16+10&&height<=409); assert.equal(book.Sheets.付款汇总明细.H2.v,note);
  await assert.rejects(Export(result([record({notes:'说明\n'.repeat(30)})])),/付款汇总明细第2行.*内容过长/);
});

test('空结果仍输出两张表、零合计、合法筛选和动态汇总行', async () => {
  const {book}=await read(result([],{classified:[],monthly:[],totals:{totalCents:0,paidCents:0,remainingCents:0}}));
  assert.equal(book.Sheets.付款汇总明细['!ref'],'A1:K1');
  assert.equal(book.Sheets.付款汇总明细['!autofilter'].ref,'A1:K1');
  const s=book.Sheets.付款预算汇总;
  assert.deepEqual([s.C2.v,s.D2.v,s.E2.v,s.B5.v,s.C5.v,s.D5.v],[0,0,0,0,0,0]);
  assert.equal(s.F2.v,''); assert.equal(s.E5.v,'');
});

test('异常数字和缺失输入阻止导出，不把未知金额写成零', async () => {
  await assert.rejects(Export({}),/缺少季度预算整理结果/);
  await assert.rejects(Export(result([record({amount:Infinity})])),/无效数字/);
  await assert.rejects(Export(result([],{totals:{totalCents:NaN,paidCents:0,remainingCents:0}})),/汇总金额无效/);
});

test('下载端再次拦截未完成核对及空金额，已确认剔除仍保留原空值',async()=>{
  await assert.rejects(Export({...result(),canExport:false}),/未完成.*核对/);
  await assert.rejects(Export({...result(),issues:[{blocking:true,resolved:false}]}),/未完成.*核对/);
  await assert.rejects(Export(result([record({amount:null,cents:null})])),/空白金额/);
  const {book}=await read(result([record({amount:null,cents:null,excluded:'是',notes:'缺少金额，已确认本次剔除'})],{classified:[],monthly:[],totals:{totalCents:0,paidCents:0,remainingCents:0}}));
  assert.equal(book.Sheets.付款汇总明细.F2.v,'');assert.equal(book.Sheets.付款汇总明细.K2.v,'是');
});

test('已校验的文本金额导出为数值，原记录内容不被改写',async()=>{
  const data=result([record({amount:'￥1,378.00',cents:137800})],{classified:[],monthly:[],totals:{totalCents:137800,paidCents:0,remainingCents:137800}});
  const {book}=await read(data);assert.equal(book.Sheets.付款汇总明细.F2.t,'n');assert.equal(book.Sheets.付款汇总明细.F2.v,1378);assert.equal(data.records[0].amount,'￥1,378.00');
});

test('手工预算对比独立第三表，合计差额与来源准确，格式范围不含多余空白',async()=>{
  const data=result();
  data.comparison={monthKey:'2026-11',rows:[
    {monthKey:'2026-11',supplier:'供应商甲',subject:'迈德斯特',category:'正常付款-预估付款',systemCents:1489200,manualCents:1489200,differenceCents:0,status:'一致',systemSources:[{file:'预估.xlsx',sheet:'邢芳芳',row:3}],manualSources:[{file:'手工.xlsx',sheet:'Sheet1',row:52}]},
    {monthKey:'2026-11',supplier:'供应商乙',subject:'国源',category:'正常付款-正常对账表',systemCents:100,manualCents:300,differenceCents:-200,status:'金额差异',systemSources:[{file:'台账.xlsx',sheet:'对账明细-李奇',row:2},{file:'台账.xlsx',sheet:'对账明细-李奇',row:3}],manualSources:[{file:'手工.xlsx',sheet:'Sheet1',row:10}]},
    {monthKey:'2026-11',supplier:'=1+1',subject:'国源',category:'正常付款-预估付款',systemCents:50,manualCents:0,differenceCents:50,status:'仅系统有',systemSources:[],manualSources:[]}
  ],totals:{systemCents:1489350,manualCents:1489500,differenceCents:-150}};
  const before=JSON.stringify(data),{book,zip}=await read(data),s=book.Sheets.下月预算对比;
  assert.deepEqual(book.SheetNames,['付款汇总明细','付款预算汇总','下月预算对比']);
  assert.deepEqual([s.E5.v,s.F5.v,s.G5.v],[14893.5,14895,-1.5]);
  assert.deepEqual([s.E3.v,s.F3.v,s.G3.v],[1,3,-2]);assert.equal(s.G2.v,0);assert.equal(s.G2.t,'n');assert.equal(s.G2.z,'#,##0.00');
  assert.equal(s.A2.v,'2026年11月');assert.equal(s.B4.v,'=1+1');assert.equal(s.B4.f,undefined);
  assert.equal(s.I3.v,'台账.xlsx / 对账明细-李奇 / 第2、3行');assert.equal(s.J2.v,'手工.xlsx / Sheet1 / 第52行');
  assert.equal(s['!ref'],'A1:J5');assert.equal(s['!autofilter'].ref,'A1:J4');assert.equal(s.K1,undefined);
  const content=await xml(zip,3),styles=stylesFrom(await zip.file('xl/styles.xml').async('string'));
  assert.match(content,/topLeftCell="A2"/);assert.doesNotMatch(content,/<filterColumn|hidden="1"/);
  assert.equal(styleAt(content,styles,'G2').color,'FFFFFFFF');assert.equal(styleAt(content,styles,'G3').color,'FFFFC7CE');
  for(const [address,value]of Object.entries(s))if(!address.startsWith('!'))assert.equal(value.f,undefined);
  assert.equal(JSON.stringify(data),before);
});

test('手工对比为空也提供表头和零合计，未知金额不能被导出为零',async()=>{
  const data=result();data.comparison={monthKey:'2026-12',rows:[],totals:{systemCents:0,manualCents:0,differenceCents:0}};
  const {book}=await read(data),s=book.Sheets.下月预算对比;assert.equal(s['!ref'],'A1:J2');assert.equal(s['!autofilter'],undefined);assert.equal(s.G2.v,0);
  data.comparison.totals.manualCents=null;await assert.rejects(Export(data),/汇总金额无效/);
});

test('月份只作展示，不参与数值求和；缺年份数值保留文本，零金额仍为数值', async () => {
  const data=result([record({amount:0,cents:0}),record({month:11,monthKey:'',amount:0,cents:0}),record({month:'2027-01-31',monthKey:'2027-01',amount:1,cents:100})]);
  const {book}=await read(data),sheet=book.Sheets.付款汇总明细;
  assert.deepEqual([sheet.A2.v,sheet.A3.v,sheet.A4.v],['2026年11月','11','2027年1月']);
  for(const address of ['A2','A3','A4']){assert.equal(sheet[address].t,'s');assert.equal(sheet[address].z,'@');assert.equal(sheet[address].f,undefined);}
  for(const address of ['F2','F3']){assert.equal(sheet[address].t,'n');assert.equal(sheet[address].v,0);}
  assert.equal([sheet.A2,sheet.A3,sheet.A4,sheet.F2,sheet.F3,sheet.F4].filter(c=>c.t==='n').reduce((sum,c)=>sum+c.v,0),1);
});

test('浏览器UMD入口可直接使用本地XLSX和JSZip导出', async () => {
  const context = {XLSX:X,JSZip:Zip};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../quarter-budget-export.js'),'utf8'),context);
  assert.equal(typeof context.QuarterBudgetExport,'function');
  const bytes = await context.QuarterBudgetExport(result());
  const book = X.read(bytes,{type:'array'});
  assert.deepEqual(book.SheetNames,['付款汇总明细','付款预算汇总']);
  assert.equal(book.Sheets.付款汇总明细.B2.v,'供应商甲');
});
