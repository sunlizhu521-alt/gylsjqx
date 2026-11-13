const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Core = require('../negotiated-payment-summary-core.js');
const Export = require('../negotiated-payment-summary-export.js');

function fixture(count = 8) {
  const dates = Array.from({length:count},(_,index) => {
    const date = new Date(Date.UTC(2026,8,1+index*7));
    const year=date.getUTCFullYear(),month=date.getUTCMonth()+1,day=date.getUTCDate();
    return {key:`${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`,label:`${year}年${month}月${day}日`};
  });
  const negotiatedCents = Array.from({length:count},(_,i) => i ? 0 : 12345);
  const normalCents = Array.from({length:count},() => 0);
  return {canExport:true,errors:[],unresolved:[],dates,rows:[{supplier:'供应商甲',negotiatedCents,normalCents,negotiatedTotalCents:12345,normalTotalCents:0}]};
}
async function read(result) {
  const bytes=await Export(result),zip=await Zip.loadAsync(bytes),book=X.read(bytes,{type:'array',cellNF:true,cellStyles:true});
  return {bytes,zip,book,sheet:book.Sheets['谈判付款整理'],validation:book.Sheets['校验'],xml:await zip.file('xl/worksheets/sheet1.xml').async('string'),validationXml:await zip.file('xl/worksheets/sheet2.xml').async('string'),styles:await zip.file('xl/styles.xml').async('string')};
}

test('八个日期生成19列整理表并附加校验表，双层表头合并及年月日文本正确',async () => {
  const result=fixture(),before=JSON.stringify(result),{book,sheet}=await read(result);
  assert.deepEqual(book.SheetNames,['谈判付款整理','校验']);
  assert.equal(sheet['!ref'],'A1:S3');
  assert.deepEqual(['A1','B1','C1','K1','L1'].map(a=>sheet[a].v),['供应商','谈判总金额','谈判付款节奏','正常付款总金额','正常付款']);
  assert.deepEqual(sheet['!merges'].map(X.utils.encode_range),['A1:A2','B1:B2','K1:K2','C1:J1','L1:S1']);
  for (let i=0;i<8;i++) {
    const left=sheet[X.utils.encode_col(i+2)+'2'],right=sheet[X.utils.encode_col(i+11)+'2'];
    assert.equal(left.t,'s');assert.equal(left.z,'@');assert.equal(left.v,result.dates[i].label);assert.equal(left.w,result.dates[i].label);
    assert.equal(right.v,left.v);assert.equal(right.w,left.w);
  }
  assert.equal(sheet.C2.v,'2026年9月1日');assert.equal(sheet.D2.v,'2026年9月8日');assert.equal(sheet.G2.w,'2026年9月29日');
  assert.equal(JSON.stringify(result),before);
});

test('合计使用动态SUM并缓存金额，保留空白、零、负数及数字类型',async () => {
  const result=fixture();
  result.rows[0].negotiatedCents[1]=-45;result.rows[0].negotiatedTotalCents=12300;
  result.rows[0].negotiatedCents[2]=null;
  result.rows[0].normalCents[7]=90001;result.rows[0].normalTotalCents=90001;
  result.rows.push({supplier:'保留日期的另一供应商',negotiatedCents:[null,null,1,null,null,null,null,null],normalCents:Array(8).fill(null),negotiatedTotalCents:1,normalTotalCents:0});
  const {sheet}=await read(result);
  assert.equal(sheet.B3.f,'SUM(C3:J3)');assert.equal(sheet.B3.v,123);assert.equal(sheet.B3.z,'#,##0.00');
  assert.equal(sheet.K3.f,'SUM(L3:S3)');assert.equal(sheet.K3.v,900.01);
  assert.equal(sheet.C3.v,123.45);assert.equal(sheet.D3.v,-.45);
  assert.equal(sheet.E3.v,'');assert.equal(sheet.L3.v,0);assert.equal(sheet.L3.t,'n');assert.equal(sheet.S3.v,900.01);
});

test('同一月不同日期分列，两类付款分别删除整列空日期且金额只计入对应日',async () => {
  const data=fixture(2);
  data.dates=[{key:'2026-11-02',label:'2026年11月2日'},{key:'2026-11-18',label:'2026年11月18日'}];
  data.rows[0]={supplier:'同月多批供应商',negotiatedCents:[10000,20000],normalCents:[null,5000],negotiatedTotalCents:30000,normalTotalCents:5000};
  const {sheet}=await read(data);
  assert.equal(sheet['!ref'],'A1:F3');
  assert.deepEqual(['C2','D2','F2'].map(a=>sheet[a].v),['2026年11月2日','2026年11月18日','2026年11月18日']);
  assert.deepEqual(['C3','D3','F3'].map(a=>sheet[a].v),[100,200,50]);
  assert.equal(sheet.B3.f,'SUM(C3:D3)');assert.equal(sheet.B3.v,300);
  assert.equal(sheet.E3.f,'SUM(F3:F3)');assert.equal(sheet.E3.v,50);
  for(const c of [2,3,5]){
    const needed=[...sheet[X.utils.encode_col(c)+'2'].v].reduce((n,ch)=>n+(ch.codePointAt(0)>255?2:1),4);
    assert(sheet['!cols'][c].wch>=needed,'完整具体日期和筛选按钮需完整显示');
  }
});

test('不对称日期保留明确零、负数及整列相抵为零，其他供应商的单格空白不删',async () => {
  const data=fixture(4);
  data.rows=[
    {supplier:'供应商甲',negotiatedCents:[0,null,null,null],normalCents:[null,700,null,null],negotiatedTotalCents:0,normalTotalCents:700},
    {supplier:'供应商乙',negotiatedCents:[null,-500,100,null],normalCents:[null,-700,null,null],negotiatedTotalCents:-400,normalTotalCents:-700}
  ];
  const {sheet,xml,styles}=await read(data);
  assert.equal(sheet['!ref'],'A1:G4');assert.equal(sheet['!autofilter'].ref,'A2:G4');
  assert.deepEqual(['C2','D2','E2','G2'].map(a=>sheet[a].v),[data.dates[0].label,data.dates[1].label,data.dates[2].label,data.dates[1].label]);
  assert.deepEqual(['C3','D3','E3','C4','D4','E4','G3','G4'].map(a=>sheet[a].v),[0,'','','',-5,1,7,-7]);
  assert.equal(sheet.B4.f,'SUM(C4:E4)');assert.equal(sheet.B4.v,-4);
  assert.equal(sheet.F3.f,'SUM(G3:G3)');assert.equal(sheet.F3.v,7);
  assert.deepEqual(sheet['!merges'].map(X.utils.encode_range),['A1:A2','B1:B2','F1:F2','C1:E1']);
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  for(const address of ['D3','E3','C4']){
    const attrs=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0];
    assert.match(xfs[Number(attrs.match(/\bs="(\d+)"/)[1])],/borderId="2"/,'有效行列内部空白仍有发丝边框');
  }
});

test('取消标红后旧已付标记不影响导出，两侧日期金额与全部正文均保持白底',async () => {
  const data=fixture(5);
  data.rows=[
    {supplier:'汇总已付和部分已付',negotiatedCents:[null,10000,null,20000,0],normalCents:[5000,null,null,null,6000],negotiatedTotalCents:30000,normalTotalCents:11000},
    {supplier:'尚未付清',negotiatedCents:[null,7000,null,null,null],normalCents:[null,null,null,null,8000],negotiatedTotalCents:7000,normalTotalCents:8000}
  ];
  data.paid={negotiated:[[false,true,false,false,true],[false,false,false,true,false]],normal:[[false,false,false,false,true],[true,false,false,false,false]]};
  const before=JSON.stringify(data),{sheet,xml,styles}=await read(data);
  assert.equal(sheet['!ref'],'A1:H4');
  assert.deepEqual(['C2','D2','E2','G2','H2'].map(a=>sheet[a].v),[data.dates[1].label,data.dates[3].label,data.dates[4].label,data.dates[0].label,data.dates[4].label]);
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  const fills=styles.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/)[1].match(/<fill\b[^>]*>[\s\S]*?<\/fill>/g);
  for(let row=1;row<=4;row++)for(let col=0;col<8;col++){
    const address=X.utils.encode_col(col)+row,attrs=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0];
    const xf=xfs[Number(attrs.match(/\bs="(\d+)"/)[1])],fill=fills[Number(xf.match(/\bfillId="(\d+)"/)[1])];
    assert.match(fill,new RegExp(`fgColor rgb="${row<=2?'FF315B9D':'FFFFFFFF'}"`),address);
    assert.match(xf,new RegExp(`borderId="${row<=2?1:2}"`));
  }
  assert.doesNotMatch(xml,/<conditionalFormatting\b/);
  assert.doesNotMatch(styles,/FFFF0000/,'不再输出已付专用红色样式');
  assert.equal(sheet.C3.v,100);assert.equal(sheet.E3.v,0);assert.equal(sheet.H3.v,60);
  assert.equal(sheet.C3.z,'#,##0.00');assert.equal(sheet.D3.z,'#,##0.00');
  assert.equal(sheet.B3.f,'SUM(C3:E3)');assert.equal(sheet.B3.v,300);assert.equal(sheet.F3.f,'SUM(G3:H3)');assert.equal(sheet.F3.v,110);
  assert.equal(sheet.D4.v,'','空白不因错误paid标记被标红或填值');
  assert.equal(JSON.stringify(data),before);
});

test('未提供或提供任意旧已付款矩阵时日期金额都保持白底',async () => {
  for(const paid of [undefined,{negotiated:[[false]],normal:[[false]]},{negotiated:[[true]],normal:[[true]]}]){
    const data=fixture(1);if(paid)data.paid=paid;
    const {sheet,xml}=await read(data);
    for(const address of ['A3','B3','C3','D3','E3'])assert.equal(sheet[address].s.fgColor.rgb,'FFFFFF',address);
    assert.doesNotMatch(xml,/<conditionalFormatting\b/);
  }
});

test('单侧没有数值日期时只保留总金额，另一侧范围、组头和筛选仍正确',async () => {
  for(const negotiatedEmpty of [true,false]){
    const data=fixture(3);
    data.rows[0]={supplier:'单侧付款',negotiatedCents:negotiatedEmpty?[null,null,null]:[null,-100,null],normalCents:negotiatedEmpty?[null,100,null]:[null,null,null],negotiatedTotalCents:negotiatedEmpty?0:-100,normalTotalCents:negotiatedEmpty?100:0};
    const {sheet}=await read(data);
    assert.equal(sheet['!ref'],'A1:D3');assert.equal(sheet['!autofilter'].ref,'A2:D3');
    assert.equal(sheet['!merges'].length,3);
    if(negotiatedEmpty){
      assert.equal(sheet.C1.v,'正常付款总金额');assert.equal(sheet.D1.v,'正常付款');assert.equal(sheet.D2.v,data.dates[1].label);
      assert.equal(sheet.B3.f,'SUM(0)');assert.equal(sheet.B3.v,0);assert.equal(sheet.C3.f,'SUM(D3:D3)');assert.equal(sheet.C3.v,1);
      assert(!Object.values(sheet).some(c=>c?.v==='谈判付款节奏'));
    }else{
      assert.equal(sheet.C1.v,'谈判付款节奏');assert.equal(sheet.D1.v,'正常付款总金额');assert.equal(sheet.C2.v,data.dates[1].label);
      assert.equal(sheet.B3.f,'SUM(C3:C3)');assert.equal(sheet.B3.v,-1);assert.equal(sheet.D3.f,'SUM(0)');assert.equal(sheet.D3.v,0);
      assert(!Object.values(sheet).some(c=>c?.v==='正常付款'));
    }
  }
});

test('两侧日期金额全部为空或没有日期时仍输出供应商与两个零总金额',async () => {
  for(const count of [0,3]){
    const data=fixture(count);
    data.rows[0]={supplier:'没有分批金额',negotiatedCents:Array(count).fill(null),normalCents:Array(count).fill(null),negotiatedTotalCents:0,normalTotalCents:0};
    const {sheet}=await read(data);
    assert.equal(sheet['!ref'],'A1:C3');assert.equal(sheet['!autofilter'].ref,'A2:C3');
    assert.deepEqual(['A1','B1','C1'].map(a=>sheet[a].v),['供应商','谈判总金额','正常付款总金额']);
    assert.deepEqual(sheet['!merges'].map(X.utils.encode_range),['A1:A2','B1:B2','C1:C2']);
    for(const a of ['B3','C3']){assert.equal(sheet[a].f,'SUM(0)');assert.equal(sheet[a].v,0);assert.equal(sheet[a].t,'n');}
  }
});

test('先校验全部原金额，不能通过删除空列绕过NaN、Infinity、缺值、字符串和小数分',async () => {
  for(const invalid of [NaN,Infinity,undefined,'123',false,1.5])for(const field of ['negotiatedCents','normalCents']){
    const data=fixture(2);data.rows[0][field][1]=invalid;
    await assert.rejects(Export(data),/金额不是有效的整数分/);
  }
});

test('真实闰日及跨年日期仍导出文本，具体日期不会变成可求和的序列数',async () => {
  const data=fixture(5);
  data.dates=['1900-02-28','2000-02-29','2024-02-29','2024-12-31','2025-01-01'].map(key=>{
    const [y,m,d]=key.split('-').map(Number);return {key,label:`${y}年${m}月${d}日`};
  });
  const {sheet}=await read(data);
  for(let i=0;i<data.dates.length;i++)for(const c of [i+2,i+8]){
    const value=sheet[X.utils.encode_col(c)+'2'];
    assert.equal(value.t,'s');assert.equal(value.z,'@');assert.equal(value.v,data.dates[i].label);assert.equal(value.f,undefined);
  }
});

test('无效日历日期、重复日期和逆序日期禁止导出，不自动滚动到下月',async () => {
  for(const key of ['2026-02-29','1900-02-29','2100-02-29','2024-04-31','2026-11-00','2026-11-32','2026-13-01','1899-12-31','2026-1-02','2026-11','not-a-date']){
    const data=fixture(1);data.dates[0].key=key;
    await assert.rejects(Export(data),/付款日期无效/,key);
  }
  for(const keys of [['2026-11-18','2026-11-18'],['2026-11-18','2026-11-02']]){
    const data=fixture(2);data.dates=keys.map(key=>({key}));
    await assert.rejects(Export(data),/重复或未按日期升序/);
  }
});

test('表头深灰细边框，正文有效行列交叉区域含中间空白和零值均保留发丝边框',async () => {
  const {xml,styles,sheet}=await read(fixture());
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  const borders=styles.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)[1].match(/<border\b[^>]*>[\s\S]*?<\/border>/g);
  for (let row=1;row<=3;row++)for(let c=0;c<19;c++) {
    const address=X.utils.encode_col(c)+row;
    const attrs=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))?.[0];
    assert(attrs,address+' should preserve its explicit cell');
    const xf=xfs[Number(attrs.match(/\bs="(\d+)"/)[1])];
    const borderId=Number(xf.match(/borderId="(\d+)"/)[1]);
    assert.equal(borderId,row<=2?1:2,address);
    for(const side of ['left','right','top','bottom']) assert.match(borders[borderId],new RegExp(`<${side} style="${row<=2?'thin':'hair'}"><color rgb="FF404040"\\/></${side}>`),address+' '+side);
    assert.match(xf,/wrapText="0" shrinkToFit="0"/);
    assert.match(xf,new RegExp(`fillId="${row<=2?3:2}"`));
  }
  assert.match(styles,/<sz val="11"/);assert.match(styles,/<name val="宋体"/);assert.match(styles,/FF315B9D/);
  assert.match(xml,/<pane ySplit="2" topLeftCell="A3"[^>]*state="frozen"/);
  assert.doesNotMatch(xml,/hidden="1"|<tableParts\b/);
  assert.match(xml,/<autoFilter ref="A2:S3"\/>/);
  assert(sheet['!cols'][10].wch>=18,'normal total heading must fit');
  assert(sheet['!cols'][2].wch>=16,'full year/month/day must fit');
  assert.equal(sheet.D3.v,0,'中间零金额日期仍为零，边框不能改变数值');
  assert.equal(sheet.L3.v,0,'零金额仍为数值零');
});

test('日期只有一个或跨越Z列时，合计范围、合并和末列仍正确',async () => {
  for (const count of [1,14]) {
    const {sheet}=await read(fixture(count));
    const total=X.utils.encode_col(count+2),start=X.utils.encode_col(count+3),end=X.utils.encode_col(count*2+2);
    assert.equal(sheet['!ref'],`A1:${end}3`);
    assert.equal(sheet.B3.f,`SUM(C3:${X.utils.encode_col(count+1)}3)`);
    assert.equal(sheet[total+'3'].f,`SUM(${start}3:${end}3)`);
    assert.equal(sheet['!merges'].length,count===1?3:5);
    assert.equal(sheet[end+'2'].w,fixture(count).dates.at(-1).label);
    if (count===1) {
      assert.equal(sheet.C2.w,'2026年9月1日');
      assert(sheet['!cols'][2].wch>=16,'single-date group heading must fit without a merge');
    }
  }
});

test('第二行普通筛选覆盖全部供应商、合计及日期列，范围随日期和明细行变化',async () => {
  for (const count of [1,8,14]) {
    const data=fixture(count);
    data.rows.push({...data.rows[0],supplier:'供应商乙'});
    const {sheet,xml,zip}=await read(data);
    const end=X.utils.encode_col(count*2+2),range=`A2:${end}4`;
    assert.equal(sheet['!autofilter'].ref,range);
    assert.match(xml,new RegExp(`<autoFilter ref="${range}"\\/>`));
    assert.doesNotMatch(xml,/<filterColumn\b|<tableParts\b/);
    assert(!Object.keys(zip.files).some(name=>name.startsWith('xl/tables/')));
    assert.deepEqual(sheet['!merges'].slice(0,3).map(X.utils.encode_range),['A1:A2','B1:B2',`${X.utils.encode_col(count+2)}1:${X.utils.encode_col(count+2)}2`]);
    assert.equal(sheet.B4.f,`SUM(C4:${X.utils.encode_col(count+1)}4)`);
  }
});

test('供应商文字不作为公式执行，列宽按长名称和金额自适应',async () => {
  const result=fixture();result.rows[0].supplier='=HYPERLINK("https://example.invalid","浙江测试医疗器械科技有限公司")';
  const {sheet}=await read(result);
  assert.equal(sheet.A3.t,'s');assert.equal(sheet.A3.f,undefined);assert.equal(sheet.A3.v,result.rows[0].supplier);
  assert(sheet['!cols'][0].wch>70);
  const next=fixture();next.rows[0].negotiatedCents[0]=123456789012345;next.rows[0].negotiatedTotalCents=123456789012345;
  const large=await read(next);
  assert(large.sheet['!cols'][2].wch>=23);
});

test('拒绝未完成核对、异常金额、范围不完整、合计不一致及重复日期',async () => {
  for (const change of [
    r=>{r.canExport=false;},r=>{r.unresolved=[{id:'G40'}];},r=>{r.errors=['坏金额'];},
    r=>{r.rows[0].negotiatedCents[0]=NaN;},r=>{r.rows[0].normalTotalCents=1;},
    r=>{r.rows[0].normalCents.pop();},r=>{r.rows[0].supplier='';},
    r=>{r.dates[1].key=r.dates[0].key;},r=>{r.dates[0].key='2026-13-01';},
    r=>{r.rows[0].normalTotalCents=Number.MAX_SAFE_INTEGER+1;}
  ]) {
    const data=fixture();change(data);await assert.rejects(Export(data));
  }
  await assert.rejects(Export({canExport:true,dates:[],rows:[]}));
  await assert.rejects(Export());
});

test('接近整数分上限而转元后损失一分时拒绝导出',async () => {
  const data=fixture();
  data.rows[0].negotiatedCents[0]=Number.MAX_SAFE_INTEGER-1;
  data.rows[0].negotiatedTotalCents=Number.MAX_SAFE_INTEGER-1;
  await assert.rejects(Export(data),/精确保存到分/);
});

test('浏览器UMD接口使用现有本地XLSX与JSZip',async () => {
  const context={XLSX:X,JSZip:Zip,NegotiatedPaymentSummaryCore:Core};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../negotiated-payment-summary-export.js'),'utf8'),context);
  assert.equal(typeof context.NegotiatedPaymentSummaryExport,'function');
  const bytes=await context.NegotiatedPaymentSummaryExport(fixture());
  assert.deepEqual(X.read(bytes,{type:'array'}).SheetNames,['谈判付款整理','校验']);
});

test('校验逐行保留来源、编号文本与分精确金额，差异允许下载且合计不参与筛选',async () => {
  const data=fixture(1);
  data.validationRows=[
    ['付款节奏.xlsx','谈判付款',8,'供应商甲','国源','00123','是',12345,12345,12345,12345,0,0,0,'一致',''],
    ['付款节奏.xlsx','谈判付款',9,'供应商甲','国源','00124','否',5000,4000,null,0,-1000,null,-5000,'不一致','核对 Q-AB 日期金额；核对月度支付金额。'],
    ['','','','合计','','','',17345,16345,12345,12345,-1000,null,-5000,'不一致','请处理上述差异。']
  ];
  const before=JSON.stringify(data),{book,validation,validationXml,styles}=await read(data);
  assert.deepEqual(book.SheetNames,['谈判付款整理','校验']);
  assert.equal(validation['!ref'],'A1:P4');
  assert.deepEqual(Array.from({length:16},(_,c)=>validation[X.utils.encode_col(c)+'1'].v),Core.VALIDATION_HEADERS);
  assert.equal(validation.C2.t,'n');assert.equal(validation.C2.v,8);
  assert.equal(validation.F2.t,'s');assert.equal(validation.F2.v,'00123');assert.equal(validation.F2.f,undefined);
  assert.equal(validation.H2.t,'n');assert.equal(validation.H2.v,123.45);assert.equal(validation.H2.z,'#,##0.00');
  assert.equal(validation.J3.v,'');assert.equal(validation.K3.v,0);assert.equal(validation.K3.t,'n');
  assert.equal(validation.L3.v,-10);assert.equal(validation.N3.v,-50);
  assert.equal(validation.D4.v,'合计');assert.equal(validation.H4.v,173.45);
  assert.equal(validation['!autofilter'].ref,'A1:P3');
  assert.match(validationXml,/<pane ySplit="1" topLeftCell="A2"[^>]*state="frozen"/);
  assert.match(validationXml,/<autoFilter ref="A1:P3"\/>/);
  assert.doesNotMatch(validationXml,/<mergeCells\b|<conditionalFormatting\b|<tableParts\b/);
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  for(let row=1;row<=4;row++)for(let c=0;c<16;c++){
    const address=X.utils.encode_col(c)+row,attrs=validationXml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))[0];
    const xf=xfs[Number(attrs.match(/\bs="(\d+)"/)[1])];
    assert.match(xf,new RegExp(`fillId="${row===1?3:2}"`),address);
    assert.match(xf,new RegExp(`borderId="${row===1?1:2}"`),address);
    assert.match(xf,/horizontal="center" vertical="center" wrapText="0" shrinkToFit="0"/);
  }
  assert.doesNotMatch(styles,/FFFF0000/);
  for(let c=0;c<16;c++)assert(validation['!cols'][c].wch>=Core.VALIDATION_HEADERS[c].split('').reduce((sum,ch)=>sum+(ch.charCodeAt(0)>255?2:1),4));
  assert.equal(JSON.stringify(data),before);
});

test('长校验说明保持原值和不自动换行，超过Excel列宽或行高限制仍允许下载',async () => {
  const data=fixture(1),message='需要核实谈判明细和分批日期金额是否遗漏。'.repeat(20);
  data.validationRows=[['输入.xlsx','谈判付款',2,'供应商甲','','MDST-0001','是',12345,12345,12345,null,0,0,null,'待核对',message]];
  const {validation}=await read(data);
  assert.equal(validation.P2.v,message);
  assert.equal(validation['!cols'][15].wch,255);
  assert.equal(validation['!autofilter'].ref,'A1:P2');
  const multiline=Array(30).fill('请核实来源金额。').join('\n');
  data.validationRows[0][15]=multiline;
  const next=await read(data);
  assert.equal(next.validation.P2.v,multiline);
  const row=next.validationXml.match(/<row\b(?=[^>]*\br="2")[^>]*>/)[0];
  assert.equal(Number(row.match(/\bht="([\d.]+)"/)[1]),409);
});

test('旧调用缺少校验记录时仍生成校验表头，异常金额不被静默转换为零',async () => {
  const {validation}=await read(fixture(1));
  assert.equal(validation['!ref'],'A1:P1');assert.equal(validation['!autofilter'].ref,'A1:P1');
  for(const value of [NaN,Infinity,'1.00',1.5]){
    const data=fixture(1);data.validationRows=[['','','','供应商','','','',value]];
    await assert.rejects(Export(data),/金额不是有效的整数分/);
  }
});
