const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Export = require('../negotiated-payment-summary-export.js');

function fixture(count = 8) {
  const months = Array.from({length:count},(_,index) => {
    const date = new Date(Date.UTC(2026,8+index,1));
    const year=date.getUTCFullYear(),month=date.getUTCMonth()+1;
    return {key:`${year}-${String(month).padStart(2,'0')}`,label:`${year}年${month}月`};
  });
  const negotiatedCents = Array.from({length:count},(_,i) => i ? null : 12345);
  const normalCents = Array.from({length:count},(_,i) => i ? null : 0);
  return {canExport:true,errors:[],unresolved:[],months,rows:[{supplier:'供应商甲',negotiatedCents,normalCents,negotiatedTotalCents:12345,normalTotalCents:0}]};
}
async function read(result) {
  const bytes=await Export(result),zip=await Zip.loadAsync(bytes),book=X.read(bytes,{type:'array',cellNF:true,cellStyles:true});
  return {bytes,zip,book,sheet:book.Sheets['谈判付款整理'],xml:await zip.file('xl/worksheets/sheet1.xml').async('string'),styles:await zip.file('xl/styles.xml').async('string')};
}

test('八个月生成单张19列表，双层表头合并及年月日期正确',async () => {
  const result=fixture(),before=JSON.stringify(result),{book,sheet}=await read(result);
  assert.deepEqual(book.SheetNames,['谈判付款整理']);
  assert.equal(sheet['!ref'],'A1:S3');
  assert.deepEqual(['A1','B1','C1','K1','L1'].map(a=>sheet[a].v),['供应商','谈判总金额','谈判付款节奏','正常付款总金额','正常付款']);
  assert.deepEqual(sheet['!merges'].map(X.utils.encode_range),['A1:A2','B1:B2','K1:K2','C1:J1','L1:S1']);
  for (let i=0;i<8;i++) {
    const left=sheet[X.utils.encode_col(i+2)+'2'],right=sheet[X.utils.encode_col(i+11)+'2'];
    assert.equal(left.t,'n');assert.equal(left.z,'yyyy"年"m"月"');assert.equal(left.w,result.months[i].label);
    assert.equal(right.v,left.v);assert.equal(right.w,left.w);
  }
  assert.equal(sheet.C2.v,46266);assert.equal(sheet.G2.w,'2027年1月');
  assert.equal(JSON.stringify(result),before);
});

test('合计使用动态SUM并缓存金额，保留空白、零、负数及数字类型',async () => {
  const result=fixture();
  result.rows[0].negotiatedCents[1]=-45;result.rows[0].negotiatedTotalCents=12300;
  result.rows[0].normalCents[7]=90001;result.rows[0].normalTotalCents=90001;
  const {sheet}=await read(result);
  assert.equal(sheet.B3.f,'SUM(C3:J3)');assert.equal(sheet.B3.v,123);assert.equal(sheet.B3.z,'#,##0.00');
  assert.equal(sheet.K3.f,'SUM(L3:S3)');assert.equal(sheet.K3.v,900.01);
  assert.equal(sheet.C3.v,123.45);assert.equal(sheet.D3.v,-.45);
  assert.equal(sheet.E3.v,'');assert.equal(sheet.L3.v,0);assert.equal(sheet.L3.t,'n');assert.equal(sheet.S3.v,900.01);
});

test('所有表头同蓝色，正文白底，全范围含空白格都有边框、不换行并冻结两行',async () => {
  const {xml,styles,sheet}=await read(fixture());
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  for (let row=1;row<=3;row++)for(let c=0;c<19;c++) {
    const address=X.utils.encode_col(c)+row;
    const attrs=xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${address}")[^>]*>`))?.[0];
    assert(attrs,address+' should have an explicit cell for its border');
    const xf=xfs[Number(attrs.match(/\bs="(\d+)"/)[1])];
    assert.match(xf,/borderId="1"/);assert.match(xf,/wrapText="0" shrinkToFit="0"/);
    assert.match(xf,new RegExp(`fillId="${row<=2?3:2}"`));
  }
  assert.match(styles,/<sz val="11"/);assert.match(styles,/<name val="宋体"/);assert.match(styles,/FF315B9D/);
  assert.match(xml,/<pane ySplit="2" topLeftCell="A3"[^>]*state="frozen"/);
  assert.doesNotMatch(xml,/<autoFilter\b|hidden="1"|<tableParts\b/);
  assert(sheet['!cols'][10].wch>=18,'normal total heading must fit');
  assert(sheet['!cols'][2].wch>=13,'full year/month must fit');
});

test('月份只有一个或跨越Z列时，合计范围、合并和末列仍正确',async () => {
  for (const count of [1,14]) {
    const {sheet}=await read(fixture(count));
    const total=X.utils.encode_col(count+2),start=X.utils.encode_col(count+3),end=X.utils.encode_col(count*2+2);
    assert.equal(sheet['!ref'],`A1:${end}3`);
    assert.equal(sheet.B3.f,`SUM(C3:${X.utils.encode_col(count+1)}3)`);
    assert.equal(sheet[total+'3'].f,`SUM(${start}3:${end}3)`);
    assert.equal(sheet['!merges'].length,count===1?3:5);
    assert.equal(sheet[end+'2'].w,fixture(count).months.at(-1).label);
    if (count===1) {
      assert.equal(sheet.C2.w,'2026年9月');
      assert(sheet['!cols'][2].wch>=16,'single-month group heading must fit without a merge');
    }
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

test('拒绝未完成核对、异常金额、范围不完整、合计不一致及重复月份',async () => {
  for (const change of [
    r=>{r.canExport=false;},r=>{r.unresolved=[{id:'G40'}];},r=>{r.errors=['坏金额'];},
    r=>{r.rows[0].negotiatedCents[0]=NaN;},r=>{r.rows[0].normalTotalCents=1;},
    r=>{r.rows[0].normalCents.pop();},r=>{r.rows[0].supplier='';},
    r=>{r.months[1].key=r.months[0].key;},r=>{r.months[0].key='2026-13';},
    r=>{r.rows[0].normalTotalCents=Number.MAX_SAFE_INTEGER+1;}
  ]) {
    const data=fixture();change(data);await assert.rejects(Export(data));
  }
  await assert.rejects(Export({canExport:true,months:[],rows:[]}));
  await assert.rejects(Export());
});

test('接近整数分上限而转元后损失一分时拒绝导出',async () => {
  const data=fixture();
  data.rows[0].negotiatedCents[0]=Number.MAX_SAFE_INTEGER-1;
  data.rows[0].negotiatedTotalCents=Number.MAX_SAFE_INTEGER-1;
  await assert.rejects(Export(data),/精确保存到分/);
});

test('浏览器UMD接口使用现有本地XLSX与JSZip',async () => {
  const context={XLSX:X,JSZip:Zip};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../negotiated-payment-summary-export.js'),'utf8'),context);
  assert.equal(typeof context.NegotiatedPaymentSummaryExport,'function');
  const bytes=await context.NegotiatedPaymentSummaryExport(fixture());
  assert.deepEqual(X.read(bytes,{type:'array'}).SheetNames,['谈判付款整理']);
});
