const test = require('node:test'), assert = require('node:assert/strict');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const Core = require('../table-compare-core.js'), Export = require('../table-compare-export.js');
const text = v => ({t:'s',v,z:'@'}), money = v => ({t:'n',v,z:'#,##0.00'});
const fixture = () => ({sheets:[{name:'比对汇总',headers:[...Core.HEADERS],rows:[
  [text('正常付款'),text('李奇'),text('供应商甲'),text('付款主体'),text('000123'),text('=1+1'),money(100),money(120.25),money(20.25),text('增加'),text('原备注'),text('旧摘要\n新摘要'),text('旧版.xlsx / Sheet1 第 2 行'),text('新版.xlsx / Sheet1 第 3 行'),text('')],
  [text('谈判付款'),text('徐英'),text('供应商乙'),text('付款主体'),text('未提交OA'),text('摘要'),money(-60),money(0),money(60),text('增加'),text(''),text(''),text('旧版.xlsx / Sheet1 第 3 行'),text('新版.xlsx / Sheet1 第 2 行'),text('')]
]}]});

test('单张15列结果完整重读，金额为数值，备注可填且源结果不变', async () => {
  const result = fixture(), before = JSON.stringify(result), bytes = await Export(result);
  const book = X.read(bytes,{type:'array',cellNF:true,cellStyles:true}), s = book.Sheets.比对汇总;
  assert.deepEqual(book.SheetNames,['比对汇总']);assert.equal(s['!ref'],'A1:O3');assert.equal(s['!autofilter'].ref,'A1:O3');
  assert.deepEqual(X.utils.sheet_to_json(s,{header:1,defval:'',raw:true}),[Core.HEADERS,...result.sheets[0].rows.map(row=>row.map(c=>c.v))]);
  for(const address of ['G2','H2','I2','G3','H3','I3']){assert.equal(s[address].t,'n');assert.equal(s[address].z,'#,##0.00');assert(!/\.\s*$/.test(s[address].w));}
  assert.equal(s.E2.t,'s');assert.equal(s.E2.v,'000123');assert.equal(s.F2.v,'=1+1');assert.equal(s.F2.f,undefined);
  assert.equal(s.O2.v,'');assert.equal(s.O3.v,'');assert.equal(s.P1,undefined);assert.equal(s.A4,undefined);
  assert.equal(JSON.stringify(result),before);
  const zip = await Zip.loadAsync(bytes), xml = await zip.file('xl/worksheets/sheet1.xml').async('string'), styles = await zip.file('xl/styles.xml').async('string');
  assert.match(xml,/<pane ySplit="1" topLeftCell="A2"[^>]*state="frozen"/);assert(!/<f[\s>]/.test(xml));assert(!Object.keys(zip.files).some(p=>p.startsWith('xl/tables/')));
  const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
  for(const cell of xml.match(/<c\b[^>]+>/g)){
    const address=cell.match(/\br="([A-Z]+)(\d+)"/),xf=xfs[Number(cell.match(/\bs="(\d+)"/)[1])],header=address[2]==='1';
    assert.match(xf,new RegExp(`borderId="${header?1:2}"`));assert.match(xf,new RegExp(`fillId="${header?3:2}"`));
    assert.match(xf,/horizontal="center" vertical="center"/);
    if(header||['E','G','H','I'].includes(address[1]))assert.match(xf,/wrapText="0"/);
  }
  assert.match(styles,/<left style="thin"><color rgb="FF404040"/);assert.match(styles,/<left style="hair"><color rgb="FF404040"/);
  assert.equal(s['!cols'].length,15);assert(s['!cols'].every(c=>Number.isFinite(c.wch)&&c.wch>=10));
});

test('空结果仍保留表头筛选，辅助列不产生空白边框区域', async () => {
  const result=fixture();result.sheets[0].rows=[];
  const s=X.read(await Export(result),{type:'array'}).Sheets.比对汇总;
  assert.equal(s['!ref'],'A1:O1');assert.equal(s['!autofilter'].ref,'A1:O1');assert.equal(s.A2,undefined);
});

test('生成范围里的日期为文本，缓存公式转值不带入公式', async () => {
  const result=fixture();result.sheets[0].rows[0][5]={t:'n',v:46352,z:'yyyy"年"mm"月"dd"日"',f:'DATE(2026,11,26)'};
  const s=X.read(await Export(result),{type:'array',cellNF:true}).Sheets.比对汇总;
  assert.equal(s.F2.t,'s');assert.equal(s.F2.v,'2026年11月26日');assert.equal(s.F2.f,undefined);assert.equal(s.F2.z,'@');
});
