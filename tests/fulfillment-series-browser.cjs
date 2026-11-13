const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js');
const HEADERS=['采购下单人','事业部','销售分渠道','系列','上期未交付数据','本期未交付数据'];
function fixture(stage){
 const book=X.utils.book_new();
 for(const [index,name]of [...C.SHEETS,'孙立柱'].entries()){
  const records=index===4?[['排除','其他采购员','其他系列',999]]:stage===1?
   [[name+'-共','采购员甲','共用系列',index+1],...(index===0?[['消失','采购员乙','旧系列',5],['小数1','小数采购员','小数系列',0.1],['小数2','小数采购员','小数系列',0.2]]:[])]:
   [[name+'-共','采购员甲','共用系列',index+2],...(index===0?[['新增','采购员乙','新系列',7],['小数1','小数采购员','小数系列',0.3],['零值','零值采购员','零值系列',0]]:[])];
  const rows=records.map(([code,owner,series,quantity])=>C.COLUMNS.map(key=>({
   '物料编码':code,'SKU':'SKU-'+code,'物料名称':'物料'+code,'采购下单人':owner,
   '事业部':'国内事业部','销售分渠道':'国内事业部*TEMU业务部','系列':series,
   '供应商简称':'合成供应商','运营':'王志豪','未交付数量':stage===3&&index===0?'错误数量':quantity,
   '下单月份':'2026年11月','下单日期':'2026年11月1日'
  })[key]??''));
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([C.COLUMNS,...rows]),name);
 }
 return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[],base=process.env.QA_URL||'http://127.0.0.1:8940/';
  page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('fulfillment.html',base).href);await page.waitForFunction(()=>!document.getElementById('file').disabled);
  const upload=async(bytes,name)=>{await page.locator('#file').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});await page.waitForFunction(()=>!document.getElementById('file').disabled);};
  const history=()=>page.evaluate(()=>FulfillmentCache.get());
  const download=async selector=>{const pending=page.waitForEvent('download');await page.locator(selector).click();const file=await pending,bytes=await fs.readFile(await file.path());return {bytes,book:X.read(bytes,{cellNF:true,cellStyles:true}),name:file.suggestedFilename()};};
  const rows=book=>X.utils.sheet_to_json(book.Sheets.系列对比,{header:1,defval:''});
  await upload(fixture(1),'上期更新版.xlsx');assert.equal((await history()).previous,null);
  const first=await download('#download');assert.deepEqual(rows(first.book),[
   HEADERS,['采购员甲','国内事业部','国内事业部*TEMU业务部','共用系列','',10],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','旧系列','',5],
   ['小数采购员','国内事业部','国内事业部*TEMU业务部','小数系列','',0.3]
  ]);
  await upload(fixture(2),'本期更新版.xlsx');const expected=[
   HEADERS,['采购员甲','国内事业部','国内事业部*TEMU业务部','共用系列',10,14],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','新系列',0,7],
   ['小数采购员','国内事业部','国内事业部*TEMU业务部','小数系列',0.3,0.3],
   ['零值采购员','国内事业部','国内事业部*TEMU业务部','零值系列',0,0],
   ['采购员乙','国内事业部','国内事业部*TEMU业务部','旧系列',5,0]
  ];
  let saved=await download('#download');assert.deepEqual(rows(saved.book),expected);
  assert.deepEqual(saved.book.SheetNames,['汇总表','合并核对','差异','系列对比',...C.SHEETS,'孙立柱']);
  const baseline=await history();assert.equal(baseline.previous.filename,'上期更新版.xlsx');
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);
  assert.deepEqual(rows((await download('#download')).book),expected);
  await upload(fixture(2),'重复副本.xlsx');const duplicate=await history();assert.deepEqual(duplicate.previous,baseline.previous);assert.deepEqual(duplicate.current.result.rows,baseline.current.result.rows);
  await upload(fixture(3),'异常更新版.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#status').innerText(),/数量校验未通过/);assert.deepEqual(await history(),duplicate);
  await page.reload();await page.waitForFunction(()=>!document.getElementById('download').disabled);saved=await download('#download');assert.deepEqual(rows(saved.book),expected);
  const division=await download('[data-download-division="国内事业部"]');assert.deepEqual(division.book.SheetNames,['汇总表']);
  const review=await download('#download-review');assert.deepEqual(review.book.SheetNames,['合并核对']);
  const sheet=saved.book.Sheets.系列对比;assert.equal(sheet['!autofilter'].ref,'A1:F6');assert.equal(sheet.A1.s.fgColor.rgb,'315B9D');
  for(let r=2;r<=6;r++)for(const col of ['E','F'])assert.equal(sheet[col+r].t,'n');
  const style=await page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=x=>p.parseFromString(x,'application/xml');
   const wb=parse(await z.file('xl/workbook.xml').async('string')),rels=parse(await z.file('xl/_rels/workbook.xml.rels').async('string')),entry=all(wb,'sheet').find(e=>e.getAttribute('name')==='系列对比');
   const rel=all(rels,'Relationship').find(e=>e.getAttribute('Id')===entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
   const doc=parse(await z.file(new URL(rel.getAttribute('Target'),'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1)).async('string')),styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],borders=all(styles,'borders')[0];
   return {formulaCount:all(doc,'f').length,hiddenRows:all(doc,'row').filter(r=>r.getAttribute('hidden')==='1').length,filterCount:all(doc,'filterColumn').length,aligned:all(doc,'c').every(c=>{const xf=xfs.children[Number(c.getAttribute('s')||0)],a=all(xf,'alignment')[0];return a?.getAttribute('horizontal')==='center'&&a.getAttribute('vertical')==='center';}),bordered:all(doc,'c').filter(c=>all(c,'v').length||all(c,'t').length).every(c=>{const xf=xfs.children[Number(c.getAttribute('s')||0)],border=borders.children[Number(xf.getAttribute('borderId')||0)];return ['left','right','top','bottom'].every(edge=>!!all(border,edge)[0]?.getAttribute('style'));})};
  },[...saved.bytes]);
  assert.deepEqual(style,{formulaCount:0,hiddenRows:0,filterCount:0,aligned:true,bordered:true});assert.deepEqual(errors,[]);
  const out=process.env.QA_OUTPUT||'/private/tmp/fulfillment-series-qa';await fs.mkdir(out,{recursive:true});await fs.writeFile(path.join(out,'系列对比-合成验收.xlsx'),saved.bytes);
  console.log('PASS: series comparison from four sources, original purchaser fields, first/prior/current totals, exact decimals, new/removed/zero groups, cache reload/duplicate/failure, division-only/review-only exports, numeric types/filter/styles and no formulas');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
