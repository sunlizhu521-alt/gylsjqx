const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const root = path.resolve(__dirname,'..'), base = new URL(process.env.QA_URL || 'https://product-dimension.test/');
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const headers = ['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组','销售产品线','销售系列','销售产品分类','财务加权平均价','不含税结算价','结算价','场景分组','一级品类','二级分类','三级分类','型号','品牌','国家','销售区域','当前日期','上架时间（月）','生命周期','是否正常备货'];
const materialHeaders = ['创建组织','使用组织','编码','领星SKU','SKU','名称','数据状态','禁用状态','物料属性','已使用','物料分组','创建人','创建日期','备案产品型号','UPC'];
const originalRow = ['OLD',0,'保留SKU','原商品名称','2024/11/30','原分组','人工分类','','',0];
const dimensionRows = [originalRow,originalRow,['1','','','短编码'],[0,'','','零编码'],['C','','','冲突旧名称'],['C','','','冲突新名称','','','人工选择分类']];
function dimensionFixture(rows=dimensionRows) {
  const book = X.utils.book_new(), sheet = X.utils.aoa_to_sheet([headers,...rows]);
  sheet.E2.z='m/d/yy'; sheet.E3.z='m/d/yy';
  sheet.U2={t:'n',f:'10+32',v:42};sheet.U3={t:'n',f:'10+32',v:42};
  sheet['!autofilter']={ref:`A1:X${rows.length+1}`}; sheet['!rows']=[{}, {hidden:true}];
  X.utils.book_append_sheet(book,sheet,'Dim-YL医疗器械商品分类');
  const other = X.utils.aoa_to_sheet([['销售系列','结算价'],['人工辅助',12]]);
  other.C2={t:'n',f:'B2*2',v:24,c:[{a:'测试',t:'原辅助表批注'}]}; other['!ref']='A1:C2';
  X.utils.book_append_sheet(book,other,'Sheet1');
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
function materialRow(code,name,organization='集团',sku=0) {
  return ['集团',organization,code,sku,'000SKU',name,'已审核','否','自制','是','新分组','测试','2026/11/14','不可自动映射型号',''];
}
const materialRows = [materialRow('OLD','物料表不要覆盖旧名称'),materialRow('0001','带前导零新增'),materialRow('0001','带前导零新增','子公司'),materialRow('NEW','<img src=x onerror=alert(1)> & 新商品'),materialRow(0,'不要覆盖零编码')];
function materialFixture(rows=materialRows) {
  const book=X.utils.book_new(); X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([materialHeaders,...rows]),'Sheet1');
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
async function materialFormulaFixture(cacheXml) {
  const zip=await Zip.loadAsync(materialFixture([materialRow('FORMULA','待替换公式')])),name='xl/worksheets/sheet1.xml';
  const xml=await zip.file(name).async('string');
  const next=xml.replace(/<c\b[^>]*\br="F2"[^>]*>[\s\S]*?<\/c>/,'<c r="F2" t="str"><f>IF(1=1,&quot;&quot;,&quot;名称&quot;)</f>'+cacheXml+'</c>');
  assert.notEqual(next,xml,'合成缺缓存公式已写入物料名称 F2');zip.file(name,next);
  return zip.generateAsync({type:'nodebuffer'});
}
async function upload(scope,slot,buffer,name) {
  await scope.locator(`#${slot}File:not([disabled])`).setInputFiles({name,mimeType:mime,buffer});
  await scope.locator(`#${slot}File:not([disabled])`).waitFor({timeout:120000});
}
async function ready(scope) { await scope.locator('#download:not([disabled])').waitFor({timeout:120000}); }
async function download(page,scope,version) {
  const pending=page.waitForEvent('download',{timeout:120000}); await scope.locator('#download').click();
  const file=await pending, chunks=[];
  for await(const chunk of await file.createReadStream())chunks.push(chunk);
  assert.match(file.suggestedFilename(),/_V[1-9]\d*\.xlsx$/);
  if(version)assert.match(file.suggestedFilename(),new RegExp('_V'+version+'\\.xlsx$'));
  return Buffer.concat(chunks);
}
async function preserved(before,after) {
  const a=await Zip.loadAsync(before),b=await Zip.loadAsync(after);
  // Only the auxiliary sheet's own relationship graph stays byte-identical;
  // main-sheet comment/VML addresses move when duplicate rows are removed.
  const pending=['xl/worksheets/sheet2.xml'],seen=new Set();
  while(pending.length){
    const name=pending.pop();if(seen.has(name))continue;seen.add(name);
    assert(a.file(name),name);assert(b.file(name),name);
    assert.equal(await b.file(name).async('base64'),await a.file(name).async('base64'),name+' must stay unchanged');
    const rels=path.posix.join(path.posix.dirname(name),'_rels',path.posix.basename(name)+'.rels');
    if(!a.file(rels))continue;
    assert(b.file(rels),rels);const xml=await a.file(rels).async('string');
    assert.equal(await b.file(rels).async('string'),xml,rels+' must stay unchanged');
    for(const match of xml.matchAll(/<Relationship\b[^>]*\/?\s*>/g)){
      if(/\bTargetMode="External"/.test(match[0]))continue;
      const target=/\bTarget="([^"]+)"/.exec(match[0])?.[1];if(!target)continue;
      const decoded=target.replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>');
      pending.push(decoded.startsWith('/')?decoded.slice(1):path.posix.normalize(path.posix.join(path.posix.dirname(name),decoded)));
    }
  }
}
function assertOutput(bytes) {
  const book=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),sheet=book.Sheets['Dim-YL医疗器械商品分类'];
  assert.deepEqual(book.SheetNames,['Dim-YL医疗器械商品分类','Sheet1']);
  assert.equal(sheet['!ref'],'A1:X7');
  const rows=new Map(); for(let r=2;r<=7;r++) {const code=String(sheet['A'+r].v);assert(!rows.has(code),'每个编码只有一行');rows.set(code,r);}
  assert.deepEqual([...rows.keys()],['OLD','1','0','C','0001','NEW']);
  const old=rows.get('OLD');assert.equal(sheet['D'+old].v,'原商品名称');assert.equal(sheet['E'+old].v,'2024/11/30');assert.equal(sheet['G'+old].v,'人工分类');assert.equal(sheet['J'+old].v,0);
  assert.equal(sheet['D'+rows.get('C')].v,'冲突旧名称');assert([undefined,null,''].includes(sheet['G'+rows.get('C')]?.v),'冲突统一保留首行，不拼接后行内容');
  assert.equal(sheet['U'+old].v,42);assert.equal(sheet['U'+old].f,undefined,'主表公式转已保存结果值');
  const added=rows.get('0001');assert.equal(sheet['A'+added].t,'s');assert.equal(sheet['B'+added].v,'0');assert.equal(sheet['C'+added].v,'000SKU');assert.equal(sheet['D'+added].v,'带前导零新增');assert.equal(sheet['E'+added].v,'2026/11/14');assert.equal(sheet['F'+added].v,'新分组');
  for(let c=6;c<24;c++)assert([undefined,null,''].includes(sheet[X.utils.encode_col(c)+added]?.v),'新增其他字段留空');
  assert.equal(sheet['D'+rows.get('NEW')].v,'<img src=x onerror=alert(1)> & 新商品');
  assert.equal(sheet['!autofilter'].ref,'A1:X7');assert(!(sheet['!rows']||[]).some(r=>r?.hidden),'所有记录可见');
  return book;
}
(async()=>{
  if((process.env.QA_DIMENSION_FILE||process.env.QA_MATERIAL_FILE)&&process.env.QA_URL&&!['localhost','127.0.0.1','[::1]'].includes(base.hostname))throw Error('真实文件仅允许本地验收');
  const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
  try{
    const context=await browser.newContext({acceptDownloads:true,viewport:{width:1440,height:1000}}),page=await context.newPage(),errors=[];
    page.setDefaultTimeout(30000);page.on('pageerror',e=>errors.push(e.message));
    await context.route('**/*',route=>{
      const req=route.request(),url=new URL(req.url());
      if(url.origin!==base.origin||!['GET','HEAD'].includes(req.method()))return route.abort();
      if(process.env.QA_URL)return route.continue();
      const relative=decodeURIComponent(url.pathname).slice(1)||'index.html';if(relative.split('/').includes('..'))return route.abort();
      const file=path.join(root,relative);if(!fs.existsSync(file)||!fs.statSync(file).isFile())return route.fulfill({status:404,body:''});
      return route.fulfill({contentType:file.endsWith('.html')?'text/html; charset=utf-8':file.endsWith('.css')?'text/css':'application/javascript',body:fs.readFileSync(file)});
    });
    await page.goto(new URL('product-dimension.html',base).href,{waitUntil:'networkidle'});
    assert.equal(await page.locator('h1').innerText(),'维护商品维度表');assert(await page.locator('#download').isDisabled());
    const original=dimensionFixture(),material=materialFixture();
    await upload(page,'dimension',original,'合成维度.xlsx');assert(await page.locator('#download').isDisabled());
    await upload(page,'material',material,'合成物料.xlsx');await ready(page);
    await page.locator('#issue-title').click();
    assert.match(await page.locator('#issues').innerText(),/C/,'冲突保留首行仍提示来源');
    assert.match(await page.locator('#preview').innerText(),/0001/);assert.match(await page.locator('#preview').innerText(),/NEW/);assert.equal(await page.locator('#preview img').count(),0);
    const bytes=await download(page,page,1);assertOutput(bytes);await preserved(original,bytes);
    if(process.env.QA_SCREENSHOT)await page.screenshot({path:process.env.QA_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:390,height:844});const dimensions=await page.locator('body').evaluate(body=>({width:body.clientWidth,scroll:body.scrollWidth}));assert(dimensions.scroll<=dimensions.width+2,'窄屏仅表格横向滚动');
    if(process.env.QA_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.QA_MOBILE_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:1440,height:1000});
    await upload(page,'dimension',bytes,'合成维度_V1.xlsx');await ready(page);assertOutput(await download(page,page,2));
    assert.equal(await page.locator('select').count(),0,'统一保留首行，无需选择');
    for(const slot of ['material','dimension']){
      await upload(page,slot,Buffer.from('broken'),'损坏.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#status').innerText(),/失败|异常|有效|格式/);
      await upload(page,slot,slot==='material'?material:bytes,slot==='material'?'合成物料.xlsx':'合成维度.xlsx');await ready(page);
    }
    await upload(page,'material',materialFixture([materialRow('','缺失编码')]),'缺编码.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/编码/);
    await upload(page,'material',await materialFormulaFixture(''),'缺公式计算结果.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/公式|缓存|计算结果/);assert.match(await page.locator('#issues').innerText(),/F2/);
    const aliasZip=await Zip.loadAsync(await materialFormulaFixture(''));
    aliasZip.file('xl/workbook.xml',(await aliasZip.file('xl/workbook.xml').async('string')).replace(/xmlns:r=/g,'xmlns:rel=').replace(/ r:id=/g,' rel:id='));
    await upload(page,'material',await aliasZip.generateAsync({type:'nodebuffer'}),'关系前缀不同且缺公式结果.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/F2/);
    await upload(page,'material',await materialFormulaFixture('<v></v>'),'公式结果为空字符串.xlsx');await ready(page);
    await upload(page,'material',materialFixture([materialRow('CONFLICT','名称甲'),materialRow('CONFLICT','名称乙')]),'冲突物料.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/冲突|不同/);
    await upload(page,'material',material,'合成物料.xlsx');await ready(page);
    await page.evaluate(()=>{window.__realDimensionExport=window.ProductDimensionExport;window.ProductDimensionExport=async()=>{throw Error('测试生成失败');};});
    await page.locator('#download').click();await page.locator('#status').filter({hasText:'测试生成失败'}).waitFor();
    await page.evaluate(()=>window.ProductDimensionExport=window.__realDimensionExport);await download(page,page,3);
    await page.reload({waitUntil:'networkidle'});await upload(page,'dimension',bytes,'合成维度.xlsx');await upload(page,'material',material,'合成物料.xlsx');await ready(page);await download(page,page,4);
    await page.locator('#clear').click();assert(await page.locator('#download').isDisabled());assert.equal(await page.locator('#dimensionFile').inputValue(),'');assert.equal(await page.locator('#materialFile').inputValue(),'');
    await page.goto(new URL('index.html',base).href,{waitUntil:'networkidle'});await page.locator('button[data-tool="product-dimension"]').click();
    const frame=page.frameLocator('iframe[data-tool="product-dimension"]');await upload(frame,'dimension',bytes,'合成维度.xlsx');await upload(frame,'material',material,'合成物料.xlsx');await ready(frame);
    await page.locator('button[data-tool="funds"]').click();await page.locator('button[data-tool="product-dimension"]').click();await ready(frame);assert.match(await frame.locator('#dimensionReference').innerText(),/合成维度.xlsx/);assertOutput(await download(page,frame,5));
    if(process.env.QA_DIMENSION_FILE&&process.env.QA_MATERIAL_FILE){
      await page.goto(new URL('product-dimension.html',base).href,{waitUntil:'networkidle'});
      const source=fs.readFileSync(process.env.QA_DIMENSION_FILE),materials=fs.readFileSync(process.env.QA_MATERIAL_FILE);
      await upload(page,'dimension',source,path.basename(process.env.QA_DIMENSION_FILE));await upload(page,'material',materials,path.basename(process.env.QA_MATERIAL_FILE));
      await ready(page);
      const output=await download(page,page),book=X.read(output,{type:'buffer',cellNF:true}),sheet=book.Sheets['Dim-YL医疗器械商品分类'];
      assert.equal(sheet['!ref'],'A1:X23719');const codes=[];for(let r=2;r<=23719;r++)codes.push(String(sheet['A'+r].v));assert.equal(new Set(codes).size,23718);assert(codes.includes('000000001'));await preserved(source,output);
      if(process.env.QA_SAMPLE_OUTPUT)fs.writeFileSync(process.env.QA_SAMPLE_OUTPUT,output);
    }
    assert.deepEqual(errors,[]);console.log('PASS: 商品维度双槽上传、重复与冲突、前导零/零值、只追加不覆盖、导出回读、辅助表保留、失败清理、版本递增、重复导入、布局及首页切换'+(process.env.QA_DIMENSION_FILE?'、真实样表本地验收':''));
  }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
