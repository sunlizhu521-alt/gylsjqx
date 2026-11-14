const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const root = path.resolve(__dirname,'..'), base = new URL(process.env.QA_URL || 'https://product-dimension.test/');
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const headers = ['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组','销售产品线','销售系列','销售产品分类','财务加权平均价','不含税结算价','结算价','场景分组','一级品类','二级分类','三级分类','型号','品牌','国家','销售区域','当前日期','上架时间（月）','生命周期','是否正常备货'];
const materialHeaders = ['创建组织','使用组织','编码','领星SKU','SKU','名称','数据状态','禁用状态','物料属性','已使用','物料分组','创建人','创建日期','备案产品型号','UPC'];
const required = ['销售产品线','销售系列','销售产品分类','场景分组','一级品类','二级分类','三级分类','销售区域','当前日期','上架时间（月）','生命周期'];
const editableRequired = required.slice(0,8), computed = required.slice(8);
const manual = {'销售产品线':'测试产品线','销售系列':'测试系列','销售产品分类':'全新品','场景分组':'测试场景','一级品类':'测试一级','二级分类':'测试二级','三级分类':'测试三级','销售区域':'中国'};
const originalRow = ['OLD',0,'保留SKU','原商品名称','2024/11/30','原分组','人工分类','','',0];
function historyRow(code,sku,model,returned=false) {
  return [code,sku,sku,(returned?'退回':'正常')+'轮椅国内','2025/01/01','轮椅组','轮椅产品线','测试1系列',returned?'全新换包装-九大产品线':'全新品',10,20,30,'九大产品线','轮椅','手动轮椅',returned?'二手商品':'低靠背',model,'测试品牌','中国','中国',46000,'21个月','稳定期','是'];
}
function conflictHistory(code,line) {
  const row=historyRow(code,'000SKU','CLASH');row[1]=0;row[5]='其他历史组';row[6]=line;row[7]='冲突系列';row[9]=row[10]=row[11]='';return row;
}
const dimensionRows = [originalRow,originalRow,['1','','','短编码'],[0,'','','零编码'],['C','','','冲突旧名称'],['C','','','冲突新名称','','','人工选择分类'],
  historyRow('BASE1','M-BASE1','MD-BASE1'),historyRow('RET1','RE-M-BASE1','RE-MD-BASE1',true),historyRow('BASE2','M-BASE2','MD-BASE2'),historyRow('RET2','RE-M-BASE2','RE-MD-BASE2',true),historyRow('BASE3','M-BASE3','MD-BASE3'),conflictHistory('CLASH1','冲突甲线'),conflictHistory('CLASH2','冲突乙线')];
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
function materialRow(code,name,organization='集团',sku=0,group='新分组',created='2026/11/14') {
  return ['集团',organization,code,sku,sku===0?'000SKU':sku,name,'已审核','否','自制','是',group,'测试',created,'不可自动映射型号',''];
}
const materialRows = [materialRow('OLD','物料表不要覆盖旧名称'),materialRow('0001','带前导零新增','集团','M-BASE1','轮椅组'),materialRow('0001','带前导零新增','子公司','M-BASE1','轮椅组'),materialRow('RETURN','第三款退回轮椅国内','集团','RE-M-BASE3','轮椅组'),materialRow('NEW','<img src=x onerror=alert(1)> & 新商品','集团',0,'轮椅组'),materialRow(0,'不要覆盖零编码')];
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
async function ready(scope,label='') {
  try{await scope.locator('#download:not([disabled])').waitFor({timeout:20000});}
  catch(error){throw Error(`${label} 下载未启用：${await scope.locator('#status').textContent()}；${await scope.locator('#pending-counts').textContent()}；${JSON.stringify(await scope.locator('#reviews input[aria-invalid="true"]').evaluateAll(inputs=>inputs.map(input=>({code:input.dataset.code,field:input.dataset.field,value:input.value,note:document.getElementById(input.getAttribute('aria-describedby'))?.textContent}))))}`,{cause:error});}
}
function field(scope,code,name) { return scope.locator(`#reviews input[data-code=${JSON.stringify(code)}][data-field=${JSON.stringify(name)}]`); }
async function setField(scope,code,name,value) {
  const card=scope.locator(`details.review[data-code=${JSON.stringify(code)}]`);
  if(await card.getAttribute('open')===null)await card.locator('summary').first().click();
  const input=field(scope,code,name);await input.fill(value);await input.dispatchEvent('change');
}
async function allReviews(scope) { await scope.locator('#review-filter').selectOption('all'); }
async function fillRequired(scope,code,values=manual) {
  await allReviews(scope);for(const name of editableRequired)await setField(scope,code,name,values[name]);
}
async function previewRow(scope,code) {
  return scope.locator('#preview tbody tr').evaluateAll((rows,key)=>rows.map(row=>Array.from(row.cells,cell=>cell.textContent)).find(values=>values[0]===key),code);
}
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
  const expected=['OLD','1','0','C','BASE1','RET1','BASE2','RET2','BASE3','CLASH1','CLASH2','0001','RETURN','NEW'];
  assert.equal(sheet['!ref'],'A1:X15');
  const rows=new Map();for(let r=2;r<=15;r++){const code=String(sheet['A'+r].v);assert(!rows.has(code),'每个编码只有一行');rows.set(code,r);}
  assert.deepEqual([...rows.keys()],expected);
  const original=X.read(dimensionFixture(),{type:'buffer',cellNF:true}).Sheets['Dim-YL医疗器械商品分类'],seen=new Set();
  dimensionRows.forEach((values,index)=>{
    const code=String(values[0]);if(seen.has(code))return;seen.add(code);const target=rows.get(code);
    for(let c=1;c<24;c++)assert.equal(sheet[X.utils.encode_col(c)+target]?.v??'',original[X.utils.encode_col(c)+(index+2)]?.v??'',code+' 原有 '+headers[c]+' 不覆盖');
  });
  const get=(code,name)=>sheet[X.utils.encode_col(headers.indexOf(name))+rows.get(code)];
  assert.equal(get('OLD','当前日期').v,42);assert.equal(get('OLD','当前日期').f,undefined,'主表公式转已保存结果值');
  assert.equal(get('0001','物料编码').t,'s');assert.equal(get('0001','领星SKU').v,'M-BASE1');assert.equal(get('0001','SKU').v,'M-BASE1');assert.equal(get('0001','金蝶名称').v,'带前导零新增');
  assert.equal(get('0001','金蝶创建时间').v,46340);assert.equal(get('0001','当前日期').v,46340);assert.equal(get('0001','上架时间（月）').v,'0个月');assert.equal(get('0001','生命周期').v,'新品期');
  assert.equal(get('0001','销售产品线').v,'轮椅产品线');assert.equal(get('0001','销售系列').v,'测试1系列');assert.equal(get('0001','销售产品分类').v,'全新品');
  assert.equal(get('RETURN','销售产品分类').v,'全新换包装-九大产品线');assert.equal(get('RETURN','三级分类').v,'二手商品');assert.equal(get('RETURN','型号').v,'RE-MD-BASE3');
  assert.equal(get('NEW','领星SKU').v,'0');assert.equal(get('NEW','SKU').v,'000SKU');assert.equal(get('NEW','销售系列').v,'手工指定系列');assert.equal(get('NEW','金蝶名称').v,'<img src=x onerror=alert(1)> & 新商品');
  for(const code of ['0001','RETURN','NEW'])for(const name of required){const cell=get(code,name);assert(cell&&cell.v!==''&&!/待确认|待补充|未知|^\/$/.test(String(cell.v)),code+' 必填 '+name);assert.equal(cell.f,undefined);}
  for(const code of ['0001','NEW'])for(const name of ['财务加权平均价','不含税结算价','结算价'])assert([undefined,null,''].includes(get(code,name)?.v),'新编码不继承历史价格');
  assert.equal(sheet['!autofilter'].ref,'A1:X15');assert(!(sheet['!rows']||[]).some(r=>r?.hidden),'所有记录可见');
  return book;
}
(async()=>{
  if((process.env.QA_DIMENSION_FILE||process.env.QA_MATERIAL_FILE)&&process.env.QA_URL&&!['localhost','127.0.0.1','[::1]'].includes(base.hostname))throw Error('真实文件仅允许本地验收');
  const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||(process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
  try{
    const context=await browser.newContext({acceptDownloads:true,viewport:{width:1440,height:1000}}),page=await context.newPage(),errors=[];
    page.setDefaultTimeout(30000);page.on('pageerror',e=>errors.push(e.message));
    await context.addInitScript(()=>{const RealDate=Date;window.Date=class extends RealDate{constructor(...args){super(...(args.length?args:['2026-11-14T04:00:00Z']));}static now(){return new RealDate('2026-11-14T04:00:00Z').getTime();}};});
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
    await upload(page,'material',material,'合成物料.xlsx');
    assert(await page.locator('#download').isDisabled(),'新增待补资料未齐全不能下载');
    assert.equal(await page.locator('#review-filter').inputValue(),'pending');
    assert.match(await page.locator('#pending-counts').innerText(),/1/);
    assert.equal(await page.locator('#preview th').count(),24,'完整24列预览');
    assert.match(await page.locator('#preview').innerText(),/0001/);assert.match(await page.locator('#preview').innerText(),/NEW/);assert.equal(await page.locator('#preview img').count(),0);
    assert.equal(await field(page,'NEW','销售系列').count(),1);assert.equal(await field(page,'0001','销售系列').count(),0,'默认仅显示待补商品');
    const auto=await previewRow(page,'0001'),returned=await previewRow(page,'RETURN');
    assert.equal(auto[6],'轮椅产品线');assert.equal(auto[7],'测试1系列');assert.equal(auto[8],'全新品');
    assert.equal(returned[8],'全新换包装-九大产品线');assert.equal(returned[15],'二手商品');assert.equal(returned[16],'RE-MD-BASE3');
    assert.match((await previewRow(page,'NEW'))[6],/^$|待确认/,'精确SKU冲突不得被宽泛同组产品线覆盖');
    await fillRequired(page,'NEW');await ready(page);
    for(const name of computed){
      assert.equal(await field(page,'0001',name).count(),0,name+'不可编辑');
      const output=page.locator(`#reviews output[data-code="0001"][data-field=${JSON.stringify(name)}]`);assert.equal(await output.count(),1);assert(await output.textContent(),name+'只读结果');
    }
    for(const name of editableRequired){
      const prior=await field(page,'0001',name).inputValue();await setField(page,'0001',name,'');assert(await page.locator('#download').isDisabled(),name+'空白阻断');
      await setField(page,'0001',name,'待确认');assert(await page.locator('#download').isDisabled(),name+'占位值阻断');
      if(name==='销售区域'){await setField(page,'0001',name,'无法区分');assert(await page.locator('#download').isDisabled(),'销售区域不能填无法区分');}
      await setField(page,'0001',name,prior);await ready(page,'恢复'+name+'='+prior);
    }
    await setField(page,'NEW','销售系列','手工指定系列');await setField(page,'NEW','金蝶创建时间','2026-11-01');await ready(page);
    assert.equal(await field(page,'NEW','销售系列').inputValue(),'手工指定系列','重算不覆盖手工值');
    assert.equal((await previewRow(page,'NEW'))[7],'手工指定系列');
    await upload(page,'material',material,'替换物料.xlsx');assert(await page.locator('#download').isDisabled(),'替换物料清除旧补录');
    await allReviews(page);assert.notEqual(await field(page,'NEW','销售系列').inputValue(),'手工指定系列');
    await fillRequired(page,'NEW');await setField(page,'NEW','销售系列','手工指定系列');await ready(page);
    await upload(page,'dimension',original,'合成维度.xlsx');assert(await page.locator('#download').isDisabled(),'替换维度清除旧补录');
    await fillRequired(page,'NEW');await setField(page,'NEW','销售系列','手工指定系列');await ready(page);
    await page.locator('#issue-title').click();assert.match(await page.locator('#issues').innerText(),/C/,'首行冲突仍提示来源');
    const bytes=await download(page,page,1);assertOutput(bytes);await preserved(original,bytes);
    if(process.env.QA_SCREENSHOT)await page.screenshot({path:process.env.QA_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:390,height:844});const dimensions=await page.locator('body').evaluate(body=>({width:body.clientWidth,scroll:body.scrollWidth}));assert(dimensions.scroll<=dimensions.width+2,'窄屏仅表格横向滚动');
    if(process.env.QA_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.QA_MOBILE_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:1440,height:1000});
    await upload(page,'dimension',bytes,'合成维度_V1.xlsx');await ready(page);assertOutput(await download(page,page,2));
    assert.equal(await page.locator('#reviews input').count(),0,'重新上传已补全结果无需重复补录');
    for(const slot of ['material','dimension']){
      await upload(page,slot,Buffer.from('broken'),'损坏.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#status').innerText(),/失败|异常|有效|格式/);
      await upload(page,slot,slot==='material'?material:bytes,slot==='material'?'合成物料.xlsx':'合成维度.xlsx');await ready(page);
    }
    await upload(page,'material',materialFixture([materialRow('','缺失编码')]),'缺编码.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/编码/);
    await upload(page,'material',await materialFormulaFixture(''),'缺公式计算结果.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/公式|缓存|计算结果/);assert.match(await page.locator('#issues').innerText(),/F2/);
    const aliasZip=await Zip.loadAsync(await materialFormulaFixture(''));
    aliasZip.file('xl/workbook.xml',(await aliasZip.file('xl/workbook.xml').async('string')).replace(/xmlns:r=/g,'xmlns:rel=').replace(/ r:id=/g,' rel:id='));
    await upload(page,'material',await aliasZip.generateAsync({type:'nodebuffer'}),'关系前缀不同且缺公式结果.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/F2/);
    await upload(page,'material',await materialFormulaFixture('<v></v>'),'公式结果为空字符串.xlsx');assert.doesNotMatch(await page.locator('#issues').textContent(),/没有计算结果|缺少.*结果/);await fillRequired(page,'FORMULA');await ready(page);
    await upload(page,'material',materialFixture([materialRow('CONFLICT','名称甲'),materialRow('CONFLICT','名称乙')]),'冲突物料.xlsx');assert(await page.locator('#download').isDisabled());assert.match(await page.locator('#issues').innerText(),/冲突|不同/);
    await upload(page,'material',material,'合成物料.xlsx');await ready(page);
    await upload(page,'material',materialFixture([materialRow('FUTURE','未来创建国内','集团','M-BASE1','轮椅组','2026/12/01')]),'未来日期.xlsx');
    assert(await page.locator('#download').isDisabled(),'未来创建日期不得计算月龄');await allReviews(page);
    assert.match(await page.locator('#pending-counts').innerText(),/[1-9]/);await setField(page,'FUTURE','金蝶创建时间','日期不详');assert(await page.locator('#download').isDisabled());
    await setField(page,'FUTURE','金蝶创建时间','');assert(await page.locator('#download').isDisabled());
    await setField(page,'FUTURE','金蝶创建时间','2026-04-14');await ready(page);const repaired=await previewRow(page,'FUTURE');assert.equal(repaired[21],'7个月');assert.equal(repaired[22],'稳定期');
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
      await upload(page,'dimension',fs.readFileSync(process.env.QA_DIMENSION_FILE),path.basename(process.env.QA_DIMENSION_FILE));await upload(page,'material',fs.readFileSync(process.env.QA_MATERIAL_FILE),path.basename(process.env.QA_MATERIAL_FILE));
      assert(await page.locator('#download').isDisabled(),'真实新增商品未完成手工核对前不能下载');
      assert.match(await page.locator('#counts').innerText(),/43/);assert.match(await page.locator('#counts').innerText(),/23,718/);assert.match(await page.locator('#pending-counts').innerText(),/[1-9]/);
      assert.equal(await page.locator('#preview tbody tr').count(),43);
      const codes=await page.locator('#preview tbody tr').evaluateAll(rows=>rows.map(row=>row.cells[0].textContent));assert.equal(new Set(codes).size,43);assert(codes.includes('000000001'));
    }
    assert.deepEqual(errors,[]);console.log('PASS: 商品维度自动补全、退回学习、11必填/手工补录、覆盖保留与替换清理、未来日期、24列预览、前导零/零值、只追加、导出回读、辅助表保留、公式缓存、版本/重复导入、布局及首页切换'+(process.env.QA_DIMENSION_FILE?'、真实样表本地验收':''));
  }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
