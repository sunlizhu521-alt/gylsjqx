const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const root = path.resolve(__dirname,'..'), base = new URL(process.env.QA_URL || 'https://negotiated-summary.test/');
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function fixture(rows) {
  const headers = ['供应商','付款主体','是否谈判分批','2027年1月支付','2026年11月支付','2026年9月支付','OA-申请付款金额'];
  const book = X.utils.book_new(); X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'谈判付款');
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['其他工作表不读取'],[999999]]),'正常付款');
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
const original = fixture([
  ['供应商甲','主体甲','是',30,20,10,999999],['供应商甲','主体乙','是',null,20,0,999999],
  ['供应商甲','主体甲','否',5,null,null,999999],['供应商乙','主体甲','',null,180000,null,999999]
]);
async function upload(scope,bytes=original,name='合成付款节奏.xlsx') {
  await scope.locator('#sourceFile:not([disabled])').setInputFiles({name,mimeType:mime,buffer:bytes});
  await scope.locator('#sourceFile:not([disabled])').waitFor();
}
async function download(page,scope) {
  const pending=page.waitForEvent('download'); await scope.locator('#download').click();
  const file=await pending,chunks=[]; for await(const chunk of await file.createReadStream()) chunks.push(chunk);
  assert.match(file.suggestedFilename(),/^谈判付款整理-\d{4}-\d{2}-\d{2}\.xlsx$/);
  return Buffer.concat(chunks);
}
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH || (process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
  try {
    const context=await browser.newContext({acceptDownloads:true,viewport:{width:1440,height:1000}}),page=await context.newPage(),errors=[];
    page.on('pageerror',e=>errors.push(e.message));
    await context.route('**/*',route=>{
      const req=route.request(),url=new URL(req.url());
      if(url.origin!==base.origin || !['GET','HEAD'].includes(req.method()))return route.abort();
      if(process.env.QA_URL)return route.continue();
      const relative=decodeURIComponent(url.pathname).slice(1)||'index.html';
      if(relative.split('/').includes('..'))return route.abort();
      const file=path.join(root,relative);
      if(!fs.existsSync(file)||!fs.statSync(file).isFile())return route.fulfill({status:404,body:''});
      return route.fulfill({contentType:file.endsWith('.html')?'text/html; charset=utf-8':file.endsWith('.css')?'text/css':'application/javascript',body:fs.readFileSync(file)});
    });
    await page.goto(new URL('index.html',base).href,{waitUntil:'networkidle'});
    await page.locator('.tool-tab[data-tool="negotiated-payment-summary"]').click();
    const frame=page.frameLocator('iframe[data-tool="negotiated-payment-summary"]');
    await upload(frame);
    await frame.locator('#review-panel:not([hidden])').waitFor();
    assert.equal(await frame.locator('#download').isDisabled(),true);
    assert.match(await frame.locator('#counts').innerText(),/4 条明细，2 家供应商/);
    assert.match(await frame.locator('#reference').innerText(),/合成付款节奏.xlsx\n引用时间：/);
    assert.match(await frame.locator('#totals').innerText(),/谈判付款 ￥80.00；正常付款 ￥5.00；待分类 ￥180,000.00/);
    await frame.locator('select[data-address="C5"]').selectOption('否');
    await frame.locator('#download:not([disabled])').waitFor();
    const preview=await frame.locator('#preview').innerText();
    assert.match(await frame.locator('#totals').innerText(),/正常付款 ￥180,005.00/);
    await page.locator('.tool-tab[data-tool="funds-budget-summary"]').click();
    await page.locator('.tool-tab[data-tool="negotiated-payment-summary"]').click();
    assert.equal(await frame.locator('#preview').innerText(),preview);
    assert.equal(await frame.locator('select[data-address="C5"]').inputValue(),'否');
    const bytes=await download(page,frame),book=X.read(bytes,{type:'buffer',cellNF:true}),s=book.Sheets.谈判付款整理;
    assert.deepEqual(book.SheetNames,['谈判付款整理']); assert.equal(s['!ref'],'A1:I4');
    assert.equal(s.B3.v,80); assert.equal(s.B3.f,'SUM(C3:E3)');assert.equal(s.F3.v,5);assert.equal(s.H4.v,180000);
    assert.equal(X.SSF.format(s.C2.z,s.C2.v),'2026年9月');assert.equal(X.SSF.format(s.E2.z,s.E2.v),'2027年1月');
    await frame.locator('select[data-address="C5"]').selectOption(''); assert(await frame.locator('#download').isDisabled());
    await frame.locator('select[data-address="C5"]').selectOption('是'); assert.match(await frame.locator('#totals').innerText(),/谈判付款 ￥180,080.00/);
    await upload(frame); assert.equal(await frame.locator('select[data-address="C5"]').inputValue(),'','新上传不沿用旧分类');
    assert(await frame.locator('#download').isDisabled());
    await upload(frame,fixture([['<img src=x onerror=alert(1)>','','是',0,'坏金额',null,0]]),'异常.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/异常.xlsx.*E2/);assert.equal(await frame.locator('#preview img').count(),0);
    // An uncached string formula must not disappear into a blank/zero on browser import.
    const brokenFormula=await Zip.loadAsync(fixture([['供应商甲','','是',null,10,null,0]]));
    let xml=await brokenFormula.file('xl/worksheets/sheet1.xml').async('string');
    xml=xml.replace(/<c\b[^>]*\br="E2"[^>]*>[\s\S]*?<\/c>/,'<c r="E2" t="str"><f>1+1</f></c>');
    brokenFormula.file('xl/worksheets/sheet1.xml',xml);
    await upload(frame,await brokenFormula.generateAsync({type:'nodebuffer'}),'缺公式结果.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/E2.*公式/);
    await upload(frame,Buffer.from('broken'),'损坏.xlsx'); await frame.locator('#status').filter({hasText:'读取失败'}).waitFor();
    assert(await frame.locator('#download').isDisabled()); assert(await frame.locator('#result').isHidden());assert(await frame.locator('#review-panel').isHidden());
    await upload(frame);await frame.locator('select[data-address="C5"]').selectOption('否');
    if(process.env.QA_SCREENSHOT)await page.screenshot({path:process.env.QA_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:390,height:844});
    const dimensions=await frame.locator('body').evaluate(body=>({width:body.clientWidth,scroll:body.scrollWidth}));
    assert(dimensions.scroll<=dimensions.width+2,'窄屏仅表格区域横向滚动');
    if(process.env.QA_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.QA_MOBILE_SCREENSHOT,fullPage:true});
    await frame.locator('#clear').click();assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#result').isHidden());
    await page.goto(new URL('negotiated-payment-summary.html',base).href,{waitUntil:'networkidle'});
    await upload(page,fixture([['供应商乙','','否',5,null,null,0]]),'新文件.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    assert.match(await page.locator('#counts').innerText(),/1 条明细，1 家供应商/);assert(await page.locator('#review-panel').isHidden());
    const replaced=X.read(await download(page,page),{type:'buffer'}).Sheets.谈判付款整理;
    assert.equal(replaced.A3.v,'供应商乙');assert.equal(replaced.F3.v,5);
    // A failed or corrupted export cannot start a download.
    await page.evaluate(()=>{window.__realExport=window.NegotiatedPaymentSummaryExport;window.NegotiatedPaymentSummaryExport=async()=>new Uint8Array([0]);});
    await page.locator('#download').click();await page.locator('#status').filter({hasText:'导出失败'}).waitFor();
    await page.evaluate(()=>window.NegotiatedPaymentSummaryExport=window.__realExport);
    await download(page,page);
    if(process.env.QA_SAMPLE_FILE){
      await upload(page,fs.readFileSync(process.env.QA_SAMPLE_FILE),path.basename(process.env.QA_SAMPLE_FILE));
      await page.locator('select[data-address="G40"]').waitFor();
      assert.match(await page.locator('#counts').innerText(),/119 条明细，47 家供应商/);
      assert.match(await page.locator('#totals').innerText(),/34,628,537.65.*10,636,995.03.*180,000.00/);
      assert(await page.locator('#download').isDisabled());
      // A temporary test choice exercises export; it is never saved into the user's source workbook.
      await page.locator('select[data-address="G40"]').selectOption('否');
      const sample=await download(page,page),out=X.read(sample,{type:'buffer',cellNF:true}).Sheets.谈判付款整理;
      assert.equal(out['!ref'],'A1:S49');
      let negotiated=0,normal=0;for(let r=3;r<=49;r++){negotiated+=Math.round(out['B'+r].v*100);normal+=Math.round(out['K'+r].v*100);}
      assert.equal(negotiated,3462853765);assert.equal(normal,1081699503);
      if(process.env.QA_SAMPLE_OUTPUT)fs.writeFileSync(process.env.QA_SAMPLE_OUTPUT,sample);
    }
    assert.deepEqual(errors,[]);
    console.log('PASS: 首页与独立页、分类核对、跨年月份、切换保留、替换重置、错误清理、公式缓存保护、预览下载一致、手机布局'+(process.env.QA_SAMPLE_FILE?'、真实样表':'') );
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
