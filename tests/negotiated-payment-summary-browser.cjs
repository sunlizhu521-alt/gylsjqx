const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), Zip = require('../vendor/jszip.min.js');
const root = path.resolve(__dirname,'..'), base = new URL(process.env.QA_URL || 'https://negotiated-summary.test/');
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function fixture(rows) {
  const headers = ['供应商','付款主体','是否谈判分批','谈判明细','2027年1月支付','2026年11月支付','2026年9月支付','OA-申请付款金额'];
  const book = X.utils.book_new(); X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'谈判付款');
  X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['其他工作表不读取'],[999999]]),'正常付款');
  return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
}
const original = fixture([
  ['供应商甲','主体甲','是','1、9月30日－10\n2、11月25日－20\n3、1月25日－30',999,999,999,999999],
  ['供应商甲','主体乙','是','1、9月30日－0\n2、11月26日－20',999,999,999,999999],
  ['供应商甲','主体甲','否','1月25日－5',999,999,999,999999],
  ['供应商乙','主体甲','','11月25日－180000',999,999,999,999999]
]);
async function upload(scope,bytes=original,name='合成付款节奏.xlsx') {
  await scope.locator('#sourceFile:not([disabled])').setInputFiles({name,mimeType:mime,buffer:bytes});
  await scope.locator('#sourceFile:not([disabled])').waitFor();
}
async function download(page,scope,version,date='2026-11-12') {
  const pending=page.waitForEvent('download'); await scope.locator('#download').click();
  const file=await pending,chunks=[]; for await(const chunk of await file.createReadStream()) chunks.push(chunk);
  assert.match(file.suggestedFilename(),/^谈判付款整理-\d{4}-\d{2}-\d{2}-V[1-9]\d*\.xlsx$/);
  if(version)assert.equal(file.suggestedFilename(),`谈判付款整理-${date}-V${version}.xlsx`);
  return Buffer.concat(chunks);
}
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH || (process.platform==='darwin'?'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome':undefined)});
  try {
    const context=await browser.newContext({acceptDownloads:true,viewport:{width:1440,height:1000}}),page=await context.newPage(),errors=[];
    await context.addInitScript(()=>{
      const RealDate=Date;
      window.Date=class extends RealDate {
        constructor(...args){super(...(args.length?args:[window.__negotiatedSummaryTestNow || '2026-11-12T02:00:00Z']));}
        static now(){return new RealDate(window.__negotiatedSummaryTestNow || '2026-11-12T02:00:00Z').getTime();}
      };
    });
    let downloadCount=0;page.on('download',()=>downloadCount++);
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
    const bytes=await download(page,frame,1),book=X.read(bytes,{type:'buffer',cellNF:true}),s=book.Sheets.谈判付款整理;
    assert.deepEqual(book.SheetNames,['谈判付款整理']); assert.equal(s['!ref'],'A1:K4');
    assert.equal(s.B3.v,80); assert.equal(s.B3.f,'SUM(C3:F3)');assert.equal(s.G3.v,5);assert.equal(s.I4.v,180000);
    assert.equal(s.C2.t,'s');assert.equal(s.C2.z,'@');assert.equal(s.C2.v,'2026年9月30日');assert.equal(X.SSF.format(s.F2.z,s.F2.v),'2027年1月25日');
    assert.deepEqual(['C','D','E','F'].map(c=>s[c+'2'].v),['2026年9月30日','2026年11月25日','2026年11月26日','2027年1月25日']);
    assert.deepEqual(['C','D','E','F'].map(c=>s[c+'3'].v),[10,20,20,30]);
    assert.equal(s['!autofilter'].ref,'A2:K4');
    for(const date of ['2026年11月25日','2026年11月26日'])assert(preview.includes(date));
    await frame.locator('select[data-address="C5"]').selectOption(''); assert(await frame.locator('#download').isDisabled());
    await frame.locator('select[data-address="C5"]').selectOption('是'); assert.match(await frame.locator('#totals').innerText(),/谈判付款 ￥180,080.00/);
    await upload(frame); assert.equal(await frame.locator('select[data-address="C5"]').inputValue(),'','新上传不沿用旧分类');
    assert(await frame.locator('#download').isDisabled());
    await upload(frame,fixture([['<img src=x onerror=alert(1)>','','是','11月25日－坏金额']]),'异常.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/异常.xlsx.*D2/);assert.equal(await frame.locator('#preview img').count(),0);
    // An uncached string formula must not disappear into a blank/zero on browser import.
    const brokenFormula=await Zip.loadAsync(fixture([['供应商甲','','是','11月25日－10']]));
    let xml=await brokenFormula.file('xl/worksheets/sheet1.xml').async('string');
    xml=xml.replace(/<c\b[^>]*\br="D2"[^>]*>[\s\S]*?<\/c>/,'<c r="D2" t="str"><f>1+1</f></c>');
    brokenFormula.file('xl/worksheets/sheet1.xml',xml);
    await upload(frame,await brokenFormula.generateAsync({type:'nodebuffer'}),'缺公式结果.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/D2.*公式/);
    await upload(frame,fixture([['供应商甲','','是','11月25日－10\n11月26日－20'],['供应商乙','','否','']]),'缺明细.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/D3.*(?:空|缺)/);
    await upload(frame,Buffer.from('broken'),'损坏.xlsx'); await frame.locator('#status').filter({hasText:'读取失败'}).waitFor();
    assert(await frame.locator('#download').isDisabled()); assert(await frame.locator('#result').isHidden());assert(await frame.locator('#review-panel').isHidden());
    await upload(frame);await frame.locator('select[data-address="C5"]').selectOption('否');
    if(process.env.QA_SCREENSHOT)await page.screenshot({path:process.env.QA_SCREENSHOT,fullPage:true});
    await page.setViewportSize({width:390,height:844});
    const dimensions=await frame.locator('body').evaluate(body=>({width:body.clientWidth,scroll:body.scrollWidth}));
    assert(dimensions.scroll<=dimensions.width+2,'窄屏仅表格区域横向滚动');
    if(process.env.QA_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.QA_MOBILE_SCREENSHOT,fullPage:true});
    await frame.locator('#clear').click();assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#result').isHidden());
    // A real date format must survive browser import; a plain numeric header is not year evidence.
    const dateBook=X.utils.book_new(),dateSheet=X.utils.aoa_to_sheet([
      ['供应商','付款主体','是否谈判分批','谈判明细',46356],['日期供应商','','是','11月25日－100']
    ]);
    dateSheet.E1.z='yyyy"年"m"月"d"日"';X.utils.book_append_sheet(dateBook,dateSheet,'谈判付款');
    await upload(frame,Buffer.from(X.write(dateBook,{type:'buffer',bookType:'xlsx'})),'日期表头.xlsx');
    await frame.locator('#download:not([disabled])').waitFor();
    assert.match(await frame.locator('#preview').innerText(),/2026年11月25日/);
    dateSheet.E1.z='General';
    await upload(frame,Buffer.from(X.write(dateBook,{type:'buffer',bookType:'xlsx'})),'普通数字表头.xlsx');
    await frame.locator('#status').filter({hasText:'修正来源问题'}).waitFor();assert(await frame.locator('#download').isDisabled());
    assert.match(await frame.locator('#issues').innerText(),/D2.*年份/);
    await page.goto(new URL('negotiated-payment-summary.html',base).href,{waitUntil:'networkidle'});
    await upload(page,fixture([['供应商乙','','否','1月25日－5']]),'新文件.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    assert.match(await page.locator('#counts').innerText(),/1 条明细，1 家供应商/);assert(await page.locator('#review-panel').isHidden());
    const replaced=X.read(await download(page,page,2),{type:'buffer'}).Sheets.谈判付款整理;
    assert.equal(replaced.A3.v,'供应商乙');assert.equal(replaced.D3.v,5);assert.equal(replaced.E2.v,'2027年1月25日');
    // A failed or corrupted export cannot start a download.
    const beforeFailure=downloadCount;
    await page.evaluate(()=>{window.__realExport=window.NegotiatedPaymentSummaryExport;window.NegotiatedPaymentSummaryExport=async()=>{throw Error('测试生成失败');};});
    await page.locator('#download').click();await page.locator('#status').filter({hasText:'测试生成失败'}).waitFor();
    assert.equal(await page.evaluate(()=>JSON.parse(localStorage.getItem('negotiated-payment-summary-export-versions'))['2026-11-12']),2,'生成失败不推进版本');
    assert.equal(downloadCount,beforeFailure,'生成失败不得触发下载');
    await page.evaluate(()=>window.NegotiatedPaymentSummaryExport=async()=>new Uint8Array([0]));
    await page.locator('#download').click();await page.locator('#status').filter({hasText:'导出失败'}).waitFor();
    assert.equal(await page.evaluate(()=>JSON.parse(localStorage.getItem('negotiated-payment-summary-export-versions'))['2026-11-12']),2,'回读校验失败不推进版本');
    assert.equal(downloadCount,beforeFailure,'回读校验失败不得触发下载');
    await page.evaluate(()=>window.NegotiatedPaymentSummaryExport=window.__realExport);
    await download(page,page,3);
    await page.reload({waitUntil:'networkidle'});
    await page.locator('#sourceFile:not([disabled])').waitFor();
    assert.equal(await page.locator('#sourceFile').evaluate(input=>input.files.length),0,'刷新不恢复上传文件');
    assert.equal(await page.locator('#reference').innerText(),'尚未引用文件');
    assert(await page.locator('#download').isDisabled());assert(await page.locator('#result').isHidden());
    await upload(page,fixture([['供应商乙','','否','1月25日－5']]),'刷新后新文件.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    await download(page,page,4);
    await page.evaluate(()=>window.__negotiatedSummaryTestNow='2026-11-12T15:59:59Z');
    await download(page,page,5);
    await page.evaluate(()=>window.__negotiatedSummaryTestNow='2026-11-12T16:00:00Z');
    await download(page,page,1,'2026-11-13');
    assert.deepEqual(await page.evaluate(()=>JSON.parse(localStorage.getItem('negotiated-payment-summary-export-versions'))),{'2026-11-12':5,'2026-11-13':1});
    await page.evaluate(()=>Object.defineProperty(window,'localStorage',{get(){throw new DOMException('测试缓存不可用','SecurityError');}}));
    await download(page,page,2,'2026-11-13');await download(page,page,3,'2026-11-13');
    assert.match(await page.locator('#status').innerText(),/本页内仍会递增/);
    if(process.env.QA_SAMPLE_FILE){
      await upload(page,fs.readFileSync(process.env.QA_SAMPLE_FILE),path.basename(process.env.QA_SAMPLE_FILE));
      await page.locator('#status').filter({hasText:'修正来源问题'}).waitFor();
      assert.match(await page.locator('#counts').innerText(),/119 条明细，47 家供应商/);
      assert(await page.locator('#download').isDisabled());
      const issues=await page.locator('#issues').innerText();
      for(const address of ['L25','L44','L78'])assert(issues.includes(address),address+' 缺明细必须提示');
      assert.match(await page.locator('#preview').innerText(),/2027年4月25日/);
    }
    assert.deepEqual(errors,[]);
    console.log('PASS: 首页与独立页、具体日分列/同月多日/跨年、明细金额/忽略月度金额、分类核对、缺明细保护、切换保留、替换重置、错误清理、公式缓存保护、预览下载一致、手机布局、版本递增/刷新/失败不递增/北京时间跨日/缓存不可用'+(process.env.QA_SAMPLE_FILE?'、真实样表':'') );
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
