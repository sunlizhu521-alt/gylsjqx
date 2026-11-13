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
    assert.deepEqual(book.SheetNames,['谈判付款整理','校验']);
    assert.equal(book.Sheets.校验['!ref'],'A1:P6');
    assert.equal(book.Sheets.校验.D2.v,'供应商甲');
    assert.equal(book.Sheets.校验.O2.v,'待核对');
    assert.equal(await frame.locator('#validation-preview tbody tr').count(),5);
    assert.match(await frame.locator('#validation-counts').innerText(),/差异不阻止下载/); assert.equal(s['!ref'],'A1:I4');
    assert.equal(s.B3.v,80); assert.equal(s.B3.f,'SUM(C3:F3)');assert.equal(s.G3.v,5);assert.equal(s.H4.v,180000);
    assert.equal(s.C2.t,'s');assert.equal(s.C2.z,'@');assert.equal(s.C2.v,'2026年9月30日');assert.equal(X.SSF.format(s.F2.z,s.F2.v),'2027年1月25日');
    assert.deepEqual(['C','D','E','F'].map(c=>s[c+'2'].v),['2026年9月30日','2026年11月25日','2026年11月26日','2027年1月25日']);
    assert.deepEqual(['C','D','E','F'].map(c=>s[c+'3'].v),[10,20,20,30]);
    assert.equal(s['!autofilter'].ref,'A2:I4');
    assert.deepEqual(['H','I'].map(c=>s[c+'2'].v),['2026年11月25日','2027年1月25日']);
    assert.equal(s.G3.f,'SUM(H3:I3)');assert.equal(s.J2,undefined);
    assert.equal(await frame.locator('#preview thead tr:nth-child(2) th').count(),6);
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
    assert(await frame.locator('#download').isDisabled()); assert(await frame.locator('#result').isHidden());assert(await frame.locator('#review-panel').isHidden());assert(await frame.locator('#validation-panel').isHidden());
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
    assert.equal(replaced.A3.v,'供应商乙');assert.equal(replaced.C3.v,5);assert.equal(replaced.D3.v,5);assert.equal(replaced.D2.v,'2027年1月25日');
    assert.equal(replaced.B3.f,'SUM(0)');assert.equal(replaced.C3.f,'SUM(D3:D3)');assert.equal(replaced['!ref'],'A1:D3');
    assert.equal(await page.locator('#preview th').filter({hasText:/^谈判付款节奏$/}).count(),0);
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
    await upload(page,fixture([
      ['零金额供应商','','否','11月25日－0'],
      ['其他供应商','','是','9月30日－-5\n11月26日－10']
    ]),'空日期列验收.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    const compact=X.read(await download(page,page),{type:'buffer',cellNF:true}).Sheets.谈判付款整理;
    assert.equal(compact['!ref'],'A1:F4');
    assert.deepEqual(['C2','D2','F2'].map(a=>compact[a].v),['2026年9月30日','2026年11月26日','2026年11月25日']);
    assert.equal(compact.F3.v,0);assert.equal(compact.F3.t,'n');assert.equal(compact.C4.v,-5);
    assert.equal(compact.E4.f,'SUM(F4:F4)');assert.equal(compact.G2,undefined);
    assert.deepEqual(await page.locator('#preview thead tr:nth-child(2) th').allTextContents(),['2026年9月30日','2026年11月26日','2026年11月25日']);
    const paidBook=X.read(fixture([
      ['混合付款供应商','','是','11月2日－100\n11月25日－100'],
      ['混合付款供应商','','是','11月2日－50'],
      ['同额分批供应商','','是','11月2日－100\n11月25日－100'],
      ['正常已付供应商','','否','11月25日－75'],
      ['正常未付供应商','','否','11月25日－25']
    ]),{type:'buffer'}),paidSheet=paidBook.Sheets.谈判付款;
    paidSheet.I1={t:'s',v:'2026年11月2日'};paidSheet.J1={t:'s',v:'2026年11月16日'};
    for(const [address,value] of [['I2',100],['J2',100],['I3',50],['I4',100],['J4',100],['J5',75],['J6',25]])paidSheet[address]={t:'n',v:value};
    for(const address of ['I2','I4','J5'])paidSheet[address].c=[{a:'核实人',t:'11月5日核实-已支付'}];
    paidSheet.J6.c=[{a:'核实人',t:'申请日期（付款节奏跟进）：2026年11月2日\n请核实是否已支付'}];
    paidSheet['!ref']='A1:J6';
    await upload(page,Buffer.from(X.write(paidBook,{type:'buffer',bookType:'xlsx'})),'批注已支付验收.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    assert.equal(await page.locator('#preview td.is-paid').count(),0,'有已支付批注也不标红');
    assert.equal(await page.locator('#preview tbody tr:nth-child(1) td:nth-child(3)').getAttribute('class'),null,'同日混合款项不红');
    assert.equal(await page.locator('#preview tbody tr:nth-child(2) td:nth-child(3)').evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(255, 255, 255)');
    const paidBytes=await download(page,page),paidOutput=X.read(paidBytes,{type:'buffer',cellStyles:true}).Sheets.谈判付款整理;
    const redAddresses=Object.entries(paidOutput).filter(([address,c])=>/^[A-Z]+\d+$/.test(address)&&/^(?:FF)?FF0000$/i.test(c.s?.fgColor?.rgb||'')).map(([address])=>address).sort();
    assert.deepEqual(redAddresses,[]);
    assert.equal(paidOutput.C3.v,150);assert.equal(paidOutput.D4.v,100);assert.equal(paidOutput.F6.v,25);
    assert.equal(paidOutput['!ref'],'A1:F6');
    const paidZip=await Zip.loadAsync(paidBytes);assert.doesNotMatch(await paidZip.file('xl/worksheets/sheet1.xml').async('string'),/<conditionalFormatting\b/);
    if(process.env.QA_PAID_SCREENSHOT)await page.screenshot({path:process.env.QA_PAID_SCREENSHOT,fullPage:true});
    // Four-way checks keep source records separate and remain downloadable on mismatch.
    const auditBook=X.utils.book_new(),auditHeaders=Array(42).fill('');
    Object.assign(auditHeaders,{0:'供应商',1:'付款主体',2:'是否谈判分批',3:'谈判明细',4:'OA-申请付款金额',5:'对公付款0A流程号'});
    for(let c=16;c<=27;c++)auditHeaders[c]=`2026年11月${c-15}日`;
    for(let c=34;c<=41;c++)auditHeaders[c]=`${2026+Math.floor((c-24)/12)}年${(c-24)%12+1}月支付`;
    const auditRows=[auditHeaders];
    for(const [i,total] of [10,12].entries()){
      const row=Array(42).fill('');
      Object.assign(row,{0:'重复供应商',1:'主体甲',2:'是',3:'11月1日－10',4:10,5:'OA00'+i,16:total,34:10});auditRows.push(row);
    }
    X.utils.book_append_sheet(auditBook,X.utils.aoa_to_sheet(auditRows),'谈判付款');
    await upload(page,Buffer.from(X.write(auditBook,{type:'buffer',bookType:'xlsx'})),'四组金额校验.xlsx');
    await page.locator('#download:not([disabled])').waitFor();
    assert.match(await page.locator('#validation-counts').innerText(),/一致 1 条；不一致 1 条；待核对 0 条/);
    assert.equal(await page.locator('#validation-preview tbody tr').count(),3);
    const auditOutput=X.read(await download(page,page),{type:'buffer',cellNF:true});
    assert.deepEqual(auditOutput.SheetNames,['谈判付款整理','校验']);
    const auditSheet=auditOutput.Sheets.校验;
    assert.equal(auditSheet.O2.v,'一致');assert.equal(auditSheet.O3.v,'不一致');
    assert.equal(auditSheet.M3.v,2);assert.equal(auditSheet.J4.v,22);assert.equal(auditSheet.C3.v,3);
    assert.equal(auditSheet.F2.t,'s');assert.equal(auditSheet.F2.v,'OA000');
    assert.equal(auditSheet['!autofilter'].ref,'A1:P3');
    assert.equal(auditOutput.Sheets.谈判付款整理.B3.v,20);
    if(process.env.QA_SAMPLE_FILE){
      const sourceBytes=fs.readFileSync(process.env.QA_SAMPLE_FILE),Core=require('../negotiated-payment-summary-core.js');
      const source={filename:path.basename(process.env.QA_SAMPLE_FILE),book:X.read(sourceBytes,{type:'buffer',cellNF:true,cellFormula:true,sheetStubs:true,sheets:'谈判付款'})};
      const initial=Core.build(source),decisions=Object.fromEntries(initial.unresolved.map(item=>[item.id,'否']));
      await upload(page,sourceBytes,source.filename);
      for(const item of initial.unresolved)await page.locator(`select[data-address="${item.id}"]`).selectOption('否');
      const expected=Core.build(source,decisions);
      if(expected.errors.length){
        assert(await page.locator('#download').isDisabled());
        assert.match(await page.locator('#status').innerText(),/修正来源问题/);
      }else{
        await page.locator('#download:not([disabled])').waitFor();
        const sampleBytes=await download(page,page),output=X.read(sampleBytes,{type:'buffer',cellNF:true,cellStyles:true}).Sheets.谈判付款整理;
        const columns=Core.dateColumns(expected),n=columns.negotiated.length,m=columns.normal.length;
        assert.equal(output['!ref'],`A1:${X.utils.encode_col(n+m+2)}${expected.rows.length+2}`);
        for(const [r,row] of expected.rows.entries()){
          assert.equal(output['A'+(r+3)].v,row.supplier);
          for(const [type,start] of [['negotiated',2],['normal',n+3]])columns[type].forEach((index,c)=>{
            const address=X.utils.encode_col(start+c)+(r+3),value=output[address];
            if(row[type+'Cents'][index]!==null)assert.equal(Math.round(value.v*100),row[type+'Cents'][index]);
          });
        }
        const actualRed=Object.entries(output).filter(([address,c])=>/^[A-Z]+\d+$/.test(address)&&/^(?:FF)?FF0000$/i.test(c.s?.fgColor?.rgb||'')).map(([address])=>address).sort();
        assert.deepEqual(actualRed,[]);
        assert.equal(await page.locator('#preview td.is-paid').count(),0);
        const validation=X.read(sampleBytes,{type:'buffer',cellNF:true}).Sheets.校验;
        assert.equal(validation['!ref'],`A1:P${expected.validationRows.length+1}`);
        assert.equal(await page.locator('#validation-preview tbody tr').count(),expected.validationRows.length);
        expected.validationRows.forEach((values,r)=>values.forEach((value,c)=>{
          const actual=validation[X.utils.encode_cell({r:r+1,c})]?.v;
          if(value==null||value==='')assert(actual==null||actual==='');
          else assert.equal(Core.VALIDATION_MONEY_COLUMNS.includes(c)?Math.round(actual*100):actual,value);
        }));
        console.log('真实样表校验',JSON.stringify({records:expected.recordCount,counts:expected.validationCounts}));
        if(process.env.QA_SAMPLE_OUTPUT)fs.writeFileSync(process.env.QA_SAMPLE_OUTPUT,sampleBytes);
      }
    }
    assert.deepEqual(errors,[]);
    console.log('PASS: 首页与独立页、各区删除空日期列/零值负值/单侧空区、取消标红/原批注不影响白底/校验表与预览一致、具体日分列/跨年、分类核对、缺明细保护、切换及替换、公式缓存保护、预览下载一致、手机布局、版本递增'+(process.env.QA_SAMPLE_FILE?'、真实样表':'') );
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
