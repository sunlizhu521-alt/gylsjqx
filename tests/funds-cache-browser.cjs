const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),{fixtures}=require('./funds-fixtures.cjs');

(async()=>{
 const data=await fixtures(),base=process.env.QA_URL||'http://127.0.0.1:8940/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.route(url=>url.pathname.endsWith('/funds.html')&&!url.searchParams.has('v'),r=>r.fulfill({contentType:'text/html',body:'<p>旧页面</p>'}));
  await page.clock.setFixedTime(new Date('2026-10-27T00:00:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  let frame=page.frameLocator('iframe[data-tool="funds"]');await frame.locator('#ledgerFile:not([disabled])').waitFor();
  assert.match(await page.locator('iframe[data-tool="funds"]').getAttribute('src'),/[?&]v=/,'首页传入版本号，避免旧工具页面缓存');
  const upload=async(kind,bytes=data[kind],name=kind+'.xlsx')=>frame.locator('#'+kind+'File').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
  const ready=async()=>{await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:60000});};
  const downloaded=async()=>{const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;return X.read(await fs.readFile(await d.path()),{cellNF:true,cellStyles:true});};
  const oaVersion=label=>{const b=X.read(data.oa);for(const row of X.utils.sheet_to_json(b.Sheets.Sheet0,{header:1}).keys())if(row)b.Sheets.Sheet0['B'+(row+1)]={t:'s',v:label};return Buffer.from(X.write(b,{type:'buffer',bookType:'xlsx'}));};
  for(const kind of ['ledger','applications','oa','negotiated'])await upload(kind);
  await frame.locator('#negotiatedSheet').selectOption(data.sheetName);await ready();
  const initialInfo={};for(const kind of ['ledger','applications','oa','negotiated']){initialInfo[kind]=await frame.locator('#'+kind+'Info').textContent();assert.match(initialInfo[kind],new RegExp('^'+kind+'\\.xlsx\\n引用时间：'));assert.match(initialInfo[kind],/08:00:00/);}
  const initialPreview=await frame.locator('#normalPreview').innerText();
  await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();
  assert.equal(await frame.locator('#normalPreview').innerText(),initialPreview,'切换工具保留本次内存结果');

  await page.clock.setFixedTime(new Date('2026-10-27T01:00:00Z'));
  await upload('oa',oaVersion('最新文件节点'));
  await frame.locator('#oaInfo').filter({hasText:'09:00:00'}).waitFor();
  assert(await frame.locator('#results').isHidden());assert(await frame.locator('#download').isDisabled(),'同名新文件先清除旧导出');
  for(const kind of ['ledger','applications','negotiated'])assert.equal(await frame.locator('#'+kind+'Info').textContent(),initialInfo[kind],'引用时间不能串到其他文件角色');
  await ready();assert.match(await frame.locator('#normalPreview').innerText(),/最新文件节点/);
  let saved=await downloaded();assert(X.utils.sheet_to_json(saved.Sheets.正常付款).every(row=>row['当前节点']==='最新文件节点'),'导出读取同名文件的新内容');

  await frame.locator('body').evaluate(()=>{
   const original=File.prototype.arrayBuffer;window.__releaseOldRead=null;window.__delayNextRead=true;
   File.prototype.arrayBuffer=function(){const p=original.call(this);if(window.__delayNextRead){window.__delayNextRead=false;return new Promise(resolve=>{window.__releaseOldRead=()=>p.then(resolve);});}return p;};
  });
  await upload('oa',oaVersion('迟到的旧文件'));
  assert(await frame.locator('#analyze').isDisabled(),'读取未完成不能使用旧输入分析');
  await upload('oa',oaVersion('最后上传的文件'));
  await frame.locator('#analyze:not([disabled])').waitFor();
  await frame.locator('body').evaluate(async()=>{await window.__releaseOldRead();});
  await ready();saved=await downloaded();assert(X.utils.sheet_to_json(saved.Sheets.正常付款).every(row=>row['当前节点']==='最后上传的文件'),'较早上传的慢读取不能覆盖较晚上传');

  const negotiatedVersion=(date='',note='')=>{
   const book=X.read(data.negotiated,{cellNF:true,cellStyles:true}),sheet=book.Sheets[data.sheetName];sheet.L2={t:'s',v:date};sheet.O2={t:'s',v:note};
   X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['核对说明'],['旧结果：付款节奏跟进-2026-11-13-V2.xlsx']]),'付款申请核对');
   return Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
  };
  const version2='付款节奏跟进-2026-11-13-V2.xlsx',version3='付款节奏跟进-2026-11-13-V3.xlsx',missingDate=negotiatedVersion(),newDate=negotiatedVersion('2026-11-12','同名V3最新内容');
  const readNegotiated=async(bytes,name)=>{await upload('negotiated',bytes,name);await frame.locator('#negotiatedSheet').selectOption(data.sheetName);};
  await readNegotiated(missingDate,version2);await ready();assert.match(await frame.locator('#reviewPreview').innerText(),/V2\.xlsx/);assert.match(await frame.locator('#sourceVersion').innerText(),/V2\.xlsx/);
  await readNegotiated(missingDate,version3);
  for(const id of ['issueList','normalPreview','negotiatedPreview','reviewPreview','progressPreview','sourceVersion'])assert.equal(await frame.locator('#'+id).textContent(),'','换文件彻底清空旧结果 '+id);
  assert(await frame.locator('#download').isDisabled());await ready();
  const currentReview=await frame.locator('#reviewPreview').innerText();assert.match(currentReview,/V3\.xlsx/);assert.doesNotMatch(currentReview,/V2\.xlsx/);assert.match(await frame.locator('#sourceVersion').innerText(),/V3\.xlsx/);
  saved=await downloaded();let exportedReviews=JSON.stringify(X.utils.sheet_to_json(saved.Sheets.付款申请核对));assert.match(exportedReviews,/V3\.xlsx/);assert.doesNotMatch(exportedReviews,/V2\.xlsx/,'更新来源文件原来已有的核对sheet，不沿用旧说明');
  await readNegotiated(newDate,version3);await ready();saved=await downloaded();
  assert.equal(saved.Sheets.谈判付款.L2.v,'2026-11-12');assert.equal(saved.Sheets.谈判付款.O2.v,'同名V3最新内容');
  exportedReviews=JSON.stringify(X.utils.sheet_to_json(saved.Sheets.付款申请核对));assert.doesNotMatch(exportedReviews,/V2\.xlsx|应付时间为空或无法识别/);
  await frame.locator('#negotiatedFile').setInputFiles([]);assert.equal(await frame.locator('#negotiatedInfo').textContent(),'尚未引用');assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#analyze').isDisabled());assert.equal(await frame.locator('#reviewPreview').textContent(),'');
  await readNegotiated(newDate,version3);await ready();
  const silentlyReplace=()=>frame.locator('#negotiatedFile').evaluate((input,{bytes,name})=>{const transfer=new DataTransfer();transfer.items.add(new File([new Uint8Array(bytes)],name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));input.files=transfer.files;},{bytes:[...newDate],name:version3});
  await silentlyReplace();await frame.locator('#analyze').click();assert.match(await frame.locator('#status').textContent(),/来源版本不一致/);assert(await frame.locator('#download').isDisabled());assert.equal(await frame.locator('#reviewPreview').textContent(),'');
  await readNegotiated(newDate,version3);await ready();let unintendedDownloads=0;const countDownload=()=>unintendedDownloads++;page.on('download',countDownload);
  await silentlyReplace();await frame.locator('#download').click();assert.match(await frame.locator('#status').textContent(),/来源文件已更换/);assert(await frame.locator('#download').isDisabled());assert.equal(unintendedDownloads,0);page.off('download',countDownload);
  // A replacement event during an in-flight export must not restore the previous result.
  await readNegotiated(missingDate,version2);
  await frame.locator('body').evaluate(()=>{const original=FundsExport;window.__holdExport=true;window.FundsExport=async(...args)=>{const blob=await original(...args);if(window.__holdExport){window.__holdExport=false;return new Promise(resolve=>{window.__releaseExport=()=>resolve(blob);});}return blob;};});
  await frame.locator('#analyze').click();await frame.locator('body').evaluate(async()=>{while(!window.__releaseExport)await new Promise(resolve=>setTimeout(resolve,10));});
  await frame.locator('#negotiatedFile').evaluate((input,{bytes,name})=>{const transfer=new DataTransfer();transfer.items.add(new File([new Uint8Array(bytes)],name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));input.files=transfer.files;input.dispatchEvent(new Event('change',{bubbles:true}));},{bytes:[...newDate],name:version3});
  await frame.locator('#negotiatedInfo').filter({hasText:version3+'\n引用时间'}).waitFor();await frame.locator('body').evaluate(()=>window.__releaseExport());
  await frame.locator('#negotiatedSheet:not([disabled])').waitFor();assert(await frame.locator('#download').isDisabled());assert.equal(await frame.locator('#reviewPreview').textContent(),'');
  await frame.locator('#negotiatedSheet').selectOption(data.sheetName);await ready();saved=await downloaded();assert.equal(saved.Sheets.谈判付款.O2.v,'同名V3最新内容');assert.doesNotMatch(JSON.stringify(X.utils.sheet_to_json(saved.Sheets.付款申请核对)),/V2\.xlsx/);

  await frame.locator('body').evaluate(()=>{
   const original=FundsCore.build;window.__buildDecisions=[];window.FundsCore={...FundsCore,build:(sources,decisions)=>{
    window.__buildDecisions.push(JSON.parse(JSON.stringify(decisions)));const result=original(sources,decisions);
    result.issues.push({id:'cache-decision-test',targetLabel:'测试人工核对',message:'请选择本次处理',resolved:Boolean(decisions['cache-decision-test']?.length),choices:[{value:'keep',label:'保留原值'}]});return result;
   }};
  });
  await frame.locator('#analyze').click();await frame.locator('[data-decision="cache-decision-test"]').selectOption('keep');
  await frame.locator('#download:not([disabled])').waitFor({timeout:60000});
  await upload('oa',oaVersion('换文件清空决定'));
  await frame.locator('#analyze:not([disabled])').click();await frame.locator('[data-decision="cache-decision-test"]').waitFor();
  assert.deepEqual(await frame.locator('body').evaluate(()=>window.__buildDecisions.at(-1)),{},'新文件不得继承旧核对决定');
  assert(await frame.locator('#download').isDisabled());
  await frame.locator('[data-decision="cache-decision-test"]').selectOption('keep');await frame.locator('#download:not([disabled])').waitFor();

  await frame.locator('body').evaluate(()=>{const original=File.prototype.arrayBuffer;window.__failNextRead=true;File.prototype.arrayBuffer=function(){if(window.__failNextRead){window.__failNextRead=false;return Promise.reject(Error('测试读取失败'));}return original.call(this);};});
  await upload('oa',oaVersion('失败文件'));
  await frame.locator('#oaInfo').filter({hasText:'读取失败'}).waitFor();assert(await frame.locator('#results').isHidden());assert(await frame.locator('#download').isDisabled());assert(await frame.locator('#analyze').isDisabled(),'读取失败不得回退到旧OA文件');
  await upload('oa',oaVersion('重新上传成功'));await frame.locator('#analyze:not([disabled])').waitFor();
  await page.evaluate(()=>localStorage.setItem('funds-export-versions',JSON.stringify({'2026-10-27':9})));
  await page.reload({waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();frame=page.frameLocator('iframe[data-tool="funds"]');await frame.locator('#ledgerFile:not([disabled])').waitFor();
  for(const kind of ['ledger','applications','oa','negotiated']){assert.equal(await frame.locator('#'+kind+'Info').textContent(),'尚未引用');assert.equal(await frame.locator('#'+kind+'File').evaluate(el=>el.files.length),0);}
  assert(await frame.locator('#analyze').isDisabled());assert(await frame.locator('#download').isDisabled());assert.equal(await frame.locator('#filenameHint').textContent(),'付款节奏跟进-2026-10-27-V10.xlsx');
  assert.deepEqual(errors,[]);
  console.log('PASS: V2 to V3 review sources and same-name latest contents; original review sheet replaced, remove/silent replacement/in-flight export block stale results; uploads are not persisted, revision and manual-decision reset guards, refresh requires all four files.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
