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
  const upload=async(kind,bytes=data[kind])=>frame.locator('#'+kind+'File').setInputFiles({name:kind+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
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
  assert(await frame.locator('#analyze').isDisabled());assert(await frame.locator('#download').isDisabled());assert.equal(await frame.locator('#filenameHint').textContent(),'资金整理-2026-10-27-V10.xlsx');
  assert.deepEqual(errors,[]);
  console.log('PASS: funds files are not persisted or restored; versioned iframe, same-name latest bytes, per-role reference time, slow old upload discarded, old decisions cleared, failed upload blocks stale results, refresh requires all four files; only download version persists.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
