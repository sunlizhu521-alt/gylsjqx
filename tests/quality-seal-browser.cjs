const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const tools=['internal-statement','funds','receipt-sign','seal','beihuo','fulfillment','contract','payment'];
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[],requests=[],downloads=[];
  page.on('download',d=>downloads.push(d));
  page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>requests.push({method:r.method(),url:r.url()}));
  await page.goto(base,{waitUntil:'networkidle'});
  assert.deepEqual(await page.locator('.tool-tab').evaluateAll(tabs=>tabs.map(t=>t.dataset.tool)),tools);
  await page.locator('.tool-tab[data-tool="seal"]').click();
  const frame=page.frameLocator('iframe[data-tool="seal"]');await frame.locator('#reportFile:not([disabled])').waitFor();
  assert.equal(await frame.locator('#generate').isDisabled(),true);assert.equal(await frame.locator('#download').isDisabled(),true);
  const makeImage=async(width,height)=>{
   const uri=await page.evaluate(({width,height})=>{
    const c=document.createElement('canvas');c.width=width;c.height=height;const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,width,height);
    for(const [color,left,top] of [['#0000ff',20,20],['#008000',width-60,20],['#800080',20,height-60],['#00c8c8',width-60,height-60]]){x.fillStyle=color;x.fillRect(left,top,40,40);}
    x.fillStyle='#111';x.font='32px sans-serif';x.fillText('合成检验报告',width/3,110);return c.toDataURL('image/png');
   },{width,height});return Buffer.from(uri.split(',')[1],'base64');
  };
  const source=await makeImage(1000,700),second=await makeImage(600,900);
  const upload=async(name,buffer=source)=>{await frame.locator('#reportFile:not([disabled])').setInputFiles({name,mimeType:'image/png',buffer});};
  const ready=()=>frame.locator('#generate:not([disabled])').waitFor();
  const snapshot=()=>frame.locator('#preview').evaluate(c=>{
   const x=c.getContext('2d'),a=x.getImageData(0,0,c.width,c.height).data;let red=0;
   for(let y=Math.floor(c.height*.55);y<c.height;y++)for(let col=Math.floor(c.width*.55);col<c.width;col++){const n=(y*c.width+col)*4;if(a[n]>120&&a[n+1]<a[n]*.7&&a[n+2]<a[n]*.7)red++;}
   return {width:c.width,height:c.height,red,corners:[[40,40],[c.width-40,40],[40,c.height-40],[c.width-40,c.height-40]].map(([left,top])=>Array.from(x.getImageData(left,top,1,1).data).slice(0,3)),uri:c.toDataURL('image/png')};
  });
  await upload('合成检验报告.png');await ready();assert.match(await frame.locator('#fileInfo').innerText(),/合成检验报告.png/);
  const original=await snapshot();assert.deepEqual([original.width,original.height],[1000,700]);assert.equal(original.red,0);
  assert.deepEqual(original.corners,[[0,0,255],[0,128,0],[128,0,128],[0,200,200]]);
  await frame.locator('#rotateRight').click();const turned=await snapshot();assert.deepEqual([turned.width,turned.height],[700,1000]);assert.equal(turned.red,0);
  assert.deepEqual(turned.corners,[[128,0,128],[0,0,255],[0,200,200],[0,128,0]]);
  const previewOnly=page.waitForEvent('download',{timeout:350}).then(()=>true,()=>false);
  await frame.locator('#generate').click();await ready();const stamped=await snapshot();assert.equal(await previewOnly,false,'旋正并加章只预览，不能自动下载');assert.equal(downloads.length,0);assert(stamped.red>100,'右下角应盖上原质检电子章');
  assert.deepEqual([stamped.width,stamped.height],[700,1000]);
  await frame.locator('#generate').click();await ready();assert.equal((await snapshot()).uri,stamped.uri,'重复生成不能叠加印章');
  const pending=page.waitForEvent('download');await frame.locator('#download').click();const download=await pending;
  assert.equal(download.suggestedFilename(),'合成检验报告_旋正已盖章.jpg');const bytes=await fs.readFile(await download.path());assert.equal(bytes[0],0xff);assert.equal(bytes[1],0xd8);
  const exported=await page.evaluate(async data=>{const i=new Image();i.src='data:image/jpeg;base64,'+data;await i.decode();const c=document.createElement('canvas');c.width=i.naturalWidth;c.height=i.naturalHeight;const x=c.getContext('2d');x.drawImage(i,0,0);const a=x.getImageData(0,0,c.width,c.height).data;let red=0;for(let y=Math.floor(c.height*.55);y<c.height;y++)for(let col=Math.floor(c.width*.55);col<c.width;col++){const n=(y*c.width+col)*4;if(a[n]>120&&a[n+1]<a[n]*.7&&a[n+2]<a[n]*.7)red++;}return {width:c.width,height:c.height,red};},bytes.toString('base64'));
  assert.deepEqual([exported.width,exported.height],[700,1000]);assert(exported.red>100);
  await page.locator('.tool-tab[data-tool="funds"]').click();await page.locator('.tool-tab[data-tool="seal"]').click();assert.equal((await snapshot()).uri,stamped.uri);assert.equal(await page.locator('iframe[data-tool="seal"]').count(),1);
  await frame.locator('#rotateLeft').click();const resetByRotation=await snapshot();assert.deepEqual([resetByRotation.width,resetByRotation.height],[1000,700]);assert.equal(resetByRotation.red,0);
  await frame.locator('#rotateLeft').click();assert.deepEqual([(await snapshot()).width,(await snapshot()).height],[700,1000]);
  await frame.locator('#resetRotation').click();assert.deepEqual([(await snapshot()).width,(await snapshot()).height],[1000,700]);
  await upload('另一份报告.png',second);await ready();const replaced=await snapshot();assert.deepEqual([replaced.width,replaced.height],[600,900]);assert.equal(replaced.red,0);
  const nextDownload=page.waitForEvent('download');await frame.locator('#download').click();assert.equal((await nextDownload).suggestedFilename(),'另一份报告_旋正已盖章.jpg');
  await frame.locator('#clearImage').click();assert.equal(await frame.locator('#generate').isDisabled(),true);assert.equal(await frame.locator('#download').isDisabled(),true);
  await upload('损坏图片.png',Buffer.from('not a valid PNG'));await frame.locator('#status').filter({hasText:/失败|无法|错误/}).waitFor();assert.equal(await frame.locator('#download').isDisabled(),true);
  await upload('恢复有效图片.png',second);await ready();assert.equal((await snapshot()).red,0);
  for(const width of [1440,390]){
   await page.setViewportSize({width,height:950});const outer=await page.evaluate(()=>({w:document.documentElement.clientWidth,s:document.documentElement.scrollWidth}));assert(outer.s<=outer.w+1);
   const inner=await frame.locator('body').evaluate(()=>({w:document.documentElement.clientWidth,s:document.documentElement.scrollWidth}));assert(inner.s<=inner.w+1,'手机内嵌页面不能横向溢出');
  }
  if(process.env.QA_OUT){await fs.mkdir(process.env.QA_OUT,{recursive:true});await fs.writeFile(process.env.QA_OUT+'/stamped.jpg',bytes);await page.setViewportSize({width:1440,height:1000});await frame.locator('#generate').click();await ready();await frame.locator('main').screenshot({path:process.env.QA_OUT+'/seal.png'});}
  await page.reload({waitUntil:'networkidle'});await page.locator('.tool-tab.active[data-tool="seal"]').waitFor();await frame.locator('#reportFile:not([disabled])').waitFor();assert.equal(await frame.locator('#download').isDisabled(),true,'刷新不保留原报告文件');
  // The file chooser must remain disabled until its script has registered upload handlers.
  const loadingPage=await browser.newPage();loadingPage.on('pageerror',e=>errors.push(e.message));let releaseScript,scriptRequested;
  const scriptGate=new Promise(resolve=>{releaseScript=resolve}),scriptSeen=new Promise(resolve=>{scriptRequested=resolve});
  await loadingPage.route('**/quality-seal.js*',async route=>{scriptRequested();await scriptGate;await route.continue();});
  await loadingPage.goto(base,{waitUntil:'networkidle'});await loadingPage.locator('.tool-tab[data-tool="seal"]').click();const loading=loadingPage.frameLocator('iframe[data-tool="seal"]');
  await scriptSeen;await loading.locator('#reportFile').waitFor();assert.equal(await loading.locator('#reportFile').isDisabled(),true,'脚本未就绪时不能选择图片');
  releaseScript();await loading.locator('#reportFile:not([disabled])').waitFor();await loading.locator('#reportFile:not([disabled])').setInputFiles({name:'加载后报告.png',mimeType:'image/png',buffer:source});
  await loading.locator('#generate:not([disabled])').waitFor();assert.match(await loading.locator('#fileInfo').innerText(),/加载后报告.png/);await loadingPage.close();
  // Hold the original seal load so a replacement or clear can invalidate in-flight generation.
  for(const mode of ['replace','clear']){
   const racePage=await browser.newPage(),raceDownloads=[];racePage.on('pageerror',e=>errors.push(e.message));racePage.on('download',d=>raceDownloads.push(d));
   let release,requested;const gate=new Promise(resolve=>{release=resolve}),requestSeen=new Promise(resolve=>{requested=resolve});
   await racePage.route('**/assets/quality-seal.png',async route=>{requested();await gate;await route.continue();});
   await racePage.goto(base,{waitUntil:'networkidle'});await racePage.locator('.tool-tab[data-tool="seal"]').click();const f=racePage.frameLocator('iframe[data-tool="seal"]');
   await f.locator('#reportFile:not([disabled])').setInputFiles({name:'旧报告.png',mimeType:'image/png',buffer:source});await f.locator('#download:not([disabled])').waitFor();
   await f.locator('#download').click();await requestSeen;
   if(mode==='replace'){await f.locator('#reportFile:not([disabled])').setInputFiles({name:'新报告.png',mimeType:'image/png',buffer:second});await f.locator('#generate:not([disabled])').waitFor();}
   else await f.locator('#clearImage').click();
   const response=racePage.waitForResponse(r=>r.url().endsWith('/assets/quality-seal.png'));release();await (await response).finished();await racePage.waitForTimeout(150);
   assert.equal(raceDownloads.length,0,'已失效的生成不得触发下载');
   const size=await f.locator('#preview').evaluate(c=>[c.width,c.height]);assert.deepEqual(size,mode==='replace'?[600,900]:[1,1]);
   if(mode==='replace'){assert.match(await f.locator('#fileInfo').innerText(),/新报告.png/);assert.match(await f.locator('#status').innerText(),/已加载/);}
   else{assert.equal(await f.locator('#previewSection').isVisible(),false);assert.equal(await f.locator('#download').isDisabled(),true);}
   await racePage.close();
  }
  // A missing built-in seal must leave a clean report preview and allow a retry.
  const retryPage=await browser.newPage();retryPage.on('pageerror',e=>errors.push(e.message));let sealRequests=0;
  await retryPage.route('**/assets/quality-seal.png',route=>++sealRequests===1?route.fulfill({status:404,body:'missing seal'}):route.continue());
  await retryPage.goto(base,{waitUntil:'networkidle'});await retryPage.locator('.tool-tab[data-tool="seal"]').click();const retry=retryPage.frameLocator('iframe[data-tool="seal"]');
  await retry.locator('#reportFile:not([disabled])').setInputFiles({name:'重试报告.png',mimeType:'image/png',buffer:source});await retry.locator('#generate:not([disabled])').waitFor();await retry.locator('#generate').click();
  await retry.locator('#status').filter({hasText:/加章失败/}).waitFor();await retry.locator('#download:not([disabled])').waitFor();
  const retried=retryPage.waitForEvent('download');await retry.locator('#download').click();assert.equal((await retried).suggestedFilename(),'重试报告_旋正已盖章.jpg');assert.equal(sealRequests,2);
  let failedDrawDownloads=0;retryPage.on('download',()=>failedDrawDownloads++);
  await retry.locator('body').evaluate(()=>{window.__qaGetContext=HTMLCanvasElement.prototype.getContext;HTMLCanvasElement.prototype.getContext=()=>null;});
  await retry.locator('#generate').click();await retry.locator('#status').filter({hasText:/加章失败/}).waitFor();assert.equal(failedDrawDownloads,0);
  await retry.locator('body').evaluate(()=>{HTMLCanvasElement.prototype.getContext=window.__qaGetContext;delete window.__qaGetContext;});
  const drawRetry=retryPage.waitForEvent('download');await retry.locator('#download').click();assert.equal((await drawRetry).suggestedFilename(),'重试报告_旋正已盖章.jpg');
  await retryPage.close();
  assert.deepEqual(errors,[]);assert(requests.every(r=>r.method==='GET'),'本地盖章过程不能上传文件');
  console.log(JSON.stringify({base,tools:tools.length,rotation:true,stamp:true,jpg:true,repeatSafe:true,replacement:true,invalidImage:true,embeddedState:true,responsive:true,previewOnly:true,inflightInvalidation:true,sealRetry:true,canvasFailureRetry:true,startupReady:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
