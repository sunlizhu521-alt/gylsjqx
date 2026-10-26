const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/',out='/private/tmp/gylsjqx-responsive-qa';await fs.mkdir(out,{recursive:true});
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
  await page.goto(base,{waitUntil:'networkidle'});
  for(const [width,height] of [[2560,1440],[1440,900],[1024,768],[390,844]]){
   await page.setViewportSize({width,height});
   for(const id of ['merge','payment','funds','receipt-sign','fulfillment','beihuo','contract']){
    await page.locator('.tool-tab[data-tool="'+id+'"]').click();
    const frameElement=page.locator('iframe[data-tool="'+id+'"]');
    if(await frameElement.count())await page.frameLocator('iframe[data-tool="'+id+'"]').locator('body').waitFor();
    const layout=await page.evaluate(()=>{const box=document.querySelector('.shell').getBoundingClientRect(),workspace=document.querySelector('.workspace').getBoundingClientRect();return {width:document.documentElement.clientWidth,scroll:document.documentElement.scrollWidth,left:box.left,right:box.right,shellWidth:box.width,workspaceRight:workspace.right};});
    assert.equal(layout.left,0);assert(Math.abs(layout.shellWidth-layout.width)<=1,`${id} ${width}px: shell must fill browser`);assert(layout.scroll<=layout.width+1,`${id} ${width}px: outer page overflow ${layout.scroll-layout.width}`);assert(layout.workspaceRight<=layout.width,`${id} ${width}px: workspace off-screen`);
    if(await frameElement.count()){
     const inner=await page.frameLocator('iframe[data-tool="'+id+'"]').locator('body').evaluate(body=>({width:document.documentElement.clientWidth,scroll:document.documentElement.scrollWidth,body:body.getBoundingClientRect().width}));
     assert(inner.scroll<=inner.width+1,`${id} ${width}px: inner page overflow ${inner.scroll-inner.width}`);
    }
   }
   await page.locator('.tool-tab[data-tool="funds"]').click();await page.locator('.workspace').scrollIntoViewIfNeeded();await page.screenshot({path:out+'/funds-'+width+'.png'});
  }
  assert.deepEqual(errors,[]);console.log('PASS: full-width homepage and all embedded tools at 2560, 1440, 1024 and 390px without page-level horizontal overflow');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
