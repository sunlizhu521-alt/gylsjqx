const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const {sample,ID}=require('./funds-batch-relocation-fixtures.cjs');
const write=book=>Buffer.from(X.write(structuredClone(book),{type:'buffer',bookType:'xlsx'}));
const read=bytes=>X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true});
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/',browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  async function upload(role,buffer,name){await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});}
  async function download(){
   await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:30000});assert.equal(await frame.locator('#pendingCount').innerText(),'0');
   const event=page.waitForEvent('download');await frame.locator('#download').click();return fs.readFile(await(await event).path());
  }
  for(const multiple of [false,true]){
   const input=sample(multiple?120000:15230,multiple?120000:49990,multiple);
   for(const role of ['ledger','applications','oa','negotiated'])await upload(role,write(input[role].book),input[role].name);
   for(let round=0;round<2;round++){
    const bytes=await download(),out=read(bytes),s=out.Sheets.谈判付款;
    assert.equal(s.L2.s.fgColor.rgb.slice(-6),'FF0000');assert.match(s.L2.c[0].t,/申请日期（付款节奏跟进）：2026年10月21日/);assert.doesNotMatch(s.L2.c[0].t,/2026年11月19日/);
    assert.equal(s.C3.v,ID(2));assert.equal(s.F3.f,'D3-E3');assert.equal(s.J3.f,'SUM(L3:N3)');
    const text=X.utils.sheet_to_json(out.Sheets.付款申请核对,{header:1}).flat().join('\n');
    if(multiple){assert.match(text,/多个同额候选/);assert(!['FF0000','FFFF00'].includes(s.M3.s?.fgColor?.rgb?.slice(-6)));assert.equal(await frame.locator('#negotiatedPreview tbody tr').nth(1).locator('td').nth(12).getAttribute('class'),null);}
    else {assert.match(text,/已按供应商、付款主体、金额及原批注唯一匹配/);assert.equal(s.M3.s.fgColor.rgb.slice(-6),'FFFF00');assert.match(s.M3.c[0].t,/申请日期（付款节奏跟进）：2026年11月19日/);assert.match(s.M3.c[0].t,/付款时间（付款节奏跟进）：未记录/);assert.equal(await frame.locator('#negotiatedPreview tbody tr').nth(1).locator('td').nth(12).getAttribute('class'),'funds-requested');}
    assert.doesNotMatch(text,/申请标红核对|申请超额/);
    if(!round)await upload('negotiated',bytes,'重传.xlsx');
   }
  }
  if(process.env.FUNDS_REAL_DIR){
   const dir=process.env.FUNDS_REAL_DIR,files={ledger:'采购付款台账.xlsx',applications:'付款单申请汇总.xlsx',oa:'MDST-对公付款-20261120100116.xlsx',negotiated:'付款节奏跟进-2026-11-20-V2.xlsx'};
   const source=read(await fs.readFile(dir+'/'+files.negotiated)).Sheets.谈判付款;
   for(const [role,name]of Object.entries(files))await upload(role,await fs.readFile(dir+'/'+name),name);
   const out=read(await download()),s=out.Sheets.谈判付款;
   assert.deepEqual(s['!merges'],source['!merges']);
   for(const [a,c]of Object.entries(source))if(/^[A-Z]+[12]$/.test(a)&&c.v!=null)assert.equal(s[a]?.v,c.v,a+' 原表头');
   for(const a of ['T42','T79']){assert.equal(s[a].s.fgColor.rgb.slice(-6),'FF0000');assert.equal(s[a].c[0].t,source[a].c[0].t);}
   assert.equal(s.V80.v,49990);assert.equal(s.V80.s.fgColor.rgb.slice(-6),'FFFF00');assert.match(s.V80.c[0].t,/2026年11月19日/);assert.equal(s.D80.v,'MDST-DGFK20261112018');
   for(const [a,c]of Object.entries(source))if(c?.f&&!/^(?:N|O|P)\d+$/.test(a))assert.equal(s[a]?.f,c.f,a+' 原计划公式');
   assert.equal(s.M80.f,source.M80.f);assert.equal(s.V43.v,120000);assert.equal(s.W44.v,120000);assert(!['FF0000','FFFF00'].includes(s.V43.s?.fgColor?.rgb?.slice(-6)));
  }
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,syntheticDownloads:4,realSource:!!process.env.FUNDS_REAL_DIR,uniqueBatchMatched:true,ambiguousBatchesPreserved:true,oldCommentsAndFillsPreserved:true,reimport:true}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e.stack);process.exitCode=1;});
