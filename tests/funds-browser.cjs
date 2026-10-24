const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const {fixtures,NORMAL_HEADERS}=require('./funds-fixtures.cjs');
const REAL_PATHS={ledger:'采购付款台账.xlsx',applications:'付款申请汇总_1024.xlsx',negotiated:'采购付款台账_1020.xlsx',oa:'MDST-对公付款-20261024190903.xlsx'};
const FILE_IDS={ledger:'ledgerFile',applications:'applicationsFile',negotiated:'negotiatedFile',oa:'oaFile'};
const fields=sheet=>Object.fromEntries((X.utils.sheet_to_json(sheet,{header:1,defval:''})[0]||[]).map((v,i)=>[v,i]));
function compareNegotiated(before,after){
 const headers=fields(before),allowed=new Set(['已付金额','是否结清','剩余付款金额','提交付款申请单日期','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号'].map(h=>headers[h]).filter(Number.isInteger));
 for(const address of Object.keys(before).filter(k=>/^[A-Z]+[0-9]+$/.test(k))){
  if(allowed.has(X.utils.decode_cell(address).c))continue;
  const a=before[address],b=after[address];assert.ok(b,'原单元格丢失 '+address);
  if(a.f)assert.equal(b.f,a.f,'原公式改变 '+address);else assert.deepEqual(b.v,a.v,'原内容改变 '+address);
  assert.equal(b.z,a.z,'原数字格式改变 '+address);
  if(a.c)assert.deepEqual(b.c,a.c,'原批注改变 '+address);
 }
 assert.deepEqual(after['!merges'],before['!merges'],'原合并关系改变');
 for(let i=0;i<(before['!rows']||[]).length;i++)if(before['!rows'][i]?.hpt)assert.ok((after['!rows']?.[i]?.hpt||before['!rows'][i].hpt)>=before['!rows'][i].hpt,'原行高缩小');
}
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/',real=process.env.REAL_FUNDS==='1',out=process.env.QA_OUTPUT||'/private/tmp/funds-qa';
 if(real)assert.equal(new URL(base).hostname,'127.0.0.1','真实台账只在本地验证');
 await fs.mkdir(out,{recursive:true});
 const data=real?Object.fromEntries(await Promise.all(Object.entries(REAL_PATHS).map(async([k,n])=>[k,await fs.readFile(path.join('/Users/qingtianzhu/Desktop/资金台账',n))]))):await fixtures();
 const selectedName=real?'采购付款台账_汇总1017（谈判付款）':data.sheetName;
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[];let downloads=0;page.on('pageerror',e=>errors.push(e.message));page.on('download',()=>downloads++);
  await page.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
  await page.route(url=>url.pathname.endsWith('/funds.html')&&!url.searchParams.has('v'),r=>r.fulfill({contentType:'text/html',body:'<p id="stale-funds-page">旧缓存页面，无法上传</p>'}));
  await page.clock.setFixedTime(new Date('2026-10-24T00:30:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});const homepage=page.url();
  await page.locator('.tool-tab[data-tool="funds"]').click();assert.equal(page.url(),homepage);
  let frame=page.frameLocator('iframe[data-tool="funds"]');
  await frame.locator('#ledgerFile').waitFor();assert.match(await page.locator('iframe[data-tool="funds"]').getAttribute('src'),/[?&]v=/);
  async function upload(key,bytes=data[key]){await frame.locator('#'+FILE_IDS[key]+':not([disabled])').waitFor();await frame.locator('#'+FILE_IDS[key]).setInputFiles({name:real?REAL_PATHS[key]:key+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});}
  if(!real){await upload('negotiated',data.noMatch);await frame.locator('#negotiatedInfo').filter({hasText:/谈判付款/}).waitFor();assert(await frame.locator('#analyze').isDisabled());await upload('negotiated',data.single);await frame.locator('#negotiatedSheet option').filter({hasText:'其他月份谈判付款'}).waitFor({state:'attached'});assert.equal(await frame.locator('#negotiatedSheet').inputValue(),'其他月份谈判付款');}
  for(const key of ['ledger','applications','oa','negotiated'])await upload(key);
  await frame.locator('#negotiatedSheet option').filter({hasText:selectedName}).first().waitFor({state:'attached'});
  assert.equal(await frame.locator('#negotiatedSheet').inputValue(),'','多候选必须等待用户选择');
  assert(await frame.locator('#analyze').isDisabled());
  await frame.locator('#negotiatedSheet').selectOption({label:selectedName});
  if(!real){
   const versionsBefore=await page.evaluate(()=>localStorage.getItem('funds-export-versions'));
   await frame.locator('body').evaluate(()=>{window.__fundsExportForQA=window.FundsExport;window.FundsExport=async()=>{throw Error('模拟导出生成失败');};});
   await frame.locator('#analyze:not([disabled])').click();
   await frame.locator('#status').filter({hasText:'模拟导出生成失败'}).waitFor();
   assert(await frame.locator('#download').isDisabled(),'生成失败时禁止下载');assert.equal(downloads,0,'生成失败不能触发下载');
   assert.equal(await page.evaluate(()=>localStorage.getItem('funds-export-versions')),versionsBefore,'生成失败不能递增版本');
   await frame.locator('body').evaluate(()=>{window.FundsExport=window.__fundsExportForQA;delete window.__fundsExportForQA;});
  }
  await frame.locator('#analyze:not([disabled])').click();
  await frame.locator('#normalPreview').waitFor();
  const pending=()=>frame.locator('#pendingCount').innerText().then(s=>Number(s.match(/\d+/)?.[0]||0));
  const initialPending=await pending();
  if(initialPending)assert(await frame.locator('#download').isDisabled(),'待核对时禁止下载');
  async function resolveKeep(){
   for(let count=0;count<800&&await pending();count++){
    const unresolved=frame.locator('[data-decision]');let applied=false;
    for(let i=0;i<await unresolved.count();i++){
     const select=unresolved.nth(i),selected=await select.evaluate(e=>[...e.selectedOptions].map(o=>o.value).filter(Boolean));if(selected.length)continue;
     const values=await select.locator('option').evaluateAll(opts=>opts.map(o=>o.value));assert(values.includes('keep'),'缺少明确保留原值操作');
     await select.selectOption(['keep']);applied=true;break;
    }
    if(!applied)throw Error('待核对事项无法通过明确保留原值解决');
   }
   assert.equal(await pending(),0,'未完成所有核对');
   await frame.locator('#download:not([disabled])').waitFor();
  }
  await resolveKeep();
  const previewText=await frame.locator('#normalPreview').innerText();
  await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();
  assert.equal(await page.locator('#toastArea .toast').count(),0,'旧工具提示不应显示在资金整理页面');
  assert.equal(await frame.locator('#normalPreview').innerText(),previewText);assert.equal(await pending(),0);
  await page.screenshot({path:out+'/page.png',fullPage:true});
  async function download(version){const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;assert.equal(d.suggestedFilename(),`资金整理-2026-10-24-V${version}.xlsx`);const target=out+'/V'+version+'.xlsx';await d.saveAs(target);return fs.readFile(target);}
  const bytes=await download(1);const saved=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),original=X.read(data.negotiated,{type:'buffer',cellNF:true,cellStyles:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款']);compareNegotiated(original.Sheets[selectedName],saved.Sheets.谈判付款);
  const normal=X.utils.sheet_to_json(saved.Sheets.正常付款,{header:1,defval:''});assert.deepEqual(normal[0],NORMAL_HEADERS);
  assert.equal(normal.length-1,real?35:5);assert.equal(saved.Sheets.正常付款['!autofilter'].ref,`A1:V${normal.length}`);
  const metadata=await frame.locator('body').evaluate(async(_,{bytes,source})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),src=await JSZip.loadAsync(new Uint8Array(source)),parse=x=>new DOMParser().parseFromString(x,'application/xml'),all=(doc,n)=>[...doc.getElementsByTagNameNS('*',n)];
   const book=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
   const normal=all(book,'sheet').find(e=>e.getAttribute('name')==='正常付款');const relation=all(rels,'Relationship').find(e=>e.getAttribute('Id')===normal.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
   const target=relation.getAttribute('Target');const file=target.startsWith('/')?target.slice(1):'xl/'+target;
   const sheet=parse(await zip.file(file).async('string'));const style=await zip.file('xl/styles.xml').async('string'),old=await src.file('xl/styles.xml').async('string');
   const styles=parse(style),oldStyles=parse(old);const prefix={};for(const tag of ['fonts','fills','borders','cellXfs']){const a=all(oldStyles,tag)[0],b=all(styles,tag)[0];prefix[tag]=a?[...a.children].every((el,i)=>el.isEqualNode(b.children[i])):true;}
   return {frozen:all(sheet,'pane')[0]?.getAttribute('ySplit'),stylesPreserved:prefix,bodyColor:style.includes('FFFFFFFF')};
  },{bytes:[...bytes],source:[...data.negotiated]});
  assert.equal(metadata.frozen,'1');assert(Object.values(metadata.stylesPreserved).every(Boolean),'原样式定义改变');
  if(!real){
   const n=saved.Sheets.谈判付款,nh=fields(n),row=(r,key)=>n[X.utils.encode_cell({r,c:nh[key]})]?.v;
   assert.equal(row(1,'已付金额'),100);assert.equal(row(1,'剩余付款金额'),200);assert.equal(row(2,'已付金额'),0);assert.equal(row(2,'剩余付款金额'),0);
   assert.equal(row(1,'是否提交付款申请单'),'是');assert.match(String(row(1,'提交付款申请单日期')),/2026.*10.*24/);
   for(const [key,value] of [['当前节点','财务审核'],['未操作者','测试审核人'],['流程状态','未归档'],['申请人','李奇'],['流程编号','MDST-DGFK20261024010']])assert.equal(row(1,key),value);
   assert.equal(row(2,'流程编号'),'MDST-DGFK20261024011','零元流程应正常匹配');
   assert.equal(normal[1][12],0);assert.equal(normal[1][14],100);assert.equal(normal[1][16],'是');assert.equal(normal[1][17],'财务审核');
   assert.equal(normal[1][21],'MDST-DGFK20261024001');
  }
  await download(2);
  await page.reload({waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const key of ['ledger','applications','oa','negotiated'])await upload(key);
  await frame.locator('#negotiatedSheet option').filter({hasText:selectedName}).first().waitFor({state:'attached'});await frame.locator('#negotiatedSheet').selectOption({label:selectedName});await frame.locator('#analyze:not([disabled])').click();await frame.locator('#normalPreview').waitFor();await resolveKeep();await download(3);
  await page.setViewportSize({width:390,height:844});await page.screenshot({path:out+'/mobile.png',fullPage:true});assert(await frame.locator('#download').isVisible());assert.deepEqual(errors,[]);
  if(!real){
   // A wrong source OA number must be reviewable, without replacing the source value.
   const wrongOA='MDST-DGFK20990101999',book=X.read(data.ledger,{type:'buffer'});
   book.Sheets['对账明细-李奇'].D2={t:'s',v:wrongOA};
   const changed={...data,ledger:Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}))};
   const reviewPage=await browser.newPage({viewport:{width:1440,height:1000}});
   try{
    reviewPage.on('pageerror',e=>errors.push(e.message));
    await reviewPage.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
    await reviewPage.clock.setFixedTime(new Date('2026-10-24T00:30:00Z'));
    await reviewPage.goto(base,{waitUntil:'networkidle'});await reviewPage.locator('.tool-tab[data-tool="funds"]').click();
    const review=reviewPage.frameLocator('iframe[data-tool="funds"]');
    for(const key of ['ledger','applications','oa','negotiated']){
     await review.locator('#'+FILE_IDS[key]+':not([disabled])').waitFor();
     await review.locator('#'+FILE_IDS[key]).setInputFiles({name:key+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:changed[key]});
    }
    await review.locator('#negotiatedSheet option').filter({hasText:selectedName}).first().waitFor({state:'attached'});
    await review.locator('#negotiatedSheet').selectOption({label:selectedName});await review.locator('#analyze:not([disabled])').click();
    const oaIssue=review.locator('.issue[data-issue-id="oa:对账明细-李奇!2"]');await oaIssue.waitFor();
    assert(await review.locator('#download').isDisabled());
    const issueText=await oaIssue.innerText();assert.match(issueText,new RegExp(wrongOA));assert.match(issueText,/正常供应商1/);assert.match(issueText,/100/);assert.match(issueText,/MDST-DGFK20261024001/);
    const matchedValue=await oaIssue.locator('option').filter({hasText:'MDST-DGFK20261024001'}).first().getAttribute('value');
    assert.match(matchedValue,/^match:/);await oaIssue.locator('select').selectOption([matchedValue]);
    for(let i=0;i<20&&Number(await review.locator('#pendingCount').innerText());i++)await review.locator('[data-decision]').first().selectOption(['keep']);
    await review.locator('#download:not([disabled])').waitFor();
    const event=reviewPage.waitForEvent('download');await review.locator('#download').click();const downloaded=await event;
    assert.equal(downloaded.suggestedFilename(),'资金整理-2026-10-24-V1.xlsx');const target=out+'/oa-mismatch-confirmed.xlsx';await downloaded.saveAs(target);
    const checked=X.read(await fs.readFile(target),{type:'buffer'}).Sheets.正常付款;
    assert.equal(checked.D2.v,wrongOA,'原付款OA号必须保留');assert.equal(checked.V2.v,'MDST-DGFK20261024001','匹配流程编号单独输出');
    assert.equal(checked.M2.v,0,'OA核对不得改变已付金额');assert.equal(checked.R2.v,'财务审核');
    assert.deepEqual(errors,[]);
   }finally{await reviewPage.close();}
  }
  console.log(JSON.stringify({status:'PASS',real,initialPending,normalRows:normal.length-1,output:out,metadata}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
