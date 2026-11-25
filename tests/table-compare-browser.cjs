const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js');
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const HEADERS=['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','备注','上周五',''];
const OUTPUT=['分类','采购跟单员','供应商','付款主体','对公付款OA流程号','摘要','旧版申请付款金额','新版申请付款金额','差额','变化','原备注','比对说明','旧版来源','新版来源','备注'];
const VERSION_KEY='table-compare-export-versions',DAY='2026-11-26',OTHER='{"2026-11-26":29}';
const flow=n=>'MDST-DGFK20261126'+String(n).padStart(3,'0');
const row=(supplier,oa,amount,summary='保留原摘要',remark='保留人工备注')=>['正常付款-正常对账表','李奇',supplier,'合成付款主体',oa,summary,amount,remark,999999,'辅助列'];
const write=book=>Buffer.from(X.write(book,{bookType:'xlsx',type:'buffer'}));
function source(rows,name='付款明细'){
 const book=X.utils.book_new(),sheet=X.utils.aoa_to_sheet([HEADERS,...rows]);
 // An auxiliary formula error must not affect the eight business columns.
 if(rows.length)sheet.J2={t:'e',v:7,f:'1/0'};
 X.utils.book_append_sheet(book,sheet,name);return {book,bytes:write(book)};
}
function fixtures(){
 const old=source([
  row('不变供应商','000123456789012345',100),row('增加金额供应商',flow(2),200),row('仅旧供应商',flow(3),300),
  row('补号供应商','未提交OA',3400,'原摘要'),row('多候选供应商',flow(5),10),row('多候选供应商',flow(5),20),
  row('重复明细供应商',flow(6),5),row('重复明细供应商',flow(6),5),row('不同编号供应商',flow(7),70),
  row('空金额供应商',flow(9),''),row('零金额供应商',flow(10),0),row('负金额供应商',flow(11),-10)
 ]),current=source([
  row('新增供应商',flow(12),90),row('不变供应商','000123456789012345',100),row('增加金额供应商',flow(2),250),
  row('补号供应商',flow(4),3400,'最新摘要'),row('多候选供应商',flow(5),30),row('多候选供应商',flow(5),40),
  row('重复明细供应商',flow(6),5),row('重复明细供应商',flow(6),5),row('不同编号供应商',flow(8),70),
  row('空金额供应商',flow(9),''),row('零金额供应商',flow(10),0),row('负金额供应商',flow(11),-5)
 ]);
 return {old,current};
}
async function replaceCell(bytes,address,xml){
 const zip=await Zip.loadAsync(bytes),entry='xl/worksheets/sheet1.xml',original=await zip.file(entry).async('string');
 const pattern=new RegExp('<c\\b(?=[^>]*\\br="'+address+'")[^>]*?(?:/>|>[\\s\\S]*?</c>)');
 assert.match(original,pattern);zip.file(entry,original.replace(pattern,xml));return zip.generateAsync({type:'nodebuffer'});
}
(async()=>{
 const base=process.env.BASE_URL||process.env.QA_URL||'http://127.0.0.1:8899/',out=process.env.OUT_DIR||process.env.QA_OUT||'/private/tmp/table-compare-browser';
 await fs.mkdir(out,{recursive:true});const fixture=fixtures(),replacement=source([row('最新替换唯一记录',flow(50),0)]),errors=[];
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000},timezoneId:'America/Los_Angeles'});
  page.on('pageerror',e=>errors.push(e.message));await page.clock.setFixedTime(new Date('2026-11-25T17:00:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});await page.evaluate(v=>localStorage.setItem('funds-ledger-export-versions',v),OTHER);
  const homepage=page.url();await page.locator('button[data-tool="table-compare"]').click();assert.equal(page.url(),homepage);
  let frame=page.frameLocator('iframe[data-tool="table-compare"]');await frame.locator('#oldFile:not([disabled])').waitFor();
  assert.equal(await frame.locator('input[type=file]').count(),2);assert(await frame.locator('#download').isDisabled());
  async function upload(bytes,name,role='old',surface=frame){
   await surface.locator('#'+role+'File:not([disabled])').setInputFiles({name,mimeType:MIME,buffer:bytes});
   await surface.locator('#'+role+'Reference').filter({hasText:name}).waitFor();
   assert.match(await surface.locator('#'+role+'Reference').innerText(),/引用时间.*2026.*11.*26/);
  }
  const ready=async(surface=frame)=>surface.locator('#download:not([disabled])').waitFor();
  const snapshot=async(surface=frame)=>surface.locator('#preview').evaluate(el=>({headers:[...el.querySelectorAll('th')].map(c=>c.textContent),rows:[...el.querySelectorAll('tbody tr')].filter(r=>r.cells.length===15).map(r=>[...r.cells].map(c=>c.textContent))}));
  async function blocked(pattern){await frame.locator('#status').filter({hasText:pattern}).waitFor();assert(await frame.locator('#download').isDisabled());assert.equal((await snapshot()).rows.length,0);}
  let version=0;
  async function download(label){
   const pending=page.waitForEvent('download');await frame.locator('#download').click();const d=await pending;
   assert.equal(d.suggestedFilename(),'付款台账表格比对-'+DAY+'-V'+(++version)+'.xlsx');
   const bytes=await fs.readFile(await d.path());await fs.writeFile(path.join(out,label+'.xlsx'),bytes);return bytes;
  }
  async function verify(bytes){
   const book=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),sheet=book.Sheets.比对汇总,snap=await snapshot();
   assert.deepEqual(book.SheetNames,['比对汇总']);assert.deepEqual(snap.headers,OUTPUT);
   const range=X.utils.decode_range(sheet['!ref']),matrix=Array.from({length:range.e.r+1},(_,r)=>Array.from({length:15},(_,c)=>sheet[X.utils.encode_cell({r,c})]?.v??''));
   assert.deepEqual(matrix[0],OUTPUT);assert.equal(range.e.c,14);assert.equal(snap.rows.length,matrix.length-1);assert.equal(sheet['!autofilter'].ref,'A1:O'+matrix.length);
   for(let r=1;r<matrix.length;r++)for(let c=0;c<15;c++){
    const cell=sheet[X.utils.encode_cell({r,c})];assert.equal(snap.rows[r-1][c],cell?X.utils.format_cell(cell):'');assert(!cell?.f);
    if([6,7,8].includes(c)&&typeof matrix[r][c]==='number'){assert.equal(cell.t,'n');assert(!/\.$/.test(X.utils.format_cell(cell)),'整数不得有孤立小数点');}
    else if(matrix[r][c]!=='')assert.equal(cell.t,'s');
    if(c===14)assert.equal(matrix[r][c],'','备注必须留给用户填写');
   }
   const zip=await Zip.loadAsync(bytes),xml=await zip.file('xl/worksheets/sheet1.xml').async('string'),styles=await zip.file('xl/styles.xml').async('string');
   assert.match(xml,/<pane\b[^>]*ySplit="1"[^>]*state="frozen"/);assert(!/<(?:f|filterColumn|tableParts)\b/.test(xml));
   assert(!/hidden="1"/.test(xml));assert.match(styles,/FF315B9D/);assert.match(styles,/FF404040/);assert.match(styles,/style="hair"/);assert.match(styles,/style="thin"/);
   const xfs=styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
   for(const address of ['A1','O1',...(matrix.length>1?['E2','G2','H2','I2']:[])]){
    const cell=xml.match(new RegExp('<c\\b(?=[^>]*\\br="'+address+'")[^>]*>'))?.[0];assert(cell,address);
    const xf=xfs[Number(cell.match(/\bs="(\d+)"/)?.[1]||0)];assert(!/wrapText="1"/.test(xf),'表头、OA及金额不能换行：'+address);
   }
   assert.equal(sheet['!cols'].length,15);assert(sheet['!cols'].every(c=>Number.isFinite(c.wch)&&c.wch>=10));return matrix.slice(1);
  }
  function verifyFacts(rows){
   assert.equal(rows.length,17);assert.equal(rows[0][2],'新增供应商');assert.deepEqual(rows.slice(-5).map(r=>r[2]),['仅旧供应商','多候选供应商','多候选供应商','不同编号供应商','空金额供应商']);
   const counts=rows.reduce((a,r)=>(a[r[9]]=(a[r[9]]||0)+1,a),{});assert.deepEqual(counts,{增加:4,不变:5,待核对:6,减少:2});
   const by=name=>rows.filter(r=>r[2]===name);
   assert.deepEqual(by('增加金额供应商')[0].slice(6,10),[200,250,50,'增加']);assert.deepEqual(by('负金额供应商')[0].slice(6,10),[-10,-5,5,'增加']);
   assert.deepEqual(by('仅旧供应商')[0].slice(6,10),[300,'',-300,'减少']);assert.deepEqual(by('新增供应商')[0].slice(6,10),['',90,90,'增加']);
   assert.equal(by('补号供应商')[0][4],flow(4));assert.equal(by('补号供应商')[0][5],'最新摘要');assert.equal(by('补号供应商')[0][9],'不变');assert.match(by('补号供应商')[0][11],/未提交OA/);assert.match(by('补号供应商')[0][11],/原摘要/);
   assert.equal(by('不变供应商')[0][4],'000123456789012345');assert.equal(by('重复明细供应商').length,2);assert(by('重复明细供应商').every(r=>r[9]==='不变'));
   assert(by('多候选供应商').every(r=>r[9]==='待核对'&&r[8]===''));assert(by('空金额供应商').every(r=>r[9]==='待核对'&&r[8]===''));
   assert.deepEqual(by('零金额供应商')[0].slice(6,10),[0,0,0,'不变']);assert(rows.every(r=>r[10]==='保留人工备注'));
   assert(rows.every(r=>r[12]||r[13]));assert(!JSON.stringify(rows).includes('999999'));assert(!JSON.stringify(rows).includes('辅助列'));
  }
  // Upload in reverse order: neither upload order nor filename guesses control the comparison direction.
  await upload(fixture.current.bytes,'新版台账.xlsx','new');assert(await frame.locator('#download').isDisabled());
  await upload(fixture.old.bytes,'旧版台账.xlsx');await ready();verifyFacts(await verify(await download('all-records')));
  assert.match(await frame.locator('#summary').innerText(),/不完整/);
  const initial=await snapshot();await page.locator('button[data-tool="funds-ledger"]').click();await page.locator('button[data-tool="table-compare"]').click();assert.deepEqual(await snapshot(),initial);
  await frame.locator('#preview').evaluate(el=>{el.scrollLeft=el.scrollWidth;});await page.screenshot({path:path.join(out,'comparison-preview.png'),fullPage:true});
  await frame.locator('body').evaluate(()=>{window.__qaExport=window.TableCompareExport;window.TableCompareExport=async()=>{throw Error('合成导出失败');};});
  const beforeVersion=await page.evaluate(k=>localStorage.getItem(k),VERSION_KEY);await frame.locator('#download').click();await frame.locator('#status').filter({hasText:/导出失败/}).waitFor();
  assert.equal(await page.evaluate(k=>localStorage.getItem(k),VERSION_KEY),beforeVersion);await frame.locator('body').evaluate(()=>{window.TableCompareExport=window.__qaExport;delete window.__qaExport;});
  verifyFacts(await verify(await download('after-export-failure')));
  await frame.locator('#clear-new').click();assert(await frame.locator('#download').isDisabled());assert.equal((await snapshot()).rows.length,0);
  await upload(Buffer.from('not an Excel workbook'),'读取失败.xlsx','new');await blocked(/失败|找不到|未找到/);
  await upload(fixture.current.bytes,'恢复新版.xlsx','new');await ready();verifyFacts(await verify(await download('recovered')));
  for(const [xml,label]of [
   ['<c r="G2"><f>1+1</f></c>','数字公式缺缓存'],
   ['<c r="G2" t="str"><f>"1"</f></c>','文本公式缺缓存'],
   ['<c r="G2" t="inlineStr"><f>"1"</f></c>','内联文本公式缺缓存'],
   ['<c r="G2" t="e"><f>1/0</f><v>#DIV/0!</v></c>','公式错误']
  ]){await upload(await replaceCell(fixture.current.bytes,'G2',xml),label+'.xlsx','new');await blocked(/公式|缓存|错误/);}
  await upload(fixture.current.bytes,'恢复公式文件.xlsx','new');await ready();
  const multi=source([row('第一候选',flow(30),1)],'候选甲');X.utils.book_append_sheet(multi.book,source([row('第二候选',flow(31),2)],'候选乙').book.Sheets.候选乙,'候选乙');
  await upload(write(multi.book),'多工作表.xlsx','new');await frame.locator('#newSheet').waitFor({state:'visible'});assert.equal(await frame.locator('#newSheet').inputValue(),'');assert(await frame.locator('#download').isDisabled());
  await frame.locator('#newSheet').selectOption('候选乙');await ready();assert.equal((await snapshot()).rows[0][2],'第二候选');
  await frame.locator('#newSheet').selectOption('候选甲');await ready();assert.equal((await snapshot()).rows[0][2],'第一候选');
  await upload(fixture.current.bytes,'单工作表恢复.xlsx','new');await ready();
  await frame.locator('body').evaluate(()=>{window.__qaRead=File.prototype.arrayBuffer;window.__qaGates={};File.prototype.arrayBuffer=function(){const read=window.__qaRead.call(this);return this.name.startsWith('延迟-')?new Promise((resolve,reject)=>{window.__qaGates[this.name]=async()=>{try{resolve(await read);}catch(e){reject(e);}};}):read;};});
  const release=async names=>frame.locator('body').evaluate(async(_,names)=>{for(const name of names){await window.__qaGates[name]();delete window.__qaGates[name];}await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));},names);
  await upload(fixture.current.bytes,'延迟-旧选择.xlsx','new');assert(await frame.locator('#download').isDisabled());await upload(replacement.bytes,'最新选择.xlsx','new');await ready();
  const latest=await snapshot();assert.equal(latest.rows[0][2],'最新替换唯一记录');await release(['延迟-旧选择.xlsx']);assert.deepEqual(await snapshot(),latest);
  await frame.locator('body').evaluate((_,uploads)=>{for(const u of uploads){const transfer=new DataTransfer();transfer.items.add(new File([Uint8Array.from(u.bytes)],u.name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const input=document.getElementById(u.id);input.files=transfer.files;input.dispatchEvent(new Event('change',{bubbles:true}));}},[
   {id:'oldFile',name:'延迟-旧版.xlsx',bytes:[...fixture.old.bytes]},{id:'newFile',name:'延迟-新版.xlsx',bytes:[...fixture.current.bytes]}
  ]);
  assert(await frame.locator('#download').isDisabled());await release(['延迟-新版.xlsx']);assert(await frame.locator('#download').isDisabled());await release(['延迟-旧版.xlsx']);await ready();verifyFacts(await verify(await download('simultaneous')));
  await upload(fixture.current.bytes,'延迟-移除.xlsx','new');await frame.locator('#clear-new').click();await release(['延迟-移除.xlsx']);assert(await frame.locator('#download').isDisabled());assert.equal((await snapshot()).rows.length,0);
  await frame.locator('body').evaluate(()=>{File.prototype.arrayBuffer=window.__qaRead;delete window.__qaRead;delete window.__qaGates;});
  await upload(fixture.current.bytes,'重新上传.xlsx','new');await ready();assert.equal(await page.evaluate(()=>localStorage.getItem('funds-ledger-export-versions')),OTHER);
  await page.reload({waitUntil:'networkidle'});await frame.locator('#oldFile:not([disabled])').waitFor();assert(await frame.locator('#download').isDisabled());
  await upload(fixture.old.bytes,'刷新旧版.xlsx');await upload(fixture.current.bytes,'刷新新版.xlsx','new');await ready();verifyFacts(await verify(await download('version-after-refresh')));
  const mobile=await browser.newPage({viewport:{width:390,height:844},isMobile:true,deviceScaleFactor:1});mobile.on('pageerror',e=>errors.push(e.message));await mobile.clock.setFixedTime(new Date('2026-11-25T17:00:00Z'));
  await mobile.goto(base,{waitUntil:'networkidle'});await mobile.locator('button[data-tool="table-compare"]').click();const mobileFrame=mobile.frameLocator('iframe[data-tool="table-compare"]');
  await upload(replacement.bytes,'手机旧版.xlsx','old',mobileFrame);await upload(replacement.bytes,'手机新版.xlsx','new',mobileFrame);await ready(mobileFrame);
  const tops=await mobileFrame.locator('.upload').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().top));assert.equal(tops.length,2);assert.equal(tops[0],tops[1]);
  assert(await mobileFrame.locator('.uploads-scroll').evaluate(el=>el.scrollWidth>el.clientWidth));assert(await mobileFrame.locator('body').evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
  assert(await mobile.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));await mobile.screenshot({path:path.join(out,'comparison-mobile.png'),fullPage:true});
  const empty=source([]);await upload(empty.bytes,'空旧版.xlsx');await upload(empty.bytes,'空新版.xlsx','new');await ready();assert.equal((await verify(await download('empty-result'))).length,0);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,out,columns:15,comparison:true,allRecords:true,duplicateMultiplicity:true,ambiguities:true,formulaCacheGuard:true,sourceSelection:true,uploadReplacement:true,latestWins:true,toolSwitch:true,previewExport:true,downloadVersions:true,beijingDate:true,mobile:true,emptyResult:true,errors}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
