const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),{fixtures}=require('./funds-fixtures.cjs');
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',TITLE='付款节奏跟进';
const SHEETS=['谈判付款','正常付款','付款申请核对','付款进度整理','辅助说明','谈判付款-反馈'];
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
(async()=>{
 const base=process.env.BASE_URL||process.env.QA_URL||'http://127.0.0.1:8899/',out=process.env.OUT_DIR||process.env.QA_OUT||'/private/tmp/funds-rename-browser';
 await fs.mkdir(out,{recursive:true});const data=await fixtures(),ledger=X.read(data.ledger,{cellNF:true}),oa=X.read(data.oa,{cellNF:true});
 Object.assign(ledger.Sheets['对账明细-李奇'],{M2:{t:'n',v:40},N2:{t:'s',v:'是'},O2:{t:'n',v:999}});
 oa.Sheets.Sheet0.B3={t:'s',v:'另一审批节点'};
 oa.Sheets.Sheet0['!ref']='A1:G'+(Math.max(...Object.keys(oa.Sheets.Sheet0).filter(a=>/^[A-Z]+\d+$/.test(a)).map(a=>X.utils.decode_cell(a).r))+1);
 data.ledger=write(ledger);data.oa=write(oa);
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.clock.setFixedTime(new Date('2026-11-08T02:00:00Z'));
  await page.addInitScript(()=>{
   if(!localStorage.getItem('funds-export-versions'))localStorage.setItem('funds-export-versions',JSON.stringify({'2026-11-08':9}));
   if(!localStorage.getItem('funds-ledger-export-versions'))localStorage.setItem('funds-ledger-export-versions','{"2026-11-08":31}');
  });
  await page.goto(base,{waitUntil:'networkidle'});
  const menu=page.locator('button[data-tool="funds"]');assert((await menu.innerText()).includes(TITLE));assert(!(await menu.innerText()).includes('资金整理'));
  await menu.click();assert.equal(await page.locator('#toolTitle').textContent(),TITLE);
  assert.equal(await page.locator('iframe[data-tool="funds"]').getAttribute('title'),TITLE);
  let frame=page.frameLocator('iframe[data-tool="funds"]');
  await frame.locator('#ledgerFile:not([disabled])').waitFor();assert.equal(await frame.locator('h1').textContent(),TITLE);
  assert.equal(await frame.locator('body').evaluate(()=>document.title),TITLE+'｜供应链数据清洗');
  assert.equal(await frame.locator('#download').textContent(),'下载'+TITLE+'表');
  assert.equal(await frame.locator('#filenameHint').textContent(),TITLE+'-2026-11-08-V10.xlsx');
  const settlement='settlement:对账明细-李奇!2',oaIssue='oa:对账明细-李奇!2';
  const decision=id=>frame.locator('select[data-decision="'+id+'"]');
  async function setup(){
   for(const role of ['ledger','applications','oa','negotiated']){await frame.locator('#'+role+'File:not([disabled])').waitFor();await frame.locator('#'+role+'File').setInputFiles({name:role+'.xlsx',mimeType:MIME,buffer:data[role]});}
   await frame.locator('#negotiatedSheet option').filter({hasText:data.sheetName}).waitFor({state:'attached'});
   await frame.locator('#negotiatedSheet').selectOption(data.sheetName);await frame.locator('#analyze:not([disabled])').click();await decision(settlement).waitFor();
  }
  async function checkChoices(){
   const options=await decision(settlement).locator('option').evaluateAll(nodes=>nodes.filter(n=>n.value).map(n=>({value:n.value,text:n.textContent,title:n.title})));
   assert.deepEqual(options,[
    {value:'keep',text:'ledger.xlsx：保留原付款情况',title:'ledger.xlsx：保留原付款情况'},
    {value:'recalculate',text:'ledger.xlsx：确认以明确金额重算并更新结清状态',title:'ledger.xlsx：确认以明确金额重算并更新结清状态'}
   ]);
   const matches=await decision(oaIssue).locator('option').evaluateAll(nodes=>nodes.filter(n=>n.value.startsWith('match:')).map(n=>({text:n.textContent,title:n.title})));
   assert.equal(matches.length,2);for(const candidate of matches){assert.match(candidate.text,/^确认匹配：oa\.xlsx \/ Sheet0 第 [23] 行/);assert.equal(candidate.title,candidate.text);assert(!candidate.text.startsWith('ledger.xlsx：'));}
   assert.equal(await decision(oaIssue).locator('option[value="keep"]').textContent(),'ledger.xlsx：保留原流程编号及流程信息');
  }
  async function choose(value){await decision(settlement).selectOption(value);if(await decision(oaIssue).count())await decision(oaIssue).selectOption('keep');await frame.locator('#download:not([disabled])').waitFor();}
  async function download(version){
   const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;
   assert.equal(d.suggestedFilename(),TITLE+'-2026-11-08-V'+version+'.xlsx');const bytes=await fs.readFile(await d.path());
   await fs.writeFile(path.join(out,'V'+version+'.xlsx'),bytes);return X.read(bytes,{cellNF:true,cellStyles:true});
  }
  async function verify(book,choice){
   assert.deepEqual(book.SheetNames,SHEETS);const sheet=book.Sheets.正常付款,range=X.utils.decode_range(sheet['!ref']);
   const headers=Array.from({length:range.e.c+1},(_,c)=>sheet[X.utils.encode_cell({r:0,c})].v);
   const supplier=headers.indexOf('供应商名称'),row=Array.from({length:range.e.r},(_,i)=>i+1).find(r=>sheet[X.utils.encode_cell({r,c:supplier})]?.v==='正常供应商1');
   assert.notEqual(row,undefined);const get=h=>sheet[X.utils.encode_cell({r:row,c:headers.indexOf(h)})];
   assert.equal(get('已付金额').v,40);assert.equal(get('剩余金额').v,choice==='keep'?999:60);assert.equal(get('是否结清').v,choice==='keep'?'是':'否');
   assert.equal(get('本次申请付款金额').v,100);assert.equal(get('对公付款0A流程号').v,'MDST-DGFK20261024001');
   if(choice==='recalculate')assert.equal(get('剩余金额').f,'G'+(row+1)+'-N'+(row+1));
   const preview=await frame.locator('#normalPreview tbody tr').evaluateAll(rows=>rows.map(r=>[...r.cells].map(c=>c.textContent)));
   const visible=preview.find(r=>r[supplier]==='正常供应商1');for(const h of ['已付金额','剩余金额','是否结清'])assert.equal(visible[headers.indexOf(h)],X.utils.format_cell(get(h)));
   for(const name of SHEETS.slice(0,4))assert(book.Sheets[name]['!autofilter']);
  }
  await setup();await checkChoices();await choose('keep');await verify(await download(10),'keep');
  await frame.locator('#onlyPending').uncheck();await checkChoices();await choose('recalculate');await verify(await download(11),'recalculate');
  const builtComments=await frame.locator('body').evaluate((_,payload)=>{
   const input={};for(const [role,bytes]of Object.entries(payload.files))input[role]={name:role+'.xlsx',book:XLSX.read(Uint8Array.from(bytes),{type:'array',cellNF:true})};
   input.negotiated.sheetName=payload.sheet;return Object.values(FundsCore.build(input).applicationComments);
  },{files:Object.fromEntries(['ledger','applications','oa','negotiated'].map(role=>[role,[...data[role]]])),sheet:data.sheetName});
  assert(builtComments.length);assert(builtComments.every(t=>t.startsWith('申请日期（付款节奏跟进）：')));
  await page.screenshot({path:path.join(out,'rename-options.png'),fullPage:true});
  const versionState=await page.evaluate(()=>JSON.parse(localStorage.getItem('funds-export-versions')));assert.equal(versionState['2026-11-08'],11);
  await page.reload({waitUntil:'networkidle'});await page.locator('button[data-tool="funds"]').click();frame=page.frameLocator('iframe[data-tool="funds"]');
  await frame.locator('#ledgerFile:not([disabled])').waitFor();assert.equal(await frame.locator('#filenameHint').textContent(),TITLE+'-2026-11-08-V12.xlsx');
  assert.equal(await frame.locator('#ledgerInfo').textContent(),'尚未引用');await setup();await checkChoices();await choose('recalculate');await verify(await download(12),'recalculate');
  assert.equal(await page.evaluate(()=>localStorage.getItem('funds-ledger-export-versions')),'{"2026-11-08":31}','保留版本键，不串改月度台账序号');

  const noteHeaders=['供应商','付款计划','备注','申请付款金额','是否提交付款申请单'],notes=X.utils.book_new();
  const noteSheet=X.utils.aoa_to_sheet([noteHeaders,['合成批注供应商',10,'原备注',30,'已经提交部分']]);
  noteSheet.B2.c=[{a:'原作者',t:'保留原人工批注\n申请日期（资金整理）：2026年10月1日'}];X.utils.book_append_sheet(notes,noteSheet,'谈判付款');
  const result={applicationComments:{B2:'申请日期（付款节奏跟进）：2026年11月8日'},normalHeaders:['供应商名称','申请付款金额'],normalRows:[],reviewHeaders:['核对说明'],reviewRows:[],negotiated:{sheetName:'谈判付款',headerRow:0,headers:noteHeaders,rows:[noteHeaders.map((_,c)=>noteSheet[X.utils.encode_cell({r:1,c})])],updates:[],columnOrder:[0,1,2,3,4]},highlights:{negotiatedCells:['B2'],normalRows:[]}};
  async function exportNotes(bytes,result){return Buffer.from(await frame.locator('body').evaluate(async(_,{bytes,result})=>[...new Uint8Array(await(await FundsExport(Uint8Array.from(bytes),result)).arrayBuffer())],{bytes:[...bytes],result}));}
  const noteInput=write(notes),migrated=await exportNotes(noteInput,result),first=X.read(migrated).Sheets.谈判付款.B2.c[0];
  assert.equal(first.t,'保留原人工批注\n申请日期（付款节奏跟进）：2026年11月8日');assert.equal(first.a,X.read(noteInput).Sheets.谈判付款.B2.c[0].a);
  const nativeAuthors=await frame.locator('body').evaluate(async(_,bytes)=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),authors=[];
   for(const name of Object.keys(zip.files).filter(p=>/comments[^/]*\.xml$/i.test(p))){const doc=new DOMParser().parseFromString(await zip.file(name).async('string'),'application/xml');authors.push(...[...doc.getElementsByTagNameNS('*','author')].map(n=>n.textContent));}
   return authors;
  },[...migrated]);assert(nativeAuthors.includes('原作者'),'原生批注XML保留中文作者');
  const next={...result,applicationComments:{B2:'申请日期（付款节奏跟进）：2026年11月9日'}};
  const repeated=await exportNotes(migrated,next);assert.equal(X.read(repeated).Sheets.谈判付款.B2.c[0].t,'保留原人工批注\n申请日期（付款节奏跟进）：2026年11月9日');
  const withdrawn=await exportNotes(repeated,{...next,highlights:undefined});assert.equal(X.read(withdrawn).Sheets.谈判付款.B2.c[0].t,'保留原人工批注');
  await fs.writeFile(path.join(out,'legacy-comment-migrated.xlsx'),migrated);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,out,title:TITLE,oldVersionContinues:[10,11,12],settlementKeep:true,settlementRecalculate:true,sourceChoicePrefix:true,matchSourcePreserved:true,legacyCommentMigration:true,fourSheets:true,errors}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});

