const assert=require('node:assert/strict'),fs=require('node:fs/promises');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,append,set,ID,PAID40}=require('./funds-progress-fixtures.cjs');
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const fill=c=>(c?.s?.fgColor?.rgb||'').replace(/^FF(?=.{6}$)/,'');
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8994/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000}}),errors=[];
  page.on('pageerror',e=>errors.push(e.message));await page.clock.setFixedTime(new Date('2026-11-15T02:00:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  const frame=page.frameLocator('iframe[data-tool="funds"]');await frame.locator('#ledgerFile').waitFor();
  const input=fixture(),apps=input.applications.book.Sheets.付款申请汇总;
  X.utils.sheet_add_aoa(apps,[['核对说明','是否OA备注','是否已付款'],['','2026-11-03 10:30','2026-11-05'],['','',''],['','是','2026-11-05'],['','','']],{origin:'G1'});
  for(const row of [2,6])set(input,'oa','签字意见',row,'');
  append(input,'applications',['2026-11-05','无流程合成供应商','迈德斯特',ID(99),12,'正常付款']);
  append(input,'applications',['2026-11-06','缺编号合成供应商','迈德斯特','',18,'正常付款']);
  append(input,'applications',['2026-11-07','多流程合成供应商','迈德斯特',ID(1)+' / '+ID(2),20,'正常付款']);
  const untouched=X.utils.aoa_to_sheet([['原内容'],['不得修改']]);untouched.A2.c=[{a:'原作者',t:'保留人工批注'}];X.utils.book_append_sheet(input.negotiated.book,untouched,'原工作表');
  async function upload(role,bytes=write(input[role].book)){
   if(role==='negotiated')input[role].book=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true});
   await frame.locator('#'+role+'File:not([disabled])').waitFor();await frame.locator('#'+role+'File').setInputFiles({name:input[role].name,mimeType:MIME,buffer:bytes});
  }
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  async function analyze(){await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();assert.equal(await frame.locator('#pendingCount').innerText(),'0');}
  async function snapshot(id){return frame.locator('#'+id).evaluate(el=>({headers:[...el.querySelectorAll('th')].map(e=>e.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(td=>({text:td.textContent,title:td.title,color:getComputedStyle(td).backgroundColor})))}));}
  async function download(version){const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;assert.equal(d.suggestedFilename(),'付款节奏跟进-2026-11-15-V'+version+'.xlsx');return fs.readFile(await d.path());}
  async function verify(bytes,states,colors){
   const out=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),core=C.build(input);
   assert.deepEqual(out.SheetNames,[...C.OUTPUT_SHEETS,'原工作表']);assert.equal(out.Sheets.原工作表.A2.v,'不得修改');assert.equal(out.Sheets.原工作表.A2.c[0].t,'保留人工批注');
   const progress=out.Sheets.付款进度整理,headers=X.utils.sheet_to_json(progress,{header:1})[0],data=X.utils.sheet_to_json(progress,{defval:''});
   assert.equal(headers.length,13);assert.deepEqual(headers,C.PROGRESS_HEADERS);assert.deepEqual(headers.slice(0,3),['申请日期','付款来源','跟单员']);
   assert.deepEqual(data.map(r=>r.是否已付款),states);assert.equal(data[0].付款来源,'谈判付款表');assert.equal(data[1].付款来源,'资金台账');assert.equal(progress['!autofilter'].ref,'A1:M'+(states.length+1));
   const snap=await snapshot('progressPreview');assert.deepEqual(snap.headers,headers);
   for(let r=0;r<states.length;r++)for(let c=0;c<headers.length;c++){
    const cell=progress[X.utils.encode_cell({r:r+1,c})],paid=headers[c]==='是否已付款'&&states[r]==='已付款';
    assert.equal(snap.rows[r][c].text,X.utils.format_cell(cell));assert.equal(fill(cell)==='FF0000',paid);if(paid)assert.equal(snap.rows[r][c].color,'rgb(255, 0, 0)');
   }
   assert.equal(progress.A2.t,'s');assert.equal(progress.F2.t,'s');assert.equal(progress.G2.t,'n');assert.equal(progress.H2.t,'s');
   const n=out.Sheets.谈判付款;assert.equal(n.F2.f,'D2-E2');assert.equal(n.F2.v,100);assert.equal(n.K2.f,'SUM(M2:N2)');assert.equal(n.I2.c[0].t,'保留原批注');
   const ns=await snapshot('negotiatedPreview');
   for(const [address,color]of Object.entries(colors)){
    assert.equal(fill(n[address]),color,address);const p=X.utils.decode_cell(address);
    assert.equal(ns.rows[p.r-1][p.c].color,color==='FF0000'?'rgb(255, 0, 0)':color==='FFFF00'?'rgb(255, 255, 0)':'rgb(255, 255, 255)');
    if(n[address].c)assert.equal(ns.rows[p.r-1][p.c].title,n[address].c.map(c=>c.t).join('\n').trim());
   }
   const normal=out.Sheets.正常付款,normalRows=X.utils.sheet_to_json(normal);assert.deepEqual(normalRows.map(r=>r.供应商名称),['合成供应商2','合成供应商3']);assert.equal(normalRows[0].本次申请付款金额,200);assert.equal(normalRows[0].应付时间,'2026年11月9日');
   for(let c=0;c<C.NORMAL_HEADERS.length;c++)assert.equal(fill(normal[X.utils.encode_cell({r:1,c})]),'FF0000','正常付款原整行红色保留');
   const total=out.Sheets.总额校验核对,totalSnap=await snapshot('totalValidationPreview');assert.deepEqual(totalSnap.headers,core.totalValidation.headers);assert.equal(X.utils.decode_range(total['!ref']).e.r,core.totalValidation.rows.length);
   for(let r=0;r<core.totalValidation.rows.length;r++)for(let c=0;c<core.totalValidation.headers.length;c++)assert.equal(totalSnap.rows[r][c].text,X.utils.format_cell(total[X.utils.encode_cell({r:r+1,c})]));
   return out;
  }
  await analyze();const first=await download(1),initial=await verify(first,['已付款','未付款','已付款','未付款','待核对','待核对','待核对'],{M2:'FF0000',N2:''});
  const note=initial.Sheets.谈判付款.M2.c[0].t;assert.match(note,/共2批，第1批/);assert.match(note,/申请日期（付款节奏跟进）：2026年11月1日；2026年11月3日/);assert.match(note,/OA备注时间（付款节奏跟进）：.*2026年11月3日.*10:30/);assert.match(note,/付款时间（付款节奏跟进）：2026年11月5日/);
  const before=await snapshot('progressPreview');await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();assert.deepEqual(await snapshot('progressPreview'),before);
  // Latest application data, not archived OA or a prior generated note, controls negotiated payment status.
  await upload('negotiated',first);for(const row of [2,4])set(input,'applications','是否已付款',row,'');await upload('applications');await analyze();
  await verify(await download(2),['待核对','未付款','待核对','未付款','待核对','待核对','待核对'],{M2:'FFFF00',N2:''});
  for(const row of [2,4])set(input,'applications','是否已付款',row,'否');await upload('applications');await analyze();
  await verify(await download(3),['未付款','未付款','未付款','未付款','待核对','待核对','待核对'],{M2:'FFFF00',N2:''});
  for(const row of [2,4]){set(input,'applications','是否OA备注',row,'是');set(input,'applications','是否已付款',row,'是');}await upload('applications');await analyze();
  const yes=await download(4);await verify(yes,['已付款','未付款','已付款','未付款','待核对','待核对','待核对'],{M2:'FF0000',N2:''});
  const yesNote=X.read(yes).Sheets.谈判付款.M2.c[0].t;assert.match(yesNote,/OA备注时间（付款节奏跟进）：已备注[，,]时间未填/);assert.match(yesNote,/付款时间（付款节奏跟进）：已付款[，,]时间未填/);
  append(input,'applications',['2026-11-08','合成供应商1','迈德斯特',ID(1),60,'谈判付款','第2批','2026-11-09','2026-11-10']);set(input,'oa','当前节点',3,'办理人执行');await upload('oa');await upload('applications');await analyze();
  const paid=await download(5);await verify(paid,['已付款','已付款','已付款','未付款','待核对','待核对','待核对','已付款'],{M2:'FF0000',N2:'FF0000'});assert.match(X.read(paid).Sheets.谈判付款.N2.c[0].t,/共2批，第2批/);
  set(input,'applications','核对说明',9,'重新申请');await upload('applications');await analyze();await verify(await download(6),['已付款','已付款','已付款','未付款','待核对','待核对','待核对','重新申请'],{M2:'FF0000'});
  input.applications.book=fixture().applications.book;await upload('applications');for(const row of [2,6])set(input,'oa','签字意见',row,PAID40);await upload('oa');await analyze();
  await verify(await download(7),['待核对','已付款','待核对','未付款'],{M2:'FFFF00'});
  if(process.env.QA_OUT){await fs.mkdir(process.env.QA_OUT,{recursive:true});await fs.writeFile(process.env.QA_OUT+'/payment-progress.xlsx',paid);await frame.locator('#progressPreviewSection').scrollIntoViewIfNeeded();await frame.locator('#progressPreviewSection').screenshot({path:process.env.QA_OUT+'/payment-progress.png'});await frame.locator('#negotiatedPreviewSection').scrollIntoViewIfNeeded();await frame.locator('#negotiatedPreview').evaluate(e=>{e.scrollLeft=e.scrollWidth;});await frame.locator('#negotiatedPreviewSection').screenshot({path:process.env.QA_OUT+'/batch-colors.png'});}
  // A guessed yellow cell must not become paid merely by re-uploading its generated date note.
  input.negotiated.book=fixture().negotiated.book;input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]].N2.v=40;
  const ambiguousApps=input.applications.book.Sheets[input.applications.book.SheetNames[0]];
  ambiguousApps.G1={t:'s',v:'是否已付款'};ambiguousApps.G2={t:'s',v:'是'};ambiguousApps.G4={t:'s',v:'是'};ambiguousApps['!ref']='A1:G5';
  await upload('negotiated');await upload('applications');await analyze();const ambiguous=await download(8),ambiguousBook=X.read(ambiguous,{cellStyles:true});
  assert.equal(fill(ambiguousBook.Sheets.谈判付款.M2),'FFFF00');assert.match(ambiguousBook.Sheets.谈判付款.M2.c[0].t,/批次归属待核对/);
  await upload('negotiated',ambiguous);await analyze();const repeated=X.read(await download(9),{cellStyles:true});
  assert.equal(fill(repeated.Sheets.谈判付款.M2),'FFFF00');assert.notEqual(fill(repeated.Sheets.谈判付款.N2),'FF0000');
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,progressColumns:13,sheets:C.OUTPUT_SHEETS,latestApplicationStatus:true,fourNotes:true,yellowRed:true,normalPreserved:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
