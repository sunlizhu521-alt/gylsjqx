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
  function sourceClasses(){for(const row of [2,4])set(input,'applications','付款分类',row,'谈判付款-谈判付款表');set(input,'applications','付款分类',3,'正常付款-正常对账表');}
  sourceClasses();let expectedDates=['2026年11月6日','2026年11月9日','2026年11月6日',''];
  const negotiatedSource=input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];
  X.utils.sheet_add_aoa(negotiatedSource,[['是否谈判分批'],['是']],{origin:'O1'});negotiatedSource.O2.f='"是"';negotiatedSource.O2.c=[{a:'人工',t:'保留分批标记说明'}];
  X.utils.sheet_add_aoa(apps,[['核对说明','是否OA备注','是否已付款'],['','2026-11-03 10:30','2026-11-05'],['','',''],['','是','2026-11-05'],['','','']],{origin:'G1'});
  for(const row of [2,6])set(input,'oa','签字意见',row,'');
  append(input,'applications',['2026-11-05','无流程合成供应商','迈德斯特',ID(99),12,'正常付款']);
  append(input,'applications',['2026-11-06','缺编号合成供应商','迈德斯特','',18,'正常付款']);
  append(input,'applications',['2026-11-07','多流程合成供应商','迈德斯特',ID(1)+' / '+ID(2),20,'正常付款']);
  const untouched=X.utils.aoa_to_sheet([['原内容'],['不得修改']]);untouched.A2.c=[{a:'原作者',t:'保留人工批注'}];X.utils.book_append_sheet(input.negotiated.book,untouched,'原工作表');
  X.utils.book_append_sheet(input.negotiated.book,X.utils.aoa_to_sheet([['旧校验结果'],['本次下载应移除']]),'总额校验核对');
  async function upload(role,bytes=write(input[role].book)){
   if(role==='negotiated')input[role].book=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true});
   await frame.locator('#'+role+'File:not([disabled])').waitFor();await frame.locator('#'+role+'File').setInputFiles({name:input[role].name,mimeType:MIME,buffer:bytes});
  }
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  async function analyze(){await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();assert.equal(await frame.locator('#pendingCount').innerText(),'0');}
  async function snapshot(id){return frame.locator('#'+id).evaluate(el=>({headers:[...el.querySelectorAll('th')].map(e=>e.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(td=>({text:td.textContent,title:td.title,color:getComputedStyle(td).backgroundColor})))}));}
  async function download(version){const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;assert.equal(d.suggestedFilename(),'付款节奏跟进-2026-11-15-V'+version+'.xlsx');return fs.readFile(await d.path());}
  async function verifySummary(out,paperDates=['2026-11-02','2026-11-09'],expectedStates){
   const summary=out.Sheets['谈判付款付款表付款节奏'],headers=['分类','创跟单员','供应商','付款主体','对公付款0A流程号','摘要','OA-申请付款金额','付款日期','申请纸质单日期','是否申请纸质单','是否已付款'],core=C.build(input);
   const source=input.negotiated.book.Sheets[C.negotiatedSheets(input.negotiated.book)[0]],sourceRows=X.utils.sheet_to_json(source,{header:1}).slice(1).filter(row=>row[0]);
   const data=X.utils.sheet_to_json(summary,{header:1}),snap=await snapshot('negotiatedSummaryPreview');
   assert.deepEqual(data[0],headers);assert.deepEqual(snap.headers,headers);assert.equal(data.length-1,core.negotiatedSummaryRows.length);assert.equal(snap.rows.length,core.negotiatedSummaryRows.length);
   assert.equal(sourceRows.length,1);assert.equal(data.length,3,'原谈判一行的两个付款日期展开为两行');assert.deepEqual(data.slice(1).map(row=>row[0]),Array(2).fill(sourceRows[0][14]==='是'?'谈判付款-谈判付款表':sourceRows[0][14]==='否'?'正常付款-谈判付款表':''));
   for(let r=0;r<core.negotiatedSummaryRows.length;r++)for(let c=0;c<headers.length;c++){
    const actual=summary[X.utils.encode_cell({r:r+1,c})],expected=core.negotiatedSummaryRows[r][c];assert.equal(actual?.v??'',expected?.v??'');assert.equal(snap.rows[r][c].text,X.utils.format_cell(actual));
   }
   assert.deepEqual([summary.G2.v,summary.G3.v],[40,60],'每期使用明细金额，不重复整笔OA总额');assert.equal(summary.G2.v+summary.G3.v,100);
   assert.deepEqual([summary.H2,summary.H3].map(c=>X.utils.format_cell(c)),['2026-11-02','2026-11-09']);assert([summary.H2,summary.H3].every(c=>c.t==='n'&&c.z==='yyyy-mm-dd'));
   if(paperDates)assert.deepEqual([summary.I2,summary.I3].map(c=>X.utils.format_cell(c)),paperDates);
   if(expectedStates)assert.deepEqual(data.slice(1).map(row=>row.slice(9)),expectedStates);
  }
  async function verify(bytes,states,colors){
   const out=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),core=C.build(input);
   assert.deepEqual(out.SheetNames,[...C.OUTPUT_SHEETS,'原工作表']);assert.equal(out.Sheets.原工作表.A2.v,'不得修改');assert.equal(out.Sheets.原工作表.A2.c[0].t,'保留人工批注');
   const paidState=value=>({'已付款':'是','未付款':'否'}[value]||'待核对');
   await verifySummary(out,undefined,[['是',paidState(states[0])],states.length===8?['是',paidState(states[7])]:['待核对','待核对']]);
   const progress=out.Sheets.付款进度整理,headers=X.utils.sheet_to_json(progress,{header:1})[0],data=X.utils.sheet_to_json(progress,{defval:''});
   assert.equal(headers.length,13);assert.deepEqual(headers,C.PROGRESS_HEADERS);assert.deepEqual(headers.slice(0,3),['申请日期','付款来源','跟单员']);
   assert.deepEqual(data.slice(0,4).map(r=>r.应付时间),expectedDates);assert.deepEqual(data.map(r=>r.是否已付款),states);assert.equal(data[0].付款来源,'谈判付款表-谈判付款');assert.equal(data[1].付款来源,'资金台账');assert.equal(progress['!autofilter'].ref,'A1:M'+(states.length+1));
   const snap=await snapshot('progressPreview');assert.deepEqual(snap.headers,headers);
   for(let r=0;r<states.length;r++)for(let c=0;c<headers.length;c++){
    const cell=progress[X.utils.encode_cell({r:r+1,c})],paid=headers[c]==='是否已付款'&&states[r]==='已付款';
    assert.equal(snap.rows[r][c].text,X.utils.format_cell(cell));assert.equal(fill(cell)==='FF0000',paid);if(paid)assert.equal(snap.rows[r][c].color,'rgb(255, 0, 0)');
   }
   assert.equal(progress.A2.t,'s');assert.equal(progress.F2.t,'s');assert.equal(progress.G2.t,'n');assert.equal(progress.H2.t,'s');
   const n=out.Sheets.谈判付款;assert.equal(n.F2.f,'D2-E2');assert.equal(n.F2.v,100);assert.equal(n.K2.f,'SUM(M2:N2)');assert.equal(n.I2.c[0].t,'保留原批注');
   assert.equal(n.O2.v,'是');assert.equal(n.O2.f,'"是"');assert.equal(n.O2.c[0].t,'保留分批标记说明');
   const ns=await snapshot('negotiatedPreview');
   for(const [address,color]of Object.entries(colors)){
    assert.equal(fill(n[address]),color,address);const p=X.utils.decode_cell(address);
    assert.equal(ns.rows[p.r-1][p.c].color,color==='FF0000'?'rgb(255, 0, 0)':color==='FFFF00'?'rgb(255, 255, 0)':'rgb(255, 255, 255)');
    if(n[address].c)assert.equal(ns.rows[p.r-1][p.c].title,n[address].c.map(c=>c.t).join('\n').trim());
   }
   const normal=out.Sheets.正常付款,normalRows=X.utils.sheet_to_json(normal);assert.deepEqual(normalRows.map(r=>r.供应商名称),['合成供应商2','合成供应商3']);assert.equal(normalRows[0].本次申请付款金额,200);assert.equal(normalRows[0].应付时间,'2026年11月9日');
   for(let c=0;c<C.NORMAL_HEADERS.length;c++)assert.equal(fill(normal[X.utils.encode_cell({r:1,c})]),'FF0000','正常付款原整行红色保留');
   assert(!out.Sheets.总额校验核对,'下载不含总额校验核对，也不保留源文件中同名旧表');
   const totalSnap=await snapshot('totalValidationPreview');assert.deepEqual(totalSnap.headers,core.totalValidation.headers);assert.equal(totalSnap.rows.length,core.totalValidation.rows.length);
   for(let r=0;r<core.totalValidation.rows.length;r++)for(let c=0;c<core.totalValidation.headers.length;c++)assert.equal(totalSnap.rows[r][c].text,X.utils.format_cell(core.totalValidation.rows[r][c]));
   return out;
  }
  await analyze();const first=await download(1),initial=await verify(first,['已付款','未付款','已付款','未付款','待核对','待核对','待核对'],{M2:'FF0000',N2:''});
  const note=initial.Sheets.谈判付款.M2.c[0].t;assert.match(note,/共2批，第1批/);assert.match(note,/申请日期（付款节奏跟进）：2026年11月1日；2026年11月3日/);assert.match(note,/OA备注时间（付款节奏跟进）：.*2026年11月3日.*10:30/);assert.match(note,/付款时间（付款节奏跟进）：2026年11月5日/);
  const before=await snapshot('progressPreview'),summaryBefore=await snapshot('negotiatedSummaryPreview');await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();assert.deepEqual(await snapshot('progressPreview'),before);assert.deepEqual(await snapshot('negotiatedSummaryPreview'),summaryBefore);
  // Latest application data, not archived OA or a prior generated note, controls negotiated payment status.
  await upload('negotiated',first);set(input,'applications','应付时间',2,'2026-12-31');set(input,'applications','应付时间',4,'2027-01-08');expectedDates=['2026年11月6日','2026年11月9日','2026年11月6日',''];for(const row of [2,4])set(input,'applications','是否已付款',row,'');await upload('applications');await analyze();
  assert.deepEqual(await frame.locator('#negotiatedSheet').evaluate(el=>[...el.options].map(option=>option.value).filter(Boolean)),['谈判付款'],'重新上传结果时，新汇总表不能成为谈判主表候选');
  await verify(await download(2),['待核对','未付款','待核对','未付款','待核对','待核对','待核对'],{M2:'FFFF00',N2:''});
  for(const row of [2,4])set(input,'applications','是否已付款',row,'否');await upload('applications');await analyze();
  await verify(await download(3),['未付款','未付款','未付款','未付款','待核对','待核对','待核对'],{M2:'FFFF00',N2:''});
  for(const row of [2,4]){set(input,'applications','是否OA备注',row,'是');set(input,'applications','是否已付款',row,'是');}await upload('applications');await analyze();
  const yes=await download(4);await verify(yes,['已付款','未付款','已付款','未付款','待核对','待核对','待核对'],{M2:'FF0000',N2:''});
  const yesNote=X.read(yes).Sheets.谈判付款.M2.c[0].t;assert.match(yesNote,/OA备注时间（付款节奏跟进）：已备注[，,]时间未填/);assert.match(yesNote,/付款时间（付款节奏跟进）：已付款[，,]时间未填/);
  append(input,'applications',['2026-11-08','合成供应商1','迈德斯特',ID(1),60,'谈判付款','第2批','2026-11-09','2026-11-10']);set(input,'oa','当前节点',3,'办理人执行');await upload('oa');await upload('applications');await analyze();
  const paid=await download(5);await verify(paid,['已付款','已付款','已付款','未付款','待核对','待核对','待核对','已付款'],{M2:'FF0000',N2:'FF0000'});assert.match(X.read(paid).Sheets.谈判付款.N2.c[0].t,/共2批，第2批/);
  set(input,'applications','核对说明',9,'重新申请');await upload('applications');await analyze();await verify(await download(6),['已付款','已付款','已付款','未付款','待核对','待核对','待核对','重新申请'],{M2:'FF0000'});
  input.applications.book=fixture().applications.book;sourceClasses();expectedDates=['2026年11月6日','2026年11月9日','2026年11月6日',''];await upload('applications');for(const row of [2,6])set(input,'oa','签字意见',row,PAID40);await upload('oa');await analyze();
  await verify(await download(7),['待核对','已付款','待核对','未付款'],{M2:'FFFF00'});
  if(process.env.QA_OUT){await fs.mkdir(process.env.QA_OUT,{recursive:true});await fs.writeFile(process.env.QA_OUT+'/payment-progress.xlsx',paid);await frame.locator('#progressPreviewSection').scrollIntoViewIfNeeded();await frame.locator('#progressPreviewSection').screenshot({path:process.env.QA_OUT+'/payment-progress.png'});await frame.locator('#negotiatedPreviewSection').scrollIntoViewIfNeeded();await frame.locator('#negotiatedPreview').evaluate(e=>{e.scrollLeft=e.scrollWidth;});await frame.locator('#negotiatedPreviewSection').screenshot({path:process.env.QA_OUT+'/batch-colors.png'});}
  // A guessed yellow cell must not become paid merely by re-uploading its generated date note.
  input.negotiated.book=fixture().negotiated.book;input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]].N2.v=40;
  const ambiguousApps=input.applications.book.Sheets[input.applications.book.SheetNames[0]];
  ambiguousApps.G1={t:'s',v:'是否已付款'};ambiguousApps.G2={t:'s',v:'是'};ambiguousApps.G4={t:'s',v:'是'};ambiguousApps['!ref']='A1:Z5';
  await upload('negotiated');await upload('applications');await analyze();const ambiguous=await download(8),ambiguousBook=X.read(ambiguous,{cellStyles:true});
  assert.equal(fill(ambiguousBook.Sheets.谈判付款.M2),'FFFF00');assert.match(ambiguousBook.Sheets.谈判付款.M2.c[0].t,/批次归属待核对/);
  await upload('negotiated',ambiguous);await analyze();const repeated=X.read(await download(9),{cellStyles:true});
  assert.equal(fill(repeated.Sheets.谈判付款.M2),'FFFF00');assert.notEqual(fill(repeated.Sheets.谈判付款.N2),'FF0000');
  X.utils.sheet_add_aoa(input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]],[['是否谈判分批'],['否']],{origin:'O1'});
  for(const row of [2,4])set(input,'applications','付款分类',row,'正常付款-谈判付款表');set(input,'applications','申请日期',2,'2026-11-19');set(input,'applications','申请日期',3,'2026-11-20');set(input,'applications','应付时间',2,'2027-01-26');set(input,'applications','应付时间',3,'');await upload('applications');await upload('negotiated');await analyze();const normalSource=X.read(await download(10),{cellNF:true}),progressRecords=X.utils.sheet_to_json(normalSource.Sheets.付款进度整理);
  assert.equal(progressRecords[0].付款来源,'谈判付款表-正常付款');assert.equal(progressRecords[0].是否已付款,'已付款');
  const finalPreview=await snapshot('progressPreview');assert.equal(finalPreview.rows[0][1].text,'谈判付款表-正常付款');
  assert.equal(normalSource.Sheets.谈判付款.O2.v,'否');assert.equal(progressRecords[0].应付时间,'2027年1月26日');assert.equal(progressRecords[1].应付时间,'');assert.equal(finalPreview.rows[0][7].text,'2027年1月26日');assert.equal(finalPreview.rows[1][7].text,'');
  await verifySummary(normalSource,null,[['待核对','待核对'],['待核对','待核对']]);
  assert.match(JSON.stringify(X.utils.sheet_to_json(normalSource.Sheets.付款申请核对)),/应付时间为空或无法识别/);
  // Moving back across the cutoff must restore the legacy source, not the last generated date.
  set(input,'applications','申请日期',2,'2026-11-18');set(input,'applications','申请日期',3,'2026-11-18');await upload('applications');await analyze();
  const legacy=X.read(await download(11)),legacyRows=X.utils.sheet_to_json(legacy.Sheets.付款进度整理),legacyPreview=await snapshot('progressPreview');
  assert.equal(legacyRows[0].应付时间,'2026年11月6日');assert.equal(legacyRows[1].应付时间,'2026年11月9日');
  assert.equal(legacyPreview.rows[0][7].text,'2026年11月6日');assert.equal(legacyPreview.rows[1][7].text,'2026年11月9日');
  assert(!X.utils.sheet_to_json(legacy.Sheets.付款申请核对).some(r=>r.来源文件==='applications.xlsx'&&[2,3].includes(r.来源行号)&&/应付时间为空或无法识别/.test(r.核对说明)));
  // Payment dates and paper dates retain their distinct sources across a month boundary.
  Object.assign(input,fixture());sourceClasses();const crossMonth=input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]];crossMonth.M1.v-=7;crossMonth.N1.v-=7;
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);await analyze();
  const crossMonthBook=X.read(await download(12),{cellNF:true});await verifySummary(crossMonthBook,['2026-10-26','2026-11-02'],[['是','待核对'],['待核对','待核对']]);
  assert.equal(X.utils.format_cell(crossMonthBook.Sheets.谈判付款.M1),'2026年10月26日');assert.equal(X.utils.format_cell(crossMonthBook.Sheets.谈判付款.N1),'2026年11月2日');
  // Excluding a payment from the summary must retain its source row, even when its summary-only fields need review.
  X.utils.sheet_add_aoa(input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]],[['不支付合成供应商','迈德斯特','不支付',25,0,25,'否',46332,'保留不支付备注','未提供计划日期',25,'',25,'','待确认']],{origin:'A3'});
  input.negotiated.book.Sheets[input.negotiated.book.SheetNames[0]].I3.c=[{a:'人工',t:'不支付原批注必须保留'}];
  async function verifyExcluded(bytes,count){
   const out=X.read(bytes,{cellNF:true,cellStyles:true}),summary=out.Sheets['谈判付款付款表付款节奏'],raw=out.Sheets.谈判付款,data=X.utils.sheet_to_json(summary,{header:1}),snap=await snapshot('negotiatedSummaryPreview');
   assert.deepEqual(out.SheetNames,C.OUTPUT_SHEETS);assert.deepEqual(data[0],C.NEGOTIATED_SUMMARY_HEADERS);assert.equal(data[0].length,11);assert.equal(data.length-1,count);assert.equal(snap.rows.length,count);assert.deepEqual(snap.headers,C.NEGOTIATED_SUMMARY_HEADERS);assert.equal(summary['!autofilter'].ref,'A1:K'+(count+1));
   assert.equal(raw.A3.v,'不支付合成供应商');assert.equal(raw.C3.v,'不支付');assert.equal(raw.I3.v,'保留不支付备注');assert.equal(raw.I3.c[0].t,'不支付原批注必须保留');assert.equal(raw.J3.v,'未提供计划日期');assert.equal(raw.O3.v,'待确认');
   assert(!JSON.stringify(data).includes('不支付合成供应商'));assert(!JSON.stringify(snap).includes('不支付合成供应商'));
   if(count){assert.deepEqual(data.slice(1).map(row=>row[2]),['合成供应商1','合成供应商1']);assert.deepEqual(data.slice(1).map(row=>row[6]),[40,60]);}
   else{assert.equal(summary['!ref'],'A1:K1');assert.equal(raw.C2.v,'不支付');assert.equal(raw.A2.v,'合成供应商1');assert.equal(raw.J2.v,'1、11月2日－40\n2、11月9日－60');}
   return out;
  }
  await upload('negotiated');await analyze();await verifyExcluded(await download(13),2);
  set(input,'negotiated','对公付款0A流程号',2,'不支付');await upload('negotiated');await analyze();const excluded=await download(14);await verifyExcluded(excluded,0);
  await upload('negotiated',excluded);await analyze();await verifyExcluded(await download(15),0);
  assert.deepEqual(await frame.locator('#negotiatedSheet').evaluate(el=>[...el.options].map(option=>option.value).filter(Boolean)),['谈判付款']);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,progressColumns:13,negotiatedSummaryColumns:11,sheets:C.OUTPUT_SHEETS,latestApplicationStatus:true,fourNotes:true,yellowRed:true,normalPreserved:true,summaryPreview:true,summaryReupload:true,summaryInstallments:true,summaryAutoStatus:true,crossMonthDates:true,notPayExcluded:true,emptySummary:true,excludedReupload:true,totalPreviewPreserved:true,totalSheetRemoved:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
