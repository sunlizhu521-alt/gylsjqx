const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {fixture,append,set,ID}=require('./funds-progress-fixtures.cjs');
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const write=b=>Buffer.from(X.write(b,{type:'buffer',bookType:'xlsx'}));
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/',browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.clock.setFixedTime(new Date('2026-11-01T02:00:00Z'));await page.goto(base,{waitUntil:'networkidle'});
  await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');await frame.locator('#ledgerFile').waitFor();
  const input=fixture();
  append(input,'applications',['2026-11-05','无流程合成供应商','迈德斯特',ID(99),12,'正常付款']);
  append(input,'applications',['2026-11-06','缺编号合成供应商','迈德斯特','',18,'正常付款']);
  append(input,'applications',['2026-11-07','多流程合成供应商','迈德斯特',ID(1)+' / '+ID(2),20,'正常付款']);
  async function upload(role,bytes=write(input[role].book)){await frame.locator('#'+role+'File:not([disabled])').waitFor();await frame.locator('#'+role+'File').setInputFiles({name:input[role].name,mimeType:MIME,buffer:bytes});}
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  async function analyze(){await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();assert.equal(await frame.locator('#pendingCount').innerText(),'0');}
  async function download(version){const pending=page.waitForEvent('download');await frame.locator('#download').click();const d=await pending;assert.equal(d.suggestedFilename(),'资金整理-2026-11-01-V'+version+'.xlsx');return fs.readFile(await d.path());}
  async function preview(){return frame.locator('#progressPreview').evaluate(el=>({headers:[...el.querySelectorAll('th')].map(e=>e.textContent),rows:[...el.querySelectorAll('tbody tr')].map(tr=>[...tr.cells].map(td=>({text:td.textContent,red:td.classList.contains('funds-applied'),color:getComputedStyle(td).backgroundColor})))}));}
  async function verify(bytes,expectedStates){
   const saved=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),core=C.build(input),sheet=saved.Sheets.付款进度整理;
   assert.deepEqual(saved.SheetNames,C.OUTPUT_SHEETS);assert.equal(core.stats.pending,0);
   const expectedReviews=core.reviewRows.map(row=>row.map((c,index)=>index===5?C.formatReviewText(c.v):c.v));
   assert.deepEqual(X.utils.sheet_to_json(saved.Sheets.付款申请核对,{header:1,defval:''}).slice(1),expectedReviews);
   assert.deepEqual(await frame.locator('#reviewPreview tbody tr').evaluateAll(rows=>rows.map(tr=>[...tr.cells].map(td=>td.textContent))),core.reviewRows.map(row=>row.map((c,index)=>index===5?C.formatReviewText(c.v):X.utils.format_cell(c))));
   for(const [index,state]of expectedStates.entries())if(state==='重新申请')assert(!core.reviewRows.some(row=>row[6].v===input.applications.name&&row[8].v===index+2),'重新申请的来源行不进入核对表');
   assert.deepEqual(X.utils.sheet_to_json(sheet,{header:1})[0],C.PROGRESS_HEADERS);assert.equal(X.utils.decode_range(sheet['!ref']).e.r,7);
   const rows=X.utils.sheet_to_json(sheet,{header:1}).slice(1);assert.deepEqual(rows.map(row=>row[11]),expectedStates);assert.deepEqual(rows.map(row=>row[5]),[40,200,40,0,12,18,20]);
   for(let r=0;r<core.progressRows.length;r++)for(let c=0;c<12;c++){
    const actual=sheet[X.utils.encode_cell({r:r+1,c})],expected=core.progressRows[r][c];assert.equal(actual.v,expected.v);assert.equal(actual.t,expected.t);if(expected.z)assert.equal(actual.z,expected.z);
   }
   assert.equal(sheet.A2.t,'n');assert.equal(sheet.G2.t,'n');assert.equal(sheet.F2.t,'n');assert.equal(sheet.E2.t,'s');assert.equal(sheet.F5.v,0);assert.equal(sheet['!autofilter'].ref,'A1:L8');
   assert.equal(X.SSF.format(sheet.G2.z,sheet.G2.v),'2026年11月6日');assert.equal(X.SSF.format(sheet.G3.z,sheet.G3.v),'2026年11月9日');
   const snap=await preview();assert.deepEqual(snap.headers,C.PROGRESS_HEADERS);assert.equal(snap.rows.length,7);
   for(let r=0;r<7;r++)for(let c=0;c<12;c++){
    const cell=sheet[X.utils.encode_cell({r:r+1,c})];assert.equal(snap.rows[r][c].text,X.utils.format_cell(cell));
    const red=c===11&&expectedStates[r]==='已付款';assert.equal(snap.rows[r][c].red,red);if(red)assert.equal(snap.rows[r][c].color,'rgb(255, 0, 0)');
    assert.equal(cell.s?.fgColor?.rgb==='FF0000',red,'仅已付款状态格使用普通红色填充');
   }
   const n=saved.Sheets.谈判付款;assert.equal(n.F2.f,'D2-E2');assert.equal(n.F2.v,100);assert.equal(n.E2.v,0);assert.equal(n.K2.f,'SUM(M2:N2)');assert.equal(n.I2.c[0].t,'保留原批注');assert.equal(n.M2.c[0].t,'申请日期（资金整理）：2026年11月1日；2026年11月3日');
   const normal=saved.Sheets.正常付款,normalRows=X.utils.sheet_to_json(normal);assert.deepEqual(normalRows.map(r=>r['供应商名称']),['合成供应商2','合成供应商3']);assert.deepEqual(normalRows.map(r=>r['申请付款的日期']),['2026/11/02周','2026/11/09周']);
   const normalHeaders=X.utils.sheet_to_json(normal,{header:1})[0];assert.deepEqual(normalHeaders.slice(1,8),['跟单员','供应商名称','付款主体','对公付款0A流程号','摘要','本次申请付款金额','应付时间']);assert.deepEqual(await frame.locator('#normalPreview th').allTextContents(),normalHeaders);
   for(let r=0;r<core.normalRows.length;r++)for(let c=0;c<normalHeaders.length;c++){const actual=normal[X.utils.encode_cell({r:r+1,c})],expected=core.normalRows[r][c];assert.equal(actual.v,expected.v);assert.equal(actual.t,expected.t);if(expected.f)assert.equal(actual.f,expected.f);}
   assert.equal(normal.G2.v,200);assert.equal(normal.G2.z,'#,##0.00');assert.equal(normal.H2.t,'n');assert.equal(X.SSF.format(normal.H2.z,normal.H2.v),'2026年11月9日');assert.equal(normal.P2.f,'G2-N2');assert.equal(normal.P2.v,normal.G2.v-normal.N2.v);
   const meta=await frame.locator('body').evaluate(async(_,bytes)=>{
    const z=await JSZip.loadAsync(Uint8Array.from(bytes)),parse=s=>new DOMParser().parseFromString(s,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
    const wb=parse(await z.file('xl/workbook.xml').async('string')),rels=all(parse(await z.file('xl/_rels/workbook.xml.rels').async('string')),'Relationship'),entry=all(wb,'sheet').find(e=>e.getAttribute('name')==='付款进度整理'),id=entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),target=rels.find(e=>e.getAttribute('Id')===id).getAttribute('Target'),path=new URL(target,'https://x.invalid/xl/workbook.xml').pathname.slice(1),doc=parse(await z.file(path).async('string')),styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0].children,app=parse(await z.file('docProps/app.xml').async('string'));
    return {pane:all(doc,'pane')[0].getAttribute('state'),split:all(doc,'pane')[0].getAttribute('ySplit'),rules:all(doc,'conditionalFormatting').length,borders:all(doc,'c').every(c=>Number(xfs[Number(c.getAttribute('s')||0)].getAttribute('borderId'))>0),headerFills:all(doc,'row')[0].children.length&&[...all(doc,'row')[0].children].map(c=>xfs[Number(c.getAttribute('s'))].getAttribute('fillId')),titles:all(all(app,'TitlesOfParts')[0],'lpstr').map(e=>e.textContent),count:all(all(app,'HeadingPairs')[0],'i4')[0].textContent,widths:all(doc,'col').map(c=>Number(c.getAttribute('width')))};
   },[...bytes]);
   assert.equal(meta.pane,'frozen');assert.equal(meta.split,'1');assert.equal(meta.rules,0);assert(meta.borders);assert.equal(new Set(meta.headerFills).size,1);assert.deepEqual(meta.titles,C.OUTPUT_SHEETS);assert.equal(meta.count,'4');assert.equal(meta.widths.length,12);assert(meta.widths[4]>40);
   return snap;
  }
  await analyze();const bytes=await download(1),snap=await verify(bytes,['已付款','未付款','已付款','未付款','待核对','待核对','待核对']);
  await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();assert.deepEqual(await preview(),snap);
  if(process.env.QA_OUT){await fs.mkdir(process.env.QA_OUT,{recursive:true});await fs.writeFile(process.env.QA_OUT+'/progress.xlsx',bytes);await frame.locator('#progressPreviewSection').scrollIntoViewIfNeeded();await frame.locator('#progressPreviewSection').screenshot({path:process.env.QA_OUT+'/progress.png'});await frame.locator('#progressPreview tbody tr').first().locator('td').last().scrollIntoViewIfNeeded();await frame.locator('#progressPreviewSection').screenshot({path:process.env.QA_OUT+'/progress-status.png'});}
  // A previous export may be reused as the negotiated source, but flow status must come from the latest OA upload.
  await upload('negotiated',bytes);set(input,'oa','流程状态',2,'未归档');set(input,'oa','流程状态',6,'未归档');await upload('oa');await analyze();const second=await download(2);await verify(second,['未付款','未付款','未付款','未付款','待核对','待核对','待核对']);
  // A per-application reapply note takes precedence over archived, unarchived and multi-flow states.
  X.utils.sheet_add_aoa(input.applications.book.Sheets.付款申请汇总,[['核对说明'],['需重新申请'],['重新申请'],[''],[''],[''],[''],['重新申请']],{origin:'G1'});
  set(input,'oa','流程状态',2,'归档');set(input,'oa','流程状态',6,'归档');await upload('oa');await upload('applications');await analyze();await verify(await download(3),['重新申请','重新申请','已付款','未付款','待核对','待核对','重新申请']);
  for(const row of [2,3,8])set(input,'applications','核对说明',row,'');await upload('applications');await analyze();await verify(await download(4),['已付款','未付款','已付款','未付款','待核对','待核对','待核对']);
  for(const row of [2,3,6]){set(input,'oa','当前节点',row,'办理人执行');set(input,'oa','流程状态',row,'未归档');}
  await upload('oa');await analyze();await verify(await download(5),['已付款','已付款','已付款','未付款','待核对','待核对','待核对']);
  assert.equal(await frame.locator('#progressPreview tbody tr').nth(1).locator('td').nth(11).getAttribute('title'),'OA当前节点：办理人执行');
  for(const row of [2,3,6])set(input,'oa','当前节点',row,'财务审批');
  await upload('oa');await analyze();await verify(await download(6),['未付款','未付款','未付款','未付款','待核对','待核对','待核对']);
  assert.deepEqual(errors,[]);console.log(JSON.stringify({base,rows:7,sheets:C.OUTPUT_SHEETS,preview:true,export:true,latestUpload:true,reapply:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
