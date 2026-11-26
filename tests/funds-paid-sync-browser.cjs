const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {fixtures}=require('./funds-fixtures.cjs');
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
(async()=>{
 const data=await fixtures(),source=X.read(data.negotiated,{cellNF:true}),sheet=source.Sheets[data.sheetName];
 // Fixed plans must stay unchanged when actual payments change.
 sheet.Q2={t:'n',v:220,f:'G2-P2'};
 sheet.N2={t:'n',v:300,f:"'"+data.sheetName+"'!G2-'"+data.sheetName+"'!E2"};
 sheet.N3={t:'n',v:0,f:"'"+data.sheetName+"'!G3-'"+data.sheetName+"'!E3"};
 const zip=await Zip.loadAsync(write(source));let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
 for(const address of ['N2','N3'])xml=xml.replace(new RegExp('(<c r="'+address+'"[^>]*>)([\\s\\S]*?)(</c>)'),(_,a,b,c)=>a+b.replace(/<f>[^<]*<\/f>/,f=>address==='N2'?f.replace('<f>','<f t="shared" si="0" ref="N2:N3">'):'<f t="shared" si="0"/>')+c);
 zip.file('xl/worksheets/sheet1.xml',xml);data.negotiated=await zip.generateAsync({type:'nodebuffer'});
 const original=Buffer.from(data.negotiated),browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const base=process.env.QA_URL||'http://127.0.0.1:8940/',page=await browser.newPage({acceptDownloads:true}),errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  const upload=(role,bytes=data[role])=>frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
  for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
  await frame.locator('#negotiatedSheet').selectOption({label:data.sheetName});
  const download=async()=>{
   await frame.locator('#analyze:not([disabled])').click();
   try{await frame.locator('#download:not([disabled])').waitFor({timeout:30000});}catch(e){throw Error(await frame.locator('#status').innerText());}
   const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;return {bytes:await fs.readFile(await d.path()),name:d.suggestedFilename()};
  };
  const verify=async(output,paid)=>{
   const saved=X.read(output.bytes,{cellNF:true}),s=saved.Sheets.谈判付款,headers=X.utils.sheet_to_json(s,{header:1})[0];
   assert.deepEqual(saved.SheetNames,['谈判付款','谈判付款付款表付款节奏','正常付款','付款申请核对','付款进度整理','辅助说明','谈判付款-反馈']);
   assert.equal(s.E2.v,paid);assert.equal(s.N2.v,300-paid);assert.equal(s.N2.f,"'谈判付款'!G2-'谈判付款'!E2");assert.equal(s.N3.f,"'谈判付款'!G3-'谈判付款'!E3");assert.equal(s.N3.v,0);assert.equal(s.K2.v,paid===300?'是':'否');
   assert.equal(s.E2.t,'n');assert.equal(s.E2.z,'#,##0.00');assert.equal(s.N2.z,'#,##0.00');
   assert.equal(s.Q2.f,'G2-P2');assert.equal(s.Q2.v,220);assert.equal(s.P2.v,80);assert.equal(s.G2.v,300);assert.equal(s.O2.v,'必须保留的人工备注');assert.deepEqual(s['!merges'],sheet['!merges']);
   const previewHeaders=await frame.locator('#negotiatedPreview th').allTextContents(),first=frame.locator('#negotiatedPreview tbody tr').first();
   for(const [h,expected] of [['已付金额',paid],['剩余付款金额',300-paid]])assert.equal(Number((await first.locator('td').nth(previewHeaders.indexOf(h)).innerText()).replace(/,/g,'')),expected);
   assert.equal(await first.locator('td').nth(previewHeaders.indexOf('是否结清')).innerText(),paid===300?'是':'否');
   assert.equal(s[X.utils.encode_cell({r:1,c:headers.indexOf('是否提交付款申请单')})].v,'已经提交部分');
   return saved;
  };
  const first=await download();await verify(first,100);
  const ledger=X.read(data.ledger),l=ledger.Sheets['对账明细-李奇'];l.M3.v=150;const latest=write(ledger);
  await upload('ledger',latest);assert(await frame.locator('#download').isDisabled());const second=await download();await verify(second,150);assert.notEqual(first.name,second.name);
  await upload('negotiated',second.bytes);await frame.locator('#negotiatedSheet').selectOption('谈判付款');const repeated=await download();await verify(repeated,150);
  l.M3={t:'s',v:''};await upload('ledger',write(ledger));const blank=await download();await verify(blank,150);
  l.M3={t:'n',v:300};await upload('ledger',write(ledger));const settled=await download();await verify(settled,300);
  // A different application total is a review, never allocation of the complete paid total to one batch.
  l.H3={t:'n',v:3000};l.M3={t:'n',v:900};await upload('ledger',write(ledger));const mismatch=await download();await verify(mismatch,150);
  const notes=X.utils.sheet_to_json(X.read(mismatch.bytes).Sheets.付款申请核对).map(r=>r.核对说明).filter(n=>n.includes('已付金额同步核对：'));
  assert.equal(notes.length,1);assert.match(notes[0],/300\.00/);assert.match(notes[0],/3000\.00/);assert.match(notes[0],/需要处理/);assert.equal(await frame.locator('#pendingCount').innerText(),'0');
  // One negotiated payment may combine all ledger lines belonging to its OA.
  const combined=X.read(original,{cellNF:true}),combinedSheet=combined.Sheets[data.sheetName],id='MDST-DGFK20260709009',supplier='德州艾丽伦医疗科技有限公司',total=2010760;
  for(const [address,value] of Object.entries({A2:supplier,B2:'迈德斯特',C2:id,D2:total,E2:0,G2:total,K2:'否',P2:90000}))combinedSheet[address]={t:typeof value==='number'?'n':'s',v:value};
  combinedSheet.N2={t:'n',f:'G2-E2',v:total};combinedSheet.Q2={t:'n',f:'G2-P2',v:total-90000};
  const splitLedger=X.read(data.ledger,{cellNF:true}),splitSheet=splitLedger.Sheets['对账明细-邢芳芳'],headers=X.utils.sheet_to_json(splitSheet,{header:1})[0];
  for(const [index,amount] of [942145,1068615].entries()){
   const record={'月份':'2026-07-01','供应商名称':supplier,'付款主体':'迈德斯特','对公付款0A流程号':id,'开票时间':'2026-07-01','开票金额':amount,'摘要':'合成拆笔'+index,'申请付款金额':amount,'应付款时间':'2026-11-25','单据号【财务应付单】':'AP-SPLIT-'+index,'OA合同编号':'HT-SPLIT-'+index,'合同总金额':amount,'已付款金额':'','是否结清':'否','剩余付款金额':amount,'是否谈判付款':'是'};
   X.utils.sheet_add_aoa(splitSheet,[headers.map(h=>record[h]??'')],{origin:-1});
  }
  const splitOA=X.read(data.oa);X.utils.sheet_add_aoa(splitOA.Sheets.Sheet0,[[id,'财务审核','测试审核人','未归档','邢芳芳',supplier,total]],{origin:-1});
  await upload('ledger',write(splitLedger));await upload('oa',write(splitOA));await upload('negotiated',write(combined));await frame.locator('#negotiatedSheet').selectOption(data.sheetName);
  const grouped=await download(),groupedBook=X.read(grouped.bytes,{cellNF:true}),groupedSheet=groupedBook.Sheets.谈判付款;
  assert(!X.utils.sheet_to_json(groupedBook.Sheets.付款申请核对).some(row=>row.供应商名称===supplier&&row.核对说明.includes('已付金额同步核对')));
  assert(!(await frame.locator('#reviewPreview').innerText()).includes('已付金额同步核对'));
  assert.equal(groupedSheet.E2.v,0);assert.equal(groupedSheet.N2.v,total);assert.equal(groupedSheet.N2.f,'G2-E2');assert.equal(groupedSheet.K2.v,'否');assert.equal(groupedSheet.G2.v,total);
  assert.equal(groupedSheet.Q2.f,'G2-P2');assert.equal(groupedSheet.Q2.v,total-90000);assert.equal(groupedSheet.O2.v,'必须保留的人工备注');assert.deepEqual(groupedSheet['!merges'],combinedSheet['!merges']);
  const groupedHeaders=await frame.locator('#negotiatedPreview th').allTextContents(),groupedPreview=frame.locator('#negotiatedPreview tbody tr').first();
  for(const [header,expected] of [['已付金额',0],['剩余付款金额',total]])assert.equal(Number((await groupedPreview.locator('td').nth(groupedHeaders.indexOf(header)).innerText()).replace(/,/g,'')),expected);
  assert.equal(await frame.locator('#pendingCount').innerText(),'0');
  await page.locator('.tool-tab[data-tool="fulfillment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();assert(await frame.locator('#download').isEnabled());
  assert(data.negotiated.equals(original),'source workbook bytes must stay unchanged');assert.deepEqual(errors,[]);
  await frame.locator('#negotiatedPreviewSection').screenshot({path:'/private/tmp/funds-paid-sync-preview.png'});
  await fs.writeFile('/private/tmp/funds-paid-sync-output.xlsx',settled.bytes);
  console.log('PASS: homepage upload/download, latest ledger replacement, repeated uploads, blank preservation, full settlement, unmatched split-amount review, matching ledger splits suppress only mismatch; preview equals reopened values, shared formulas, plans, merges and notes retained');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
