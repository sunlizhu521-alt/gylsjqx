const assert=require('node:assert/strict'),fs=require('node:fs/promises'),X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js'),Core=require('../funds-core.js');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const {fixtures}=require('./funds-fixtures.cjs');
(async()=>{
 const data=await fixtures(),book=X.utils.book_new();
 const headers=['供应商','付款主体','对公付款0A流程号','合同总金额','已付金额','摘要','申请付款金额','单据号【财务应付单】','创建人','是否已确认付款','是否结清','应付款时间','账期','剩余付款金额','贾总反馈备注','孙立柱备注','与张培培沟通记录','谈判明细','分批总金额','核对',46349,46377,46384,46391,46398,46405,46412,46419,46426,46433,46440,46447,46454,'是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号'];
 const sheet=X.utils.aoa_to_sheet([headers,
 ['谈判甲公司','迈德斯特','MDST-DGFK20261024010',300,0,'原摘要',300,'AP00030010','李奇','是','否',46349,'票到30天',300,'保留合并备注','','原人工沟通','1、11月25日－80.00\n2、12月25日－220.00',300,'一致',80,220],
 ['谈判乙公司','迈德斯特','MDST-DGFK20261024011',0,null,'保留空已付',0,'AP00030011','徐英','是','原结清信息',46349,'',null,'','','人工核对','','',null,0],
 ['手工状态冲突供应商','迈德斯特','',40,0,'保留矛盾状态',40,'','孙立柱','','是',46349,'',40,'','','','1、11月25日－40.00',40,'人工确认',40]]);
 sheet.N2={t:'n',v:300,f:'G2-E2'};sheet.S2={t:'n',v:300,f:'SUM(U2:AG2)'};sheet.T2={t:'s',v:'一致',f:'IF(ROUND(S2,2)=ROUND(G2,2),"一致","不一致")'};
 sheet.S3={t:'n',v:0,f:'SUM(U3:AG3)'};sheet.S4={t:'n',v:40,f:'SUM(U4:AG4)'};sheet.T3={t:'s',v:'一致',f:'IF(ROUND(S3,2)=ROUND(G3,2),"一致","不一致")'};
 for(let c=20;c<33;c++)sheet[X.utils.encode_cell({r:0,c})].z=Core.DATE_FORMAT;
 sheet['!merges']=[{s:{r:1,c:14},e:{r:2,c:14}}];sheet['!rows']=[{hpt:30},{hpt:60},{hpt:30},{hpt:30}];sheet['!cols']=headers.map(()=>({wch:24}));sheet['!cols'][17]={wch:52};sheet.XFD117={t:'z'};sheet['!ref']='A1:XFD117';
 X.utils.book_append_sheet(book,sheet,'谈判付款');
 data.negotiated=process.env.FUNDS_V19_FILE?await fs.readFile(process.env.FUNDS_V19_FILE):Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
 if(!process.env.FUNDS_V19_FILE){
  const zip=await Zip.loadAsync(data.negotiated);let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
  for(const [address,attrs,master] of [['S2','t="shared" si="0" ref="S2:S4"',true],['S3','t="shared" si="0"',false],['S4','t="shared" si="0"',false],['T2','t="shared" si="1" ref="T2:T3"',true],['T3','t="shared" si="1"',false]])xml=xml.replace(new RegExp('(<c r="'+address+'"[^>]*>)([\\s\\S]*?)(</c>)'),(_,start,body,end)=>start+body.replace(/<f>[^<]*<\/f>/,f=>master?f.replace('<f>','<f '+attrs+'>'):'<f '+attrs+'/>')+end);
  const dateStyleIds=new Set([...xml.matchAll(/<c r="(?:U|V|W|X|Y|Z|AA|AB|AC|AD|AE|AF|AG)1"[^>]*s="(\d+)"/g)].map(m=>Number(m[1])));
  let styleXml=await zip.file('xl/styles.xml').async('string');
  styleXml=styleXml.replace(/(<cellXfs[^>]*>)([\s\S]*?)(<\/cellXfs>)/,(_,a,body,b)=>{let id=0;return a+body.replace(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g,xf=>dateStyleIds.has(id++)?xf.replace(/numFmtId="\d+"/,'numFmtId="31"'):xf)+b;});
  zip.file('xl/styles.xml',styleXml);
  zip.file('xl/worksheets/sheet1.xml',xml);data.negotiated=await zip.generateAsync({type:'nodebuffer'});
 }
 if(process.env.FUNDS_CLASSIFICATION_QA){
  const appBook=X.read(data.applications),app=appBook.Sheets.付款申请汇总;
  app.K1={t:'s',v:'付款分类'};
  for(let r=2;r<=X.utils.decode_range(app['!ref']).e.r+1;r++){
   if(app['D'+r]?.v==='MDST-DGFK20261024001'){app['K'+r]={t:'s',v:'谈判付款'};app['B'+r]={t:'s',v:'不同申请供应商'};app['J'+r]={t:'n',v:1};}
   if(app['D'+r]?.v==='MDST-DGFK20261024010'){app['K'+r]={t:'s',v:'谈判付款'};app['J'+r]={t:'n',v:400};app['A'+r]={t:'s',v:'异常日期'};}
  }
  app['!ref']='A1:K'+(X.utils.decode_range(app['!ref']).e.r+1);data.applications=Buffer.from(X.write(appBook,{type:'buffer',bookType:'xlsx'}));
 }
 const source=X.read(data.negotiated,{cellStyles:true,cellNF:true}),input={};
 for(const key of ['ledger','applications','oa','negotiated'])input[key]={name:key+'.xlsx',book:X.read(data[key],{cellNF:true})};input.negotiated.sheetName='谈判付款';
 const expected=Core.build(input);assert.equal(expected.negotiated.batchLayout,true);assert.equal(expected.negotiated.headers.length,39);assert.equal(expected.stats.pending,0);
 if(process.env.FUNDS_CLASSIFICATION_QA){assert.equal(expected.normalRows.length,4);assert.equal(expected.reviewRows.length,0);assert.equal(expected.negotiated.rows.length,3);}
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true});const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.goto(process.env.QA_URL||'http://127.0.0.1:8940/',{waitUntil:'networkidle'});
  await page.locator('.tool-tab[data-tool="funds"]').click();const frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const key of ['ledger','applications','oa','negotiated'])await frame.locator('#'+key+'File:not([disabled])').setInputFiles({name:key+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:data[key]});
  await frame.locator('#analyze:not([disabled])').click();try{await frame.locator('#download:not([disabled])').waitFor({timeout:60000});}catch(e){throw Error(await frame.locator('#status').innerText());}
  assert.equal(await frame.locator('#negotiatedPreview th').count(),39);assert.equal(await frame.locator('#negotiatedPreview .batch-detail').first().evaluate(e=>getComputedStyle(e).textAlign),'left');
  assert(await frame.locator('th').evaluateAll(cells=>cells.every(c=>getComputedStyle(c).whiteSpace==='nowrap'&&c.scrollWidth<=c.clientWidth+1)));
  const downloadPromise=page.waitForEvent('download');await frame.locator('#download').click();const download=await downloadPromise;
  const output=await fs.readFile(await download.path()),saved=X.read(output,{cellStyles:true,cellNF:true}),s=saved.Sheets.谈判付款,src=source.Sheets.谈判付款;
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);assert.equal(s['!ref'],'A1:AM'+(expected.negotiated.sourceRange.e.r+1));
  if(process.env.FUNDS_CLASSIFICATION_QA){
   const normal=X.utils.sheet_to_json(saved.Sheets.正常付款);assert.equal(normal.length,4);assert(!normal.some(r=>r['对公付款0A流程号']==='MDST-DGFK20261024001'));
   assert.equal(X.utils.sheet_to_json(saved.Sheets.付款申请核对).length,0);assert.equal(s['!ref'],'A1:AM4');
  }
  const names=X.utils.sheet_to_json(s,{header:1})[0];for(const h of ['提交付款申请单日期','提交时间汇总','付款自然周'])assert(!names.includes(h));
  const updates=new Set(expected.negotiated.updates.map(u=>u.address));
  for(const [address,c] of Object.entries(src)){
   if(!/^[A-Z]+[1-9]\d*$/.test(address)||updates.has(address)||c.v==null&&!c.f)continue;
   assert.equal(s[address]?.v??'',c.v??'',address+' value');assert.equal(s[address]?.f,c.f,address+' formula');
  }
  assert.deepEqual(s['!merges'],src['!merges']);
  if(!process.env.FUNDS_V19_FILE){
   for(const address of ['L2','L3']){assert.equal(s[address].v,46329);assert.equal(s[address].t,'n');assert.equal(X.SSF.format(s[address].z,s[address].v),'2026年11月3日');}
   assert.equal(await frame.locator('#negotiatedPreview tbody tr').first().locator('td').nth(11).innerText(),'2026年11月3日');
  }
  for(let c=20;c<33;c++){
   const address=X.utils.encode_cell({r:0,c}),d=X.SSF.parse_date_code(s[address].v),label=`${d.y}年${d.m}月${d.d}日`;
   assert.equal(X.SSF.format(s[address].z,s[address].v),label,address+' full Chinese date');
   assert(s['!cols'][c].width>=20,address+' date and filter must fit');
  }
  const layout=await frame.locator('body').evaluate(async(_,{bytes})=>{
   const z=await JSZip.loadAsync(new Uint8Array(bytes)),p=new DOMParser(),parse=x=>p.parseFromString(x,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
   const doc=parse(await z.file('xl/worksheets/sheet1.xml').async('string')),styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c]));
   const a=all(xfs.children[Number(cells.get('R2').getAttribute('s'))],'alignment')[0];
   return {left:a.getAttribute('horizontal'),wrap:a.getAttribute('wrapText'),height:all(doc,'row').find(r=>r.getAttribute('r')==='2').getAttribute('ht'),filter:all(doc,'autoFilter')[0].getAttribute('ref')};
  },{bytes:[...output]});
  assert.equal(layout.left,'left');assert.equal(layout.wrap,'1');assert.equal(layout.filter,s['!ref']);assert(Number(layout.height)>=30);if(!process.env.FUNDS_V19_FILE)assert(Number(layout.height)>=60);
  // Switching tools keeps the uploaded V19 and its prepared download.
  await page.locator('.tool-tab:not([data-tool="funds"])').first().click();
  await page.locator('.tool-tab[data-tool="funds"]').click();assert(await frame.locator('#download').isEnabled());
  assert.deepEqual(errors,[]);await fs.writeFile('/private/tmp/funds-v19-browser-output.xlsx',output);
  await frame.locator('#negotiatedPreviewSection').screenshot({path:'/private/tmp/funds-v19-browser-preview.png'});
  console.log('PASS V19 homepage upload → 39 columns → download; original financial values/formulas/status/merges, typed week headers, detail layout and tool switching');
 }finally{await browser.close();}
})().catch(e=>{console.error(e.message);process.exit(1);});
