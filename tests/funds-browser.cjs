const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixtures,NORMAL_HEADERS,REVIEW_HEADERS}=require('./funds-fixtures.cjs');
const REAL_PATHS={ledger:'采购付款台账.xlsx',applications:'付款申请汇总_1024.xlsx',negotiated:'采购付款台账_1020.xlsx',oa:'MDST-对公付款-20261024190903.xlsx'};
const FILE_IDS={ledger:'ledgerFile',applications:'applicationsFile',negotiated:'negotiatedFile',oa:'oaFile'};
const fields=sheet=>Object.fromEntries(Array.from({length:X.utils.decode_range(sheet['!ref']).e.c+1},(_,i)=>[sheet[X.utils.encode_cell({r:0,c:i})]?.v??'',i]));
function compareNegotiated(before,after){
 assert.equal(X.utils.decode_range(after['!ref']).e.r,X.utils.decode_range(before['!ref']).e.r,'谈判表原行范围必须完整保留');
 const headers=fields(before),allowed=new Set(['对公付款0A流程号','对公付款OA流程号','付款OA号','应付款时间','应付时间','提交付款申请单日期','提交时间汇总','是否提交付款申请单','当前节点','未操作者','流程状态','申请人','流程编号','付款自然周'].map(h=>headers[h]).filter(Number.isInteger));
 const supplementOnly=new Set(['对公付款0A流程号','对公付款OA流程号','付款OA号','付款主体','单据号【财务应付单】'].map(h=>headers[h]).filter(Number.isInteger));
 for(const address of Object.keys(before).filter(k=>/^[A-Z]+[0-9]+$/.test(k))){
  const col=X.utils.decode_cell(address).c;if(allowed.has(col))continue;
  const a=before[address],b=after[address];assert.ok(b,'原单元格丢失 '+address);
  const mayFill=supplementOnly.has(col)&&!a.f&&a.t!=='e'&&!String(a.v??'').trim();
  if(a.f)assert.equal(b.f,a.f,'原公式改变 '+address);else if(!mayFill)assert.deepEqual(b.v,a.v,'原内容改变 '+address);
  assert.equal(b.z,a.z,'原数字格式改变 '+address);
  if(a.c)assert.deepEqual(b.c,a.c,'原批注改变 '+address);
 }
 assert.deepEqual(after['!merges'],before['!merges'],'原合并关系改变');
 for(const row of after['!rows']||[])if(row)assert.equal(row.hpt,30,'所有行统一30磅');
}
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/',real=process.env.REAL_FUNDS==='1',out=process.env.QA_OUTPUT||'/private/tmp/funds-qa';
 if(real)assert.equal(new URL(base).hostname,'127.0.0.1','真实台账只在本地验证');
 await fs.mkdir(out,{recursive:true});
 const data=real?Object.fromEntries(await Promise.all(Object.entries(REAL_PATHS).map(async([k,n])=>[k,await fs.readFile(path.join('/Users/qingtianzhu/Desktop/资金台账',n))]))):await fixtures();
 const selectedName=real?'采购付款台账_汇总1017（谈判付款）':data.sheetName;
 const input={};for(const key of ['ledger','applications','negotiated','oa'])input[key]={name:real?REAL_PATHS[key]:key+'.xlsx',book:X.read(data[key],{cellNF:true,cellStyles:true})};input.negotiated.sheetName=selectedName;const coreResult=C.build(input);
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
  assert.equal(await frame.locator('.issue[data-issue-id^="classification:"]').count(),0,'空名称或空金额应直接纳入正常付款，无需分类核对');
  assert.equal(await frame.locator('.issue[data-issue-id^="payment-match:"],.issue[data-issue-id^="identity:"]').count(),0,'谈判表付款信息不依赖台账匹配；名称或主体空白不阻挡整理');
  assert.equal(await frame.locator('.issue[data-issue-id^="application:"]').count(),0,'申请状态自动按供应商和金额填写，不要求人工确认');
  assert.equal(await frame.locator('.issue[data-issue-id^="date:"]').count(),0,'无法识别的登记月份或应付款时间保留原值，不要求人工确认');
  const pending=()=>frame.locator('#pendingCount').innerText().then(s=>Number(s.match(/\d+/)?.[0]||0));
  const initialPending=await pending();
  assert.equal(initialPending,0,'不同OA号不列冲突核对，无匹配申请自动标记未提交');
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
  const classificationText=await frame.locator('#classificationSummary').innerText();
  assert.match(classificationText,new RegExp(`${coreResult.stats.negotiatedLedgerRows} 条[^；。\\n]*识别为谈判付款`));
  assert.match(classificationText,new RegExp(`${real?coreResult.stats.normalRows:5} 条正常付款`));
  assert.match(classificationText,new RegExp(`${real?coreResult.stats.excludedNormalRows:0} 条无对公付款 OA 号未纳入正常付款`));
  assert.match(classificationText,/0 条暂未分类/);
  const previewText=await frame.locator('#normalPreview').innerText();
  assert.deepEqual((await frame.locator('#normalPreview th').allTextContents()).slice(17,20),['是否提交付款申请单','提交时间汇总','当前节点']);
  const previewSubmissions=await frame.locator('#normalPreview tbody tr td:nth-child(19)').allTextContents();
  for(const id of ['normalPreview','negotiatedPreview'])assert(await frame.locator('#'+id+' .flow-code').evaluateAll(cells=>cells.length>0&&cells.every(c=>getComputedStyle(c).whiteSpace==='nowrap'&&getComputedStyle(c).maxWidth==='none'&&!/[\r\n]/.test(c.textContent))),'流程编号预览必须单行自适应');
  assert.equal(await frame.locator('#normalPreview th').nth(24).textContent(),'是否对账');
  assert.equal(await frame.locator('#normalPreview th').nth(25).innerText(),'申请付款的日期');assert.equal(await frame.locator('#normalPreview th').nth(1).innerText(),'跟单员');
  const previewWeeks=await frame.locator('#normalPreview tbody tr td:nth-child(26)').allTextContents();
  const previewReconciled=await frame.locator('#normalPreview tbody tr td:nth-child(25)').allTextContents();
  assert.equal(previewReconciled.filter(v=>v==='已对账').length,real?coreResult.stats.normalRows:5);
  assert.equal(previewReconciled.filter(v=>v==='未对账').length,0);
  await page.locator('.tool-tab[data-tool="payment"]').click();await page.locator('.tool-tab[data-tool="funds"]').click();
  assert.equal(await page.locator('#toastArea .toast').count(),0,'旧工具提示不应显示在资金整理页面');
  assert.equal(await frame.locator('#normalPreview').innerText(),previewText);assert.equal(await pending(),0);
  await page.screenshot({path:out+'/page.png',fullPage:true});
  async function download(version){const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event;assert.equal(d.suggestedFilename(),`资金整理-2026-10-24-V${version}.xlsx`);const target=out+'/V'+version+'.xlsx';await d.saveAs(target);return fs.readFile(target);}
  const bytes=await download(1);const saved=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),original=X.read(data.negotiated,{type:'buffer',cellNF:true,cellStyles:true});
  assert.deepEqual(saved.SheetNames,['谈判付款','正常付款','付款申请核对']);compareNegotiated(original.Sheets[selectedName],saved.Sheets.谈判付款);if(!real){assert.equal(saved.Sheets.谈判付款.L2.v,original.Sheets[selectedName].L2.v);assert.equal(saved.Sheets.谈判付款.L2.z,original.Sheets[selectedName].L2.z,'缺原OA号的日期及格式保留');assert.equal(saved.Sheets.谈判付款.L3.v,46329);assert.equal(X.SSF.format(saved.Sheets.谈判付款.L3.z,46329),'2026年11月3日');}
  const negotiatedHeaders=fields(saved.Sheets.谈判付款),previewColumnNames=await frame.locator('#negotiatedPreview th').allTextContents();assert.equal(Math.abs(negotiatedHeaders['提交付款申请单日期']-negotiatedHeaders['提交时间汇总']),1);for(const key of ['提交付款申请单日期','提交时间汇总','是否提交付款申请单','流程编号'])assert.equal(previewColumnNames.indexOf(key),negotiatedHeaders[key]);
  for(const sheet of Object.values(saved.Sheets)){assert.equal(sheet['!autofilter'].ref,sheet['!ref'],'三个sheet首行筛选覆盖全部列');for(const row of sheet['!rows']||[])if(row){assert(sheet===saved.Sheets.付款申请核对?row.hpt>=30:row.hpt===30);assert(!row.hidden);}assert((sheet['!cols']||[]).slice(0,X.utils.decode_range(sheet['!ref']).e.c+1).every((c,i)=>c.width>=12&&c.width<=(sheet===saved.Sheets.付款申请核对?255:X.utils.sheet_to_json(sheet,{header:1})[0][i]&&['流程编号','付款自然周','申请付款的日期','对公付款0A流程号','开票时间','OA合同编号'].includes(X.utils.sheet_to_json(sheet,{header:1})[0][i])?255:80)),'列宽自适应且有合理范围');}
  for(const [col,min] of [[4,32],[5,22],[11,32]])assert(saved.Sheets.正常付款['!cols'][col].width>=min,'正常付款 E/F/L 必须充分加宽');
  const normal=X.utils.sheet_to_json(saved.Sheets.正常付款,{header:1,defval:''});assert.deepEqual(normal[0],NORMAL_HEADERS);assert.deepEqual(normal.slice(1).map(row=>row[1]),coreResult.normalRows.map(row=>row[1].v));
  assert.deepEqual(normal.slice(1).map(row=>row[18]),previewSubmissions,'提交时间汇总的预览与导出一致');
  assert.deepEqual(normal.slice(1).map(row=>row[25]),previewWeeks,'申请付款的日期预览与导出一致');
  const negotiatedRows=X.utils.sheet_to_json(saved.Sheets.谈判付款,{header:1,defval:''}),weekIndex=negotiatedHeaders['付款自然周'];
  assert.equal(previewColumnNames[weekIndex],'付款自然周');
  assert.deepEqual(negotiatedRows.slice(1,101).map(row=>row[weekIndex]),await frame.locator('#negotiatedPreview tbody tr td:nth-child('+(weekIndex+1)+')').allTextContents(),'谈判付款自然周预览与导出一致');
  if(!real){for(const row of normal.slice(1))assert.equal(row[25],'2026/10/26周');assert.equal(negotiatedRows[1][weekIndex],'2026/10/26周；2026/11/02周');assert.equal(negotiatedRows[2][weekIndex],'日期待确认');}
  const applicationReview=X.utils.sheet_to_json(saved.Sheets.付款申请核对,{header:1,defval:''});assert.deepEqual(applicationReview[0],REVIEW_HEADERS);
  assert.equal(applicationReview.length-1,coreResult.reviewRows.length,'核对导出与核心结果一致');
  assert.deepEqual(applicationReview.slice(1).reduce((counts,row)=>(counts[row[4]]=(counts[row[4]]||0)+1,counts),{}),coreResult.reviewRows.reduce((counts,row)=>(counts[row[4].v]=(counts[row[4].v]||0)+1,counts),{}));
  assert.equal(saved.Sheets.付款申请核对['!autofilter'].ref,`A1:J${applicationReview.length}`);
  assert(applicationReview.slice(1).every(row=>row[5]&&row[6]&&row[7]&&Number.isInteger(row[8])),'付款申请核对须包含说明和来源文件、工作表、行号');
  assert.equal(normal.length-1,real?coreResult.stats.normalRows:5);assert.equal(saved.Sheets.正常付款['!autofilter'].ref,`A1:${X.utils.encode_col(NORMAL_HEADERS.length-1)}${normal.length}`);
  assert(normal.slice(1).every(row=>row[16]==='否'),'正常付款分类应按谈判表重新计算');
  assert.deepEqual(normal.slice(1).map(row=>row[24]),previewReconciled,'导出对账状态必须与预览一致');
  assert(normal.slice(1).every(row=>String(row[4]??'').trim()&&row[24]==='已对账'),'正常付款仅保留原付款 OA 流程号有值的记录');
  const metadata=await frame.locator('body').evaluate(async(_,{bytes,source})=>{
   const zip=await JSZip.loadAsync(new Uint8Array(bytes)),src=await JSZip.loadAsync(new Uint8Array(source)),parse=x=>new DOMParser().parseFromString(x,'application/xml'),all=(doc,n)=>[...doc.getElementsByTagNameNS('*',n)];
   const book=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
   const normal=all(book,'sheet').find(e=>e.getAttribute('name')==='正常付款');const relation=all(rels,'Relationship').find(e=>e.getAttribute('Id')===normal.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
   const target=relation.getAttribute('Target');const file=target.startsWith('/')?target.slice(1):'xl/'+target;
   const sheet=parse(await zip.file(file).async('string'));const style=await zip.file('xl/styles.xml').async('string'),old=await src.file('xl/styles.xml').async('string');
   const styles=parse(style),oldStyles=parse(old);const prefix={};for(const tag of ['fonts','fills','borders','cellXfs']){const a=all(oldStyles,tag)[0],b=all(styles,tag)[0];prefix[tag]=a?[...a.children].every((el,i)=>el.isEqualNode(b.children[i])):true;}
   const xfs=all(styles,'cellXfs')[0],borders=all(styles,'borders')[0],fills=all(styles,'fills')[0],shared=zip.file('xl/sharedStrings.xml')?all(parse(await zip.file('xl/sharedStrings.xml').async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[],layoutErrors=[];
   const text=c=>c.getAttribute('t')==='s'?shared[Number(all(c,'v')[0]?.textContent)]||'':c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent).join(''):all(c,'v')[0]?.textContent??'';
   for(const entry of all(book,'sheet')){
    const relation=all(rels,'Relationship').find(e=>e.getAttribute('Id')===entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')),doc=parse(await zip.file(new URL(relation.getAttribute('Target'),'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1)).async('string')),cells=all(doc,'c'),flowColumns=new Set(cells.filter(c=>/^[A-Z]+1$/.test(c.getAttribute('r'))&&['流程编号','付款自然周','申请付款的日期',...(entry.getAttribute('name')==='正常付款'?['对公付款0A流程号','开票时间','OA合同编号']:[])].includes(text(c))).map(c=>c.getAttribute('r').replace(/\d+/,'')));
    for(const c of cells){const ref=c.getAttribute('r'),xf=xfs.children[Number(c.getAttribute('s')||0)],border=borders.children[Number(xf.getAttribute('borderId')||0)],fill=fills.children[Number(xf.getAttribute('fillId')||0)],header=/^[A-Z]+1$/.test(ref);
     if((text(c)!==''||all(c,'f').length)&&['left','right','top','bottom'].some(edge=>!all(border,edge)[0]?.getAttribute('style')||all(border,edge)[0]?.getAttribute('style')==='none'))layoutErrors.push(entry.getAttribute('name')+'!'+ref+' 缺少边框');
     if((text(c)!==''||all(c,'f').length)&&['left','right','top','bottom'].some(edge=>all(all(border,edge)[0],'color')[0]?.getAttribute('rgb')!=='FF000000'))layoutErrors.push(entry.getAttribute('name')+'!'+ref+' 边框不是清晰黑色');
     if(header&&all(fill,'fgColor')[0]?.getAttribute('rgb')!=='FF315B9D')layoutErrors.push(entry.getAttribute('name')+'!'+ref+' 表头底色不同');
     if(flowColumns.has(ref.replace(/\d+/,''))&&(all(xf,'alignment')[0]?.getAttribute('wrapText')!=='0'||/[\r\n]/.test(text(c))))layoutErrors.push(entry.getAttribute('name')+'!'+ref+' 流程编号换行');
    }
    const populatedColumns=new Set(cells.filter(c=>text(c)!==''||all(c,'f').length).map(c=>c.getAttribute('r').replace(/\d+/,''))),byRef=new Map(cells.map(c=>[c.getAttribute('r'),c])),lastRow=XLSX.utils.decode_range(all(doc,'dimension')[0].getAttribute('ref')).e.r+1;
    for(const col of populatedColumns)for(let r=1;r<=lastRow;r++){const cell=byRef.get(col+r),xf=cell&&xfs.children[Number(cell.getAttribute('s')||0)],border=xf&&borders.children[Number(xf.getAttribute('borderId')||0)];if(!border||['left','right','top','bottom'].some(edge=>!all(border,edge)[0]?.getAttribute('style')||all(border,edge)[0]?.getAttribute('style')==='none'))layoutErrors.push(entry.getAttribute('name')+'!'+col+r+' 空白格边框不连续');}
   }
   return {frozen:all(sheet,'pane')[0]?.getAttribute('ySplit'),stylesPreserved:prefix,bodyColor:style.includes('FFFFFFFF'),layoutErrors};
  },{bytes:[...bytes],source:[...data.negotiated]});
  assert.equal(metadata.frozen,'1');assert(Object.values(metadata.stylesPreserved).every(Boolean),'原样式定义改变');
  assert.deepEqual(metadata.layoutErrors,[],'导出边框、表头颜色、流程编号单行格式');
  if(!real){
   assert.deepEqual(normal.slice(1).map(row=>row[2]),['正常供应商1','正常供应商2','正常供应商3','正常供应商4','正常供应商5'],'否/空标记且未命中谈判的正常记录应保留，谈判甲乙不应混入正常表');
   const n=saved.Sheets.谈判付款,nh=fields(n),row=(r,key)=>n[X.utils.encode_cell({r,c:nh[key]})]?.v;
   assert.equal(row(1,'已付金额'),0,'不得用台账已付 100 覆盖谈判原已付 0');assert.equal(row(1,'剩余付款金额'),300);assert.equal(row(1,'是否结清'),'否');assert.equal(row(2,'已付金额'),0);assert.equal(row(2,'剩余付款金额'),0);assert.equal(row(2,'是否结清'),'是');
   for(const [key,value] of [['付款主体','迈德斯特'],['对公付款0A流程号','MDST-DGFK20261024010'],['单据号【财务应付单】','AP00030010']])assert.equal(row(1,key),value,'唯一且无冲突的台账来源应补充空字段 '+key);
   assert.equal(n.Q2.f,'N2-P2');assert.equal(n.Q2.v,220,'原付款计划公式结果应基于谈判表本身金额');
   assert.equal(row(1,'是否提交付款申请单'),'已经提交部分');assert.equal(row(1,'提交付款申请单日期'),'2026-10-23\n2026-10-24');
   assert.equal(row(1,'提交时间汇总'),'2026-10-23-100.00\n2026-10-24-80.00','按日期汇总每笔金额，不叠加原999.00历史摘要');
   assert.equal(row(2,'是否提交付款申请单'),'已经提交全部');assert.equal(row(2,'提交时间汇总'),'2026-10-24-0.00');
   const previewHeaders=await frame.locator('#negotiatedPreview th').allTextContents();assert.equal(await frame.locator('#negotiatedPreview tbody tr').first().locator('td').nth(previewHeaders.indexOf('提交时间汇总')).innerText(),'2026-10-23-100.00\n2026-10-24-80.00');
   for(const [key,value] of [['当前节点','财务审核'],['未操作者','测试审核人'],['流程状态','未归档'],['申请人','李奇'],['流程编号','MDST-DGFK20261024010']])assert.equal(row(1,key),value);
   assert.equal(row(2,'流程编号'),'MDST-DGFK20261024011','零元流程应正常匹配');
   assert.equal(normal[1][13],0);assert.equal(normal[1][15],100);assert.equal(normal[1][17],'已提交');assert.equal(normal[1][19],'财务审核');
   assert.equal(normal[1][23],'MDST-DGFK20261024001');
   assert.deepEqual(normal.slice(1).map(row=>row[18]),['2026-10-24-100.00','2026-10-24-101.00','2026-10-24-102.00','',''],'正常付款只显示名称与金额精确匹配的申请日期及金额');
   assert.equal(applicationReview.length,2,'原谈判OA空白的唯一匹配补齐列一条复核');
   assert.equal(await frame.locator('#reviewPreview tbody tr').count(),1,'未申请本身不显示，OA补齐列入复核');
   assert.equal(normal[4][17],'未申请','未匹配时按最新规则填写未申请，覆盖原是');assert.equal(normal[5][17],'未申请','部分金额未精确匹配时填写未申请');
  }else{
   const nh=fields(saved.Sheets.谈判付款),states=X.utils.sheet_to_json(saved.Sheets.谈判付款,{header:1,defval:''}).slice(1).map(row=>row[nh['是否提交付款申请单']]).filter(Boolean);
   assert.deepEqual(states.reduce((counts,state)=>(counts[state]=(counts[state]||0)+1,counts),{}),coreResult.negotiated.rows.reduce((counts,row)=>{const state=row[coreResult.negotiated.headers.indexOf('是否提交付款申请单')]?.v;if(state)counts[state]=(counts[state]||0)+1;return counts;},{}));
   assert.equal(applicationReview.slice(1).filter(row=>row[5].includes('同一流程对应多笔付款')).length,0,'同一流程下唯一名称金额可区分两笔申请，不重复分摊');
   assert.equal(applicationReview.slice(1).filter(row=>row[5].includes('申请超额')).length,0,'明确分类为谈判付款的申请不列入核对表');
   assert.equal(applicationReview.slice(1).filter(row=>row[5].includes('OA 号冲突')).length,0);assert.equal(applicationReview.slice(1).filter(row=>row[5].startsWith('1、重复提交：')).length,1);
   assert.equal(applicationReview[0].includes('候选付款申请'),false);
  }
  await download(2);
  await page.reload({waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const key of ['ledger','applications','oa','negotiated'])await upload(key);
  await frame.locator('#negotiatedSheet option').filter({hasText:selectedName}).first().waitFor({state:'attached'});await frame.locator('#negotiatedSheet').selectOption({label:selectedName});await frame.locator('#analyze:not([disabled])').click();await frame.locator('#normalPreview').waitFor();await resolveKeep();await download(3);
  await page.setViewportSize({width:390,height:844});await page.screenshot({path:out+'/mobile.png',fullPage:true});assert(await frame.locator('#download').isVisible());assert.deepEqual(errors,[]);
  if(!real){
   // Installments linked by the application's payment OA reference are accumulated once per source row.
   const applicationBook=X.read(data.applications,{type:'buffer'}),applicationName=applicationBook.SheetNames[0];
   const fullRows=X.utils.sheet_to_json(applicationBook.Sheets[applicationName],{header:1,defval:''});
   fullRows.find(row=>row[1]==='谈判甲公司'&&row[9]===80)[9]=200;
   for(const [version,extra,state,summary] of [[4,false,'已经提交全部','2026-10-23-100.00\n2026-10-24-200.00'],[5,true,'已经提交全部','2026-10-23-100.00\n2026-10-24-200.00\n2026-10-24-10.00']]){
    const rows=fullRows.map(row=>row.slice());if(extra){const duplicate=rows.find(row=>row[1]==='谈判甲公司'&&row[9]===200).slice();duplicate[9]=10;rows.push(duplicate);}
    const changed=X.utils.book_new();X.utils.book_append_sheet(changed,X.utils.aoa_to_sheet(rows),applicationName);
    await upload('applications',Buffer.from(X.write(changed,{type:'buffer',bookType:'xlsx'})));await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();assert.equal(await pending(),0,'申请超额记入核对表，不阻挡下载');
    const checked=X.read(await download(version),{type:'buffer',cellNF:true,cellStyles:true}),n=checked.Sheets.谈判付款,h=fields(n);
    assert.equal(n[X.utils.encode_cell({r:1,c:h['是否提交付款申请单']})].v,state);assert.equal(n[X.utils.encode_cell({r:1,c:h['提交时间汇总']})].v,summary);compareNegotiated(original.Sheets[selectedName],n);
    const notes=X.utils.sheet_to_json(checked.Sheets.付款申请核对);assert.equal(notes.length,extra?2:1);
    if(extra){const note=notes.find(row=>row['核对说明'].includes('申请超额'));assert.equal(note['是否提交付款申请单'],'已提交');assert.match(note['核对说明'],/310\.00/);assert.match(note['核对说明'],/300\.00/);}
   }
   // A wrong nonempty OA may match a new flow; a blank original OA excludes the normal record entirely.
   for(const [sourceOA,unmatched,negotiatedMissing] of [['MDST-DGFK20990101999',false],['',false],['MDST-DGFK20261024001',true,false],['MDST-DGFK20261024001',true,true]]){
   const book=X.read(data.ledger,{type:'buffer'});
   book.Sheets['对账明细-李奇'].D2={t:'s',v:sourceOA};
   const retained=['原审核节点','原未操作者','原流程状态','原申请人','原流程编号'];
   if(unmatched)X.utils.sheet_add_aoa(book.Sheets['对账明细-李奇'],[['当前节点','未操作者','流程状态','申请人','流程编号'],retained],{origin:'R1'});
   const changed={...data,ledger:Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}))};
   if(unmatched){
    const oa=X.read(data.oa,{type:'buffer'}),sheet=oa.Sheets[oa.SheetNames[0]];
    for(const address of Object.keys(sheet).filter(key=>/^A\d+$/.test(key))){
     if(sheet[address].v==='MDST-DGFK20261024001')sheet['G'+address.slice(1)]={t:'n',v:999};
     if(sheet[address].v==='MDST-DGFK20261024010'){
      if(negotiatedMissing){sheet[address]={t:'s',v:'MDST-DGFK20990101999'};sheet['G'+address.slice(1)]={t:'n',v:999};}
      else{sheet['G'+address.slice(1)]={t:'n',v:999};sheet['F'+address.slice(1)]={t:'s',v:'不同收款单位'};}
     }
    }
    changed.oa=Buffer.from(X.write(oa,{type:'buffer',bookType:'xlsx'}));
    const apps=X.read(data.applications,{type:'buffer'}),appSheet=apps.Sheets[apps.SheetNames[0]];for(const address of Object.keys(appSheet).filter(key=>/^B\d+$/.test(key)))if(appSheet[address].v==='谈判甲公司')appSheet[address]={t:'s',v:'申请表中不同供应商名称'};changed.applications=Buffer.from(X.write(apps,{type:'buffer',bookType:'xlsx'}));
    const negotiated=X.read(data.negotiated,{type:'buffer'}),n=negotiated.Sheets[selectedName],col=X.utils.decode_range(n['!ref']).e.c+1;
    X.utils.sheet_add_aoa(n,[['当前节点','未操作者','流程状态','申请人','流程编号'],retained],{origin:{r:0,c:col}});
    changed.negotiated=Buffer.from(X.write(negotiated,{type:'buffer',bookType:'xlsx'}));
   }
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
    await review.locator('#download:not([disabled])').waitFor();
    assert.equal(await review.locator('.issue[data-issue-id="oa:对账明细-李奇!2"]').count(),0,'原OA错误可精确匹配，原OA空白记录直接排除');
    assert.equal(Number(await review.locator('#pendingCount').innerText()),0);
    const previewRows=review.locator('#normalPreview tbody tr');
    assert.equal(await previewRows.count(),sourceOA?5:4);
    assert.match(await review.locator('#classificationSummary').innerText(),new RegExp(`${sourceOA?0:1} 条无对公付款 OA 号未纳入正常付款`));
    if(sourceOA)assert.equal(await previewRows.first().locator('td').nth(23).innerText(),unmatched?'原流程编号':'MDST-DGFK20261024001');
    else assert.equal(await previewRows.filter({hasText:'正常供应商1'}).count(),0,'有匹配新流程也不能纳入原OA空白的记录');
    assert.equal(await previewRows.first().locator('td').nth(24).textContent(),'已对账');
    const event=reviewPage.waitForEvent('download');await review.locator('#download').click();const downloaded=await event;
    assert.equal(downloaded.suggestedFilename(),'资金整理-2026-10-24-V1.xlsx');const target=out+(negotiatedMissing?'/negotiated-oa-missing.xlsx':unmatched?'/oa-unmatched-preserved.xlsx':sourceOA?'/oa-mismatch-auto.xlsx':'/oa-empty-excluded.xlsx');await downloaded.saveAs(target);
    const downloadedBook=X.read(await fs.readFile(target),{type:'buffer',cellNF:true,cellStyles:true}),checked=downloadedBook.Sheets.正常付款;
    assert.equal(checked.Y2.v,'已对账');
    if(sourceOA){
     assert.equal(checked.E2.v,unmatched?sourceOA:'MDST-DGFK20261024001','唯一匹配纠正OA号，未匹配保留');assert.equal(checked.X2.v,unmatched?'原流程编号':'MDST-DGFK20261024001','精确匹配才更新流程编号，未匹配保留原值');
     assert.equal(checked.N2.v,0,'OA核对不得改变已付金额');assert.equal(checked.T2?.v??'',unmatched?'原审核节点':'财务审核');
    }else{
     const rows=X.utils.sheet_to_json(checked);assert.equal(rows.length,4);assert.equal(rows.some(row=>row['供应商名称']==='正常供应商1'),false);
     assert(rows.every(row=>row['对公付款0A流程号']&&row['是否对账']==='已对账'));
     assert.equal(checked.P2.f,'I2-N2','排除原首行后，余额公式使用实际输出行');
     assert.equal(X.utils.sheet_to_json(downloadedBook.Sheets.付款申请核对).some(row=>row['供应商名称']==='正常供应商1'),false);
     compareNegotiated(original.Sheets[selectedName],downloadedBook.Sheets.谈判付款);
    }
    if(unmatched){
     assert.equal(await review.locator('.issue[data-issue-id^="oa:"]').count(),0);
     assert.deepEqual(['当前节点','未操作者','流程状态','申请人','流程编号'].map(key=>checked[X.utils.encode_cell({r:1,c:NORMAL_HEADERS.indexOf(key)})].v),retained);
     const n=downloadedBook.Sheets.谈判付款,h=fields(n),expected=retained;
     const previewHeaders=await review.locator('#negotiatedPreview th').allTextContents();
     for(const [index,key] of ['当前节点','未操作者','流程状态','申请人','流程编号'].entries()){
      assert.equal(n[X.utils.encode_cell({r:1,c:h[key]})].v,expected[index]);
      assert.equal(await review.locator('#negotiatedPreview tbody tr').first().locator('td').nth(previewHeaders.indexOf(key)).innerText(),expected[index]);
     }
     assert.equal(n.C2?.v??'',negotiatedMissing?'MDST-DGFK20261024010':'','跨供应商冲突不能沿用台账补入的OA');
     assert.equal(n[X.utils.encode_cell({r:1,c:h['是否提交付款申请单']})]?.v??'',negotiatedMissing?'已经提交部分':'','跨供应商OA冲突保留原提交状态');assert.equal(n[X.utils.encode_cell({r:1,c:h['提交时间汇总']})].v,negotiatedMissing?'2026-10-23-100.00\n2026-10-24-80.00':'2026-09-01-999.00');if(!negotiatedMissing)assert.match(await review.locator('#reviewPreview').innerText(),/OA供应商冲突/);
     assert.equal(n.E2.v,0);assert.equal(n.N2.v,300);assert.equal(n.K2.v,'否');
    }
    assert.deepEqual(errors,[]);
   }finally{await reviewPage.close();}
   }
  }
  if(!real){
   const invalidBook=X.read(data.applications),invalidSheet=invalidBook.Sheets.付款申请汇总;
   for(let r=2;r<=X.utils.decode_range(invalidSheet['!ref']).e.r+1;r++)if(invalidSheet['D'+r]?.v==='MDST-DGFK20261024010')invalidSheet['J'+r]={t:'s',v:'金额待确认'};
   for(const key of ['ledger','applications','oa','negotiated'])await upload(key,key==='applications'?Buffer.from(X.write(invalidBook,{type:'buffer',bookType:'xlsx'})):data[key]);
   await frame.locator('#negotiatedSheet').selectOption({label:selectedName});await frame.locator('#analyze:not([disabled])').click();
   const choice=frame.locator('.issue[data-issue-id^="submission-status:"] select');await choice.waitFor();assert(await frame.locator('#download').isDisabled());
   assert.deepEqual(await choice.locator('option').evaluateAll(opts=>opts.map(o=>o.value).filter(Boolean)),['已经提交全部','已经提交部分','未提交']);
   await frame.locator('#onlyPending').uncheck();
   for(const status of ['已经提交全部','已经提交部分','未提交']){
    await choice.selectOption(status);await frame.locator('#download:not([disabled])').waitFor();
    const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event,checked=X.read(await fs.readFile(await d.path())),n=checked.Sheets.谈判付款,h=fields(n);
    assert.equal(n[X.utils.encode_cell({r:1,c:h['是否提交付款申请单']})].v,status);
    const states=X.utils.sheet_to_json(n).filter(row=>row['供应商']!=='合计').map(row=>row['是否提交付款申请单']);assert(states.every(v=>['已经提交全部','已经提交部分','未提交'].includes(v)));
    assert.equal(n.E2.v,0);assert.equal(n.N2.v,300);assert.equal(n.K2.v,'否');
   }
  }
  if(!real){
   const flagged=X.read(data.ledger,{type:'buffer'});flagged.Sheets['对账明细-李奇'].P2={t:'s',v:'是'};
   const classified=X.read(data.applications,{type:'buffer'}),app=classified.Sheets.付款申请汇总;
   app.K1={t:'s',v:'付款分类'};for(let r=2;r<=X.utils.decode_range(app['!ref']).e.r+1;r++)app['K'+r]={t:'s',v:String(app['B'+r]?.v).startsWith('谈判')?'谈判付款':'正常付款'};app['!ref']='A1:K'+(X.utils.decode_range(app['!ref']).e.r+1);
   for(const key of ['ledger','applications','oa','negotiated'])await upload(key,key==='ledger'?Buffer.from(X.write(flagged,{type:'buffer',bookType:'xlsx'})):key==='applications'?Buffer.from(X.write(classified,{type:'buffer',bookType:'xlsx'})):data[key]);
   await frame.locator('#negotiatedSheet').selectOption({label:selectedName});await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor();
   const rows=frame.locator('#normalPreview tbody tr');assert.equal(await rows.count(),4);assert.equal(await rows.filter({hasText:'正常供应商1'}).count(),0,'台账改为是后即使申请分类正常也应排除');
   assert.match(await frame.locator('#classificationSummary').innerText(),/3 条[^；。\n]*识别为谈判付款/);
   const event=page.waitForEvent('download');await frame.locator('#download').click();const d=await event,checked=X.read(await fs.readFile(await d.path()),{type:'buffer',cellNF:true,cellStyles:true});
   assert.deepEqual(X.utils.sheet_to_json(checked.Sheets.正常付款).map(row=>row['供应商名称']),['正常供应商2','正常供应商3','正常供应商4','正常供应商5']);
   compareNegotiated(original.Sheets[selectedName],checked.Sheets.谈判付款);assert.equal(X.utils.sheet_to_json(checked.Sheets.谈判付款).some(row=>row['供应商']==='正常供应商1'),false,'不得自动新增谈判明细');
   assert.equal(checked.Sheets.正常付款.P2.f,'I2-N2','排除记录后余额公式仍对应实际行');assert.deepEqual(errors,[]);
  }
  console.log(JSON.stringify({status:'PASS',real,initialPending,normalRows:normal.length-1,applicationReviewRows:applicationReview.length-1,output:out,metadata}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1);});
