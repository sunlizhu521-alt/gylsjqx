const assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const HEADERS=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','摘要','应付款时间','预估付款时间','是否提交付款申请单','是否结清','是否谈判付款','开票时间','所属月份','账期'];
const OUTPUT_HEADERS=['采购跟单员','供应商','付款主体','对公付款OA流程号','申请付款金额','摘要','应付款时间','当前节点','未操作者','流程状态','是否已付款','申请付款月份','需要提交付款审单的日期','是否提交纸质付款审单'];
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const VERSION_KEY='funds-ledger-export-versions',OTHER_VERSIONS='{"2026-11-08":37}';
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const value=cell=>cell&&typeof cell==='object'?cell.v??'':cell??'';
const serial=iso=>(Date.parse(iso+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const flow=n=>'MDST-DGFK20261108'+String(n).padStart(3,'0');
async function replaceSourceCell(bytes,address,cellXml){
 const zip=await Zip.loadAsync(bytes),file='xl/worksheets/sheet1.xml',xml=await zip.file(file).async('string');
 const pattern=new RegExp('<c\\b[^>]*\\br="'+address+'"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)');
 assert.match(xml,pattern,'公式用例必须修改实际单元格');zip.file(file,xml.replace(pattern,cellXml));return zip.generateAsync({type:'nodebuffer'});
}
function expectedSubmission(due,today){
 const previousMonday=serial(due)-((new Date(due+'T00:00:00Z').getUTCDay()+6)%7)-7;
 const thisMonday=serial(today)-((new Date(today+'T00:00:00Z').getUTCDay()+6)%7);
 return Math.max(thisMonday,previousMonday);
}
function source(headers,rows){const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),'原始OA导出');return {book,bytes:write(book)};}
function applications(codes){
 const book=X.utils.book_new();
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['合同号','流程编码','供应商','金额','是否已签字'],...codes.map(code=>[flow(2),code,'无关供应商',999,'']),[flow(2),'','空流程不得回退',0,'签完字']]),'付款申请汇总');
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([['流程编码'],[flow(2)]]),'忽略其他申请');return {book,bytes:write(book)};
}
const withPaper=(rows,codes)=>rows.map(row=>[...row.slice(0,13),row[3]?(codes.includes(row[3])?'已提交':'未提交'):'待核对']).sort((a,b)=>['已提交','未提交','待核对',''].indexOf(a[13])-['已提交','未提交','待核对',''].indexOf(b[13]));
function fixture(replacement=false){
 const book=X.utils.book_new(),rows=PEOPLE.map(()=>[]),entries=[];
 function add(person,supplier,changes={},expectedDate='2026-11-20',inferred=false){
  const row=[supplier,'合成付款主体',flow(1),100,'保留合成摘要',serial('2026-11-20'),serial('2026-12-01'),'原人工提交状态','否','否',serial('2026-10-01'),'2026-08','60天'];
  for(const [header,v]of Object.entries(changes))row[HEADERS.indexOf(header)]=v;
  rows[person].push(row);if(expectedDate!==null)entries.push({person:PEOPLE[person],row,date:expectedDate,inferred,index:entries.length});
 }
 if(replacement)add(0,'合成替换唯一记录',{'申请付款金额':0,'应付款时间':serial('2026-11-12')},'2026-11-12');
 else{
  add(0,'合成已结清仍保留',{'是否结清':'是'});
  for(let i=0;i<2;i++)add(0,'合成重复原明细',{'对公付款0A流程号':flow(2),'应付款时间':serial('2026-11-02')},'2026-11-02');
  add(0,'合成空应付用预估',{'对公付款0A流程号':'','申请付款金额':0,'应付款时间':'','预估付款时间':serial('2026-11-03')},'2026-11-03');
  add(0,'合成已有预估列双空',{'应付款时间':'','预估付款时间':''},'2026-11-06',true);
  add(0,'保留非法应付不换预估',{'应付款时间':'待定','预估付款时间':serial('2026-11-04')},'');
  add(0,'保留非法预估不推算',{'应付款时间':'','预估付款时间':'待定'},'');
  add(0,'保留无时间无所属月份',{'应付款时间':'','预估付款时间':'','所属月份':''},'',true);
  add(0,'合成应付优先跨月',{'应付款时间':serial('2026-10-31'),'预估付款时间':serial('2026-11-05')},'2026-10-31');
  add(0,'合成十二月甲',{'应付款时间':serial('2026-12-01')},'2026-12-01');
  add(0,'合成OA冲突',{'对公付款0A流程号':flow(3),'申请付款金额':{t:'n',v:125,f:'100+25'},'应付款时间':{t:'n',v:serial('2026-11-03'),f:'DATE(2026,11,3)',z:'yyyy-mm-dd'}},'2026-11-03');
  add(0,'合成OA未匹配',{'对公付款0A流程号':'000123456789012345','申请付款金额':'','应付款时间':serial('2026-11-04')},'2026-11-04');
  for(const flag of ['是','','否待确认'])add(0,'排除谈判状态-'+flag,{'是否谈判付款':flag},null);
  for(let i=0;i<60;i++)add(0,'合成长预览-'+String(i).padStart(2,'0'),{'申请付款金额':i,'摘要':i===0?'=1+1':'完整预览 '+i,'是否结清':i%2?'是':'','应付款时间':serial('2026-11-10')},'2026-11-10');
  add(1,'合成OA一致重复',{'对公付款0A流程号':flow(4),'应付款时间':serial('2026-11-01')},'2026-11-01');
  add(2,'合成缺列60天',{'应付款时间':'','所属月份':'2026-08','账期':'60天'},'2026-11-06',true);
  add(2,'合成缺列五个月',{'应付款时间':'','所属月份':'2026-05','账期':'5个月账期(月底付款)50%承兑'},'2026-11-04',true);
  add(2,'合成缺列应付优先',{'应付款时间':serial('2026-11-07'),'所属月份':'2026-05','账期':'5个月'},'2026-11-07');
  add(3,'合成月末已归档',{'对公付款0A流程号':flow(5),'应付款时间':serial('2026-11-30')},'2026-11-30');
  add(3,'合成十二月乙',{'应付款时间':serial('2026-12-01'),'预估付款时间':serial('2026-11-01')},'2026-12-01');
  add(4,'合成文字日期',{'应付款时间':'2026-11-08','预估付款时间':serial('2026-11-09')},'2026-11-08');
 }
 for(let p=0;p<PEOPLE.length;p++){
  const indices=HEADERS.map((_,c)=>c).filter(c=>p!==2||HEADERS[c]!=='预估付款时间'),headers=indices.map(c=>HEADERS[c]);
  const sheet=X.utils.aoa_to_sheet([headers,...rows[p].map(row=>indices.map(c=>row[c]))]);
  for(let r=2;r<=rows[p].length+1;r++)for(const header of ['应付款时间','预估付款时间','开票时间']){
   const c=headers.indexOf(header);if(c<0)continue;const cell=sheet[X.utils.encode_cell({r:r-1,c})];if(cell?.t==='n')cell.z='yyyy-mm-dd';
  }
  sheet['!rows']=[{},...rows[p].map((_,r)=>({hidden:r===0}))];sheet['!cols']=headers.map((_,c)=>({wch:8,hidden:c===3}));
  X.utils.book_append_sheet(book,sheet,'对账明细-'+PEOPLE[p]+(p===4?'20261108':''));
 }
 X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([HEADERS,['历史干扰记录','合成主体',flow(1),999999,'不得纳入',serial('2026-11-01'),'','','否','否']]),'对账明细-李奇-历史');
 const oaHeaders=['流程编号','当前节点','未操作者','流程状态'],oaRows=[
  [flow(1),'归档节点','归档人员','归档'],[flow(2),'审批节点','审批人员','未归档'],
  [flow(3),'冲突节点甲','','归档'],[flow(3),'冲突节点乙','','未归档'],
  [flow(4),'一致节点','一致人员','归档'],[flow(4),'一致节点','一致人员','归档'],[flow(5),'完成节点','','已归档']
 ];
 const oa=source(oaHeaders,oaRows),newOA=source(oaHeaders,[[flow(1),'最新节点','最新人员','未归档']]);
 const fullInfo={[flow(1)]:['归档节点','归档人员','归档','已付款'],[flow(2)]:['审批节点','审批人员','未归档','未付款'],[flow(4)]:['一致节点','一致人员','归档','已付款'],[flow(5)]:['完成节点','','已归档','已付款']};
 function expected(mode='',today='2026-11-08'){
  const info=mode==='full'?fullInfo:mode==='new'?{[flow(1)]:['最新节点','最新人员','未归档','未付款']}:{};
  return entries.toSorted((a,b)=>(a.date||'9999').localeCompare(b.date||'9999')||a.index-b.index).map(({person,row,date,inferred})=>{
   const submission=inferred||!date?'':expectedSubmission(date,today),parts=submission===''?null:X.SSF.parse_date_code(submission);
   return [person,...row.slice(0,5).map(value),inferred?'未对账待定':date?serial(date):value(row[5]||row[6]),...(mode?(info[value(row[2])]||['','','','待核对']):['','','','']),parts?serial(parts.y+'-'+String(parts.m).padStart(2,'0')+'-01'):'',submission,''];
  });
 }
 return {book,bytes:write(book),oa,newOA,expected};
}
(async()=>{
 const base=process.env.BASE_URL||process.env.QA_URL||'http://127.0.0.1:8899/',out=process.env.OUT_DIR||process.env.QA_OUT||'/private/tmp/funds-ledger-browser';
 await fs.mkdir(out,{recursive:true});const input=fixture(),replacement=fixture(true),paperCodes=[flow(1),'000123456789012345'],paper=applications([...paperCodes,flow(1)]),newPaper=applications([flow(2)]);
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({viewport:{width:1600,height:1000},timezoneId:'America/Los_Angeles'}),errors=[],versions=new WeakMap();
  page.on('pageerror',e=>errors.push(e.message));await page.clock.setFixedTime(new Date('2026-11-08T02:00:00Z'));
  await page.goto(base,{waitUntil:'networkidle'});await page.evaluate(v=>localStorage.setItem('funds-export-versions',v),OTHER_VERSIONS);
  const homepage=page.url();await page.locator('button[data-tool="funds-ledger"]').click();assert.equal(page.url(),homepage);
  const frame=page.frameLocator('iframe[data-tool="funds-ledger"]');await frame.locator('#file:not([disabled])').waitFor();
  assert.equal(await frame.locator('#month').count(),0);assert(await frame.locator('#download').isDisabled());
  assert.equal(await frame.locator('#progressFile,#applied-preview,#unapplied-preview').count(),0);
  assert.equal(await frame.locator('#applicationsFile').count(),1);
  const slots=await frame.locator('.upload').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().top));assert.equal(slots.length,3);assert(slots.every(y=>y===slots[0]));
  async function upload(bytes,name,role='ledger',surface=frame){
   const [id,reference]={ledger:['file','reference'],oa:['oaFile','oaReference'],applications:['applicationsFile','applicationsReference']}[role];
   await surface.locator('#'+id+':not([disabled])').waitFor();await surface.locator('#'+id).setInputFiles({name,mimeType:MIME,buffer:bytes});
   await surface.locator('#'+reference).filter({hasText:name}).waitFor();assert.match(await surface.locator('#'+reference).innerText(),/2026.*11.*0?8/);
  }
  async function ready(surface=frame){await surface.locator('#download:not([disabled])').waitFor();}
  async function preview(surface=frame){return surface.locator('#preview').evaluate(el=>({headers:[...el.querySelectorAll('th')].map(c=>c.textContent),rows:[...el.querySelectorAll('tbody tr')].filter(tr=>tr.cells.length===14).map(tr=>[...tr.cells].map(c=>c.textContent))}));}
  async function blocked(pattern){await frame.locator('#status').filter({hasText:pattern}).waitFor();assert(await frame.locator('#download').isDisabled());assert.equal((await preview()).rows.length,0);}
  async function download(label,surface=frame,owner=page,date='2026-11-08'){
   const event=owner.waitForEvent('download');await surface.locator('#download').click();const d=await event;
   let days=versions.get(owner.context());if(!days){days=new Map();versions.set(owner.context(),days);}
   const version=(days.get(date)||0)+1,name='资金台账整理-'+date+'-V'+version+'.xlsx';
   assert.equal(d.suggestedFilename(),name);days.set(date,version);await surface.locator('#status').filter({hasText:name}).waitFor();
   const bytes=await fs.readFile(await d.path());await fs.writeFile(path.join(out,label+'.xlsx'),bytes);return bytes;
  }
  async function verify(bytes,expected,surface=frame){
   const book=X.read(bytes,{type:'buffer',cellNF:true,cellStyles:true}),name='正常付款',sheet=book.Sheets[name],snap=await preview(surface);
   assert.deepEqual(book.SheetNames,[name]);const range=X.utils.decode_range(sheet['!ref']);
   const matrix=Array.from({length:range.e.r+1},(_,r)=>Array.from({length:range.e.c+1},(_,c)=>sheet[X.utils.encode_cell({r,c})]?.v??''));
   assert.deepEqual(matrix,[OUTPUT_HEADERS,...expected]);assert.deepEqual(snap.headers,OUTPUT_HEADERS);assert.equal(snap.rows.length,expected.length);
   assert.equal(Number((await surface.locator('#record-count').innerText()).replace(/[^\d]/g,'')),expected.length);
   assert.equal(sheet['!autofilter'].ref,'A1:N'+(expected.length+1));assert(!book.Workbook?.Sheets?.[0]?.Hidden);
   const previewFills=await surface.locator('#preview').evaluate(el=>({
    headers:[...el.querySelectorAll('th')].map(cell=>getComputedStyle(cell).backgroundColor),
    rows:[...el.querySelectorAll('tbody tr')].filter(row=>row.cells.length===14).map(row=>[...row.cells].map(cell=>getComputedStyle(cell).backgroundColor))
   }));
   assert.deepEqual(previewFills.headers,Array(14).fill('rgb(49, 91, 157)'));
   assert.deepEqual(previewFills.rows,expected.map(row=>Array(14).fill(row[10]==='已付款'?'rgb(255, 153, 153)':'rgb(255, 255, 255)')),'预览整行颜色须与是否已付款一致，包含空白格');
   for(let r=0;r<expected.length;r++)for(let c=0;c<14;c++){
    const cell=sheet[X.utils.encode_cell({r:r+1,c})],v=expected[r][c];assert.equal(cell?.v??'',v);assert(!cell?.f);
    if(v!=='')assert.equal(cell.t,typeof v==='number'?'n':'s');if(c===4&&v!=='')assert.match(cell.z,/0\.00/);
    if(c===6){
     if(typeof v==='number'){assert.equal(cell.z,'yyyy"年"mm"月"dd"日"');assert.match(snap.rows[r][c],/^\d{4}年\d{2}月\d{2}日$/);}
     else{assert(['未对账待定','待定'].includes(v));assert.equal(cell.t,'s');assert.equal(snap.rows[r][c],v);}
    }
    if(c===11&&v!==''){assert.equal(cell.z,'yyyy"年"mm"月"');assert.equal(cell.t,'n');assert.equal(X.SSF.parse_date_code(cell.v).d,1);assert.match(snap.rows[r][c],/^\d{4}年\d{2}月$/);}
    if(c===12&&v!==''){assert.equal(cell.z,'yyyy"年"mm"月"dd"日"');assert.equal(cell.t,'n');assert(Number.isInteger(cell.v));assert.match(snap.rows[r][c],/^\d{4}年\d{2}月\d{2}日$/);}
    assert.equal(snap.rows[r][c],cell?X.utils.format_cell(cell):'');
   }
   const meta=await surface.locator('body').evaluate(async(_,data)=>{
    const zip=await JSZip.loadAsync(Uint8Array.from(data.bytes)),parse=s=>new DOMParser().parseFromString(s,'application/xml'),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
    const styles=parse(await zip.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0].children,fonts=all(styles,'fonts')[0].children,borders=all(styles,'borders')[0].children,fills=all(styles,'fills')[0].children;
    const doc=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),problems=[],names=new Set(),sizes=new Set(),widths=[];
    const pane=all(doc,'pane')[0];if(pane?.getAttribute('state')!=='frozen'||pane?.getAttribute('ySplit')!=='1')problems.push('首行未冻结');
    if(all(doc,'f').length||all(doc,'filterColumn').length||all(doc,'conditionalFormatting').length)problems.push('存在公式、旧筛选条件或条件格式');
    if(all(doc,'row').some(r=>r.getAttribute('hidden')==='1'||all(r,'c').length!==14))problems.push('隐藏行或缺少矩形单元格');
    for(const col of all(doc,'col')){if(col.getAttribute('hidden')==='1')problems.push('隐藏列');for(let c=+col.getAttribute('min');c<=+col.getAttribute('max');c++)widths[c-1]=+col.getAttribute('width');}
    for(const cell of all(doc,'c')){
     const xf=xfs[+(cell.getAttribute('s')||0)],font=fonts[+(xf.getAttribute('fontId')||0)],border=borders[+(xf.getAttribute('borderId')||0)];
     names.add(all(font,'name')[0]?.getAttribute('val'));sizes.add(all(font,'sz')[0]?.getAttribute('val'));
     if(['left','right','top','bottom'].some(edge=>!all(border,edge)[0]?.getAttribute('style')))problems.push('缺少边框 '+cell.getAttribute('r'));
     const row=Number(cell.getAttribute('r').match(/\d+$/)[0]),fill=fills[+(xf.getAttribute('fillId')||0)],expectedFill=row===1?'FF315B9D':data.paid[row-2]?'FFFF9999':'FFFFFFFF';
     if(all(fill,'patternFill')[0]?.getAttribute('patternType')!=='solid'||all(fill,'fgColor')[0]?.getAttribute('rgb')!==expectedFill||xf.getAttribute('applyFill')!=='1')problems.push('原生填充颜色不符 '+cell.getAttribute('r'));
     if((row===1||/^[AML]\d+$/.test(cell.getAttribute('r')))&&all(xf,'alignment')[0]?.getAttribute('wrapText')==='1')problems.push('表头或提交日期换行');
    }
    return {problems,names:[...names],sizes:[...sizes],widths};
   },{bytes:[...bytes],paid:expected.map(row=>row[10]==='已付款')});
   assert.deepEqual(meta.problems,[]);assert.equal(meta.names.length,1);assert.equal(meta.sizes.length,1);assert.equal(meta.widths.length,14);assert(meta.widths.every(w=>w>=12&&w<=255));return snap;
  }
  await upload(input.bytes,'合成全月份台账.xlsx');await ready();
  const baseSnap=await verify(await download('monthly-base'),input.expected());
  assert.equal(baseSnap.rows.find(r=>r[1]==='合成月末已归档')[12],'2026年11月23日');
  assert.equal(baseSnap.rows.find(r=>r[1]==='合成已结清仍保留')[12],'2026年11月09日');
  assert.equal(baseSnap.rows.find(r=>r[1]==='合成空应付用预估')[12],'2026年11月02日');
  assert(baseSnap.rows.filter(r=>r[6]==='未对账待定').every(r=>r[12]===''));
  assert(baseSnap.rows.some(r=>r[1]==='合成已结清仍保留'));assert.equal(baseSnap.rows.filter(r=>r[1]==='合成重复原明细').length,2);
  assert.deepEqual(baseSnap.rows.filter(r=>r[6]==='未对账待定').map(r=>r[1]),['合成缺列五个月','合成已有预估列双空','合成缺列60天','保留无时间无所属月份']);
  assert.deepEqual(baseSnap.rows.slice(6,11).map(r=>[r[1],r[6]]),[
   ['合成OA未匹配','2026年11月04日'],['合成缺列五个月','未对账待定'],['合成已有预估列双空','未对账待定'],['合成缺列60天','未对账待定'],['合成缺列应付优先','2026年11月07日']
  ]);
  assert.equal(baseSnap.rows.find(r=>r[1]==='合成十二月甲')[11],'2026年11月');assert.equal(baseSnap.rows.find(r=>r[1]==='合成十二月甲')[12],'2026年11月23日');
  assert(baseSnap.rows.some(r=>r[6]==='2026年10月31日'));assert.equal(baseSnap.rows.filter(r=>r[6]==='2026年12月01日').length,2);
  assert.deepEqual(baseSnap.rows.slice(-3).map(r=>r[1]),['保留非法应付不换预估','保留非法预估不推算','保留无时间无所属月份']);assert(baseSnap.rows.slice(-3).every(r=>r[11]===''&&r[12]===''));
  assert(!JSON.stringify(baseSnap).includes('排除'));assert(!JSON.stringify(baseSnap).includes('历史干扰记录'));assert(baseSnap.rows.length>60);
  await upload(input.oa.bytes,'合成原始OA.xlsx','oa');await ready();const fullSnap=await verify(await download('monthly-oa'),input.expected('full'));
  const warningPanel=frame.locator('#warning-panel');await warningPanel.waitFor({state:'visible'});if(!await warningPanel.evaluate(el=>el.open))await warningPanel.locator('summary').click();
  const warnings=await frame.locator('#warnings').innerText();assert.match(warnings,/待定|日期|空白/);assert.match(warnings,/冲突|待核对/);
  await page.locator('button[data-tool="funds"]').click();await page.frameLocator('iframe[data-tool="funds"]').locator('#ledgerFile').waitFor();
  await page.locator('button[data-tool="funds-ledger"]').click();assert.equal(await frame.locator('#month').count(),0);assert.deepEqual(await preview(),fullSnap);
  assert((await frame.locator('#oaReference').innerText()).includes('合成原始OA.xlsx'));await page.screenshot({path:path.join(out,'monthly-home.png'),fullPage:true});
  await upload(paper.bytes,'合成申请单汇总.xlsx','applications');await ready();
  const paperSnap=await verify(await download('paper-submission'),withPaper(input.expected('full'),paperCodes));
  assert.match(await frame.locator('#warnings').innerText(),/合成申请单汇总.xlsx.*付款申请汇总!B5/);
  await page.locator('button[data-tool="funds"]').click();await page.locator('button[data-tool="funds-ledger"]').click();assert.deepEqual(await preview(),paperSnap);
  assert.match(await frame.locator('#applicationsReference').innerText(),/合成申请单汇总.xlsx.*\n引用时间/);
  await upload(newPaper.bytes,'合成替换申请单.xlsx','applications');await ready();await verify(await download('paper-replaced'),withPaper(input.expected('full'),[flow(2)]));
  await frame.locator('#clear-applications').click();await ready();assert.deepEqual(await preview(),fullSnap);
  await verify(await download('paper-removed'),input.expected('full'));
  const ambiguous=source(['流程编码'],[[flow(1)]]);X.utils.book_append_sheet(ambiguous.book,X.utils.aoa_to_sheet([['合同号'],[flow(2)]]),'另一申请');
  for(const [bytes,name,pattern]of [
   [write(ambiguous.book),'合成多候选申请.xlsx',/来源不明确/],
   [source(['摘要'],[['无编号表头']]).bytes,'合成错误申请.xlsx',/未.*找到/],
   [await replaceSourceCell(paper.bytes,'B2','<c r="B2" t="str"><f>"OA001"</f></c>'),'合成缺缓存申请.xlsx',/公式缺少.*缓存/],
   [await replaceSourceCell(paper.bytes,'B2','<c r="B2" t="e"><f>1\/0</f><v>#DIV\/0!</v></c>'),'合成公式错误申请.xlsx',/Excel.*错误/]
  ]){
   await upload(bytes,name,'applications');await blocked(pattern);
   await upload(input.bytes,'合成失败后换台账.xlsx');await blocked(pattern);
   await upload(paper.bytes,'合成恢复申请.xlsx','applications');await ready();assert.deepEqual(await preview(),paperSnap);
  }
  await page.screenshot({path:path.join(out,'monthly-paper.png'),fullPage:true});
  await frame.locator('#clear-applications').click();await ready();
  await frame.locator('#clear-oa').click();await ready();await verify(await download('remove-paid-oa'),input.expected());
  const badOA=X.read(input.oa.bytes,{type:'buffer'});badOA.Sheets.原始OA导出.D1.v='无效状态表头';
  await upload(write(badOA),'合成错误OA.xlsx','oa');await blocked(/缺少|缺失|找不到|未找到|失败|无法/);
  await upload(input.bytes,'合成台账重选.xlsx');await blocked(/缺少|缺失|找不到|未找到|失败|无法/);
  await frame.locator('#clear-oa').click();await ready();await verify(await download('remove-failed-oa'),input.expected());
  const badLedger=X.read(input.bytes,{type:'buffer'});badLedger.Sheets['对账明细-李奇'].D1.v='无效金额表头';
  await upload(write(badLedger),'合成错误台账.xlsx');await blocked(/缺少|缺失|找不到|未找到|失败|无法/);
  for(const [address,xml]of [['D2','<c r="D2"><f>1+1</f></c>'],['J2','<c r="J2" t="str"><f>IF(1=1,&quot;否&quot;,&quot;是&quot;)</f></c>'],['F2','<c r="F2"><f>DATE(2026,11,20)</f></c>'],['G6','<c r="G6"><f>DATE(2026,11,6)</f></c>']]){
   let bytes=await replaceSourceCell(input.bytes,address,xml);if(address==='J2')bytes=await replaceSourceCell(bytes,'I2','<c r="I2"/>');
   await upload(bytes,'合成缺缓存-'+address+'.xlsx');await blocked(/公式缺少.*缓存/);
  }
  for(const address of ['F2','G6']){
   const bytes=await replaceSourceCell(input.bytes,address,'<c r="'+address+'" t="e"><f>1/0</f><v>#DIV/0!</v></c>');
   await upload(bytes,'合成日期公式错误-'+address+'.xlsx');await blocked(/Excel.*错误/);
  }
  await upload(input.bytes,'合成恢复台账.xlsx');await ready();
  const badOAFormula=await replaceSourceCell(input.oa.bytes,'B2','<c r="B2" t="str"><f>IF(1=1,&quot;节点&quot;,&quot;&quot;)</f></c>');
  await upload(badOAFormula,'合成OA缺缓存.xlsx','oa');await blocked(/公式缺少.*缓存/);await frame.locator('#clear-oa').click();await ready();
  const emptyFormula=await replaceSourceCell(input.bytes,'E2','<c r="E2" t="str"><f>IF(1=1,&quot;&quot;,&quot;摘要&quot;)</f><v/></c>'),emptyExpected=input.expected().map(r=>[...r]);
  emptyExpected.find(r=>r[1]==='合成已结清仍保留')[5]='';await upload(emptyFormula,'合成有效空缓存.xlsx');await ready();await verify(await download('valid-empty-formula'),emptyExpected);
  await upload(input.bytes,'合成竞态台账.xlsx');await upload(input.newOA.bytes,'合成替换OA.xlsx','oa');await ready();await verify(await download('replace-oa'),input.expected('new'));
  await frame.locator('body').evaluate(()=>{
   window.__qaOriginalRead=File.prototype.arrayBuffer;window.__qaGates={};
   File.prototype.arrayBuffer=function(){const reading=window.__qaOriginalRead.call(this);if(!this.name.startsWith('延迟-'))return reading;return new Promise((resolve,reject)=>{window.__qaGates[this.name]=async()=>{try{resolve(await reading);}catch(e){reject(e);}};});};
  });
  async function release(names){await frame.locator('body').evaluate(async(_,names)=>{for(const name of names){await window.__qaGates[name]();delete window.__qaGates[name];}await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));},names);}
  await frame.locator('body').evaluate((_,uploads)=>{for(const u of uploads){const transfer=new DataTransfer();transfer.items.add(new File([Uint8Array.from(u.bytes)],u.name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const el=document.getElementById(u.id);el.files=transfer.files;el.dispatchEvent(new Event('change',{bubbles:true}));}},[
   {id:'file',name:'延迟-台账.xlsx',bytes:[...input.bytes]},{id:'oaFile',name:'延迟-OA.xlsx',bytes:[...input.oa.bytes]},{id:'applicationsFile',name:'延迟-申请.xlsx',bytes:[...paper.bytes]}
  ]);
  assert(await frame.locator('#download').isDisabled());assert.equal((await preview()).rows.length,0);
  await release(['延迟-OA.xlsx']);assert(await frame.locator('#download').isDisabled());await release(['延迟-台账.xlsx']);assert(await frame.locator('#download').isDisabled());await release(['延迟-申请.xlsx']);await ready();await verify(await download('simultaneous-inputs'),withPaper(input.expected('full'),paperCodes));
  await upload(paper.bytes,'延迟-旧申请.xlsx','applications');assert(await frame.locator('#download').isDisabled());
  await upload(newPaper.bytes,'合成最新申请.xlsx','applications');await ready();const newestPaper=await preview();await release(['延迟-旧申请.xlsx']);assert.deepEqual(await preview(),newestPaper);
  await verify(await download('latest-paper-wins'),withPaper(input.expected('full'),[flow(2)]));
  await upload(paper.bytes,'延迟-移除申请.xlsx','applications');await frame.locator('#clear-applications').click();await ready();await release(['延迟-移除申请.xlsx']);assert.deepEqual(await preview(),fullSnap);
  await upload(input.oa.bytes,'延迟-旧OA.xlsx','oa');assert(await frame.locator('#download').isDisabled());
  await upload(input.newOA.bytes,'合成最新OA.xlsx','oa');await ready();const newest=await preview();await release(['延迟-旧OA.xlsx']);assert.deepEqual(await preview(),newest);
  await verify(await download('latest-oa-wins'),input.expected('new'));
  await frame.locator('body').evaluate(()=>{File.prototype.arrayBuffer=window.__qaOriginalRead;delete window.__qaOriginalRead;delete window.__qaGates;});
  await frame.locator('#clear-oa').click();await ready();await upload(replacement.bytes,'合成替换台账.xlsx');await ready();await verify(await download('replace-ledger'),replacement.expected());
  assert.equal(await page.evaluate(()=>localStorage.getItem('funds-export-versions')),OTHER_VERSIONS);

  const mobile=await browser.newPage({viewport:{width:390,height:844},isMobile:true,deviceScaleFactor:1});mobile.on('pageerror',e=>errors.push(e.message));
  await mobile.clock.setFixedTime(new Date('2026-11-08T02:00:00Z'));await mobile.goto(base,{waitUntil:'networkidle'});await mobile.locator('button[data-tool="funds-ledger"]').click();
  const mobileFrame=mobile.frameLocator('iframe[data-tool="funds-ledger"]');await upload(replacement.bytes,'合成手机台账.xlsx','ledger',mobileFrame);await ready(mobileFrame);
  assert.equal((await preview(mobileFrame)).rows.length,1);
  assert(await mobileFrame.locator('.uploads-scroll').evaluate(el=>el.scrollWidth>el.clientWidth),'窄屏上传区应独立横向滚动');
  const mobileSlots=await mobileFrame.locator('.upload').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().top));assert.equal(mobileSlots.length,3);assert(mobileSlots.every(y=>y===mobileSlots[0]));
  assert(await mobileFrame.locator('body').evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'窄屏工具主体不得横向溢出');
  assert(await mobile.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'手机外层页面不得横向溢出');
  await mobile.screenshot({path:path.join(out,'monthly-mobile.png'),fullPage:true});await mobile.close();
  const standalone=await browser.newPage({viewport:{width:1440,height:1000},timezoneId:'America/Los_Angeles'});standalone.on('pageerror',e=>errors.push(e.message));
  await standalone.clock.setFixedTime(new Date('2026-11-08T02:00:00Z'));await standalone.goto(new URL('funds-ledger.html',base).href,{waitUntil:'networkidle'});
  await standalone.evaluate(v=>localStorage.setItem('funds-export-versions',v),OTHER_VERSIONS);
  await upload(input.oa.bytes,'合成先选OA.xlsx','oa',standalone);assert(await standalone.locator('#download').isDisabled());
  await upload(paper.bytes,'合成先选申请.xlsx','applications',standalone);assert(await standalone.locator('#download').isDisabled());
  await upload(replacement.bytes,'合成独立页面.xlsx','ledger',standalone);await ready(standalone);
  await verify(await download('oa-first',standalone,standalone),withPaper(replacement.expected('full'),paperCodes),standalone);
  await standalone.screenshot({path:path.join(out,'monthly-standalone.png')});
  await standalone.reload({waitUntil:'networkidle'});await upload(replacement.bytes,'合成版本验收.xlsx','ledger',standalone);await ready(standalone);
  await verify(await download('version-after-refresh',standalone,standalone),replacement.expected(),standalone);
  const beforeFailure=await standalone.evaluate(key=>localStorage.getItem(key),VERSION_KEY),failedDownloads=[];
  const recordFailedDownload=d=>failedDownloads.push(d.suggestedFilename());standalone.on('download',recordFailedDownload);
  for(const stage of ['export','click']){
   await standalone.evaluate(stage=>{
    if(stage==='export'){window.__qaOriginalExport=FundsLedgerExport;window.FundsLedgerExport=async()=>{throw Error('合成导出失败');};}
    else{window.__qaOriginalAnchorClick=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.download)throw Error('合成下载发起失败');return window.__qaOriginalAnchorClick.call(this);};}
   },stage);
   await standalone.locator('#download').click();await standalone.locator('#status').filter({hasText:stage==='export'?'合成导出失败':'合成下载发起失败'}).waitFor();
   assert.deepEqual(failedDownloads,[],'生成或发起下载失败不能产生文件');
   assert.equal(await standalone.evaluate(key=>localStorage.getItem(key),VERSION_KEY),beforeFailure,'失败不能递增保存的版本号');
   await standalone.evaluate(stage=>{
    if(stage==='export'){window.FundsLedgerExport=window.__qaOriginalExport;delete window.__qaOriginalExport;}
    else{HTMLAnchorElement.prototype.click=window.__qaOriginalAnchorClick;delete window.__qaOriginalAnchorClick;}
   },stage);
  }
  standalone.off('download',recordFailedDownload);
  await download('version-after-failures',standalone,standalone);
  await standalone.clock.setFixedTime(new Date('2026-11-08T15:59:59Z'));
  await verify(await download('version-before-beijing-midnight',standalone,standalone),replacement.expected(),standalone);
  const sundayPreview=await preview(standalone);assert.equal(sundayPreview.rows[0][12],'2026年11月02日');
  await standalone.clock.setFixedTime(new Date('2026-11-08T16:00:00Z'));
  // Download must update the stale Sunday preview without another upload or another action.
  assert.deepEqual(await preview(standalone),sundayPreview);
  await verify(await download('version-after-beijing-midnight',standalone,standalone,'2026-11-09'),replacement.expected('','2026-11-09'),standalone);
  const mondayPreview=await preview(standalone);assert.equal(mondayPreview.rows[0][12],'2026年11月09日');
  assert.deepEqual(mondayPreview.rows.map(r=>r.filter((_,c)=>c!==11&&c!==12)),sundayPreview.rows.map(r=>r.filter((_,c)=>c!==11&&c!==12)));
  await download('version-new-day-second',standalone,standalone,'2026-11-09');
  const persisted=await standalone.evaluate(key=>JSON.parse(localStorage.getItem(key)),VERSION_KEY);
  assert.equal(persisted['2026-11-08'],4);assert.equal(persisted['2026-11-09'],2);
  await standalone.evaluate(key=>localStorage.setItem(key,'{broken-json'),VERSION_KEY);
  await download('version-corrupt-storage',standalone,standalone,'2026-11-09');
  await download('version-corrupt-storage-second',standalone,standalone,'2026-11-09');
  await standalone.evaluate(()=>{
   window.__qaStorageDescriptor=Object.getOwnPropertyDescriptor(window,'localStorage');
   Object.defineProperty(window,'localStorage',{configurable:true,get(){throw new DOMException('合成存储禁用','SecurityError');}});
  });
  await download('version-disabled-storage',standalone,standalone,'2026-11-09');
  await verify(await download('version-disabled-storage-second',standalone,standalone,'2026-11-09'),replacement.expected('','2026-11-09'),standalone);
  await standalone.evaluate(()=>{Object.defineProperty(window,'localStorage',window.__qaStorageDescriptor);delete window.__qaStorageDescriptor;});
  assert.equal(await standalone.evaluate(()=>localStorage.getItem('funds-export-versions')),OTHER_VERSIONS,'内存回退也不能改动其他工具的版本号');
  await standalone.close();assert.deepEqual(errors,[]);
  console.log(JSON.stringify({base,out,columns:14,allMonthRows:input.expected().length,submissionDate:true,weeklyDownloadRefresh:true,paperSubmission:true,paperReplacementRemoval:true,allMonthSheet:true,applicationMonth:true,undatedRetained:true,datePriority:true,negotiatedOnly:true,oaMatching:true,fullPreview:true,styles:true,followerBySource:true,lightRedPaidRows:true,toolSwitch:true,invalidRecovery:true,formulaCacheGuard:true,oaFirst:true,simultaneous:true,latestWins:true,mobile:true,downloadNames:true,versionRefresh:true,versionFailureGuard:true,beijingRollover:true,memoryFallback:true,errors}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
