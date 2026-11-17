const assert=require('node:assert/strict'),fs=require('node:fs/promises'),crypto=require('node:crypto'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js'),C=require('../funds-core.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const roles=['ledger','applications','oa','negotiated'],write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const hash=bytes=>crypto.createHash('sha256').update(bytes).digest('hex'),read=bytes=>X.read(bytes,{cellFormula:true,cellNF:true,cellStyles:true});
const serial=s=>(Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const days=['2026-09-02','2026-09-16','2026-10-02','2026-10-16','2026-11-02','2026-11-09','2026-11-16'];
const bands=[{startCol:16,endCol:17,label:'2026年9月'},{startCol:18,endCol:19,label:'2026年10月'},{startCol:20,endCol:22,label:'2026年11月'}];
const sourceName='十一月谈判付款主表';

function sample(headerRow=0,withDates=true){
 const input=fixture();
 for(const role of roles.filter(role=>role!=='negotiated'))for(const name of input[role].book.SheetNames){
  const headers=X.utils.sheet_to_json(input[role].book.Sheets[name],{header:1})[0];
  input[role].book.Sheets[name]=X.utils.aoa_to_sheet([headers]);
 }
 const headers=['供应商名称','付款主体','对公付款0A流程号','OA-申请付款金额','已付金额','剩余付款金额','是否结清','应付款时间','备注','谈判明细','分批总金额','核对','人工公式','手工合并备注','保留文本','人工标识'];
 headers.push(...days.map((day,i)=>withDates?serial(day):'人工金额'+(i+17)));
 while(headers.length<34)headers.push('保留字段'+(headers.length+1));
 headers.push(...['2026年9月支付','2026年10月支付','2026年11月支付','2026年12月支付','2027年1月支付','2027年2月支付','2027年3月支付','2027年4月支付']);
 const amounts=[100,0,25],plans=[[10,10,20,20,10,10,20],[0,0,0,0,0,0,0],[0,0,0,0,25,0,0]];
 const data=amounts.map((amount,i)=>{
  const values=Array(42).fill('');
  values.splice(0,16,'表头合成'+(i+1),'迈德斯特','未提交OA',amount,0,amount,amount?'否':'是',serial('2026-11-30'),'保留人工备注'+i,days.map((day,c)=>X.SSF.format(C.DATE_FORMAT,serial(day))+'－'+plans[i][c]).join('\n'),amount,'一致',amount*2,i===0?'手工合并备注':'','引用字符串D2','人工'+i);
  values.splice(16,7,...plans[i]);
  values.splice(34,8,plans[i][0]+plans[i][1],plans[i][2]+plans[i][3],plans[i][4]+plans[i][5]+plans[i][6],0,0,0,0,0);
  return values;
 });
 const prefix=Array.from({length:headerRow},(_,i)=>['人工标题'+(i+1)]),sheet=X.utils.aoa_to_sheet([...prefix,headers,...data]),h=headerRow+1,first=h+1;
 if(withDates)for(let c=16;c<=22;c++)sheet[X.utils.encode_cell({r:headerRow,c})].z=C.DATE_FORMAT;
 for(let i=0;i<data.length;i++){
  const r=first+i,amount=amounts[i];
  sheet['F'+r]={t:'n',v:amount,f:'D'+r+'-E'+r,z:'#,##0.00'};
  sheet['K'+r]={t:'n',v:amount,f:'SUM(Q'+r+':W'+r+')'};
  sheet['M'+r]={t:'n',v:amount*2,f:'SUM(Q'+r+':W'+r+')+$D$'+r};
  for(const [column,range] of [['AI','Q'+r+':R'+r],['AJ','S'+r+':T'+r],['AK','U'+r+':W'+r]])sheet[column+r].f='SUM('+range+')';
  sheet['H'+r].z=C.DATE_FORMAT;
 }
 sheet['Q'+first].c=[{a:'人工',t:'金额批注跟随原记录'}];sheet['N'+first].c=[{a:'人工',t:'合并备注批注'}];
 sheet['!merges']=[{s:{r:first-1,c:13},e:{r:first,c:13}}];
 sheet['!autofilter']={ref:'A'+h+':AP'+(first+2)};
 sheet['!rows']=Array.from({length:first+2},(_,i)=>({hpt:i===headerRow?31:47}));
 const book=X.utils.book_new();X.utils.book_append_sheet(book,sheet,sourceName);
 const aux=X.utils.aoa_to_sheet([['引用保留'],[7,100],['不改字符串']]);
 aux.B2={t:'n',v:107,f:"'"+sourceName+"'!$D$"+first+'+A2'};
 aux.B3={t:'s',v:sourceName+'!D'+first,f:'"'+sourceName+'!D'+first+'"'};
 aux.A2.c=[{a:'人工',t:'其他工作表原批注'}];
 X.utils.book_append_sheet(book,aux,'辅助引用');
 book.Workbook={Names:[{Name:'谈判金额范围',Ref:"'"+sourceName+"'!$D$"+first+':$D$'+(first+2)}]};
 input.negotiated={book,name:'月份表头合成.xlsx',sheetName:sourceName};
 return {input,headerRow,withDates};
}

function verifyBook(bytes,sourceBook,sourceHeaderRow,inserted,withDates){
 const saved=read(bytes),sheet=saved.Sheets.谈判付款,source=sourceBook.Sheets[sourceBook.SheetNames.find(name=>name.includes('谈判付款'))];
 const lower=sourceHeaderRow+1+inserted,first=lower+1,sourceFirst=sourceHeaderRow+2,end=first+2;
 assert.equal(X.utils.decode_range(sheet['!ref']).e.r+1,end,'只新增表头行，不新增金额记录');
 const filter=X.utils.decode_range(sheet['!autofilter'].ref);assert.equal(filter.s.r+1,lower,'筛选位于下方明细表头');assert.equal(filter.e.r+1,end);
 const merges=new Set((sheet['!merges']||[]).map(X.utils.encode_range));
 if(withDates){
  for(const band of bands){
   const start=X.utils.encode_col(band.startCol),finish=X.utils.encode_col(band.endCol),top=lower-1;
   assert.equal(sheet[start+top].v,band.label);assert.equal(sheet[start+top].t,'s');assert(merges.has(start+top+':'+finish+top),'同月日期横向合并');
  }
  for(let c=0;c<=X.utils.decode_range(sheet['!ref']).e.c;c++)if(c<16||c>22){const col=X.utils.encode_col(c);assert(merges.has(col+(lower-1)+':'+col+lower),col+' 普通表头上下合并');}
  for(let c=16;c<=22;c++){const cell=sheet[X.utils.encode_col(c)+lower];assert.equal(cell.v,X.SSF.format(C.DATE_FORMAT,serial(days[c-16])));assert.equal(cell.t,'s','日期表头不参与金额求和');}
 }else assert.equal([...merges].filter(range=>X.utils.decode_range(range).s.r===sourceHeaderRow).length,0,'无日期列不生成分组表头');
 assert(merges.has('N'+first+':N'+(first+1)),'人工跨数据行合并跟随原记录');
 const numericColumns=[3,4,5,10,...Array.from({length:7},(_,i)=>i+16),...Array.from({length:8},(_,i)=>i+34)];
 for(let i=0;i<3;i++){
  const r=first+i,sr=sourceFirst+i;assert.equal(sheet['A'+r].v,source['A'+sr].v);
  for(const c of numericColumns){const col=X.utils.encode_col(c);assert.equal(sheet[col+r].v,source[col+sr].v,col+r+' 原金额保持');assert.equal(sheet[col+r].t,'n');}
  assert.equal(sheet['F'+r].f,'D'+r+'-E'+r);assert.equal(sheet['K'+r].f,'SUM(Q'+r+':W'+r+')');assert.equal(sheet['M'+r].f,'SUM(Q'+r+':W'+r+')+$D$'+r);
  for(const [column,range] of [['AI','Q'+r+':R'+r],['AJ','S'+r+':T'+r],['AK','U'+r+':W'+r]])assert.equal(sheet[column+r].f,'SUM('+range+')');
 }
 for(let c=34;c<=41;c++){const col=X.utils.encode_col(c),titleRow=withDates?lower-1:lower;assert.equal(sheet[col+titleRow].v,source[col+(sourceHeaderRow+1)]?.v||source[col+sourceHeaderRow]?.v,col+' 原AI–AP表头保留');}
 assert.match(sheet['Q'+first].c[0].t,/^金额批注跟随原记录(?:\n|$)/);assert.equal(sheet['N'+first].c[0].t,'合并备注批注');
 assert.equal(saved.Sheets.辅助引用.B2.f,"'谈判付款'!$D$"+first+'+A2','仅目标sheet引用下移，辅助表本地A2不变');
 assert.equal(saved.Sheets.辅助引用.B3.f,sourceBook.Sheets.辅助引用.B3.f,'公式中的字符串地址不移动');assert.equal(saved.Sheets.辅助引用.A2.c[0].t,'其他工作表原批注');
 assert.equal(saved.Workbook.Names.find(name=>name.Name==='谈判金额范围').Ref,"'谈判付款'!$D$"+first+':$D$'+end);
 for(let r=1;r<=sourceHeaderRow-(withDates&&!inserted?1:0);r++)assert.equal(sheet['A'+r]?.v,source['A'+r]?.v,'前置人工标题保留');
 return saved;
}

(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8963/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  const frame=page.frameLocator('iframe[data-tool="funds"]');
  async function upload(role,buffer){await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:'表头合成-'+role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});}
  async function download(withDates){
   await frame.locator('#analyze:not([disabled])').click();
   try{await frame.locator('#download:not([disabled])').waitFor({timeout:30000});}catch(error){throw Error(await frame.locator('#status').innerText()+'；待处理：'+await frame.locator('#pendingCount').innerText());}
   assert.equal(await frame.locator('#pendingCount').innerText(),'0');assert.equal(await frame.locator('#negotiatedPreview tbody tr').count(),3);
   const head=await frame.locator('#negotiatedPreview thead').evaluate(el=>[...el.rows].map(row=>[...row.cells].map(cell=>({text:cell.textContent,rowSpan:cell.rowSpan,colSpan:cell.colSpan}))));
   assert.equal(head.length,withDates?2:1,'页面预览与下载都使用分组表头');
   if(withDates){for(const band of bands)assert(head[0].some(cell=>cell.text===band.label&&cell.colSpan===band.endCol-band.startCol+1));assert(head[0].some(cell=>cell.text==='供应商名称'&&cell.rowSpan===2));}
   const event=page.waitForEvent('download');await frame.locator('#download').click();return fs.readFile(await(await event).path());
  }
  for(const [headerRow,withDates] of [[0,true],[2,true],[0,false]]){
   const {input}=sample(headerRow,withDates),buffers=Object.fromEntries(roles.map(role=>[role,write(input[role].book)]));
   if(headerRow===0&&withDates){
    // Add the legacy empty formatting only after writing the small fixture.
    const zip=await Zip.loadAsync(buffers.negotiated),file='xl/worksheets/sheet1.xml';let xml=await zip.file(file).async('string');
    const style=xml.match(/<c\b[^>]*\bs="(\d+)"/)?.[1]||'0';
    xml=xml.replace(/<dimension\b[^>]*\bref="[^"]+"[^>]*\/>/,'<dimension ref="A1:AP1048576"/>').replace('</sheetData>','<row r="1048576"><c r="AP1048576" s="'+style+'"/></row></sheetData>');
    assert.match(xml,/<dimension ref="A1:AP1048576"\/>/);assert.match(xml,/<c r="AP1048576" s="\d+"\/>/);
    zip.file(file,xml);buffers.negotiated=await zip.generateAsync({type:'nodebuffer'});
   }
   const originalJSON=JSON.stringify(input),hashes=Object.fromEntries(roles.map(role=>[role,hash(buffers[role])])),result=C.build(input);
   assert.equal(result.stats.pending,0,'合成输入无需人工处理即可下载');
   for(const role of roles)await upload(role,buffers[role]);
   const output=await download(withDates),saved=verifyBook(output,input.negotiated.book,headerRow,withDates?1:0,withDates);
   if(headerRow===0&&withDates){
    const zip=await Zip.loadAsync(output),xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
    assert.equal(saved.Sheets.谈判付款.AP1048576,undefined,'先清除末行空格式，再插入月份表头');
    assert.doesNotMatch(xml,/<(?:row|c)\b[^>]*\br="(?:[A-Z]+)?104857[67]"/,'无残留末行或越界行');
   }
   if(headerRow===0&&withDates&&process.env.QA_OUT){
    await fs.mkdir(process.env.QA_OUT,{recursive:true});await fs.writeFile(path.join(process.env.QA_OUT,'funds-month-header.xlsx'),output);
    await frame.locator('#negotiatedPreview').evaluate(el=>{const band=[...el.querySelectorAll('th')].find(cell=>cell.textContent==='2026年9月');el.scrollLeft=band.offsetLeft;});
    await frame.locator('#negotiatedPreview').screenshot({path:path.join(process.env.QA_OUT,'funds-month-header.png')});
   }
   assert.equal(JSON.stringify(input),originalJSON,'核心处理不修改源输入对象');for(const role of roles)assert.equal(hash(buffers[role]),hashes[role],'上传下载不修改源字节');
   if(headerRow===0&&withDates){
    const zip=await Zip.loadAsync(output),file='xl/worksheets/sheet1.xml';let xml=await zip.file(file).async('string');
    for(const [address,value] of [['W3',21],['D3',101],['F3',101],['K3',101],['AK3',41],['M3',202]]){
     let matches=0;xml=xml.replace(new RegExp('(<c\\b[^>]*\\br="'+address+'"[^>]*>)([\\s\\S]*?)(</c>)'),(_,start,body,end)=>{matches++;assert.match(body,/<v>[^<]*<\/v>/);return start+body.replace(/<v>[^<]*<\/v>/,'<v>'+value+'</v>')+end;});assert.equal(matches,1,address+' 手工金额修改');
    }
    zip.file(file,xml);const edited=await zip.generateAsync({type:'nodebuffer'}),editedHash=hash(edited),editedBook=read(edited);
    await upload('negotiated',edited);const repeated=await download(true),again=verifyBook(repeated,editedBook,1,0,true);
    assert.equal(again.Sheets.谈判付款.W3.v,21,'重复上传保留人工修改后的金额');assert.equal(again.Sheets.谈判付款['!ref'],saved.Sheets.谈判付款['!ref'],'重复导出不再增加表头行');assert.equal(hash(edited),editedHash);
   }
  }
  function referenceBook(formula){
   const headers=Array.from({length:42},(_,c)=>c===16?'2026年9月2日':c===17?'2026年9月16日':'保留标题'+c),sheet=X.utils.aoa_to_sheet([headers,['甲'],['乙']]);
   sheet.Q2={t:'n',v:10};sheet.R2={t:'n',v:20};sheet.Q3={t:'n',v:30};sheet.R3={t:'n',v:40};sheet.S2={t:'n',v:100,f:formula};sheet['!autofilter']={ref:'A1:AP3'};
   const book=X.utils.book_new();X.utils.book_append_sheet(book,sheet,'谈判付款');
   book.Workbook={Names:[{Name:'_xlnm.Print_Area',Sheet:0,Ref:"'谈判付款'!$A$1:$AP$3"},{Name:'_xlnm.Print_Titles',Sheet:0,Ref:"'谈判付款'!$1:$1"},{Name:'_xlnm._FilterDatabase',Sheet:0,Ref:"'谈判付款'!$A$1:$AP$3"}]};
   return write(book);
  }
  const addHeader=async bytes=>Buffer.from(await frame.locator('body').evaluate(async(_,{bytes})=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes));
   await FundsMonthHeader(zip,{sheetName:'谈判付款',headerRow:0,headerTopRow:null,headerBands:[{startCol:16,endCol:17,label:'2026年9月'}]});
   return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...bytes]}));
  for(const formula of ['COUNTA(A1:AP1)','SUM(1:1)','COUNTBLANK(A1:A3)','COUNTA(Q:Q)'])await assert.rejects(()=>addHeader(referenceBook(formula)),error=>{
   assert.match(error.message,/月份表头/);assert.match(error.message,/表头.*范围|范围.*表头/);assert.match(error.message,/安全/);return true;
  },formula+' 表头范围变形不能静默改变计算含义');
  for(const [formula,expected] of [['SUM(Q2:R3)','SUM(Q3:R4)'],['COUNTA(Q1:R1)','COUNTA(Q2:R2)'],['COUNTA(A:B)','COUNTA(A:B)']]){
   const saved=read(await addHeader(referenceBook(formula)));assert.equal(saved.Sheets.谈判付款.S3.f,expected,'安全范围引用正确下移');
   const names=saved.Workbook.Names;assert.equal(names.find(name=>name.Name==='_xlnm.Print_Area').Ref,"'谈判付款'!$A$1:$AP$4");assert.equal(names.find(name=>name.Name==='_xlnm.Print_Titles').Ref,"'谈判付款'!$1:$2");assert.equal(names.find(name=>name.Name==='_xlnm._FilterDatabase').Ref,"'谈判付款'!$A$2:$AP$4");
  }
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,groupedHeader:true,filterOnLowerHeader:true,amountsAndAItoAPPreserved:true,formulasCommentsMergesAndReferencesShifted:true,reimportAndManualEdit:true,shiftedHeader:true,noDatesNoInsert:true,lastRowFormattingCleanedBeforeInsert:true,unsafeHeaderFormulaRangesBlocked:true,safeRangesAndBuiltinNamesShifted:true,sourceUnchanged:true}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error.name+': '+error.message);process.exitCode=1;});
