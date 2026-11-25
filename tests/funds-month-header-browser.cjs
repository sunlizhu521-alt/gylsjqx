const assert=require('node:assert/strict'),fs=require('node:fs/promises'),crypto=require('node:crypto'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js'),Zip=require('../vendor/jszip.min.js'),C=require('../funds-core.js');
const {fixture}=require('./funds-progress-fixtures.cjs');
const roles=['ledger','applications','oa','negotiated'],write=book=>Buffer.from(X.write(structuredClone(book),{type:'buffer',bookType:'xlsx'}));
const hash=bytes=>crypto.createHash('sha256').update(bytes).digest('hex'),read=bytes=>X.read(bytes,{cellFormula:true,cellNF:true,cellStyles:true});
const serial=s=>(Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
const days=['2026-09-02','2026-09-16','2026-10-02','2026-10-16','2026-11-02','2026-11-09','2026-11-30'];
const paymentDays=[...days.slice(0,6),'2026-12-10'];
const bands=[{startCol:16,endCol:17,label:'2026年9月付款'},{startCol:18,endCol:19,label:'2026年10月付款'},{startCol:20,endCol:21,label:'2026年11月付款'},{startCol:22,endCol:22,label:'2026年12月付款'}];
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
 headers[23]='是否谈判分批';
 headers.push(...['2026年9月支付','2026年10月支付','2026年11月支付','2026年12月支付','2027年1月支付','2027年2月支付','2027年3月支付','2027年4月支付']);
 const amounts=[100,0,25],plans=[[1,2,3,4,5,6,79],[0,0,0,0,0,0,0],[0,0,0,0,25,0,0]];
 const data=amounts.map((amount,i)=>{
  const values=Array(42).fill('');
  values.splice(0,16,'表头合成'+(i+1),'迈德斯特','未提交OA',amount,0,amount,amount?'否':'是',serial('2026-11-30'),'保留人工备注'+i,paymentDays.map((day,c)=>X.SSF.format(C.DATE_FORMAT,serial(day))+'－'+plans[i][c]).join('\n'),amount,'一致',amount*2,i===0?'手工合并备注':'','引用字符串D2','人工'+i);
  values.splice(16,7,...plans[i]);
  values.splice(34,8,plans[i][0]+plans[i][1],plans[i][2]+plans[i][3],plans[i][4]+plans[i][5],plans[i][6],0,0,0,0);
  values[23]='是';
  return values;
 });
 const prefix=Array.from({length:headerRow},(_,i)=>['人工标题'+(i+1)]),sheet=X.utils.aoa_to_sheet([...prefix,headers,...data]),h=headerRow+1,first=h+1;
 if(withDates)for(let c=16;c<=22;c++)sheet[X.utils.encode_cell({r:headerRow,c})].z=C.DATE_FORMAT;
 for(let i=0;i<data.length;i++){
  const r=first+i,amount=amounts[i];
  sheet['F'+r]={t:'n',v:amount,f:'D'+r+'-E'+r,z:'#,##0.00'};
  sheet['K'+r]={t:'n',v:amount,f:'SUM(Q'+r+':W'+r+')'};
  sheet['M'+r]={t:'n',v:amount*2,f:'SUM(Q'+r+':W'+r+')+$D$'+r};
  for(const [column,range] of [['AI','Q'+r+':R'+r],['AJ','S'+r+':T'+r],['AK','U'+r+':V'+r],['AL','W'+r]])sheet[column+r].f='SUM('+range+')';
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

function addOriginalBands(input,headerRow){
 const source=input.negotiated.book.Sheets[sourceName],sheet={};
 for(const [a,c]of Object.entries(source)){
  if(!/^[A-Z]+[1-9]\d*$/.test(a)){sheet[a]=c;continue;}
  const p=X.utils.decode_cell(a);sheet[X.utils.encode_cell({r:p.r>=headerRow?p.r+1:p.r,c:p.c})]={...c,...(c.f?{f:c.f.replace(/(\$?[A-Z]+\$?)(\d+)/g,(_,col,row)=>col+(Number(row)+1))}:{})};
 }
 const range=X.utils.decode_range(source['!ref']);range.e.r++;sheet['!ref']=X.utils.encode_range(range);
 sheet['!merges']=source['!merges'].map(m=>({s:{r:m.s.r+1,c:m.s.c},e:{r:m.e.r+1,c:m.e.c}}));
 for(let c=0;c<=range.e.c;c++)if(c<16||c>22){const top=X.utils.encode_cell({r:headerRow,c}),bottom=X.utils.encode_cell({r:headerRow+1,c});sheet[top]=sheet[bottom];delete sheet[bottom];sheet['!merges'].push({s:{r:headerRow,c},e:{r:headerRow+1,c}});}
 // Deliberately different from the dates/details: the user's labels must survive.
 for(const band of bands){sheet[X.utils.encode_cell({r:headerRow,c:band.startCol})]={t:'s',v:band.label.replace('付款','')};if(band.endCol>band.startCol)sheet['!merges'].push({s:{r:headerRow,c:band.startCol},e:{r:headerRow,c:band.endCol}});}
 input.negotiated.book.Sheets[sourceName]=sheet;
 input.negotiated.book.Sheets.辅助引用.B2.f="'"+sourceName+"'!$D$"+(headerRow+3)+'+A2';
 input.negotiated.book.Workbook.Names[0].Ref="'"+sourceName+"'!$D$"+(headerRow+3)+':$D$'+(headerRow+5);
}
function verify(bytes,input,headerRow,double){
 const saved=read(bytes),sheet=saved.Sheets.谈判付款,source=input.negotiated.book.Sheets[sourceName],first=headerRow+2;
 assert.equal(X.utils.decode_range(sheet['!ref']).e.r,X.utils.decode_range(source['!ref']).e.r);
 assert.deepEqual(sheet['!merges'],source['!merges'],'原合并范围保留');
 for(const [a,c]of Object.entries(source))if(/^[A-Z]+[1-9]\d*$/.test(a)){
  const p=X.utils.decode_cell(a),actual=sheet[a];
  if(p.r<=headerRow){
   if(c.v==null)continue;
   const expected=typeof c.v==='number'&&c.z&&X.SSF.is_date(c.z)?X.SSF.format(c.z,c.v):c.v;
   assert.equal(actual?.v,expected,a+' 原表头保留');
  }else if(!['E','F','G'].includes(X.utils.encode_col(p.c))){
   assert.equal(actual?.f,c.f,a+' 原公式保留');
   if(c.z&&X.SSF.is_date(c.z))continue;
   assert.equal(actual?.v,c.v,a+' 原金额及文字保留');
  }
 }
 assert.match(sheet['Q'+first].c[0].t,/金额批注跟随原记录/);
 assert.equal(sheet['N'+first].c[0].t,'合并备注批注');
 assert.equal(saved.Sheets.辅助引用.B2.f,"'谈判付款'!$D$"+first+'+A2');
 assert.equal(saved.Sheets.辅助引用.B3.f,input.negotiated.book.Sheets.辅助引用.B3.f);
 assert.equal(saved.Workbook.Names.find(n=>n.Name==='谈判金额范围').Ref,"'谈判付款'!$D$"+first+':$D$'+(first+2));
 assert.equal(X.utils.decode_range(sheet['!autofilter'].ref).s.r,headerRow);
 assert.equal(C.build({...input,negotiated:{book:saved,name:'重传.xlsx',sheetName:'谈判付款'}}).negotiated.headerRow,headerRow);
 return saved;
}
(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8940/';
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage({acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/*',r=>new URL(r.request().url()).origin===new URL(base).origin?r.continue():r.abort());
  await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
  const frame=page.frameLocator('iframe[data-tool="funds"]');
  for(const [offset,double]of [[0,false],[2,false],[0,true],[2,true]]){
   const {input}=sample(offset);if(double)addOriginalBands(input,offset);
   const headerRow=offset+Number(double),before=JSON.stringify(input);
   for(const role of roles)await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:'合成-'+role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:write(input[role].book)});
   for(let round=0;round<2;round++){
    await frame.locator('#analyze:not([disabled])').click();await frame.locator('#download:not([disabled])').waitFor({timeout:30000});
    assert.equal(await frame.locator('#pendingCount').innerText(),'0');
    assert.equal(await frame.locator('#negotiatedPreview thead tr').count(),double?2:1);
    assert.equal(await frame.locator('#negotiatedPreview tbody tr').count(),3);
    const event=page.waitForEvent('download');await frame.locator('#download').click();const output=await fs.readFile(await(await event).path());
    verify(output,input,headerRow,double);
    if(!round)await frame.locator('#negotiatedFile:not([disabled])').setInputFiles({name:'重传.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:output});
   }
   assert.equal(JSON.stringify(input),before);
  }
  assert.deepEqual(errors,[]);console.log(JSON.stringify({status:'PASS',base,downloads:8,sourceMonthHeadersPreserved:true,noHeaderInsertion:true,formulasCommentsMergesPreserved:true,reimport:true}));
 }finally{await browser.close();}
})().catch(e=>{console.error(e.stack);process.exitCode=1;});
