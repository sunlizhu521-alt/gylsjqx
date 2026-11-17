const assert=require('node:assert/strict');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');

(async()=>{
 const base=process.env.QA_URL||'http://127.0.0.1:8963/',browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const page=await browser.newPage(),errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('funds.html',base).href);
  assert.equal(await page.evaluate(()=>typeof FundsCleanup),'function','已发布页面加载清理器');
  const book=X.utils.book_new(),source=X.utils.aoa_to_sheet([
   ['供应商','零值','公式','人工空白','布尔值','合并说明'],
   ['合成供应商',0,0,'',false,'保留合并'],
   ['合计',0,'']
  ]);
  source.C2={t:'n',v:0,f:'B2+0'};source.C3={t:'s',v:'',f:'IF(B3=0,"","有值")'};
  source.H5={t:'s',v:'',c:[{a:'原作者',t:'只有批注的单元格仍须保留'}]};
  source.G4={t:'s',v:'链接',l:{Target:'#A1'}};
  source['!ref']='A1:H5';source['!merges']=[{s:{r:1,c:5},e:{r:2,c:6}}];
  source['!rows']=[{hpt:33},{hpt:47,hidden:true},{hpt:30},{hpt:24}];
  source['!cols']=Array.from({length:8},(_,c)=>({wch:14+c,...(c===3?{hidden:true}:{})}));
  source['!autofilter']={ref:'A1:H5'};
  X.utils.book_append_sheet(book,source,'谈判付款');
  const auxiliary=X.utils.aoa_to_sheet([]);X.utils.sheet_add_aoa(auxiliary,[['无需变化的辅助表'],[0]],{origin:'C3'});auxiliary['!ref']='C3:C4';X.utils.book_append_sheet(book,auxiliary,'辅助表');
  const input=Buffer.from(await page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
   const all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=t=>p.parseFromString(t,'application/xml'),make=(d,n,a={})=>{const e=d.createElementNS(ns,n);for(const [k,v]of Object.entries(a))e.setAttribute(k,String(v));return e;};
   const d=parse(await z.file('xl/worksheets/sheet1.xml').async('string')),styles=parse(await z.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],xf=xfs.children[0].cloneNode(true),alignment=make(styles,'alignment',{horizontal:'center',wrapText:1});xf.append(alignment);xf.setAttribute('applyAlignment','1');const style=xfs.children.length;xfs.append(xf);xfs.setAttribute('count',xfs.children.length);
   const data=all(d,'sheetData')[0],rows=new Map(all(d,'row').map(r=>[Number(r.getAttribute('r')),r]));
   function blank(address){const r=Number(address.match(/\d+$/)[0]);let row=rows.get(r);if(!row){row=make(d,'row',{r,s:style,customFormat:1,ht:30,customHeight:1});data.append(row);rows.set(r,row);}const old=[...row.children].find(c=>c.getAttribute('r')===address);if(old)old.remove();const c=make(d,'c',{r:address,s:style});row.append(c);return c;}
   blank('D2');blank('XFD121');blank('XFC24');blank('A1048576');blank('XFD1048576');
   const shared=parse(await z.file('xl/sharedStrings.xml').async('string'));
   for(const [address,value]of [['XFD120',''],['H4','   ']]){const index=all(shared,'si').length,si=make(shared,'si'),t=make(shared,'t');t.textContent=value;t.setAttribute('xml:space','preserve');si.append(t);shared.documentElement.append(si);const c=blank(address),v=make(d,'v');c.setAttribute('t','s');v.textContent=String(index);c.append(v);}
   shared.documentElement.setAttribute('uniqueCount',all(shared,'si').length);shared.documentElement.setAttribute('count',Number(shared.documentElement.getAttribute('count'))+2);z.file('xl/sharedStrings.xml',s.serializeToString(shared));
   const custom=blank('G5'),exts=make(d,'extLst'),ext=make(d,'ext',{uri:'urn:blank-format-test'});ext.textContent='人工元信息';exts.append(ext);custom.append(exts);
   for(const row of rows.values())row.setAttribute('spans','1:16384');
   all(d,'dimension')[0].setAttribute('ref','A1:XFD1048576');all(d,'autoFilter')[0].setAttribute('ref','A1:XFD1048576');
   all(d,'cols')[0].append(make(d,'col',{min:9,max:16384,width:9,style,customWidth:1}));
   all(d,'mergeCells')[0].append(make(d,'mergeCell',{ref:'XFC500:XFD501'}));all(d,'mergeCells')[0].setAttribute('count','2');
   z.file('xl/styles.xml',s.serializeToString(styles));z.file('xl/worksheets/sheet1.xml',s.serializeToString(d));return [...await z.generateAsync({type:'uint8array'})];
  },[...X.write(book,{type:'buffer',bookType:'xlsx',bookSST:true})]));
  const clean=async bytes=>Buffer.from(await page.evaluate(async bytes=>[...await FundsCleanup(Uint8Array.from(bytes))],[...bytes]));
  const inspect=async bytes=>page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)];
   const d=p.parseFromString(await z.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),attrs=e=>Object.fromEntries([...e.attributes].map(a=>[a.name,a.value]));
   const parts={};for(const name of Object.keys(z.files).sort())if(!z.files[name].dir)parts[name]=await z.file(name).async('base64');
   return {parts,dimension:all(d,'dimension')[0]?.getAttribute('ref'),filter:all(d,'autoFilter')[0]?.getAttribute('ref'),cells:Object.fromEntries(all(d,'c').map(c=>[c.getAttribute('r'),s.serializeToString(c)])),rows:all(d,'row').map(attrs),cols:all(d,'col').map(attrs),merges:all(d,'mergeCell').map(e=>e.getAttribute('ref'))};
  },[...bytes]);
  const before=await inspect(input),output=await clean(input),after=await inspect(output);
  assert.equal(after.dimension,'A1:H5');assert.equal(after.filter,'A1:H5');
  for(const a of ['XFD120','XFD121','XFC24','A1048576','XFD1048576'])assert.equal(after.cells[a],undefined,a+'尾部空白样式（含空共享字符串）已移除');
  assert(after.rows.every(r=>Number(r.r)<=5));assert(after.cols.every(c=>Number(c.max)<=8));
  assert(after.rows.every(r=>!r.spans||Number(r.spans.split(':')[1])<=8));
  assert.deepEqual(after.merges,['F2:G3']);
  for(const a of ['A1','A2','B2','C2','C3','D2','E2','G4','G5','H4','H5'])assert.equal(after.cells[a],before.cells[a],a+'有效区域原内容、公式、样式或元信息保持');
  for(const r of after.rows){const old=before.rows.find(v=>v.r===r.r);for(const key of ['ht','hidden','customHeight','s','customFormat'])assert.equal(r[key],old[key]);}
  assert.deepEqual(after.cols,before.cols.filter(c=>Number(c.max)<=8));
  for(const [name,value]of Object.entries(before.parts))if(/comments|\.vml$|xl\/worksheets\/sheet2\.xml/.test(name))assert.equal(after.parts[name],value,name+'原生批注、图形或未改变辅助表保持字节一致');
  const reopened=X.read(output,{cellNF:true,cellStyles:true}).Sheets.谈判付款;
  assert.equal(reopened.B2.v,0);assert.equal(reopened.E2.v,false);assert.equal(reopened.C2.f,'B2+0');assert.equal(reopened.C3.f,'IF(B3=0,"","有值")');assert.equal(reopened.H4.v,'   ');assert.equal(reopened.H5.c[0].t,'只有批注的单元格仍须保留');
  assert.deepEqual((await inspect(await clean(output))).parts,after.parts,'重复清理幂等');assert.deepEqual(await inspect(input),before,'输入文件不改变');
  for(const [address,kind]of [['XFD10','zero'],['A1048576','formula'],['XFD10','spaces']]){
   const edge=Buffer.from(await page.evaluate(async({bytes,address,kind})=>{
    const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',d=p.parseFromString(await z.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),data=d.getElementsByTagNameNS('*','sheetData')[0],r=d.createElementNS(ns,'row'),c=d.createElementNS(ns,'c'),v=d.createElementNS(ns,'v');r.setAttribute('r',address.match(/\d+$/)[0]);c.setAttribute('r',address);c.setAttribute('t',kind==='formula'?'str':kind==='spaces'?'s':'n');v.textContent=kind==='formula'?'':'0';if(kind==='formula'){const f=d.createElementNS(ns,'f');f.textContent='""';c.append(f);}if(kind==='spaces'){const shared=p.parseFromString(await z.file('xl/sharedStrings.xml').async('string'),'application/xml');v.textContent=String([...shared.getElementsByTagNameNS('*','si')].findIndex(si=>[...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join('')==='   '));}c.append(v);r.append(c);data.append(r);z.file('xl/worksheets/sheet1.xml',s.serializeToString(d));return [...await z.generateAsync({type:'uint8array'})];
   },{bytes:[...output],address,kind}));
   const edgeBefore=await inspect(edge),edgeAfter=await inspect(await clean(edge));assert.equal(edgeAfter.cells[address],edgeBefore.cells[address],'边界真实零值/空结果公式/空格字符串不能当残留删除');
   assert.equal(edgeAfter.dimension,address==='XFD10'?'A1:XFD10':'A1:H1048576');
  }
  const drawingInput=Buffer.from(await page.evaluate(async bytes=>{
   const z=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',rn='http://schemas.openxmlformats.org/officeDocument/2006/relationships',parse=text=>p.parseFromString(text,'application/xml'),d=parse(await z.file('xl/worksheets/sheet1.xml').async('string')),all=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)],make=(doc,name,attrs={})=>{const e=doc.createElementNS(ns,name);for(const [k,v]of Object.entries(attrs))e.setAttribute(k,String(v));return e;};
   all(d,'cols')[0].append(make(d,'col',{min:9,max:30,width:41,customWidth:1}));
   for(const [r,ht]of [[10,88],[20,64],[30,90]])all(d,'sheetData')[0].append(make(d,'row',{r,ht,customHeight:1}));
   all(d,'dimension')[0].setAttribute('ref','A1:AD30');
   const drawing=make(d,'drawing');drawing.setAttributeNS(rn,'r:id','rIdBlankDrawing');d.documentElement.append(drawing);
   const relPath='xl/worksheets/_rels/sheet1.xml.rels',rels=parse(await z.file(relPath).async('string')),rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');rel.setAttribute('Id','rIdBlankDrawing');rel.setAttribute('Type',rn+'/drawing');rel.setAttribute('Target','../drawings/drawing-test.xml');rels.documentElement.append(rel);
   const types=parse(await z.file('[Content_Types].xml').async('string')),override=types.createElementNS(types.documentElement.namespaceURI,'Override');override.setAttribute('PartName','/xl/drawings/drawing-test.xml');override.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawing+xml');types.documentElement.append(override);
   z.file('xl/drawings/drawing-test.xml','<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>19</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="保留图形"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>');
   z.file('xl/worksheets/sheet1.xml',s.serializeToString(d));z.file(relPath,s.serializeToString(rels));z.file('[Content_Types].xml',s.serializeToString(types));return [...await z.generateAsync({type:'uint8array'})];
  },[...output]));
  const drawingBefore=await inspect(drawingInput),drawingAfter=await inspect(await clean(drawingInput));
  assert.equal(drawingAfter.dimension,'A1:T20','图形边界扩展保留范围');
  for(const r of ['10','20'])assert.deepEqual(drawingAfter.rows.find(row=>row.r===r),drawingBefore.rows.find(row=>row.r===r),'图形锚点依赖的自定义行高保留');
  assert(!drawingAfter.rows.some(row=>row.r==='30'));
  assert.deepEqual(drawingAfter.cols.find(col=>col.min==='9'),{min:'9',max:'20',width:'41',customWidth:'1'},'图形依赖列宽保留，边界外仍清除');
  assert.equal(drawingAfter.parts['xl/drawings/drawing-test.xml'],drawingBefore.parts['xl/drawings/drawing-test.xml'],'图形本体保持字节一致');
  const exported=Buffer.from(await page.evaluate(async bytes=>{
   const result={normalHeaders:['供应商名称'],normalRows:[],reviewHeaders:['核对说明'],reviewRows:[],progressHeaders:['供应商名称'],progressRows:[],totalValidation:{headers:['供应商名称'],rows:[]},negotiated:{sheetName:'谈判付款',headerRow:0,sourceRange:XLSX.utils.decode_range('A1:H5'),headers:['供应商','零值','公式','人工空白','布尔值','合并说明','',''],managedColumns:[0],columnOrder:[0,1,2,3,4,5,6,7],updates:[]}};
   const blob=await FundsExport(Uint8Array.from(bytes),result);return [...new Uint8Array(await blob.arrayBuffer())];
  },[...input]));
  const exportedSnapshot=await inspect(exported);assert.equal(exportedSnapshot.dimension,'A1:H5');assert.equal(exportedSnapshot.cells.XFD121,undefined,'正式导出调用清理器');assert.equal(exportedSnapshot.cells.A1048576,undefined);
  assert.deepEqual(errors,[]);console.log('PASS: 清除 XFD/末行空白格式、尾部行列与独立空白合并，保留零值/false/空结果公式/批注/有效区格式/元信息，清理幂等及正式导出集成');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
