const assert=require('node:assert/strict'),path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const X=require('../vendor/xlsx.full.min.js');
const MARKER='N("GYLSJQX_FUNDS_APPLIED_V1")=0';

(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
 try{
  const base=process.env.QA_URL||'http://127.0.0.1:8940/',page=await browser.newPage();
  await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
  await page.goto(new URL('funds.html',base).href);
  await page.addScriptTag({path:path.resolve('funds-export.js')});

  const headers=['供应商','已付金额','剩余金额','2026-11-11','提交付款申请单日期','提交时间汇总','手写空白','手写零值','付款自然周','手写标记','手写批注','手写合并','手写历史批注','手写无图批注','手写孤立图形','手写复杂公式'];
  const source=X.utils.aoa_to_sheet([headers,['供应商甲',10,21,50,'保留换行\n人工填写','富文本占位','',0,20,'原工具红色','手工批注载体','', '历史工具行也保留','没有批注图形','孤立图形保留','非零'],['供应商乙',8,17,60,'第二笔人工内容',null,null,false,16,16,'合并后的手写内容']]);
  source.C2={t:'n',v:21,f:'I2+1'};
  source.I2={t:'n',v:20,f:'B2*2'};
  source.P2={t:'s',v:'非零',f:'IF(B2>0,"非零","零")'};
  source.D2.c=[{a:'业务原作者',t:'保留受管人工说明\n申请日期（资金整理）：2026年10月1日'}];
  for(const [address,text]of [['K2','人工确认：不改此批注'],['M2','手工保存记录\n申请日期（付款节奏跟进）：2026年10月1日'],['N2','保留没有图形的原批注'],['O2','仅用于建立原图形']])source[address].c=[{a:'人工原作者',t:text}];
  source['!ref']='A1:P3';
  source['!merges']=[{s:{r:2,c:10},e:{r:2,c:11}}];
  source['!rows']=[{hpt:44},{hpt:90,hidden:true},{hpt:70}];
  source['!cols']=Array.from({length:16},(_,c)=>({wch:12+c*2,...(c===11?{hidden:true}:{})}));
  source['!autofilter']={ref:'A1:P3'};
  const book=X.utils.book_new();X.utils.book_append_sheet(book,source,'谈判付款');
  const input=Buffer.from(await page.evaluate(async({bytes,marker})=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
   const all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=t=>p.parseFromString(t,'application/xml');
   const make=(doc,name,attrs={})=>{const e=doc.createElementNS(ns,name);for(const [k,v]of Object.entries(attrs))e.setAttribute(k,String(v));return e;};
   const styles=parse(await zip.file('xl/styles.xml').async('string')),sheet=parse(await zip.file('xl/worksheets/sheet1.xml').async('string'));
   const cells=new Map(all(sheet,'c').map(c=>[c.getAttribute('r'),c])),rows=new Map(all(sheet,'row').map(r=>[r.getAttribute('r'),r]));
   const fonts=all(styles,'fonts')[0],fills=all(styles,'fills')[0],borders=all(styles,'borders')[0],xfs=all(styles,'cellXfs')[0],baseXfs=all(styles,'cellStyleXfs')[0],namedStyles=all(styles,'cellStyles')[0];
   const append=(list,node)=>{const id=list.children.length;list.append(node);list.setAttribute('count',list.children.length);return id;};
   const font=fonts.children[0].cloneNode(true);all(font,'sz')[0].setAttribute('val','17');font.append(make(styles,'i'));const fontId=append(fonts,font);
   const fill=color=>{const f=make(styles,'fill'),pattern=make(styles,'patternFill',{patternType:'solid'});pattern.append(make(styles,'fgColor',{rgb:color}),make(styles,'bgColor',{rgb:color}));f.append(pattern);return f;};
   const greenId=append(fills,fill('FFB6D7A8')),redId=append(fills,fill('FFFF0000')),border=make(styles,'border');
   for(const edge of ['left','right','top','bottom']){const e=make(styles,edge,{style:'double'});e.append(make(styles,'color',{rgb:'FF7030A0'}));border.append(e);}const borderId=append(borders,border);
   const custom=make(styles,'xf',{numFmtId:4,fontId,fillId:greenId,borderId,xfId:0,applyFont:1,applyFill:1,applyBorder:1,applyAlignment:1,applyProtection:1});
   custom.append(make(styles,'alignment',{horizontal:'left',vertical:'top',wrapText:1,indent:2}),make(styles,'protection',{locked:0}));
   const customId=append(xfs,custom),ownedBase=append(baseXfs,custom.cloneNode(true));
   append(namedStyles,make(styles,'cellStyle',{name:'GYLSJQX_FUNDS_APPLIED_V2_'+ownedBase,xfId:ownedBase,hidden:1}));
   const owned=custom.cloneNode(true);owned.setAttribute('fillId',redId);owned.setAttribute('xfId',ownedBase);const ownedId=append(xfs,owned);
   for(const [address,c]of cells)if(XLSX.utils.decode_cell(address).c>=4)c.setAttribute('s',customId);
   cells.get('J2').setAttribute('s',ownedId);cells.get('B3').setAttribute('s',ownedId);
   for(const col of all(sheet,'col'))if(Number(col.getAttribute('min'))>=5){col.setAttribute('style',customId);col.setAttribute('outlineLevel','1');}
   const rich=cells.get('F2');rich.replaceChildren();rich.setAttribute('t','inlineStr');const inline=make(sheet,'is');
   for(const [value,format]of [['人工加粗\n','b'],['人工斜体，勿改','i']]){const run=make(sheet,'r'),props=make(sheet,'rPr'),t=make(sheet,'t');props.append(make(sheet,format));t.textContent=value;t.setAttribute('xml:space','preserve');run.append(props,t);inline.append(run);}rich.append(inline);
   let blank=cells.get('G2');if(!blank){blank=make(sheet,'c',{r:'G2',s:customId});rows.get('2').insertBefore(blank,cells.get('H2'));cells.set('G2',blank);}blank.replaceChildren();blank.removeAttribute('t');
   const extensions=make(sheet,'extLst'),ext=make(sheet,'ext',{uri:'urn:funds-custom-column-test'}),flag=sheet.createElementNS('urn:funds-custom-column-test','custom:flag');flag.textContent='人工元信息';ext.append(flag);extensions.append(ext);blank.append(extensions);
   // Preserve an untouched shared-formula group entirely within the custom columns.
   for(const [address,formula]of [['I3','B3*2'],['J3','']]){const c=cells.get(address),f=make(sheet,'f',{t:'shared',si:11,...(address==='I3'?{ref:'I3:J3'}:{})});f.textContent=formula;c.insertBefore(f,c.firstChild);}
   const filter=all(sheet,'autoFilter')[0],filterColumn=make(sheet,'filterColumn',{colId:4}),filters=make(sheet,'filters'),filterValue=make(sheet,'filter',{val:'保留换行\n人工填写'});filters.append(filterValue);filterColumn.append(filters);filter.append(filterColumn);
   let dxfs=all(styles,'dxfs')[0];if(!dxfs){dxfs=make(styles,'dxfs',{count:0});styles.documentElement.append(dxfs);}const dxf=make(styles,'dxf');dxf.append(fill('FFFF0000'));const dxfId=append(dxfs,dxf);
   const addRule=(ref,formula,priority)=>{const cf=make(sheet,'conditionalFormatting',{sqref:ref}),rule=make(sheet,'cfRule',{type:'expression',dxfId,priority}),f=make(sheet,'formula');f.textContent=formula;rule.append(f);cf.append(rule);sheet.documentElement.insertBefore(cf,[...sheet.documentElement.children].find(n=>['dataValidations','hyperlinks','pageMargins','ignoredErrors','legacyDrawing','extLst'].includes(n.localName))||null);};
   addRule('B2:J3',marker,1);addRule('M2:P3',marker,2);addRule('H2:H3','$H$2=0',3);
   for(const name of Object.keys(zip.files).filter(n=>/comments[^/]*\.xml$/i.test(n))){const doc=parse(await zip.file(name).async('string'));all(doc,'comment').find(c=>c.getAttribute('ref')==='O2').remove();zip.file(name,s.serializeToString(doc));}
   for(const name of Object.keys(zip.files).filter(n=>n.endsWith('.vml'))){const doc=parse(await zip.file(name).async('string'));for(const shape of all(doc,'shape')){const data=all(shape,'ClientData')[0];if(data?.getAttribute('ObjectType')!=='Note')continue;const col=Number(all(data,'Column')[0].textContent);if(col===13){shape.remove();continue;}if(col>=4){all(data,'Anchor')[0].textContent='30, 17, 80, 9, 34, 20, 84, 10';shape.setAttribute('style','position:absolute;margin-left:2000pt;margin-top:1600pt;width:250pt;height:120pt;visibility:hidden');}}zip.file(name,s.serializeToString(doc));}
   zip.file('xl/styles.xml',s.serializeToString(styles));zip.file('xl/worksheets/sheet1.xml',s.serializeToString(sheet));return [...await zip.generateAsync({type:'uint8array'})];
  },{bytes:[...X.write(book,{type:'buffer',bookType:'xlsx'})],marker:MARKER}));

  const managedColumns=[0,1,2,3,16],allHeaders=[...headers,'流程编号'];
  const result={normalHeaders:['供应商名称'],normalRows:[],reviewHeaders:['核对说明'],reviewRows:[],applicationComments:{D2:'申请日期（付款节奏跟进）：2026年11月11日'},highlights:{negotiatedCells:['D2'],normalRows:[]},negotiated:{sheetName:'谈判付款',headerRow:0,headers:allHeaders,sourceRange:X.utils.decode_range('A1:P3'),managedColumns,columnOrder:allHeaders.map((_,c)=>c),rows:[1,2].map(r=>allHeaders.map((_,c)=>source[X.utils.encode_cell({r,c})])),updates:[{address:'B2',cell:{t:'n',v:20}},{address:'Q1',cell:{t:'s',v:'流程编号'}},{address:'Q2',cell:{t:'s',v:'FLOW-1'}}]}};
  const exportBytes=async(bytes,data)=>{
   const saved=await page.evaluate(async({bytes,result})=>{const blob=await FundsExport(Uint8Array.from(bytes),result);return {bytes:[...new Uint8Array(await blob.arrayBuffer())],warnings:FundsExport.lastWarnings};},{bytes:[...bytes],result:data});
   assert.deepEqual(saved.warnings,[],'Custom formulas must not cause export-cache warnings');return Buffer.from(saved.bytes);
  };
  const inspect=async bytes=>page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],parse=t=>p.parseFromString(t,'application/xml');
   const attrs=n=>Object.fromEntries([...n.attributes].filter(a=>a.namespaceURI!=='http://www.w3.org/2000/xmlns/').map(a=>[a.name,a.value]).sort(([a],[b])=>a.localeCompare(b)));
   const tree=n=>!n?null:n.nodeType===3?n.textContent:{name:n.localName,namespace:n.namespaceURI,attrs:attrs(n),children:[...n.childNodes].filter(n=>[1,3].includes(n.nodeType)).map(tree)};
   const sheet=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),styles=parse(await zip.file('xl/styles.xml').async('string')),xfs=all(styles,'cellXfs')[0],fonts=all(styles,'fonts')[0],fills=all(styles,'fills')[0],borders=all(styles,'borders')[0],custom=c=>c>=4&&c<=15;
   const snapshot={cells:{},styles:{},columns:{},rows:all(sheet,'row').map(attrs),format:tree(all(sheet,'sheetFormatPr')[0]),merges:all(sheet,'mergeCell').map(attrs),filter:all(sheet,'autoFilter').map(n=>[...n.children].map(tree)),rules:{},comments:{},noteShapes:{}};
   for(const cell of all(sheet,'c'))if(custom(XLSX.utils.decode_cell(cell.getAttribute('r')).c)){snapshot.cells[cell.getAttribute('r')]=tree(cell);const id=Number(cell.getAttribute('s')||0),xf=xfs.children[id];snapshot.styles[id]={xf:tree(xf),font:tree(fonts.children[Number(xf.getAttribute('fontId')||0)]),fill:tree(fills.children[Number(xf.getAttribute('fillId')||0)]),border:tree(borders.children[Number(xf.getAttribute('borderId')||0)])};}
   for(let c=4;c<=15;c++){const col=all(sheet,'col').find(n=>Number(n.getAttribute('min'))<=c+1&&Number(n.getAttribute('max'))>=c+1);if(!col){snapshot.columns[c]=null;continue;}const details=attrs(col);delete details.min;delete details.max;snapshot.columns[c]=details;}
   for(let c=4;c<=15;c++)for(let r=1;r<=2;r++){const address=XLSX.utils.encode_cell({r,c});snapshot.rules[address]=all(sheet,'conditionalFormatting').filter(cf=>(cf.getAttribute('sqref')||'').split(/\s+/).some(ref=>{const range=XLSX.utils.decode_range(ref);return r>=range.s.r&&r<=range.e.r&&c>=range.s.c&&c<=range.e.c;})).flatMap(cf=>all(cf,'cfRule').map(tree));}
   for(const name of Object.keys(zip.files).filter(n=>/comments[^/]*\.xml$/i.test(n))){const doc=parse(await zip.file(name).async('string')),authors=all(doc,'author');for(const comment of all(doc,'comment'))if(custom(XLSX.utils.decode_cell(comment.getAttribute('ref')).c))snapshot.comments[comment.getAttribute('ref')]={xml:tree(comment),author:tree(authors[Number(comment.getAttribute('authorId'))])};}
   for(const name of Object.keys(zip.files).filter(n=>n.endsWith('.vml'))){const doc=parse(await zip.file(name).async('string'));for(const shape of all(doc,'shape')){const data=all(shape,'ClientData')[0];if(data?.getAttribute('ObjectType')!=='Note')continue;const c=Number(all(data,'Column')[0].textContent),r=Number(all(data,'Row')[0].textContent);if(custom(c))snapshot.noteShapes[XLSX.utils.encode_cell({r,c})]=tree(shape);}}
   return snapshot;
  },[...bytes]);

  const before=await inspect(input),output=await exportBytes(input,result),after=await inspect(output),saved=X.read(output,{cellStyles:true,cellNF:true});
  assert.deepEqual(after,before,'Custom cells, native styles, column layout, row layout, filters, merges, rules and notes stay unchanged');
  assert.equal(saved.Sheets.谈判付款.B2.v,20);assert.equal(saved.Sheets.谈判付款.C2.v,41,'Managed cache follows a changed value through a custom-column formula');
  assert.equal(saved.Sheets.谈判付款.I2.v,20,'The custom formula cache is not rewritten');assert.equal(saved.Sheets.谈判付款.I2.f,'B2*2');
  assert.equal(saved.Sheets.谈判付款.P2.v,'非零');assert.equal(saved.Sheets.谈判付款.H2.v,0);assert.equal(saved.Sheets.谈判付款.H3.v,false);
  assert.equal(saved.Sheets.谈判付款.G3,undefined,'An originally absent custom cell stays absent');
  assert.equal(saved.Sheets.谈判付款.J2.s.fgColor.rgb,'FF0000','Owned red in a custom column is retained');
  assert.equal(saved.Sheets.谈判付款.B3.s.fgColor.rgb,'B6D7A8','Owned red in a managed column can still be withdrawn');
  assert.equal(saved.Sheets.谈判付款.D2.s.fgColor.rgb,'FF0000');assert.equal(saved.Sheets.谈判付款.D2.c[0].t,'保留受管人工说明\n申请日期（付款节奏跟进）：2026年11月11日');
  assert.equal(saved.Sheets.谈判付款.Q2.v,'FLOW-1');assert.equal(saved.Sheets.谈判付款['!ref'],'A1:Q3');
  assert.equal(before.noteShapes.N2,undefined);assert(before.noteShapes.O2);assert.equal(before.comments.O2,undefined);

  const next={...result,applicationComments:{D2:'申请日期（付款节奏跟进）：2026年11月12日'},negotiated:{...result.negotiated,sourceRange:X.utils.decode_range('A1:Q3'),updates:[{address:'B2',cell:{t:'n',v:30}},{address:'Q2',cell:{t:'s',v:'FLOW-2'}}]}};
  const repeated=await exportBytes(output,next),again=X.read(repeated,{cellNF:true});
  assert.deepEqual(await inspect(repeated),before,'Re-uploading retains the same custom-column content and native metadata');
  assert.equal(again.Sheets.谈判付款.C2.v,61);assert.equal(again.Sheets.谈判付款.I2.v,20);assert.equal(again.Sheets.谈判付款.Q2.v,'FLOW-2');

  await assert.rejects(()=>exportBytes(input,{...result,negotiated:{...result.negotiated,updates:[{address:'E2',cell:{t:'s',v:'越界覆盖'}}]}}),/更新超出业务列：E2/);
  await assert.rejects(()=>exportBytes(input,{...result,applicationComments:{},highlights:{negotiatedCells:['E2'],normalRows:[]}}),/标红超出业务列：E2/);
  await assert.rejects(()=>exportBytes(input,{...result,applicationComments:{E2:'申请日期（付款节奏跟进）：2026年11月11日'},highlights:{negotiatedCells:['E2'],normalRows:[]}}),/批注超出业务列：E2/);
  await assert.rejects(()=>exportBytes(input,{...result,negotiated:{...result.negotiated,columnOrder:[0,1,2,4,3,...allHeaders.map((_,c)=>c).slice(5)]}}),/原列顺序不能变更/);
  const shared=Buffer.from(await page.evaluate(async bytes=>{
   const zip=await JSZip.loadAsync(Uint8Array.from(bytes)),p=new DOMParser(),s=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',doc=p.parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml'),cells=new Map([...doc.getElementsByTagNameNS('*','c')].map(c=>[c.getAttribute('r'),c]));
   for(const address of ['C2','D2','E2']){const c=cells.get(address);c.replaceChildren();c.removeAttribute('t');const f=doc.createElementNS(ns,'f'),v=doc.createElementNS(ns,'v');f.setAttribute('t','shared');f.setAttribute('si','7');if(address==='C2'){f.setAttribute('ref','C2:E2');f.textContent='B2*2';}v.textContent='20';c.append(f,v);}
   zip.file('xl/worksheets/sheet1.xml',s.serializeToString(doc));return [...await zip.generateAsync({type:'uint8array'})];
  },[...input]));
  await assert.rejects(()=>exportBytes(shared,{...result,negotiated:{...result.negotiated,updates:[{address:'C2',cell:{t:'n',v:40}}]}}),/共享公式跨越业务列和保留列/);
  assert.deepEqual(await inspect(input),before,'Source bytes remain unchanged after successful and rejected exports');
  console.log('PASS: custom-column native cells/rich text/blanks/zero/formula caches/styles/widths/hidden states/merges/notes/owned fills/conditional formats, repeat export, dependencies through custom formulas, and out-of-scope writes/highlights/comments/reordering/shared-formula rejection');
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
