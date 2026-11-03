/* The XLSX and PDF share values, row heights and page boundaries. */
window.InternalStatementDocument=(()=>{
'use strict';
const C=InternalStatement,W=595.28,H=841.89,M=24,widths=[70,160,65,50,80,122.28],FONT='"Songti SC","SimSun","Noto Serif CJK SC",serif';
const measure=document.createElement('canvas').getContext('2d');
function wrap(value,width,font){
 measure.font=font+'px '+FONT;const lines=[];
 for(const paragraph of String(value??'').split(/\r?\n/)){let line='';for(const ch of paragraph){if(line&&measure.measureText(line+ch).width>width){lines.push(line);line=ch;}else line+=ch;}lines.push(line);}
 return lines;
}
const amount=value=>value===''?'':Number(value),number=value=>({v:Number(value),t:'n',format:'0'+(String(value).includes('.')?'.'+'0'.repeat(String(value).split('.')[1].length):'')}),money=value=>({v:Number(value),t:'n',format:'#,##0.00'});
function layout(group,settings){
 const f=C.footer(group,settings),rows=[];
 function row(sourceRow,values,merges=[],config={}){
  const cells=Array.from({length:6},(_,i)=>{const v=values[i]??'';return v&&typeof v==='object'?{...v}:{v:String(v),t:'s'};});
  const font=config.font||(sourceRow>=44?8:9),objects=[];let height=config.height||(config.detail?16:sourceRow>=44?14:18);
  for(let c=0;c<6;c++){
   const merged=merges.find(([a,b])=>c>=a&&c<=b);if(merged&&merged[0]!==c)continue;
   const span=(merged?.[1]??c)-c+1,width=widths.slice(c,c+span).reduce((a,b)=>a+b,0),cell=cells[c];
   const value=cell.t==='n'?XLSX.SSF.format(cell.format||'General',cell.v):cell.v;
   const single=config.detail&&[0,2,3,4].includes(c);let size=font;
   if(single){measure.font=size+'px '+FONT;size=Math.min(size,(width-8)/Math.max(1,measure.measureText(value).width)*size);if(size<7)throw Error('物料 '+group.items[rows.length-5]?.code+' 的编码或数值过长，无法完整排版');}
   const lines=wrap(value,width-8,size);if(value!=='')height=Math.max(height,lines.length*(size*1.25)+4);
   objects.push({c,span,value,lines,font:size,single,bold:!!config.bold,align:config.align||(config.detail?(c===1||c===5?'left':c===0?'center':'right'):'left'),border:!!config.border||!!config.invoice&&c>=3});
  }
  rows.push({sourceRow,cells,merges,height,objects});
 }
 row(1,[group.buyer],[[0,5]],{font:15,bold:true,align:'center',height:27,border:true});
 row(2,[`${group.month.slice(0,4)}年${Number(group.month.slice(5))}月 交易明细对帐单`],[[0,5]],{font:12,bold:true,align:'center',height:22,border:true});
 row(3,['对账日期：'+C.dateLabel(f.reconcileDate)],[[0,5]],{align:'right',height:18,border:true});
 row(4,['供应商名称：'+group.supplier,'','','','付款方式：'+f.terms],[[0,3],[4,5]],{border:true});
 row(5,C.HEADERS,[],{bold:true,align:'center',border:true});
 for(const item of group.items)row(6,[item.code,item.name,{...number(item.price),format:'0.00'},number(item.quantity),money(C.amountText(item.amount)),item.note],[],{detail:true,border:true});
 const totalRow=rows.length+1;
 row(43,['数量及金额合计','','',{...number(group.quantity),f:`SUM(D6:D${totalRow-1})`},{...money(C.amountText(group.amount)),f:`SUM(E6:E${totalRow-1})`}],[[0,2]],{bold:true,border:true});
 const footerStart=rows.length;
 const val=s=>s===''?'':money(s),footerFormulaRow=footerStart+1;
 const invoiceAt=index=>{const a=f.invoices[index];return [a.month?a.month.replace('-','年')+'月':'',a.status,val(a.amount)];};
 row(44,['上期余额：','',val(f.balance)],[[0,1]]);
 row(45,['加：本期发货金额：','',{...money(C.amountText(group.amount)),f:`E${totalRow}`},'货款汇总表'],[[0,1],[3,5]],{invoice:true});
 row(46,['减：本期收款：','',val(f.receipts),'月份','是否开票','开具金额'],[[0,1]],{invoice:true});
 row(47,['减：售后异常费用：','',val(f.afterSales),...invoiceAt(0)],[[0,1]],{invoice:true});
 row(48,['减：本期退货金额：','',val(f.returns),...invoiceAt(1)],[[0,1]],{invoice:true});
 const r=footerFormulaRow;
 row(49,['合计：','',{v:amount(f.total),t:f.total===''?'str':'n',format:'#,##0.00',f:`IF(COUNT(C${r},C${r+2}:C${r+4})=4,C${r}+C${r+1}-SUM(C${r+2}:C${r+4}),"")`},...invoiceAt(2)],[[0,1]],{bold:true,invoice:true});
 row(50,['','','',...invoiceAt(3)],[],{invoice:true});
 row(51,[`截止${f.cutoff?C.dateLabel(f.cutoff):'____年__月__日'}，贵司尚欠我司已开票货款￥${f.invoiced}`,'','',...invoiceAt(4)],[[0,2]],{invoice:true});
 row(52,['未开票货款￥'+f.uninvoiced],[[0,2]]);
 row(53,[],[],{height:8});
 row(54,['备注：'+f.note],[[0,5]]);
 row(55,[],[],{height:8});
 row(56,['注：1.货已发到贵司请查收'],[[0,5]]);
 row(57,['2.数量、金额如有误请在空白之处注明'],[[0,5]]);
 row(58,['3.如对以上对账内容无异议，请回传确认；如有异议请电询。 TEL：'+f.phone],[[0,5]]);
 row(59,[],[],{height:10});
 row(60,['客户确认（盖章）：'],[[0,5]],{height:20});
 row(61,['经收人（签名）：'],[[0,5]],{height:20});
 row(62,['日期：'],[[0,5]],{height:20});
 const headerHeight=rows.slice(0,5).reduce((n,r)=>n+r.height,0),max=H-M-18,pages=[],breaks=[];
 let page=[],y=M;const pushHeader=()=>{page=[];y=M;for(let i=0;i<5;i++){page.push({row:rows[i],index:i,y});y+=rows[i].height;}};pushHeader();
 const footerHeight=rows.slice(footerStart).reduce((n,r)=>n+r.height,0);
 for(let i=5;i<rows.length;i++){
  const row=rows[i];if(row.height>max-M-headerHeight)throw Error('第 '+(i+1)+' 行内容过长，无法在一页内完整展示，请缩短该行备注');
  if(y+row.height>max||(i===footerStart&&footerHeight<=max-M-headerHeight&&y+footerHeight>max)){
   pages.push(page);breaks.push(i);pushHeader();
  }
  page.push({row,index:i,y});y+=row.height;
 }
 pages.push(page);return {group,settings:f,rows,pages,breaks,totalRow};
}
function draw(model,pageIndex,canvas){
 const scale=2.5;canvas.width=Math.ceil(W*scale);canvas.height=Math.ceil(H*scale);const ctx=canvas.getContext('2d');ctx.scale(scale,scale);ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);ctx.fillStyle='#111';ctx.strokeStyle='#333';ctx.lineWidth=.5;
 for(const {row,y}of model.pages[pageIndex])for(const obj of row.objects){
  const x=M+widths.slice(0,obj.c).reduce((a,b)=>a+b,0),width=widths.slice(obj.c,obj.c+obj.span).reduce((a,b)=>a+b,0);
  if(obj.border)ctx.strokeRect(x,y,width,row.height);
  ctx.font=(obj.bold?'bold ':'')+obj.font+'px '+FONT;ctx.textBaseline='middle';ctx.textAlign=obj.align;
  const height=obj.lines.length*obj.font*1.25,start=y+(row.height-height)/2+obj.font*1.25/2;
  for(let i=0;i<obj.lines.length;i++)ctx.fillText(obj.lines[i],obj.align==='center'?x+width/2:obj.align==='right'?x+width-4:x+4,start+i*obj.font*1.25);
 }
 ctx.font='8px '+FONT;ctx.textAlign='right';ctx.fillStyle='#555';ctx.fillText(`第 ${pageIndex+1} / ${model.pages.length} 页`,W-M,H-M+5);
}
async function pdf(models,progress=()=>{}){
 const doc=await PDFLib.PDFDocument.create(),canvas=document.createElement('canvas');let done=0,total=models.reduce((s,m)=>s+m.pages.length,0);
 for(const model of models)for(let p=0;p<model.pages.length;p++){
  draw(model,p,canvas);const image=await doc.embedPng(canvas.toDataURL('image/png')),page=doc.addPage([W,H]);page.drawImage(image,{x:0,y:0,width:W,height:H});progress(++done,total);await new Promise(resolve=>setTimeout(resolve,0));
 }
 return doc.save();
}
async function xlsx(templateBytes,models){
 const zip=await JSZip.loadAsync(templateBytes),S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const parse=s=>new DOMParser().parseFromString(s,'application/xml'),serialize=d=>new XMLSerializer().serializeToString(d),all=(d,n)=>[...d.getElementsByTagNameNS('*',n)],first=(d,n)=>all(d,n)[0];
 const base=parse(await zip.file('xl/worksheets/sheet1.xml').async('string')),wb=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string')),types=parse(await zip.file('[Content_Types].xml').async('string')),styles=parse(await zip.file('xl/styles.xml').async('string'));
 const create=(doc,name,attrs={})=>{const e=doc.createElementNS(S,name);for(const [k,v]of Object.entries(attrs))e.setAttribute(k,v);return e;};
 const xfs=first(styles,'cellXfs'),fonts=first(styles,'fonts'),numFmts=first(styles,'numFmts')||styles.documentElement.insertBefore(create(styles,'numFmts'),styles.documentElement.firstChild),styleCache=new Map();
 const baseCells=new Map(all(base,'c').map(c=>[c.getAttribute('r'),c])),formatIds=new Map(all(numFmts,'numFmt').map(f=>[f.getAttribute('formatCode'),Number(f.getAttribute('numFmtId'))]));
 let nextFormat=Math.max(164,...[...formatIds.values()].map(n=>n+1));
 const formatId=format=>{if(!formatIds.has(format)){numFmts.append(create(styles,'numFmt',{numFmtId:nextFormat,formatCode:format}));formatIds.set(format,nextFormat++);}return formatIds.get(format);};
 function styleFor(source,obj,cell){
  const original=Number(baseCells.get(source)?.getAttribute('s')||0),format=cell.format||'@',key=JSON.stringify([original,obj.font,obj.bold,obj.align,obj.single,format]);if(styleCache.has(key))return styleCache.get(key);
  const xf=xfs.children[original].cloneNode(true),font=fonts.children[Number(xf.getAttribute('fontId')||0)].cloneNode(true);
  let sz=first(font,'sz');if(!sz){sz=create(styles,'sz');font.append(sz);}sz.setAttribute('val',obj.font);
  for(const b of [...all(font,'b'),...all(font,'u')])b.remove();if(obj.bold)font.append(create(styles,'b'));xf.setAttribute('fontId',fonts.children.length);fonts.append(font);xf.setAttribute('applyFont','1');
  let alignment=first(xf,'alignment');if(!alignment){alignment=create(styles,'alignment');xf.append(alignment);}alignment.setAttribute('horizontal',obj.align);alignment.setAttribute('vertical','center');alignment.setAttribute('wrapText',obj.single?'0':'1');alignment.setAttribute('shrinkToFit','0');xf.setAttribute('applyAlignment','1');
  xf.setAttribute('numFmtId',format==='@'?49:formatId(format));xf.setAttribute('applyNumberFormat','1');
  const index=xfs.children.length;xfs.append(xf);styleCache.set(key,index);return index;
 }
 first(wb,'sheets').replaceChildren();let names=first(wb,'definedNames');if(!names){names=create(wb,'definedNames');wb.documentElement.insertBefore(names,first(wb,'calcPr'));}names.replaceChildren();
 for(const rel of all(rels,'Relationship'))if(rel.getAttribute('Type')===R+'/worksheet')rel.remove();
 for(const type of all(types,'Override'))if(type.getAttribute('ContentType').includes('worksheet+xml'))type.remove();
 for(const [index,model]of models.entries()){
  const doc=base.cloneNode(true),data=first(doc,'sheetData');data.replaceChildren();
  const merges=first(doc,'mergeCells');merges.replaceChildren();
  for(const [r,row]of model.rows.entries()){
   const tr=create(doc,'row',{r:r+1,ht:row.height,customHeight:1});
   for(let c=0;c<6;c++){
    const cell=row.cells[c],obj=row.objects.find(o=>c>=o.c&&c<o.c+o.span),address=XLSX.utils.encode_cell({r,c}),source=XLSX.utils.encode_col(c)+row.sourceRow;
    const el=create(doc,'c',{r:address,s:styleFor(source,obj,cell)});
    if(cell.f){const f=create(doc,'f');f.textContent=cell.f;el.append(f);}
    if(cell.t==='n'){const v=create(doc,'v');v.textContent=cell.v;el.append(v);}else if(cell.f){el.setAttribute('t','str');const v=create(doc,'v');v.textContent=cell.v;el.append(v);}else{el.setAttribute('t','inlineStr');const is=create(doc,'is'),t=create(doc,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=cell.v;is.append(t);el.append(is);}
    tr.append(el);
   }
   data.append(tr);for(const [a,b]of row.merges)merges.append(create(doc,'mergeCell',{ref:`${XLSX.utils.encode_col(a)}${r+1}:${XLSX.utils.encode_col(b)}${r+1}`}));
  }
  merges.setAttribute('count',merges.children.length);first(doc,'dimension').setAttribute('ref','A1:F'+model.rows.length);
  const cols=first(doc,'cols');cols.replaceChildren();widths.forEach((w,c)=>cols.append(create(doc,'col',{min:c+1,max:c+1,width:(w*96/72-5)/7,customWidth:1})));
  const setup=first(doc,'pageSetup');setup.removeAttribute('scale');setup.setAttribute('fitToWidth','1');setup.setAttribute('fitToHeight','0');setup.setAttribute('orientation','portrait');setup.setAttribute('paperSize','9');
  const sheetPr=first(doc,'sheetPr');let pagePr=first(sheetPr,'pageSetUpPr');if(!pagePr){pagePr=create(doc,'pageSetUpPr');sheetPr.append(pagePr);}pagePr.setAttribute('fitToPage','1');
  const margins=first(doc,'pageMargins');for(const edge of ['top','bottom','left','right'])margins.setAttribute(edge,M/72);
  for(const old of all(doc,'rowBreaks'))old.remove();if(model.breaks.length){const breaks=create(doc,'rowBreaks',{count:model.breaks.length,manualBreakCount:model.breaks.length});model.breaks.forEach(id=>breaks.append(create(doc,'brk',{id,max:16383,man:1})));doc.documentElement.append(breaks);}
  const name=(index+1)+'-'+model.group.supplier.replace(/[\[\]:*?\/\\]/g,'').slice(0,18)+'-'+model.group.month;
  const entry=create(wb,'sheet',{name,sheetId:index+1});entry.setAttributeNS(R,'r:id','internal'+index);first(wb,'sheets').append(entry);
  const rel=rels.createElementNS(rels.documentElement.namespaceURI,'Relationship');for(const [k,v]of Object.entries({Id:'internal'+index,Type:R+'/worksheet',Target:'worksheets/sheet'+(index+1)+'.xml'}))rel.setAttribute(k,v);rels.documentElement.append(rel);
  const type=types.createElementNS(types.documentElement.namespaceURI,'Override');type.setAttribute('PartName','/xl/worksheets/sheet'+(index+1)+'.xml');type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');types.documentElement.append(type);
  for(const [key,value]of [['_xlnm.Print_Area',`'${name.replace(/'/g,"''")}'!$A$1:$F$${model.rows.length}`],['_xlnm.Print_Titles',`'${name.replace(/'/g,"''")}'!$1:$5`]]){const n=create(wb,'definedName',{name:key,localSheetId:index});n.textContent=value;names.append(n);}
  zip.file('xl/worksheets/sheet'+(index+1)+'.xml',serialize(doc));
 }
 for(const list of [xfs,fonts,numFmts])list.setAttribute('count',list.children.length);
 first(wb,'calcPr').setAttribute('fullCalcOnLoad','1');
 zip.file('xl/styles.xml',serialize(styles));zip.file('xl/workbook.xml',serialize(wb));zip.file('xl/_rels/workbook.xml.rels',serialize(rels));zip.file('[Content_Types].xml',serialize(types));
 // Office rebuilds these optional display properties; avoid retaining the template's one-sheet list.
 zip.remove('docProps/app.xml');for(const path of ['_rels/.rels','[Content_Types].xml']){const d=parse(await zip.file(path).async('string'));for(const node of [...d.documentElement.children])if((node.getAttribute('Target')||node.getAttribute('PartName')||'').includes('docProps/app.xml'))node.remove();zip.file(path,serialize(d));}
 return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
return {layout,draw,pdf,xlsx};
})();
