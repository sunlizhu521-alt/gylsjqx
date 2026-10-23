(()=>{'use strict';
const $=id=>document.getElementById(id),W=1190,H=1684,PAD=55;
let book=null,fileName='',documentCount=0,signature=null,busy=false;
const layouts=new Map(),positions=new Map(),preview=$('preview'),ctx=preview.getContext('2d');
const text=cell=>String(cell?.w??cell?.v??'');
function wrap(c,value,width){
 const lines=[];for(const paragraph of value.replace(/\r/g,'').split('\n')){let line='';for(const char of paragraph){if(line&&c.measureText(line+char).width>width){lines.push(line);line=char;}else line+=char;}lines.push(line);}return lines;
}
function layout(sheet){
 if(!sheet['!ref'])throw Error('工作表没有内容');
 const range=XLSX.utils.decode_range(sheet['!ref']);
 if((range.e.r-range.s.r+1)*(range.e.c-range.s.c+1)>20000)throw Error('工作表范围过大，请先删除无用途的空行空列后重试');
 const merges=sheet['!merges']||[],columns=[],rows=[],cells=[];
 for(let c=range.s.c;c<=range.e.c;c++)columns.push(Math.max(8,sheet['!cols']?.[c]?.wpx||80));
 const total=columns.reduce((a,b)=>a+b,0),width=W-PAD*2,scale=width/total;
 const xs=[0];columns.forEach(v=>xs.push(xs.at(-1)+v*scale));
 for(let r=range.s.r;r<=range.e.r;r++)rows.push(Math.max(30,(sheet['!rows']?.[r]?.hpt||22)*1.7));
 let header=-1,detailEnd=-1;
 for(let r=range.s.r;r<=range.e.r;r++){
  const values=[];for(let c=range.s.c;c<=range.e.c;c++)values.push(text(sheet[XLSX.utils.encode_cell({r,c})]));
  if(values.some(v=>/物料编码/.test(v))&&values.some(v=>/数量/.test(v)))header=r;
  if(header>=0&&r>header&&values.some(v=>/^合计$/.test(v.trim())))detailEnd=r;
 }
 for(let r=range.s.r;r<=range.e.r;r++)for(let c=range.s.c;c<=range.e.c;c++){
  const merged=merges.find(m=>r>=m.s.r&&r<=m.e.r&&c>=m.s.c&&c<=m.e.c);
  if(merged&&(r!==merged.s.r||c!==merged.s.c))continue;
  const endC=Math.min(merged?.e.c??c,range.e.c),endR=Math.min(merged?.e.r??r,range.e.r);
  const original=sheet[XLSX.utils.encode_cell({r,c})];
  if(original?.t==='e'||(original?.f&&original.v==null))throw Error('单元格 '+XLSX.utils.encode_cell({r,c})+' 的公式结果不可用，请在 Excel/WPS 中重算后上传');
  const value=text(original);
  const isTitle=/采购入库单/.test(value),font=isTitle?32:20;
  ctx.font=`${isTitle||r===header?'bold ':''}${font}px "Songti SC","SimSun",serif`;
  const cellWidth=xs[endC-range.s.c+1]-xs[c-range.s.c];
  const lines=wrap(ctx,value,Math.max(6,cellWidth-12)),needed=lines.length*font*1.35+14;
  const current=rows.slice(r-range.s.r,endR-range.s.r+1).reduce((a,b)=>a+b,0);
  if(needed>current){const add=(needed-current)/(endR-r+1);for(let j=r;j<=endR;j++)rows[j-range.s.r]+=add;}
  cells.push({r,c,endR,endC,value,font,lines,title:isTitle,bold:isTitle||r===header,border:header>=0&&r>=header&&(detailEnd<0?r===header:r<=detailEnd)});
 }
 const ys=[0];rows.forEach(v=>ys.push(ys.at(-1)+v));
 // Reserve a clear footer for signatures instead of placing them over the form.
 const factor=Math.min(1,(H-PAD*2-180)/ys.at(-1));
 const objects=cells.map(cell=>({...cell,x:PAD+xs[cell.c-range.s.c]*factor,y:PAD+ys[cell.r-range.s.r]*factor,width:(xs[cell.endC-range.s.c+1]-xs[cell.c-range.s.c])*factor,height:(ys[cell.endR-range.s.r+1]-ys[cell.r-range.s.r])*factor,font:cell.font*factor}));
 return {cells:objects,bottom:PAD+ys.at(-1)*factor,minFont:20*factor*72/(W/210*25.4),factor};
}
function signatureSize(){const width=Number($('signatureWidth').value);if(!Number.isFinite(width)||width<15||width>70)throw Error('签名宽度请填写15至70毫米');return {w:width*W/210,h:width*W/210*signature.height/signature.width};}
function defaultPosition(model,size){return {x:W-PAD-size.w,y:Math.min(H-PAD-size.h,model.bottom+35)};}
function validPosition(model,pos,size){return pos.x>=PAD&&pos.y>=PAD&&pos.x+size.w<=W-PAD&&pos.y+size.h<=H-PAD&&!model.cells.some(c=>c.value.trim()&&pos.x<c.x+c.width&&pos.x+size.w>c.x&&pos.y<c.y+c.height&&pos.y+size.h>c.y);}
function draw(name,canvas=preview){
 const c=canvas.getContext('2d'),model=layouts.get(name);canvas.width=W;canvas.height=H;c.fillStyle='#fff';c.fillRect(0,0,W,H);if(!model)return;
 for(const cell of model.cells){
  if(cell.border){c.strokeStyle='#222';c.lineWidth=.9;c.strokeRect(cell.x,cell.y,cell.width,cell.height);}
  if(!cell.value)continue;c.save();c.font=`${cell.bold?'bold ':''}${cell.font}px "Songti SC","SimSun",serif`;c.fillStyle='#111';c.textBaseline='middle';
  const lines=cell.lines,lineHeight=cell.font*1.35,start=cell.y+(cell.height-lines.length*lineHeight)/2+lineHeight/2;
  c.textAlign=cell.title?'center':'left';lines.forEach((line,i)=>c.fillText(line,cell.title?cell.x+cell.width/2:cell.x+6*model.factor,start+i*lineHeight));c.restore();
 }
 if(signature){const size=signatureSize(),pos=positions.get(name)||defaultPosition(model,size);if(!validPosition(model,pos,size))throw Error('签名位置超出纸张或与单据文字重叠，请调整签名大小或点击下方空白处');c.drawImage(signature,pos.x,pos.y,size.w,size.h);}
}
function refresh(message=''){
 $('download').disabled=true;try{draw($('sheet').value);if(book){const model=layouts.get($('sheet').value);$('status').textContent=message||`${documentCount} 个文件、${book.SheetNames.length} 个工作表，每表一页；当前最小字号约 ${model.minFont.toFixed(1)} pt。${signature?'签名已添加，可点击空白处调整。':'默认签名正在加载。'}`;if(model.minFont<6)$('status').textContent+=' 内容较多，打印字号较小，请检查预览。';} $('download').disabled=busy||!book||!signature;}catch(e){$('status').textContent=e.message;}
}
async function loadImage(file){if(file.size>15*1024*1024)throw Error('签名图片不能超过15MB');const url=URL.createObjectURL(file);try{const image=new Image();image.src=url;await image.decode();if(image.width*image.height>20000000)throw Error('签名图片尺寸过大');return image;}finally{URL.revokeObjectURL(url);}}
function extractSignature(image){
 const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const g=c.getContext('2d');g.drawImage(image,0,0);
 const data=g.getImageData(0,0,c.width,c.height),p=data.data;
 // Ignore the dark screenshot frame: use the bounding area of opaque white paper.
 let left=c.width,top=c.height,right=-1,bottom=-1;
 for(let y=0;y<c.height;y++)for(let x=0;x<c.width;x++){const i=(y*c.width+x)*4;if(p[i+3]>240&&Math.min(p[i],p[i+1],p[i+2])>240){left=Math.min(left,x);right=Math.max(right,x);top=Math.min(top,y);bottom=Math.max(bottom,y);}}
 if(right<left)throw Error('默认签名图片缺少白底');
 // Keep ink only where white paper exists on all four sides, excluding rounded black corners.
 let minX=c.width,minY=c.height,maxX=-1,maxY=-1;
 const rowBounds=[],colBounds=[];
 for(let y=top;y<=bottom;y++){let a=c.width,b=-1;for(let x=left;x<=right;x++){const i=(y*c.width+x)*4;if(p[i+3]>240&&Math.min(p[i],p[i+1],p[i+2])>240){a=Math.min(a,x);b=x;}}rowBounds[y]=[a,b];}
 for(let x=left;x<=right;x++){let a=c.height,b=-1;for(let y=top;y<=bottom;y++){const i=(y*c.width+x)*4;if(p[i+3]>240&&Math.min(p[i],p[i+1],p[i+2])>240){a=Math.min(a,y);b=y;}}colBounds[x]=[a,b];}
 for(let y=0;y<c.height;y++)for(let x=0;x<c.width;x++){const i=(y*c.width+x)*4;const inside=x>left&&x<right&&y>top&&y<bottom&&x>rowBounds[y]?.[0]&&x<rowBounds[y]?.[1]&&y>colBounds[x]?.[0]&&y<colBounds[x]?.[1];const alpha=inside?255-Math.min(p[i],p[i+1],p[i+2]):0;p[i]=p[i+1]=p[i+2]=0;p[i+3]=alpha>25?alpha:0;if(p[i+3]){minX=Math.min(minX,x);maxX=Math.max(maxX,x);minY=Math.min(minY,y);maxY=Math.max(maxY,y);}}
 if(maxX<minX)throw Error('默认签名图片未能提取到笔迹');g.putImageData(data,0,0);const out=document.createElement('canvas');out.width=maxX-minX+9;out.height=maxY-minY+9;out.getContext('2d').drawImage(c,minX,minY,maxX-minX+1,maxY-minY+1,4,4,maxX-minX+1,maxY-minY+1);return out;
}
// Always use the approved built-in signature, including in a fresh browser.
const signatureReady=(async()=>{
 try{
  const response=await fetch('assets/receipt-default-signature.png');
  if(!response.ok)throw Error('HTTP '+response.status);
  signature=extractSignature(await loadImage(await response.blob()));
  refresh();return true;
 }catch(e){$('status').textContent='默认签名加载失败，请刷新页面重试：'+e.message;return false;}
})();
$('file').onchange=async()=>{
 const files=[...$('file').files];if(!files.length)return;
 book=null;layouts.clear();positions.clear();$('download').disabled=true;$('file').disabled=true;
 try{
  $('status').textContent='正在读取单据并自动签字…';
  if(!await signatureReady)throw Error('默认签名加载失败，请刷新页面重试');
  if(files.reduce((sum,f)=>sum+f.size,0)>100*1024*1024)throw Error('本批文件合计不能超过100MB');
  const next={SheetNames:[],Sheets:{}};
  for(const [index,file] of files.entries()){
   if(file.size>30*1024*1024)throw Error(file.name+' 超过30MB');
   const workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellStyles:true,cellNF:true});
   if(!workbook.SheetNames.length)throw Error(file.name+' 没有工作表');
   for(const sheetName of workbook.SheetNames){
    const name=files.length===1?sheetName:`${index+1}. ${file.name} / ${sheetName}`;
    layouts.set(name,layout(workbook.Sheets[sheetName]));next.SheetNames.push(name);next.Sheets[name]=workbook.Sheets[sheetName];
   }
  }
  book=next;documentCount=files.length;fileName=files.length===1?files[0].name:'采购入库单-批量';
  $('sheet').replaceChildren(...book.SheetNames.map(name=>new Option(name,name)));$('sheet').disabled=false;refresh();
 }catch(e){layouts.clear();$('sheet').replaceChildren();$('sheet').disabled=true;$('status').textContent='读取失败：'+e.message;ctx.clearRect(0,0,W,H);}
 finally{$('file').disabled=false;}
};
$('sheet').onchange=()=>refresh();$('signatureWidth').oninput=()=>{positions.clear();refresh();};
preview.onclick=event=>{if(!book||!signature||busy)return;try{const rect=preview.getBoundingClientRect(),size=signatureSize(),pos={x:(event.clientX-rect.left)*W/rect.width-size.w/2,y:(event.clientY-rect.top)*H/rect.height-size.h/2};if(!validPosition(layouts.get($('sheet').value),pos,size))throw Error('此处会遮挡文字或超出纸张，请点击空白处');positions.set($('sheet').value,pos);refresh();}catch(e){$('status').textContent=e.message;}};
$('download').onclick=async()=>{if(!book||!signature||busy)return;busy=true;for(const el of document.querySelectorAll('input,select,button'))el.disabled=true;try{const pdf=await PDFLib.PDFDocument.create();for(const name of book.SheetNames){const canvas=document.createElement('canvas');draw(name,canvas);const png=await pdf.embedPng(canvas.toDataURL('image/png'));const page=pdf.addPage([595.28,841.89]);page.drawImage(png,{x:0,y:0,width:595.28,height:841.89});}const blob=new Blob([await pdf.save()],{type:'application/pdf'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=fileName.replace(/\.[^.]+$/,'')+'-签字.pdf';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);$('status').textContent=`已生成 ${book.SheetNames.length} 页签字 PDF。`;}catch(e){$('status').textContent='生成失败：'+e.message;}finally{busy=false;for(const el of document.querySelectorAll('input,select,button'))el.disabled=false;$('download').disabled=!book||!signature;}};
refresh();
})();
