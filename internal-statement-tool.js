(()=>{'use strict';
const $=id=>document.getElementById(id),C=InternalStatement,D=InternalStatementDocument,fields=['reconcileDate','cutoff','terms','phone','balance','receipts','afterSales','returns','invoiced','uninvoiced','note'];
let template,source=null,settings=new Map(),models=[],selected=0,page=0,busy=false;
for(let i=0;i<5;i++){const tr=$('invoiceTable').tBodies[0].insertRow();for(const kind of ['month','status','amount']){const td=tr.insertCell(),input=document.createElement(kind==='status'?'select':'input');input.dataset.field=kind;input.setAttribute('aria-label',`货款汇总第 ${i+1} 行${{month:'月份',status:'是否开票',amount:'开具金额'}[kind]}`);if(kind==='status')for(const v of ['','是','否']){const o=document.createElement('option');o.textContent=v;o.value=v;input.append(o);}else if(kind==='month')input.type='month';else input.inputMode='decimal';td.append(input);}}
function current(){return source?.groups[selected];}
function syncPreviewSelectors(){
 const group=current(),byBuyer=source.groups.filter(g=>g.buyer===group.buyer),bySupplier=byBuyer.filter(g=>g.supplier===group.supplier);
 for(const [id,values,value]of [['previewBuyer',source.groups.map(g=>g.buyer),group.buyer],['previewSupplier',byBuyer.map(g=>g.supplier),group.supplier],['previewMonth',bySupplier.map(g=>g.month),group.month]]){
  $(id).replaceChildren();for(const v of new Set(values)){const option=document.createElement('option');option.value=v;option.textContent=id==='previewMonth'?v.slice(0,4)+'年'+Number(v.slice(5))+'月':v;$(id).append(option);}$(id).value=value;
 }
 $('previewContext').textContent=`${group.buyer} / ${group.supplier} / ${group.month.slice(0,4)}年${Number(group.month.slice(5))}月`;
}
function selectStatement(index){
 if(!source?.groups[index]||busy)return;readForm();selected=index;$('statement').value=String(index);syncPreviewSelectors();writeForm();page=0;regenerate();
}
function readForm(){const group=current();if(!group)return;const data=Object.fromEntries(fields.map(k=>[k,$(k).value]));data.invoices=[...$('invoiceTable').tBodies[0].rows].map(row=>Object.fromEntries([...row.querySelectorAll('input,select')].map(el=>[el.dataset.field,el.value])));settings.set(group.key,data);}
function writeForm(){const data=settings.get(current().key);for(const key of fields)$(key).value=data[key];[...$('invoiceTable').tBodies[0].rows].forEach((row,i)=>row.querySelectorAll('input,select').forEach(el=>{el.value=data.invoices[i][el.dataset.field];}));}
function clearPreview(message=''){
 const canvas=$('preview');canvas.getContext('2d').clearRect(0,0,canvas.width,canvas.height);canvas.hidden=true;$('previewMessage').textContent=message;$('pageInfo').textContent='';$('previous').disabled=true;$('next').disabled=true;
}
function preview(){const model=models[selected];if(!model){clearPreview($('previewMessage').textContent);return;}page=Math.min(page,model.pages.length-1);D.draw(model,page,$('preview'));$('preview').hidden=false;$('previewMessage').textContent='';$('pageInfo').textContent=`第 ${page+1} / ${model.pages.length} 页`;$('previous').disabled=busy||page===0;$('next').disabled=busy||page===model.pages.length-1;}
function regenerate(){
 clearPreview();
 try{readForm();models=source.groups.map(g=>{try{return D.layout(g,settings.get(g.key));}catch(e){throw Error(`${g.buyer} / ${g.supplier} / ${g.month}：${e.message}`);}});preview();$('status').textContent=`已读取 ${source.count} 条明细，合并为 ${source.groups.reduce((n,g)=>n+g.items.length,0)} 行；共 ${new Set(source.groups.map(g=>g.buyer)).size} 个采购组织、${source.groups.length} 份对账单。含税金额合计 ￥${C.amountText(source.amount)}。`;$('buyers').querySelectorAll('button').forEach(b=>b.disabled=false);return true;}
 catch(e){models=[];clearPreview('预览失败，请核实：'+e.message);$('status').textContent='请核实：'+e.message;$('buyers').querySelectorAll('button').forEach(b=>b.disabled=true);return false;}
}
function save(bytes,name,type){const url=URL.createObjectURL(new Blob([bytes],{type})),a=document.createElement('a');a.href=url;a.download=name.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_');a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
async function download(buyer,format){
 if(busy||!regenerate())return;busy=true;document.querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=true);
 try{
  const selected=models.filter(m=>m.group.buyer===buyer),name=C.fileName(buyer,selected.map(m=>m.group));$('status').textContent='正在生成 '+buyer+' 的 '+format.toUpperCase()+'…';
  const bytes=format==='pdf'?await D.pdf(selected,(done,total)=>{$('status').textContent=`正在生成 PDF：${done} / ${total} 页`;}):await D.xlsx(template,selected);
  if(format==='xlsx'){const saved=XLSX.read(bytes,{type:'array',cellNF:true});if(saved.SheetNames.length!==selected.length)throw Error('导出工作表数量不一致');for(const [i,m]of selected.entries()){const s=saved.Sheets[saved.SheetNames[i]];if(s['E'+m.totalRow]?.v!==Number(C.amountText(m.group.amount))||s.A1?.v!==buyer)throw Error('导出金额或采购组织校验失败');}}
  save(bytes,name+'.'+format,format==='pdf'?'application/pdf':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');$('status').textContent='已生成：'+name+'.'+format;
 }catch(e){models=[];clearPreview('生成失败：'+e.message);$('status').textContent='生成失败：'+e.message;}finally{busy=false;document.querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=false);preview();}
}
function buyers(){
 $('buyers').replaceChildren();for(const buyer of new Set(source.groups.map(g=>g.buyer))){const groups=source.groups.filter(g=>g.buyer===buyer),card=document.createElement('div');card.className='buyer';const title=document.createElement('strong');title.textContent=buyer;card.append(title);const info=document.createElement('small');info.textContent=groups.length+' 份对账单，含税金额 ￥'+C.amountText(groups.reduce((n,g)=>n+BigInt(g.amount),0n));card.append(info);
 for(const format of ['xlsx','pdf']){const button=document.createElement('button');button.textContent=format==='xlsx'?'下载 Excel':'下载 PDF';button.dataset.format=format;button.dataset.buyer=buyer;button.onclick=()=>download(buyer,format);card.append(button);}$('buyers').append(card);}
}
$('sourceFile').onchange=async()=>{
 const file=$('sourceFile').files[0];if(!file)return;$('sourceFile').disabled=true;source=null;models=[];selected=0;page=0;clearPreview();settings.clear();$('buyers').replaceChildren();for(const id of ['formSection','downloadSection','previewSection'])$(id).hidden=true;$('warnings').textContent='';$('fileInfo').textContent='正在读取：'+file.name;
 try{if(file.size>30*1024*1024)throw Error('文件超过 30 MB');$('status').textContent='正在读取并整理…';await new Promise(r=>setTimeout(r,0));const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellNF:true});source=C.parse(book,file.name);
  for(const g of source.groups)settings.set(g.key,C.defaults(g.month));$('statement').replaceChildren();source.groups.forEach((g,index)=>{const o=document.createElement('option');o.value=index;o.textContent=`${g.buyer} / ${g.supplier} / ${g.month}`;$('statement').append(o);});
  syncPreviewSelectors();writeForm();buyers();for(const id of ['formSection','downloadSection','previewSection'])$(id).hidden=false;
  $('fileInfo').textContent=file.name+'｜引用时间：'+new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false});$('warnings').textContent=source.warnings.join('\n');regenerate();
 }catch(e){source=null;models=[];$('status').textContent='读取失败：'+e.message;$('fileInfo').textContent=file.name+'（未成功引用）';}finally{$('sourceFile').disabled=false;$('sourceFile').value='';}
};
$('statement').onchange=()=>selectStatement(Number($('statement').value));
$('previewBuyer').onchange=()=>selectStatement(source.groups.findIndex(g=>g.buyer===$('previewBuyer').value));
$('previewSupplier').onchange=()=>selectStatement(source.groups.findIndex(g=>g.buyer===current().buyer&&g.supplier===$('previewSupplier').value));
$('previewMonth').onchange=()=>selectStatement(source.groups.findIndex(g=>g.buyer===current().buyer&&g.supplier===current().supplier&&g.month===$('previewMonth').value));
for(const key of fields)$(key).oninput=()=>{readForm();$('buyers').querySelectorAll('button').forEach(b=>b.disabled=true);};
$('invoiceTable').oninput=()=>{readForm();$('buyers').querySelectorAll('button').forEach(b=>b.disabled=true);};
$('update').onclick=()=>{page=0;regenerate();};$('previous').onclick=()=>{page--;preview();};$('next').onclick=()=>{page++;preview();};
(async()=>{try{const response=await fetch('assets/internal-statement-template.xlsx?v=20261103-1');if(!response.ok)throw Error('HTTP '+response.status);template=await response.arrayBuffer();$('status').textContent='请上传采购入库明细。';$('sourceFile').disabled=false;}catch(e){$('status').textContent='内置模板加载失败，请刷新重试：'+e.message;}})();
})();
