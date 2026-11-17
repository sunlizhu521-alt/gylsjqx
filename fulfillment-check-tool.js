(() => {
'use strict';
const $=id=>document.getElementById(id),C=window.FulfillmentCheckCore;
const sources={purchase:null,fulfillment:null},versionKey='fulfillment-check-export-versions',memoryVersions={};
let result=null,busy=false,decisions={},page=0;
function controls(){for(const kind of Object.keys(sources)){$(kind+'File').disabled=busy;$(kind+'Clear').disabled=busy;}for(const select of $('choices').querySelectorAll('select'))select.disabled=busy;$('download').disabled=busy||!result?.canExport;}
function clearResult(){result=null;decisions={};page=0;$('preview').replaceChildren();$('choices').replaceChildren();$('result').hidden=$('choicePanel').hidden=true;controls();}
function table(){
 const detail=$('view').value==='detail',rows=detail?result.detail:result.summary,headers=detail?C.DETAIL:C.SUMMARY;
 const pages=Math.max(1,Math.ceil(rows.length/100));page=Math.max(0,Math.min(page,pages-1));
 const table=document.createElement('table'),head=table.createTHead().insertRow(),body=table.createTBody();
 for(const label of headers){const th=document.createElement('th');th.scope='col';th.textContent=label;head.append(th);}
 for(const values of rows.slice(page*100,(page+1)*100)){const row=body.insertRow();values.forEach((v,c)=>{const td=row.insertCell();td.textContent=v??'';if(detail&&[8,13,14,16].includes(c))td.className='note';});}
 $('preview').replaceChildren(table);$('pageInfo').textContent=`第 ${page+1} / ${pages} 页，共 ${rows.length} 条`;$('previous').disabled=page===0;$('next').disabled=page===pages-1;
}
function choices(){
 $('choices').replaceChildren();$('choicePanel').hidden=!result.choices.length;$('choiceTitle').textContent=`供应商对应待确认（${result.choices.length} 项，可保留待核对下载）`;
 for(const item of result.choices){
  const card=document.createElement('div'),label=document.createElement('label'),select=document.createElement('select'),p=document.createElement('p');label.textContent=item.label+'：'+item.reason;
  select.append(new Option('保留待核对',''));for(const candidate of item.candidates)select.append(new Option(candidate.label,candidate.value));
  select.setAttribute('aria-label',item.label+'对应采购供应商');label.append(select);card.append(label);
  p.textContent='履约来源：'+item.sources.slice(0,8).join('；')+(item.sources.length>8?` 等 ${item.sources.length} 行`:'')+'\n'+item.candidates.map(c=>c.label+'：'+(c.sources.join('；')||'名称唯一包含候选')).join('\n');card.append(p);
  select.addEventListener('change',()=>{if(select.value)decisions[item.supplier]=select.value;else delete decisions[item.supplier];calculate();});$('choices').append(card);
 }
 if(Object.keys(decisions).length){const reset=document.createElement('button'),confirmed=document.createElement('p');confirmed.textContent='本次已确认：'+Object.entries(decisions).map(([from,to])=>from+' → '+to).join('；');reset.className='secondary';reset.textContent='清除本次供应商确认';reset.onclick=()=>{decisions={};calculate();};$('choices').prepend(confirmed,reset);$('choicePanel').hidden=false;}
}
function calculate(){
 if(!sources.purchase||!sources.fulfillment){$('status').textContent='请上传两份文件。';return;}
 try{
  result=C.build({purchase:sources.purchase,fulfillment:sources.fulfillment,decisions});page=0;
  const s=result.stats;$('counts').textContent=`采购 ${s.purchase} 条（未关闭 ${s.open}，已关闭 ${s.closed}）；履约 ${s.fulfillment} 条；一致 ${s.matched} 项，差异或待核对 ${result.detail.length} 项。`;
  $('result').hidden=false;choices();table();$('status').textContent='核对完成，可下载。待核对数量及单边记录未按零处理。';
 }catch(e){clearResult();$('status').textContent='核对失败：'+e.message;}controls();
}
for(const kind of Object.keys(sources)){
 $(kind+'File').addEventListener('change',async()=>{
  clearResult();sources[kind]=null;const file=$(kind+'File').files[0];$(kind+'Clear').hidden=!file;
  if(!file){$(kind+'Reference').textContent='尚未引用文件';calculate();return;}
  busy=true;controls();$(kind+'Reference').textContent=file.name+'\n引用时间：'+new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false});$('status').textContent='正在读取并核对…';
  try{
   if(!/\.xlsx$/i.test(file.name))throw Error('请上传 .xlsx 原文件，以保留原工作表');if(file.size>60*1024*1024)throw Error('文件超过 60 MB，请缩小后上传');
   const bytes=await file.arrayBuffer(),zip=await JSZip.loadAsync(bytes);if(!zip.file('xl/workbook.xml'))throw Error('不是有效的 Excel 工作簿');
   const book=XLSX.read(bytes,{type:'array',cellNF:true,cellFormula:true,sheetStubs:true});sources[kind]={bytes,book,filename:file.name};calculate();
  }catch(e){clearResult();$('status').textContent='读取失败：'+e.message;}finally{busy=false;controls();}
 });
 $(kind+'Clear').onclick=()=>{clearResult();sources[kind]=null;$(kind+'File').value='';$(kind+'Reference').textContent='尚未引用文件';$(kind+'Clear').hidden=true;calculate();};
}
$('view').onchange=()=>{page=0;if(result)table();};$('previous').onclick=()=>{page--;table();};$('next').onclick=()=>{page++;table();};
$('download').onclick=async()=>{
 if(busy||!result)return;busy=true;controls();$('status').textContent='正在保留原表并生成核对文件…';let url;
 try{
  const bytes=await window.FulfillmentCheckExport(sources.purchase.bytes,sources.fulfillment.bytes,result);
  const date=new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());let versions={};
  try{const saved=JSON.parse(localStorage.getItem(versionKey)||'{}');if(saved&&typeof saved==='object'&&!Array.isArray(saved))versions=saved;}catch(_){}
  const old=Number(versions[date]),version=Math.max(Number.isSafeInteger(old)&&old>=0?old:0,memoryVersions[date]||0)+1;if(!Number.isSafeInteger(version))throw Error('版本号无效，请核实浏览器保存的版本');
  const name=`订单履约表核对-${date}-V${version}.xlsx`;url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();memoryVersions[date]=version;
  let saved=true;try{localStorage.setItem(versionKey,JSON.stringify({...versions,[date]:version}));}catch(_){saved=false;}
  $('status').textContent='已生成：'+name+(saved?'':'。版本号仅在本页保存。');
 }catch(e){$('status').textContent='导出失败：'+e.message;}finally{if(url)setTimeout(()=>URL.revokeObjectURL(url),1000);busy=false;controls();}
};
if(!C||!window.FulfillmentCheckExport||!window.JSZip){$('status').textContent='工具加载失败，请刷新重试。';return;}controls();$('status').textContent='请上传采购订单和订单履约表。';
})();
