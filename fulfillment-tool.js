(()=>{'use strict';const C=FulfillmentCore,V=FulfillmentValidation,$=id=>document.getElementById(id);let result=null,filename='',sourceBytes=null,validation=null,busy=false;
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const divisionButtons=[...document.querySelectorAll('[data-download-division]')];
const divisionRows=division=>(result?.rows||[]).filter(row=>String(row[C.COLUMNS.indexOf('事业部')]?.v??'').trim()===division);
function updateDownloads(){
 const blocked=busy||$('file').disabled||!validation?.passed;
 $('download').disabled=blocked||!result||!sourceBytes;
 for(const button of divisionButtons){const hasRows=divisionRows(button.dataset.downloadDivision).length>0;button.disabled=blocked||!hasRows;button.title=hasRows?'仅下载该事业部的汇总表':'暂无该事业部数据';}
}
async function readWorkbook(bytes){
 const book=XLSX.read(bytes,{type:'array',cellDates:false,cellNF:true,cellStyles:true});
 const signature=new Uint8Array(bytes);
 if(signature[0]!==80||signature[1]!==75)return book;
 const zip=await JSZip.loadAsync(bytes),parse=text=>new DOMParser().parseFromString(text,'application/xml'),all=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)];
 const workbook=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 const targets=new Map(all(rels,'Relationship').map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));
 for(const node of all(workbook,'sheet')){
  const sheet=book.Sheets[node.getAttribute('name')],id=node.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),target=targets.get(id);
  if(!sheet||!target)continue;
  const path=new URL(target,'https://workbook.local/xl/workbook.xml').pathname.slice(1),part=zip.file(path);if(!part)continue;
  const doc=parse(await part.async('string'));
  sheet['!fulfillmentFilterActive']=all(doc,'filterColumn').length>0||all(doc,'sheetPr').some(n=>['1','true'].includes(n.getAttribute('filterMode')));
 }
 return book;
}
function resetCurrent(){result=null;sourceBytes=null;validation=null;$('result').hidden=true;$('quantityValidation').hidden=true;$('warnings').textContent='';updateDownloads();}
function render(){
 const number=value=>value==null?'待核验':Number(value).toLocaleString('zh-CN',{maximumFractionDigits:12});
 const groupedIssues=new Map();
 for(const issue of validation?.issues||[]){const key=[issue.sheet,issue.cell,issue.field].join('\0');if(!groupedIssues.has(key))groupedIssues.set(key,{...issue,messages:[]});const group=groupedIssues.get(key);if(!group.messages.includes(issue.message))group.messages.push(issue.message);}
 $('quantityValidation').hidden=false;
 $('quantityStatus').textContent=validation?(validation.passed?'六个数量字段校验通过，可以下载。':`数量校验未通过：${groupedIssues.size} 处，请修正源表后重新上传。所有下载已暂停。`):'缺少原文件，无法校验数量，请重新上传。所有下载已暂停。';
 $('quantityChecks').innerHTML='<table><thead><tr>'+['来源工作表','数量字段','原表核算合计','汇总合计','差额（汇总－原表）','结果'].map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+(validation?.checks||[]).map(c=>'<tr>'+[c.sheet,c.field,number(c.sourceTotal),number(c.summaryTotal),number(c.difference),c.status].map(v=>'<td>'+esc(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
 $('quantityIssues').innerHTML=[...groupedIssues.values()].map(i=>'<li><strong>'+esc(i.sheet+(i.cell?'!'+i.cell:'')+' · '+i.field)+'</strong><br>'+i.messages.map(esc).join('<br>')+'</li>').join('');
 $('warnings').textContent=result.warnings.join('\n');
 $('count').textContent=`共 ${result.rows.length} 条，预览前 ${Math.min(100,result.rows.length)} 条；下载将汇总表放在首位，并保留差异表和全部原始工作表。`;
 const table=(columns,rows)=>'<table><thead><tr>'+columns.map(v=>'<th>'+esc(v)+'</th>').join('')+'</tr></thead><tbody>'+rows.slice(0,100).map(row=>'<tr>'+row.map(c=>'<td>'+esc(c.z?XLSX.SSF.format(c.z,c.v):c.v)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
 $('preview').innerHTML=table(C.SUMMARY_COLUMNS,C.summaryRows(result.rows));$('diff-preview').innerHTML=validation?.passed?table(C.DIFF_COLUMNS,result.differences):'<p>本次数量未通过校验，未更新历史和差异。</p>';
 $('result').hidden=false;updateDownloads();
}
async function show(history){
 if(!history){$('status').textContent='等待首次上传';$('history').textContent='尚无历史数据。首次上传保存为基准，差异表仅有表头。';return;}
 filename=history.current.filename;sourceBytes=history.current.sourceBytes||null;
 validation=sourceBytes?V.validate(await readWorkbook(sourceBytes),history.current.result,C):null;
 result={...history.current.result,differences:validation?.passed?C.difference(history.previous?.result,history.current.result):[]};
 const describe=s=>`${s.filename}（${new Date(s.savedAt).toLocaleString('zh-CN')}）`;
 $('history').textContent=`本次：${describe(history.current)}\n上次：${history.previous?describe(history.previous):'尚无，当前数据作为首次基准'}\n${validation?.passed?`差异：${result.differences.length} 个物料；数量为本次减上次。`:'缓存数据需通过数量校验后才能继续使用。'}`;
 $('status').textContent=result.counts.map(x=>`${x.sheet}：${x.count} 条`).join('　');
 render();
}
$('file').disabled=true;$('status').textContent='正在读取历史数据…';
const ready=FulfillmentCache.get().then(show).catch(e=>{resetCurrent();$('status').textContent='历史读取失败：'+e.message;}).finally(()=>{$('file').disabled=false;updateDownloads();});
$('file').onchange=async()=>{
 await ready;const file=$('file').files[0];if(!file)return;
 resetCurrent();filename=file.name;
 if(file.size>60*1024*1024){$('status').textContent='文件超过60MB，请缩小后重试';return;}
 $('status').textContent='正在读取、对比并保存历史…';$('file').disabled=true;updateDownloads();
 try{
  const bytes=await file.arrayBuffer();const book=await readWorkbook(bytes),merged=C.merge(book,file.name);
  if(!merged.rows.length)throw Error('没有可合并的数据，未更新历史');
  validation=V.validate(book,merged,C);
  if(!validation.passed){sourceBytes=bytes;result={...merged,differences:[]};$('status').textContent='数量校验未通过，请查看下方差异。';$('history').textContent='本次文件未保存为历史基准，上次和本次已保存的历史保持不变。';render();return;}
  const history=await FulfillmentCache.save({filename:file.name,savedAt:new Date().toISOString(),result:merged,sourceBytes:bytes});
  await show(history);
 }catch(e){resetCurrent();$('status').textContent='处理失败：'+e.message+'。保留原有历史，请重新上传。';}
 finally{$('file').disabled=false;$('file').value='';updateDownloads();}
};
async function styledBytes(result){const zip=await JSZip.loadAsync(XLSX.write(C.workbook(result),{type:'array',bookType:'xlsx'}));const parser=new DOMParser(),serializer=new XMLSerializer(),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';const styles=parser.parseFromString(await zip.file('xl/styles.xml').async('string'),'application/xml');const el=(name,attrs)=>{const e=styles.createElementNS(ns,name);for(const [k,v] of Object.entries(attrs||{}))e.setAttribute(k,v);return e;};const borders=styles.getElementsByTagName('borders')[0],borderId=borders.children.length,b=el('border');for(const side of ['left','right','top','bottom']){const e=el(side,{style:'thin'});e.append(el('color',{rgb:'FFD9DEE7'}));b.append(e);}borders.append(b);borders.setAttribute('count',borders.children.length);const xfs=styles.getElementsByTagName('cellXfs')[0],existing=[...xfs.children],offset=existing.length;for(const old of existing){const xf=old.cloneNode(true);xf.setAttribute('borderId',borderId);xf.setAttribute('applyBorder','1');xf.setAttribute('applyAlignment','1');xf.setAttribute('fillId','0');for(const a of [...xf.getElementsByTagName('alignment')])a.remove();xf.append(el('alignment',{horizontal:'center',vertical:'center',wrapText:'0',shrinkToFit:'0'}));xfs.append(xf);}const fills=styles.getElementsByTagName('fills')[0],fillId=fills.children.length,fill=el('fill'),pattern=el('patternFill',{patternType:'solid'});pattern.append(el('fgColor',{rgb:'FF315B9D'}),el('bgColor',{indexed:'64'}));fill.append(pattern);fills.append(fill);fills.setAttribute('count',fills.children.length);const fonts=styles.getElementsByTagName('fonts')[0];for(const f of fonts.children){let size=f.getElementsByTagName('sz')[0];if(!size){size=el('sz');f.append(size);}size.setAttribute('val','11');}const fontId=fonts.children.length,font=fonts.children[0].cloneNode(true);for(const c of [...font.getElementsByTagName('color')])c.remove();font.append(el('b'),el('color',{rgb:'FFFFFFFF'}));fonts.append(font);fonts.setAttribute('count',fonts.children.length);const headerStyle=xfs.children.length,header=xfs.children[offset].cloneNode(true);header.setAttribute('fillId',fillId);header.setAttribute('fontId',fontId);header.setAttribute('applyFill','1');header.setAttribute('applyFont','1');xfs.append(header);xfs.setAttribute('count',xfs.children.length);zip.file('xl/styles.xml',serializer.serializeToString(styles));for(const path of Object.keys(zip.files).filter(p=>/^xl\/worksheets\/sheet\d+\.xml$/.test(p))){const sheet=parser.parseFromString(await zip.file(path).async('string'),'application/xml');for(const c of sheet.getElementsByTagName('c'))c.setAttribute('s',/^[A-Z]+1$/.test(c.getAttribute('r'))?headerStyle:offset+Number(c.getAttribute('s')||0));for(const r of sheet.getElementsByTagName('row')){r.setAttribute('ht',r.getAttribute('r')==='1'?'22':'18');r.setAttribute('customHeight','1');}zip.file(path,serializer.serializeToString(sheet));}return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});}
async function download(division=''){
 if(!result||!validation?.passed||busy||$('file').disabled||(!division&&!sourceBytes))return;
 const rows=division?divisionRows(division):null;if(division&&!rows.length)return;
 busy=true;$('file').disabled=true;updateDownloads();
 try{
  const blob=division?await styledBytes({rows}):await FulfillmentExport(await styledBytes(result),sourceBytes);
  const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;
  a.download=division?C.exportFilename(filename).replace(/\.xlsx$/, '-'+division+'.xlsx'):C.exportFilename(filename);
  a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 }catch(e){$('status').textContent='导出失败：'+e.message;}
 finally{busy=false;$('file').disabled=false;updateDownloads();}
}
$('download').onclick=()=>download();
for(const button of divisionButtons)button.onclick=()=>download(button.dataset.downloadDivision);
window.FulfillmentToolReady=true;
})();
