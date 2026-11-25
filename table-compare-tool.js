(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.TableCompareCore;
  const sources = {
    old:{input:'oldFile',reference:'oldReference',select:'oldSheet',label:'旧版付款台账'},
    new:{input:'newFile',reference:'newReference',select:'newSheet',label:'新版付款台账'}
  };
  for (const source of Object.values(sources)) Object.assign(source,{book:null,filename:'',sheetName:'',candidates:[],state:'empty',version:0,error:''});
  let result = null, exporting = false;
  const versionKey = 'table-compare-export-versions', memoryVersions = {};
  const esc = value => String(value ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const money = value => typeof value === 'number' && Number.isFinite(value) ? value.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}) : '待核对';

  function savedVersions() {
    try { const value = JSON.parse(localStorage.getItem(versionKey) || '{}'); return value && typeof value === 'object' && !Array.isArray(value) ? value : {}; }
    catch (_) { return {}; }
  }
  function nextDownload() {
    const date = new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
    const stored = Number(savedVersions()[date]);
    const previous = Math.max(Number.isSafeInteger(stored) && stored >= 0 ? stored : 0,memoryVersions[date] || 0);
    const version = previous + 1;
    if (!Number.isSafeInteger(version)) throw Error('版本号超出范围，请核实浏览器保存的版本号');
    return {date,version,name:C.exportFilename(date,version)};
  }
  function rememberDownload({date,version}) {
    memoryVersions[date] = version;
    try { localStorage.setItem(versionKey,JSON.stringify({...savedVersions(),[date]:version})); return true; }
    catch (_) { return false; }
  }
  async function readBook(bytes) {
    const signature = new Uint8Array(bytes,0,Math.min(2,bytes.byteLength));
    if (signature[0] === 0x50 && signature[1] === 0x4b) {
      const zip = await JSZip.loadAsync(bytes);
      let changed = false;
      for (const path of Object.keys(zip.files).filter(p=>/^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
        const xml = await zip.file(path).async('string');
        // Preserve a missing formula cache rather than treating it as a valid empty string.
        const next = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,(cell,attrs,content)=>{
          if (content === undefined || !/<f(?:\s|>|\/)/.test(content) || /<v(?:\s|>|\/)/.test(content)) return cell;
          if (!/\bt="str"/.test(attrs) && !(/\bt="inlineStr"/.test(attrs) && !/<is(?:\s|>|\/)/.test(content))) return cell;
          changed = true;
          return '<c' + attrs.replace(/\s+t="(?:str|inlineStr)"/,'') + '>' + content + '</c>';
        });
        if (next !== xml) zip.file(path,next);
      }
      if (changed) bytes = await zip.generateAsync({type:'uint8array'});
    }
    return XLSX.read(bytes,{type:'array',cellNF:true,cellFormula:true,cellDates:false,sheetStubs:true});
  }
  function clearResult() {
    result = null;
    $('result').hidden = true;
    $('download').disabled = true;
    $('warning-panel').hidden = true;
    $('warnings').replaceChildren();
    $('warning-count').textContent = '0';
    $('record-count').textContent = '0';
    $('preview').replaceChildren();
    for (const id of ['oldAmount','newAmount','difference','counts']) $(id).textContent = '';
    $('incomplete').hidden = true;
  }
  function sheetOptions(source) {
    const select = $(source.select);
    select.replaceChildren();
    if (!source.candidates.length) select.append(new Option(source.state === 'reading' ? '正在读取工作表…' : source.state === 'error' ? '请重新上传文件' : '请先上传文件',''));
    else {
      if (source.candidates.length > 1) select.append(new Option('请选择需要比对的工作表',''));
      for (const candidate of source.candidates) select.append(new Option(candidate.name,candidate.name));
    }
    select.value = source.sheetName;
  }
  function controls() {
    for (const [role,source] of Object.entries(sources)) {
      $(source.input).disabled = exporting;
      $(source.select).disabled = exporting || !source.candidates.length;
      $('clear-' + role).hidden = source.state === 'empty';
      $('clear-' + role).disabled = exporting;
    }
    $('download').disabled = exporting || !result;
  }
  function render() {
    const sheet = result.sheets[0], summary = result.summary;
    const wrap = header => /摘要|备注|比对说明|来源/.test(header);
    $('sheet-title').textContent = sheet.name;
    $('record-count').textContent = String(sheet.rows.length);
    $('preview').innerHTML = '<table><thead><tr>' + sheet.headers.map(header=>'<th scope="col">'+esc(header)+'</th>').join('') + '</tr></thead><tbody>' + sheet.rows.map(row=>'<tr>'+row.map((cell,c)=>'<td'+(wrap(sheet.headers[c])?' class="wrap"':'')+'>'+esc(C.formatCell(cell))+'</td>').join('')+'</tr>').join('') + '</tbody></table>' + (sheet.rows.length ? '' : '<p class="empty">没有付款明细，下载仍保留表头。</p>');
    $('oldAmount').textContent = money(summary.oldAmount);
    $('newAmount').textContent = money(summary.newAmount);
    $('difference').textContent = money(summary.difference);
    $('incomplete').hidden = summary.complete !== false;
    $('counts').textContent = ['增加','减少','不变','待核对'].map(name=>name+' '+(summary.counts[name] || 0)+' 条').join('　');
    $('warnings').innerHTML = result.warnings.map(warning=>'<li>'+esc(warning)+'</li>').join('');
    $('warning-count').textContent = String(result.warnings.length);
    $('warning-panel').hidden = !result.warnings.length;
    $('result').hidden = false;
    $('status').textContent = `比对完成：${sheet.rows.length} 条。` + (summary.complete === false ? '合计不完整，请核实待核对记录；仍可下载。' : result.warnings.length ? '存在需留意的内容，详见下方提示；仍可下载。' : '可下载填写备注。');
  }
  function refresh() {
    clearResult();
    const all = Object.values(sources), reading = all.filter(source=>source.state === 'reading'), failed = all.filter(source=>source.state === 'error');
    if (reading.length) $('status').textContent = '正在读取：'+reading.map(source=>source.label).join('、')+'…';
    else if (failed.length) $('status').textContent = failed.map(source=>source.label+'读取失败：'+source.error).join('\n');
    else if (all.some(source=>!source.book)) $('status').textContent = '请分别上传旧版和新版付款台账。';
    else if (all.some(source=>!source.sheetName)) $('status').textContent = '请选择'+all.filter(source=>!source.sheetName).map(source=>source.label+'的工作表').join('、')+'，不默认选择第一张。';
    else {
      try {
        const input = source => ({book:source.book,filename:source.filename,sheetName:source.sheetName});
        result = C.build(input(sources.old),input(sources.new));
        render();
      } catch (error) { clearResult(); $('status').textContent = '比对失败：'+error.message; }
    }
    controls();
  }
  if (!C || !window.XLSX || !window.TableCompareExport || !window.JSZip) {
    $('status').textContent = '工具加载失败，请刷新页面后重试。';
    return;
  }
  for (const [role,source] of Object.entries(sources)) {
    $(source.input).addEventListener('change',async()=>{
      const file = $(source.input).files[0];
      if (!file || exporting) return;
      const current = ++source.version;
      Object.assign(source,{book:null,filename:file.name,sheetName:'',candidates:[],state:'reading',error:''});
      $(source.reference).textContent = file.name+'\n引用时间：'+new Date().toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false});
      sheetOptions(source);
      refresh();
      try {
        if (!/\.xlsx?$/i.test(file.name)) throw Error('请选择 .xlsx 或 .xls 文件');
        if (file.size > 60*1024*1024) throw Error('文件超过 60 MB，请缩小后重新上传');
        const bytes = await file.arrayBuffer();
        if (current !== source.version) return;
        const book = await readBook(bytes);
        if (current !== source.version) return;
        const candidates = C.discover(book,file.name);
        if (!candidates.length) throw Error('未找到可比对的付款台账表头，请核实文件内容');
        Object.assign(source,{book,candidates,sheetName:candidates.length === 1 ? candidates[0].name : '',state:'ready'});
      } catch (error) {
        if (current !== source.version) return;
        Object.assign(source,{book:null,sheetName:'',candidates:[],state:'error',error:error.message});
        $(source.reference).textContent += '\n读取失败，请重新选择文件。';
      } finally {
        if (current === source.version) { $(source.input).value = ''; sheetOptions(source); refresh(); }
      }
    });
    $(source.select).addEventListener('change',()=>{
      if (exporting) return;
      const selected = $(source.select).value;
      source.sheetName = source.candidates.some(candidate=>candidate.name === selected) ? selected : '';
      refresh();
    });
    $('clear-' + role).addEventListener('click',()=>{
      if (exporting) return;
      source.version++;
      Object.assign(source,{book:null,filename:'',sheetName:'',candidates:[],state:'empty',error:''});
      $(source.input).value = '';
      $(source.reference).textContent = '尚未引用文件';
      sheetOptions(source);
      refresh();
    });
  }
  $('download').addEventListener('click',async()=>{
    if (!result || exporting) return;
    exporting = true;
    controls();
    let url;
    $('status').textContent = '正在生成 Excel…';
    try {
      const bytes = await TableCompareExport(result), download = nextDownload();
      url = URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link = document.createElement('a');
      link.href = url; link.download = download.name;
      document.body.append(link); link.click(); link.remove();
      const saved = rememberDownload(download);
      $('status').textContent = '已生成：'+download.name+(saved ? '' : '。浏览器未能保存版本号，本次页面内仍会递增。');
    } catch (error) { $('status').textContent = '导出失败：'+error.message+'。可以重试下载。'; }
    finally { if (url) setTimeout(()=>URL.revokeObjectURL(url),1000); exporting = false; controls(); }
  });
  refresh();
  window.TableCompareToolReady = true;
})();
