(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.QuarterBudgetCore, I = window.QuarterBudgetImport, Cache = window.QuarterBudgetCache;
  const roles = ['negotiated','ledger','budget','evidence','existing'];
  const labels = {negotiated:'谈判付款',ledger:'资金台账整理',budget:'预估预算',evidence:'付款状态依据',existing:'已有付款汇总明细'};
  const sources = Object.fromEntries(roles.map(role => [role,{book:null,filename:'',fingerprint:'',importedAt:'',state:'empty',error:'',version:0,sheets:[]} ]));
  let result = null, decisions = {}, matches = {}, exporting = false, restoring = true;
  const versionKey = 'quarter-budget-export-versions', memoryVersions = {};
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const message = value => typeof value === 'string' ? value : value?.message || String(value ?? '');
  const month = key => key ? String(key).replace(/^(\d{4})-(\d{2})$/, '$1年$2月') : '月份待核对';
  const money = cents => cents === null || cents === undefined ? '' : C.formatMoney(cents);
  const sourceText = record => {
    const s = record?.source || {};
    return [s.file, s.sheet, s.row != null ? '第' + s.row + '行' : '', s.part != null && s.part !== '' ? '分期' + s.part : ''].filter(Boolean).join(' / ');
  };
  function settings() {
    return {mode:$('mode').value,decisions:structuredClone(decisions),matches:{...matches},fingerprints:Object.fromEntries(roles.filter(role => sources[role].book).map(role => [role,sources[role].fingerprint]))};
  }
  async function cacheWork(operation) {
    try { await operation(); $('cache-status').textContent = '文件与核对选择已保存在本机浏览器；不上传服务器。'; return true; }
    catch (error) { $('cache-status').textContent = '本机缓存未能更新：' + message(error) + '。本次文件仍可使用；刷新后可能恢复上次缓存。'; return false; }
  }
  function clearChoices() { decisions = {}; matches = {}; }
  function clearResult() {
    result = null;
    $('result').hidden = true;
    $('download').disabled = true;
    $('warning-panel').hidden = true;
    for (const id of ['preview','warnings','review-list','deleted-list','classified-summary','monthly-summary']) $(id).replaceChildren();
    $('record-count').textContent = $('visible-count').textContent = '0';
  }
  function controls() {
    for (const role of roles) {
      $(role + 'File').disabled = restoring || exporting;
      $('clear-' + role).hidden = sources[role].state === 'empty';
      $('clear-' + role).disabled = restoring || exporting;
      $('required-' + role).textContent = ($('mode').value === 'new' ? ['negotiated','ledger','budget'].includes(role) : role === 'existing') ? '必需' : '可选';
    }
    $('mode').disabled = $('clear-cache').disabled = restoring || exporting;
    $('mode-help').textContent = $('mode').value === 'new' ? '新建整理需上传前三项文件。' : '已有付款汇总明细必需，其他文件可用于补充核对。';
    $('download').disabled = restoring || exporting || !result || result.canExport === false;
    for (const el of $('result').querySelectorAll('[data-action]')) el.disabled = restoring || exporting;
  }
  function reference(role) {
    const s = sources[role];
    let text = (s.state === 'ready' ? '已读取：' : '') + (s.filename || '尚未引用文件');
    if (s.state === 'reading') text += '\n正在读取…';
    if (s.importedAt) text += '\n引用时间：' + new Date(s.importedAt).toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false});
    const sheets = s.sheets.filter(sheet => typeof sheet === 'string' || !sheet.skipped).map(sheet => typeof sheet === 'string' ? sheet : sheet.name || sheet.sheet || '').filter(Boolean);
    if (sheets.length) text += '\n工作表：' + sheets.join('、');
    if (s.state === 'error') text += '\n读取失败：' + s.error;
    $('reference-' + role).textContent = text;
  }
  async function readBook(input) {
    let bytes = input;
    const signature = new Uint8Array(bytes.buffer || bytes, bytes.byteOffset || 0, Math.min(2,bytes.byteLength));
    if (signature[0] === 0x50 && signature[1] === 0x4b) {
      const zip = await JSZip.loadAsync(bytes);
      let changed = false;
      for (const path of Object.keys(zip.files).filter(p => /^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
        const xml = await zip.file(path).async('string');
        const next = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (cell, attrs, content) => {
          if (content === undefined || !/\bt="str"/.test(attrs) || !/<f(?:\s|>|\/)/.test(content) || /<v(?:\s|>|\/)/.test(content)) return cell;
          changed = true;
          return '<c' + attrs.replace(/\s+t="str"/, '') + '>' + content + '</c>';
        });
        if (next !== xml) zip.file(path, next);
      }
      if (changed) bytes = await zip.generateAsync({type:'uint8array'});
    }
    return XLSX.read(bytes,{type:'array',cellStyles:true,cellNF:true,cellFormula:true,cellDates:false,sheetStubs:true});
  }
  async function parseFile(role, entry) {
    if (!/\.xlsx?$/i.test(entry.filename || '')) throw new Error('请选择 .xlsx 或 .xls 文件');
    if (!entry.bytes || !entry.bytes.byteLength) throw new Error('文件为空，请重新选择');
    if (entry.bytes.byteLength > 60 * 1024 * 1024) throw new Error('文件超过 60 MB，请缩小后重新上传');
    const digest = await crypto.subtle.digest('SHA-256',entry.bytes);
    const fingerprint = [...new Uint8Array(digest)].map(v => v.toString(16).padStart(2,'0')).join('');
    if (entry.fingerprint && entry.fingerprint !== fingerprint) throw new Error('缓存文件校验失败，请重新上传');
    const book = await readBook(entry.bytes);
    const parsed = I.parse({book,filename:entry.filename,fingerprint},role,{mode:$('mode').value});
    if (parsed.errors?.length) throw new Error(parsed.errors.map(error => {
      const at = [error?.source?.file,error?.source?.sheet,error?.cell || (error?.source?.row ? '第' + error.source.row + '行' : '')].filter(Boolean).join(' / ');
      return (at ? at + '：' : '') + message(error);
    }).join('\n'));
    return {book,fingerprint,sheets:parsed.sheets || [],warnings:parsed.warnings || []};
  }
  function updateFilter(id, values, label) {
    const select = $(id), previous = select.value;
    select.innerHTML = '<option value="">全部</option>' + [...new Set(values)].map(value => '<option value="' + esc(value) + '">' + esc(label(value)) + '</option>').join('');
    if ([...select.options].some(option => option.value === previous)) select.value = previous;
  }
  function decisionControl(record) {
    const value = decisions[record.id]?.excluded || '';
    return '<select aria-label="' + esc(record.supplier + '：保留或剔除') + '" data-action="decision" data-record-id="' + esc(record.id) + '"><option value=""' + (!value ? ' selected' : '') + '>按当前规则（' + (record.excluded === '是' ? '剔除' : '保留') + '）</option><option value="否"' + (value === '否' ? ' selected' : '') + '>保留</option><option value="是"' + (value === '是' ? ' selected' : '') + '>剔除</option></select><input type="text" maxlength="2000" aria-label="' + esc(record.supplier + '：核对说明') + '" data-action="reason" data-record-id="' + esc(record.id) + '" placeholder="核对说明（可选）" value="' + esc(decisions[record.id]?.reason || '') + '">';
  }
  function renderPreview() {
    if (!result) return;
    const rows = result.records.filter(r => (!$('filter-month').value || (r.monthKey || '@undated') === $('filter-month').value) && (!$('filter-category').value || (r.category || '@empty') === $('filter-category').value) && (!$('filter-status').value || (r.status || '@empty') === $('filter-status').value) && (!$('filter-excluded').value || r.excluded === $('filter-excluded').value));
    $('visible-count').textContent = rows.length;
    $('preview').innerHTML = '<table><thead><tr>' + C.HEADERS.map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '<th scope="col">核对处理</th></tr></thead><tbody>' + rows.map(record => {
      const values = C.toRow(record).slice();
      values[0] = C.monthLabel(record); values[5] = money(record.cents);
      const style = record.excluded === '是' ? 'excluded' : record.related ? 'related' : record.status === '已支付' ? 'paid' : '';
      return '<tr class="' + style + '" data-record-id="' + esc(record.id) + '">' + values.map((v,i) => '<td' + (i === 5 ? ' class="detail-money"' : '') + '>' + esc(v?.v !== undefined ? v.v : v) + '</td>').join('') + '<td class="row-action">' + decisionControl(record) + '<p class="row-source">' + esc(sourceText(record)) + '</p></td></tr>';
    }).join('') + '</tbody></table>' + (!rows.length ? '<p class="empty">当前筛选没有记录。</p>' : '');
    controls();
  }
  function renderReview() {
    const records = new Map(result.records.map(r => [r.id,r]));
    const issues = result.issues || [];
    const ids = new Set([...result.records.filter(r => r.related || r.excluded === '是').map(r => r.id), ...issues.flatMap(i => i.recordIds || []), ...Object.keys(decisions), ...Object.keys(matches)]);
    const content = issues.filter(i => !(i.recordIds || []).some(id => records.has(id))).map(i => '<p>' + esc(i.message) + '</p>');
    for (const id of ids) {
      const record = records.get(id);
      if (!record) continue;
      const relevant = issues.filter(i => (i.recordIds || []).includes(id));
      const candidates = new Map(relevant.flatMap(i => i.candidates || []).map(c => [c.id,c.label]));
      if (matches[id] && !candidates.has(matches[id])) candidates.set(matches[id],'已确认的付款凭据');
      const evidence = candidates.size ? '<label>付款凭据<select data-action="evidence" data-record-id="' + esc(id) + '"><option value="">暂不指定，按现有证据核对</option>' + [...candidates].map(([key,label]) => '<option value="' + esc(key) + '"' + (matches[id] === key ? ' selected' : '') + '>' + esc(label) + '</option>').join('') + '</select></label>' : '';
      content.push('<article class="review-card"><p><strong>' + esc(record.supplier) + '</strong> · ' + esc(C.monthLabel(record)) + ' · ' + esc(money(record.cents)) + ' 元</p><p class="row-source">' + esc(sourceText(record)) + '</p>' + relevant.map(i => '<p>' + esc(i.message) + '</p>').join('') + '<div class="review-controls"><label>本笔处理<span>' + decisionControl(record) + '</span></label>' + evidence + '</div></article>');
    }
    $('review-list').innerHTML = content.join('');
    $('review-panel').hidden = !content.length;
    $('review-count').textContent = '（' + [...ids].filter(id => records.has(id)).length + ' 条）';
  }
  function summaryTable(rows, classified, totals) {
    const headers = ['付款月份',...(classified ? ['分类'] : []),'申请总金额','已付款金额','剩余付款金额','备注'];
    const cells = r => [month(r.monthKey),...(classified ? [r.category || '分类待核对'] : []),money(r.totalCents),money(r.paidCents),money(r.remainingCents),''];
    const total = ['合计',...(classified ? [''] : []),money(totals.totalCents),money(totals.paidCents),money(totals.remainingCents),''];
    return '<table class="summary-table"><thead><tr>' + headers.map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + rows.map(r => '<tr>' + cells(r).map(v => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('') + '<tr class="total">' + total.map(v => '<td>' + esc(v) + '</td>').join('') + '</tr></tbody></table>';
  }
  function render() {
    $('record-count').textContent = result.records.length;
    updateFilter('filter-month',result.records.map(r => r.monthKey || '@undated').sort(),key => key === '@undated' ? '月份待核对' : month(key));
    updateFilter('filter-category',result.records.map(r => r.category || '@empty'),v => v === '@empty' ? '分类未填' : v);
    updateFilter('filter-status',result.records.map(r => r.status || '@empty'),v => v === '@empty' ? '状态未填' : v);
    const warnings = [...new Set((result.warnings || []).map(message))];
    $('warnings').innerHTML = warnings.map(w => '<li>' + esc(w) + '</li>').join('');
    $('warning-count').textContent = warnings.length; $('warning-panel').hidden = !warnings.length;
    renderReview();
    const deleted = result.deleted || [];
    $('deleted-count').textContent = deleted.length; $('deleted-panel').hidden = !deleted.length;
    $('deleted-list').innerHTML = deleted.map(d => '<article class="review-card"><p>' + esc(d.record?.supplier) + ' · ' + esc(money(d.record?.cents)) + ' 元</p><p class="row-source">删除来源：' + esc(sourceText(d.record)) + '</p><p>' + esc(d.reason) + '</p><p class="row-source">对应保留：' + esc((d.kept || []).map(sourceText).join('\n')) + '</p></article>').join('');
    const summary = result.summary || {classified:[],monthly:[],totals:{}};
    $('classified-summary').innerHTML = summaryTable(summary.classified || [],true,summary.totals || {});
    $('monthly-summary').innerHTML = summaryTable(summary.monthly || [],false,summary.totals || {});
    $('result').hidden = false;
    renderPreview();
    $('status').textContent = result.canExport === false ? '已列出核对结果，请先处理提示的问题。' : '整理完成：' + result.records.length + ' 条明细；自动删除 ' + deleted.length + ' 条跨表重复记录。';
  }
  function refresh() {
    clearResult();
    const reading = roles.filter(role => sources[role].state === 'reading');
    const failed = roles.filter(role => sources[role].state === 'error');
    const required = $('mode').value === 'new' ? ['negotiated','ledger','budget'] : ['existing'];
    const missing = required.filter(role => !sources[role].book);
    if (restoring) $('status').textContent = '正在恢复本机文件…';
    else if (reading.length) $('status').textContent = '正在读取：' + reading.map(role => labels[role]).join('、') + '…';
    else if (failed.length) $('status').textContent = failed.map(role => labels[role] + '读取失败：' + sources[role].error).join('\n');
    else if (missing.length) $('status').textContent = '请上传：' + missing.map(role => labels[role]).join('、') + '。';
    else {
      try {
        const inputs = Object.fromEntries(roles.filter(role => sources[role].book).map(role => [role,{book:sources[role].book,filename:sources[role].filename,fingerprint:sources[role].fingerprint}]));
        result = C.build(inputs,{mode:$('mode').value,decisions,matches});
        const ids = new Set(result.records.map(r => r.id));
        for (const selections of [decisions,matches]) for (const id of Object.keys(selections)) if (!ids.has(id)) delete selections[id];
        render();
      } catch (error) { clearResult(); $('status').textContent = '整理失败：' + message(error); }
    }
    controls();
  }
  if (!C || !I || !Cache || !window.XLSX || !window.JSZip || !window.QuarterBudgetExport) {
    $('status').textContent = '工具加载失败，请刷新页面后重试。';
    return;
  }
  for (const role of roles) {
    $(role + 'File').addEventListener('change',async () => {
      const file = $(role + 'File').files[0];
      if (!file || exporting || restoring) return;
      const s = sources[role], current = ++s.version;
      clearChoices();
      Object.assign(s,{book:null,filename:file.name,fingerprint:'',importedAt:'',state:'reading',error:'',sheets:[]});
      reference(role); refresh();
      // Remove the previous file first, so a failed replacement cannot restore stale decisions.
      void cacheWork(() => Cache.removeFile(role,settings()));
      try {
        if (file.size > 60 * 1024 * 1024) throw new Error('文件超过 60 MB，请缩小后重新上传');
        if (!/\.xlsx?$/i.test(file.name)) throw new Error('请选择 .xlsx 或 .xls 文件');
        const entry = {filename:file.name,bytes:await file.arrayBuffer(),importedAt:new Date().toISOString()};
        if (s.version !== current) return;
        const parsed = await parseFile(role,entry);
        if (s.version !== current) return;
        Object.assign(s,parsed,{filename:entry.filename,importedAt:entry.importedAt,state:'ready'});
        entry.fingerprint = parsed.fingerprint;
        reference(role); refresh();
        await cacheWork(() => Cache.saveFile(role,entry,settings()));
      } catch (error) {
        if (s.version !== current) return;
        Object.assign(s,{book:null,fingerprint:'',state:'error',error:message(error)});
        clearChoices(); reference(role); refresh();
        await cacheWork(() => Cache.removeFile(role,settings()));
      } finally {
        if (s.version === current) { $(role + 'File').value = ''; reference(role); refresh(); }
      }
    });
    $('clear-' + role).addEventListener('click',() => {
      if (exporting || restoring) return;
      const s = sources[role]; s.version++;
      Object.assign(s,{book:null,filename:'',fingerprint:'',importedAt:'',state:'empty',error:'',sheets:[]});
      clearChoices(); $(role + 'File').value = ''; reference(role); refresh();
      void cacheWork(() => Cache.removeFile(role,settings()));
    });
  }
  $('mode').addEventListener('change',() => { clearChoices(); refresh(); void cacheWork(() => Cache.saveState(settings())); });
  $('clear-cache').addEventListener('click',async () => {
    if (exporting || restoring) return;
    for (const role of roles) {
      const s = sources[role]; s.version++;
      Object.assign(s,{book:null,filename:'',fingerprint:'',importedAt:'',state:'empty',error:'',sheets:[]});
      $(role + 'File').value = ''; reference(role);
    }
    clearChoices(); refresh();
    if (await cacheWork(() => Cache.clear())) $('cache-status').textContent = '已清空本工具的本机文件和核对选择。';
  });
  $('result').addEventListener('change',event => {
    const el = event.target, id = el.dataset.recordId, action = el.dataset.action;
    if (!action || !id || !result || exporting) return;
    const record = result.records.find(r => r.id === id);
    if (!record) return;
    if (action === 'decision') {
      if (!el.value) delete decisions[id];
      else decisions[id] = {excluded:el.value,reason:decisions[id]?.reason || ''};
    } else if (action === 'reason') decisions[id] = {excluded:decisions[id]?.excluded || record.excluded || '否',reason:el.value.trim()};
    else if (action === 'evidence') { if (el.value) matches[id] = el.value; else delete matches[id]; }
    refresh(); void cacheWork(() => Cache.saveState(settings()));
  });
  for (const id of ['filter-month','filter-category','filter-status','filter-excluded']) $(id).addEventListener('change',renderPreview);
  $('reset-filters').addEventListener('click',() => { for (const id of ['filter-month','filter-category','filter-status','filter-excluded']) $(id).value = ''; renderPreview(); });
  function nextDownload() {
    const date = new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
    let versions = {};
    try { const saved = JSON.parse(localStorage.getItem(versionKey) || '{}'); if (saved && typeof saved === 'object' && !Array.isArray(saved)) versions = saved; } catch (_) {}
    const stored = Number(versions[date]);
    const version = Math.max(Number.isSafeInteger(stored) && stored >= 0 && stored < Number.MAX_SAFE_INTEGER ? stored : 0,memoryVersions[date] || 0) + 1;
    return {date,version,versions,name:'季度预算整理-' + date + '_V' + version + '.xlsx'};
  }
  $('download').addEventListener('click',async () => {
    if (!result || result.canExport === false || exporting || restoring) return;
    const snapshot = result;
    exporting = true; controls(); $('status').textContent = '正在生成 Excel…';
    let url;
    try {
      const bytes = await window.QuarterBudgetExport(snapshot);
      if (!bytes?.byteLength) throw new Error('没有生成有效的文件内容');
      const download = nextDownload();
      url = URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      const link = document.createElement('a'); link.href = url; link.download = download.name;
      document.body.append(link); link.click(); link.remove();
      memoryVersions[download.date] = download.version;
      let saved = true;
      try { localStorage.setItem(versionKey,JSON.stringify({...download.versions,[download.date]:download.version})); } catch (_) { saved = false; }
      $('status').textContent = '已生成：' + download.name + (saved ? '' : '。本机版本号未能保存，本页内仍会递增。');
    } catch (error) { $('status').textContent = '导出失败：' + message(error) + '。可以重试下载。'; }
    finally { if (url) setTimeout(() => URL.revokeObjectURL(url),1000); exporting = false; controls(); }
  });
  async function restore() {
    try {
      const saved = await Cache.load();
      if (['new','existing'].includes(saved.settings?.mode)) $('mode').value = saved.settings.mode;
      let invalid = false;
      await Promise.all(roles.map(async role => {
        const entry = saved.files?.[role];
        if (!entry) return;
        const s = sources[role]; Object.assign(s,{filename:entry.filename,importedAt:entry.importedAt,state:'reading'}); reference(role);
        try { const parsed = await parseFile(role,entry); Object.assign(s,parsed,{state:'ready'}); }
        catch (error) { invalid = true; Object.assign(s,{book:null,state:'error',error:message(error)}); }
        reference(role);
      }));
      const fingerprints = saved.settings?.fingerprints || {};
      const unchanged = roles.every(role => (sources[role].fingerprint || '') === (fingerprints[role] || ''));
      if (!invalid && unchanged) {
        const storedDecisions = saved.settings?.decisions, storedMatches = saved.settings?.matches;
        if (storedDecisions && typeof storedDecisions === 'object' && !Array.isArray(storedDecisions)) decisions = Object.fromEntries(Object.entries(storedDecisions).filter(([,v]) => v && ['是','否'].includes(v.excluded) && typeof v.reason === 'string'));
        if (storedMatches && typeof storedMatches === 'object' && !Array.isArray(storedMatches)) matches = Object.fromEntries(Object.entries(storedMatches).filter(([,v]) => typeof v === 'string'));
      }
      if (invalid) {
        clearChoices();
        for (const role of roles.filter(role => sources[role].state === 'error')) await Cache.removeFile(role,settings());
      } else if (!unchanged) await Cache.saveState(settings());
      $('cache-status').textContent = roles.some(role => sources[role].book) ? '已恢复本机文件和仍有效的核对选择，未上传服务器。' : '文件与核对选择仅缓存在当前浏览器，不上传服务器。';
    } catch (error) { $('cache-status').textContent = '未能读取本机缓存：' + message(error) + '。可以重新上传，本次整理仍可使用。'; }
    finally { restoring = false; refresh(); window.QuarterBudgetToolReady = true; }
  }
  controls(); void restore();
})();
