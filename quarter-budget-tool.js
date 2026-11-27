(() => {
  'use strict';
  const $ = id => document.getElementById(id), C = window.QuarterBudgetCore, I = window.QuarterBudgetImport, Cache = window.QuarterBudgetCache;
  const roles = ['negotiated','ledger','budget','manualBudget'];
  const labels = {negotiated:'付款节奏',ledger:'资金台账',budget:'预估预算',manualBudget:'手工下月预算'};
  const sources = Object.fromEntries(roles.map(role => [role,{book:null,filename:'',fingerprint:'',importedAt:'',state:'empty',error:'',version:0,sheets:[]} ]));
  let result = null, decisions = {}, matches = {}, exporting = false, restoring = true, monthEdited = false;
  const versionKey = 'quarter-budget-export-versions', memoryVersions = {};
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const message = value => typeof value === 'string' ? value : value?.message || String(value ?? '');
  const month = key => key ? String(key).replace(/^(\d{4})-(\d{2})$/, '$1年$2月') : '月份待核对';
  const money = cents => cents === null || cents === undefined ? '' : C.formatMoney(cents);
  const sourceText = record => {
    const s = record?.source || {};
    return [s.file, s.sheet, s.row != null ? '第' + s.row + '行' : '', s.part != null && s.part !== '' ? '分期' + s.part : ''].filter(Boolean).join(' / ');
  };
  function currentInputs() {
    return Object.fromEntries(roles.filter(role => sources[role].book).map(role => [role,{book:sources[role].book,filename:sources[role].filename,fingerprint:sources[role].fingerprint}]));
  }
  function syncBudgetMonth() {
    if (!monthEdited) {
      const next = C.resolveBudgetStartMonth(currentInputs());
      if ($('budgetStartMonth').value && $('budgetStartMonth').value !== next) clearChoices();
      $('budgetStartMonth').value = next;
    }
  }
  function settings() {
    return {schema:2,budgetStartMonth:$('budgetStartMonth').value,monthEdited,decisions:structuredClone(decisions),matches:{...matches},fingerprints:Object.fromEntries(roles.filter(role => sources[role].book).map(role => [role,sources[role].fingerprint]))};
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
    for (const id of ['preview','warnings','review-list','deleted-list','classified-summary','monthly-summary','comparison-totals']) $(id).replaceChildren();
    $('record-count').textContent = $('visible-count').textContent = '0';
    $('comparison-panel').hidden = true;
  }
  function controls() {
    for (const role of roles) {
      $(role + 'File').disabled = restoring || exporting;
      $('clear-' + role).hidden = sources[role].state === 'empty';
      $('clear-' + role).disabled = restoring || exporting;
      $('required-' + role).textContent = role !== 'manualBudget' ? '必需' : '可选';
    }
    $('budgetStartMonth').disabled = $('clear-cache').disabled = restoring || exporting;
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
    const parsed = I.parse({book,filename:entry.filename,fingerprint},role,{budgetStartMonth:$('budgetStartMonth').value});
    const fatal = (parsed.errors || []).filter(error => !['AMOUNT_FORMULA_ERROR','INVALID_AMOUNT','MISSING_AMOUNT'].includes(error.code));
    if (fatal.length) throw new Error(fatal.map(error => {
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
    return '<select aria-label="' + esc(record.supplier + '：保留或剔除') + '" data-action="decision" data-record-id="' + esc(record.id) + '"><option value=""' + (!value ? ' selected' : '') + '>按当前规则（' + (record.excluded === '是' ? '剔除' : '保留') + '）</option><option value="否"' + (value === '否' ? ' selected' : '') + '>保留</option><option value="是"' + (value === '是' ? ' selected' : '') + '>剔除</option></select><input type="text" maxlength="2000" aria-label="' + esc(record.supplier + '：核对说明') + '" data-action="reason" data-record-id="' + esc(record.id) + '" placeholder="剔除缺失金额或保留重复时必填原因" value="' + esc(decisions[record.id]?.reason || '') + '">';
  }
  function renderPreview() {
    if (!result) return;
    const rows = result.records.filter(r => (!$('filter-month').value || (r.monthKey || '@undated') === $('filter-month').value) && (!$('filter-category').value || (r.category || '@empty') === $('filter-category').value) && (!$('filter-status').value || (r.status || '@empty') === $('filter-status').value) && (!$('filter-excluded').value || r.excluded === $('filter-excluded').value));
    const comparison = result.comparison;
    const extraHeaders = comparison ? ['系统预算金额','手工预算金额','差异金额','核对结果','系统来源','手工来源'] : [];
    const detailRow = record => {
      const values = C.toRow(record).slice();
      values[0] = C.monthLabel(record); values[5] = money(record.cents);
      const style = record.excluded === '是' ? 'excluded' : record.related ? 'related' : record.status === '已支付' ? 'paid' : '';
      return '<tr class="' + style + '" data-record-id="' + esc(record.id) + '">' + values.map((v,i) => '<td' + (i === 5 ? ' class="detail-money"' : '') + '>' + esc(v?.v !== undefined ? v.v : v) + '</td>').join('') + '<td class="row-action">' + decisionControl(record) + '<p class="row-source">' + esc(sourceText(record)) + '</p></td>' + extraHeaders.map(() => '<td></td>').join('') + '</tr>';
    };
    const content = [], included = new Set(), groups = new Map();
    for (const record of rows) {
      const key = C.comparisonKey(record);
      if (!groups.has(key)) groups.set(key,[]);
      groups.get(key).push(record);
    }
    for (const group of comparison?.rows || []) {
      const key = C.comparisonKey(group), details = groups.get(key) || [];
      if (!details.length && (($('filter-month').value && $('filter-month').value !== group.monthKey) || ($('filter-category').value && $('filter-category').value !== (group.category || '@empty')) || $('filter-status').value || $('filter-excluded').value)) continue;
      const values = [C.monthLabel(group),group.supplier,group.subject,'','分组对比（显示' + details.length + '条明细）','','','',group.category,'','','',...[group.systemCents,group.manualCents,group.differenceCents].map(v => v == null ? '金额待核对' : money(v)),group.status,...[group.systemSources,group.manualSources].map(list => (list || []).map(source => sourceText({source})).join('\n'))];
      content.push('<tr class="comparison-row' + (group.status !== '一致' ? ' comparison-difference' : '') + '" data-comparison-key="' + esc(key) + '">' + values.map((v,i) => '<td' + (i >= 12 && i <= 14 ? ' class="detail-money"' : '') + '>' + esc(v) + '</td>').join('') + '</tr>');
      for (const record of details) { included.add(record.id); content.push(detailRow(record)); }
    }
    for (const record of rows) if (!included.has(record.id)) content.push(detailRow(record));
    $('visible-count').textContent = rows.length;
    $('preview').innerHTML = '<table' + (comparison ? ' class="combined-table"' : '') + ' aria-label="' + (comparison ? '付款明细与下月预算对比' : '付款明细') + '"><thead><tr>' + [...C.HEADERS,'核对处理',...extraHeaders].map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + content.join('') + '</tbody></table>' + (!content.length ? '<p class="empty">当前筛选没有记录。</p>' : '');
    controls();
  }
  function renderReview() {
    const records = new Map(result.records.map(r => [r.id,r]));
    const issues = result.issues || [];
    const ids = new Set([...issues.flatMap(i => i.recordIds || []), ...result.records.filter(r => r.related || r.excluded === '是').map(r => r.id), ...Object.keys(decisions), ...Object.keys(matches)]);
    const content = issues.filter(i => !(i.recordIds || []).some(id => records.has(id))).map(i => '<p>' + esc(i.message) + '</p>');
    const rows = [];
    const issueLabels = {duplicate:'疑似重复',amount:'金额异常',evidence:'付款依据待核对',error:'读取异常'};
    for (const id of ids) {
      const record = records.get(id);
      if (!record) continue;
      const relevant = issues.filter(i => (i.recordIds || []).includes(id));
      const candidates = new Map(relevant.flatMap(i => i.candidates || []).map(c => [c.id,c.label]));
      if (matches[id] && !candidates.has(matches[id])) candidates.set(matches[id],'已确认的付款凭据');
      const evidence = candidates.size ? '<select aria-label="' + esc(record.supplier + '：付款凭据') + '" data-action="evidence" data-record-id="' + esc(id) + '"><option value="">暂不指定，按现有证据核对</option>' + [...candidates].map(([key,label]) => '<option value="' + esc(key) + '"' + (matches[id] === key ? ' selected' : '') + '>' + esc(label) + '</option>').join('') + '</select>' : '—';
      const notes = relevant.map(i => '<details><summary>' + esc(issueLabels[i.type] || '核对提示') + (i.blocking ? (i.resolved ? '（已处理）' : '（待处理）') : '') + '</summary><p>' + esc(i.message) + '</p></details>').join('') || (record.excluded === '是' ? '已剔除' : '已保留');
      const values = [record.supplier,record.subject,C.monthLabel(record),record.category,record.cents == null ? '金额待核对' : money(record.cents),record.oa,record.status];
      rows.push('<tr' + (record.excluded === '是' ? ' class="excluded"' : '') + ' data-record-id="' + esc(id) + '">' + values.map((v,i) => '<td' + (i === 4 ? ' class="detail-money"' : '') + '>' + esc(v) + '</td>').join('') + '<td class="review-source">' + esc(sourceText(record)) + '</td><td class="review-notes">' + notes + '</td><td class="row-action">' + decisionControl(record) + '</td><td class="row-action">' + evidence + '</td></tr>');
    }
    if (rows.length) content.push('<div class="preview review-scroll" tabindex="0" role="region" aria-label="关联与付款核对表，可横向滚动"><table class="review-table" aria-label="关联与付款核对"><thead><tr>' + ['供应商','付款主体','付款月份','分类','金额','OA流程号','付款状态','来源','核对提示','核对处理','付款凭据'].map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + rows.join('') + '</tbody></table></div>');
    $('review-list').innerHTML = content.join('');
    $('review-panel').hidden = !content.length;
    $('review-count').textContent = '（' + rows.length + ' 条）';
  }
  function summaryTable(rows, classified, totals) {
    const headers = ['付款月份',...(classified ? ['分类'] : []),'申请总金额','已付款金额','剩余付款金额','备注'];
    const cells = r => [month(r.monthKey),...(classified ? [r.category || '分类待核对'] : []),money(r.totalCents),money(r.paidCents),money(r.remainingCents),''];
    const total = ['合计',...(classified ? [''] : []),money(totals.totalCents),money(totals.paidCents),money(totals.remainingCents),''];
    return '<table class="summary-table"><thead><tr>' + headers.map(h => '<th scope="col">' + esc(h) + '</th>').join('') + '</tr></thead><tbody>' + rows.map(r => '<tr>' + cells(r).map(v => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('') + '<tr class="total">' + total.map(v => '<td>' + esc(v) + '</td>').join('') + '</tr></tbody></table>';
  }
  function renderComparison() {
    const comparison = result.comparison;
    $('comparison-panel').hidden = !comparison;
    $('detail-title').textContent = comparison ? '付款明细与下月预算对比' : '付款明细';
    if (!comparison) return;
    $('comparison-month').textContent = '（' + month(comparison.monthKey) + '）';
    const amount = value => value == null ? '待核对' : money(value);
    const total = comparison.totals;
    $('comparison-totals').innerHTML = [['系统预算金额',total.systemCents],['手工预算金额',total.manualCents],['差异金额',total.differenceCents]].map(([label,value]) => '<span>' + esc(label) + '<strong>' + esc(amount(value)) + '</strong></span>').join('');
  }
  function render() {
    $('record-count').textContent = result.records.length;
    const filterRows = [...result.records,...(result.comparison?.rows || [])];
    updateFilter('filter-month',filterRows.map(r => r.monthKey || '@undated').sort(),key => key === '@undated' ? '月份待核对' : month(key));
    updateFilter('filter-category',filterRows.map(r => r.category || '@empty'),v => v === '@empty' ? '分类未填' : v);
    updateFilter('filter-status',result.records.map(r => r.status || '@empty'),v => v === '@empty' ? '状态未填' : v);
    const warnings = [...new Set([...(result.warnings || []),...(result.comparison?.warnings || [])].map(message))];
    $('warnings').innerHTML = warnings.map(w => '<li>' + esc(w) + '</li>').join('');
    $('warning-count').textContent = warnings.length; $('warning-panel').hidden = !warnings.length;
    renderReview();
    const deleted = result.deleted || [];
    $('deleted-count').textContent = deleted.length; $('deleted-panel').hidden = !deleted.length;
    $('deleted-list').innerHTML = deleted.map(d => '<article class="review-card"><p>' + esc(d.record?.supplier) + ' · ' + esc(money(d.record?.cents)) + ' 元</p><p class="row-source">删除来源：' + esc(sourceText(d.record)) + '</p><p>' + esc(d.reason) + '</p><p class="row-source">对应保留：' + esc((d.kept || []).map(sourceText).join('\n')) + '</p></article>').join('');
    const summary = result.summary || {classified:[],monthly:[],totals:{}};
    $('classified-summary').innerHTML = summaryTable(summary.classified || [],true,summary.totals || {});
    $('monthly-summary').innerHTML = summaryTable(summary.monthly || [],false,summary.totals || {});
    renderComparison();
    $('result').hidden = false;
    renderPreview();
    $('status').textContent = result.canExport === false ? '已列出核对结果，请处理未确认的重复或金额问题后下载。' : '整理完成：' + result.records.length + ' 条明细；自动删除 ' + deleted.length + ' 条跨表重复记录。';
  }
  function refresh() {
    clearResult();
    const reading = roles.filter(role => sources[role].state === 'reading');
    const failed = roles.filter(role => sources[role].state === 'error');
    syncBudgetMonth();
    const required = ['negotiated','ledger','budget'];
    const missing = required.filter(role => !sources[role].book);
    if (restoring) $('status').textContent = '正在恢复本机文件…';
    else if (reading.length) $('status').textContent = '正在读取：' + reading.map(role => labels[role]).join('、') + '…';
    else if (failed.length) $('status').textContent = failed.map(role => labels[role] + '读取失败：' + sources[role].error).join('\n');
    else if (!/^\d{4}-(0[1-9]|1[0-2])$/.test($('budgetStartMonth').value)) $('status').textContent = '请选择有效的预算起月。';
    else if (missing.length) $('status').textContent = '请上传：' + missing.map(role => labels[role]).join('、') + '。';
    else {
      try {
        const inputs = currentInputs();
        result = C.build(inputs,{budgetStartMonth:$('budgetStartMonth').value,decisions,matches});
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
  $('budgetStartMonth').addEventListener('change',() => { monthEdited = true; clearChoices(); refresh(); void cacheWork(() => Cache.saveState(settings())); });
  $('clear-cache').addEventListener('click',async () => {
    if (exporting || restoring) return;
    for (const role of roles) {
      const s = sources[role]; s.version++;
      Object.assign(s,{book:null,filename:'',fingerprint:'',importedAt:'',state:'empty',error:'',sheets:[]});
      $(role + 'File').value = ''; reference(role);
    }
    monthEdited = false; clearChoices(); refresh();
    if (await cacheWork(() => Cache.clear())) $('cache-status').textContent = '已清空本工具的本机文件和核对选择。';
  });
  $('result').addEventListener('change',event => {
    const el = event.target, id = el.dataset.recordId, action = el.dataset.action;
    if (!action || !id || !result || exporting) return;
    const record = result.records.find(r => r.id === id);
    if (!record) return;
    const scroll = el.closest('.preview');
    const position = scroll ? {selector:scroll.id ? '#' + scroll.id : '.review-scroll',left:scroll.scrollLeft,top:scroll.scrollTop} : null;
    if (action === 'decision') {
      if (!el.value) delete decisions[id];
      else decisions[id] = {excluded:el.value,reason:decisions[id]?.reason || ''};
    } else if (action === 'reason') decisions[id] = {excluded:decisions[id]?.excluded || record.excluded || '否',reason:el.value.trim()};
    else if (action === 'evidence') { if (el.value) matches[id] = el.value; else delete matches[id]; }
    refresh();
    if (position) document.querySelector(position.selector)?.scrollTo(position.left,position.top);
    void cacheWork(() => Cache.saveState(settings()));
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
      if (saved.settings?.schema === 2 && /^\d{4}-(0[1-9]|1[0-2])$/.test(saved.settings.budgetStartMonth || '')) {
        $('budgetStartMonth').value = saved.settings.budgetStartMonth; monthEdited = saved.settings.monthEdited === true;
      }
      let invalid = false;
      await Promise.all(roles.map(async role => {
        const entry = saved.files?.[role];
        if (!entry) return;
        const s = sources[role]; Object.assign(s,{filename:entry.filename,importedAt:entry.importedAt,state:'reading'}); reference(role);
        try { const parsed = await parseFile(role,entry); Object.assign(s,parsed,{state:'ready'}); }
        catch (error) { invalid = true; Object.assign(s,{book:null,state:'error',error:message(error)}); }
        reference(role);
      }));
      syncBudgetMonth();
      const fingerprints = saved.settings?.fingerprints || {};
      const unchanged = saved.settings?.schema === 2 && saved.settings.budgetStartMonth === $('budgetStartMonth').value && roles.every(role => (sources[role].fingerprint || '') === (fingerprints[role] || ''));
      if (!invalid && unchanged) {
        const storedDecisions = saved.settings?.decisions, storedMatches = saved.settings?.matches;
        if (storedDecisions && typeof storedDecisions === 'object' && !Array.isArray(storedDecisions)) decisions = Object.fromEntries(Object.entries(storedDecisions).filter(([,v]) => v && ['是','否'].includes(v.excluded) && typeof v.reason === 'string'));
        if (storedMatches && typeof storedMatches === 'object' && !Array.isArray(storedMatches)) matches = Object.fromEntries(Object.entries(storedMatches).filter(([,v]) => typeof v === 'string'));
      }
      if (invalid) {
        clearChoices();
        for (const role of roles.filter(role => sources[role].state === 'error')) await Cache.removeFile(role,settings());
      } else if (!unchanged) { syncBudgetMonth(); await Cache.saveState(settings()); }
      $('cache-status').textContent = roles.some(role => sources[role].book) ? saved.migrated ? '已恢复原有三个来源；旧付款依据、已有明细和旧核对选择已清除，请重新核对。' : '已恢复本机文件和仍有效的核对选择，未上传服务器。' : '文件与核对选择仅缓存在当前浏览器，不上传服务器。';
    } catch (error) { $('cache-status').textContent = '未能读取本机缓存：' + message(error) + '。可以重新上传，本次整理仍可使用。'; }
    finally { restoring = false; refresh(); window.QuarterBudgetToolReady = true; }
  }
  syncBudgetMonth(); controls(); void restore();
})();
