(function () {
  'use strict';
  const C = PaymentCore, D = PaymentDocument;
  const state = { revision: 0, workbook: null, sheet: '', fileName: '', template: null, templateName: '内置付款申请模板', analysis: null, outputs: null, error: '', busy: false,
    times: { detail: '', template: '', directory: '', budget: '', oa: '' }, loads: { detail: 0, template: 0, directory: 0, budget: 0, oa: 0 }, pending: {}, templateInitialized: false, restoring: false, cacheMessage: '',
    budgetWorkbook: null, budgetFile: null, budgetName: '', budgetSheet: '', budget: null, budgetFailure: '', budgetConfirmed: new Set(),
    oaWorkbook: null, oaName: '', oaSheet: '', oa: null, oaFailure: '',
    directoryWorkbook: null, directoryName: '', directorySheet: '', directoryHeaders: [], directoryRows: [], directoryErrors: [], directoryMap: { supplier: '', bank: '', unit: '' }, directoryFailure: '', bankPolicy: 'directory' };
  const now = () => new Date().toLocaleString('zh-CN', { hour12: false });
  const reference = kind => state.pending[kind] ? '正在读取…' : state.times[kind] || '未引用';
  function beginLoad(kind) { invalidate(); if (['detail', 'budget', 'oa'].includes(kind)) state.budgetConfirmed.clear(); state.pending[kind] = true; state.times[kind] = ''; return ++state.loads[kind]; }
  const esc = s => C.text(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let root;
  function invalidate() {
    state.revision++; state.outputs = null;
    root?.querySelectorAll('[data-payment-download]').forEach(b => b.disabled = true);
    const preview = root?.querySelector('#paymentPages'); if (preview) preview.replaceChildren();
  }
  function status(message, error = false) {
    state.error = error ? message : '';
    const box = root?.querySelector('#paymentStatus');
    if (box) { box.textContent = message; box.classList.toggle('payment-error', error); }
  }
  async function builtin() {
    const response = await fetch('assets/payment-template.docx');
    if (!response.ok) throw new Error('内置模板读取失败，请刷新或上传模板');
    return D.readTemplate(await response.arrayBuffer());
  }
  async function cacheChange(kind, file, time, remove = false, extra = {}) {
    try {
      if (remove) await PaymentCache.remove(kind);
      else await PaymentCache.put(kind, { blob: file, name: file?.name || '', time, ...extra });
    } catch (_) {
      state.cacheMessage = '浏览器未能保存文件缓存，本次仍可使用；下次打开需重新上传。';
      const note = root?.querySelector('#paymentCacheMessage'); if (note) note.textContent = state.cacheMessage;
    }
  }
  async function restoreFiles() {
    state.restoring = true; render();
    try {
      for (const [kind, upload] of [['template', uploadTemplate], ['directory', uploadDirectory], ['budget', uploadBudget], ['detail', uploadExcel]]) {
        const saved = await PaymentCache.get(kind);
        if (saved) await upload(saved.blob ? new File([saved.blob], saved.name) : null, saved.time, saved.sheet);
        else if (kind === 'template') await uploadTemplate(null);
      }
    } catch (_) {
      state.cacheMessage = '无法读取浏览器文件缓存，请重新上传文件。';
      if (!state.template) await uploadTemplate(null);
    } finally { state.restoring = false; render(); }
  }
  async function clearCachedFiles() {
    invalidate(); state.budgetConfirmed.clear();
    for (const kind of ['detail', 'template', 'directory', 'budget', 'oa']) { state.loads[kind]++; state.pending[kind] = false; }
    state.restoring = true; render();
    try { await PaymentCache.clear(); }
    catch (_) { state.restoring = false; state.cacheMessage = '缓存清除失败，请重试或在浏览器设置中清除此网站的数据。'; render(); return; }
    for (const kind of ['detail', 'template', 'directory', 'budget', 'oa']) { state.loads[kind]++; state.pending[kind] = false; state.times[kind] = ''; }
    invalidate(); state.workbook = null; state.analysis = null; state.fileName = ''; state.sheet = '';
    state.directoryWorkbook = null; state.directoryName = ''; state.directoryFailure = ''; state.directoryRows = []; state.directoryHeaders = []; state.directoryErrors = [];
    state.budgetWorkbook = null; state.budgetFile = null; state.budgetName = ''; state.budgetSheet = ''; state.budget = null; state.budgetFailure = ''; state.budgetConfirmed.clear();
    state.oaWorkbook = null; state.oaName = ''; state.oaSheet = ''; state.oa = null; state.oaFailure = '';
    state.cacheMessage = '已清除上传文件缓存。';
    await uploadTemplate(null);
    state.restoring = false; render();
  }
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    URL.revokeObjectURL(url);
  }
  function mount(element) {
    root = element; render();
    if (!state.templateInitialized) { state.templateInitialized = true; restoreFiles(); }
  }
  function render() {
    root.innerHTML = `<div class="payment-tool">
      <p class="payment-intro">按付款主体整理，优先排入一页 A4 横向；明细较多时自动分页，保证字体清晰。文件仅在本机浏览器处理。</p>
      <div class="payment-upload-grid">
        <label class="payment-upload">付款明细（Excel）<input id="paymentExcel" type="file" accept=".xlsx,.xls" /><span>${esc(state.fileName || '请选择付款明细文件')}</span><small data-reference="detail">引用时间：${esc(reference('detail'))}</small></label>
        <label class="payment-upload">付款申请模板（Word，可选）<input id="paymentTemplate" type="file" accept=".docx" /><span>${esc(state.templateName)}</span><small data-reference="template">引用时间：${esc(reference('template'))}</small></label>
        <label class="payment-upload">供应商名录（Excel，可选）<input id="paymentDirectory" type="file" accept=".xlsx,.xls" /><span>${esc(state.directoryName || '按供应商全称获取银行信息')}</span><small data-reference="directory">引用时间：${esc(reference('directory'))}</small></label>
        <label class="payment-upload">月份预算表（Excel，可选）<input id="paymentBudget" type="file" accept=".xlsx,.xls" /><span>${esc(state.budgetName || '选择本次核对的月份预算文件')}</span><small data-reference="budget">引用时间：${esc(reference('budget'))}</small></label>
        <label class="payment-upload">对公付款OA（Excel，可选，不缓存）<input id="paymentOA" type="file" accept=".xlsx,.xls" /><span>${esc(state.oaName || '每次打开页面需重新上传')}</span><small data-reference="oa">引用时间：${esc(reference('oa'))}</small></label>
      </div>
      <div class="payment-toolbar"><span id="paymentCacheMessage">${esc(state.cacheMessage || (state.restoring ? '正在恢复本机缓存…' : '除对公付款OA外，其他上传文件在本机缓存，下次打开自动恢复。OA文件仅本次页面会话可用，刷新后需重新上传。文件不会上传服务器。'))}</span><button id="paymentClearCache" class="secondary-button">清除上传文件缓存</button></div>
      <div id="paymentDirectorySettings"></div>
      <div id="paymentBudgetSettings"></div>
      <div id="paymentOASettings"></div>
      <div class="payment-toolbar"><label>工作表 <select id="paymentSheet" ${state.workbook ? '' : 'disabled'}>${(state.workbook?.SheetNames || []).map(n => `<option ${n === state.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label><button id="paymentResetTemplate" class="secondary-button">恢复内置付款申请模板</button></div>
      <p class="payment-fixed">付款类型：<strong>采购货款</strong>　付款方式：<strong>对公转账</strong>（固定）</p>
      <p class="payment-hint">引用供应商名录时，“开户银行及账号”显示银行名称和卡号；未引用名录时整理原明细的单位和账号。账号空格保留，原始明细不改。</p>
      <p class="payment-hint">付款明细备注：付款分类-经办人-原明细备注，空白项略过。</p>
      <div id="paymentStatus" role="status" aria-live="polite"></div>
      <div id="paymentGroups"></div>
      <div id="paymentBudgetChecks"></div>
      <div class="payment-actions"><button id="paymentGenerate" class="primary-button" ${state.busy ? 'disabled' : ''}>${state.busy ? '正在排版…' : '核对并生成预览'}</button><button id="paymentDownloadWord" class="secondary-button" ${state.busy ? 'disabled' : ''}>一键下载 Word</button><button class="secondary-button" data-payment-download="docx" disabled>下载 Word</button><button class="secondary-button" data-payment-download="pdf" disabled>下载 PDF 打印版</button></div>
      <p class="payment-hint">Word 和 PDF 均为 A4 横向，已留打印边距。打印请选择 A4、横向、“实际大小 / 100%”；明细字号最低 8 磅，放不下时自动分页。Word/WPS 字体替换可能影响分页，固定打印可用 PDF。</p>
      <div id="paymentPages"></div>
    </div>`;
    root.querySelector('#paymentClearCache').addEventListener('click', clearCachedFiles);
    root.querySelector('#paymentExcel').addEventListener('change', e => uploadExcel(e.target.files[0]));
    root.querySelector('#paymentTemplate').addEventListener('change', e => uploadTemplate(e.target.files[0]));
    root.querySelector('#paymentDirectory').addEventListener('change', e => uploadDirectory(e.target.files[0]));
    root.querySelector('#paymentBudget').addEventListener('change', e => uploadBudget(e.target.files[0]));
    root.querySelector('#paymentOA').addEventListener('change', e => uploadOA(e.target.files[0]));
    root.querySelector('#paymentResetTemplate').addEventListener('click', () => uploadTemplate(null));
    root.querySelector('#paymentSheet').addEventListener('change', e => { invalidate(); state.budgetConfirmed.clear(); state.sheet = e.target.value; analyzeSheet(); renderGroups(); });
    root.querySelector('#paymentGenerate').addEventListener('click', generate);
    root.querySelector('#paymentDownloadWord').addEventListener('click', downloadAllWord);
    root.onclick = e => {
      const button = e.target.closest('[data-payment-download]');
      if (!button || button.disabled) return;
      const kind = button.dataset.paymentDownload, index = Number(button.dataset.paymentGroup || 0);
      const file = state.outputs?.files[index], group = state.outputs?.groups[index];
      if (!file || !group) return;
      try { downloadBlob(file[kind], C.downloadName([group], kind)); }
      catch (e) { status(e.message, true); }
    };
    renderDirectorySettings();
    renderBudgetSettings();
    renderOASettings();
    renderGroups();
    if (state.outputs) showOutputs();
    else if (state.error) status(state.error, true);
    if (state.restoring) root.querySelectorAll('input, select, button').forEach(e => e.disabled = true);
  }
  async function uploadExcel(file, savedTime) {
    if (!file) return;
    const load = beginLoad('detail');
    if (!savedTime) void cacheChange('detail', null, '', true);
    state.workbook = null; state.analysis = null; state.fileName = file.name; render();
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name) || file.size > 30 * 1024 * 1024) throw new Error('请选择 30 MB 以内的 Excel 文件');
      const bytes = await file.arrayBuffer();
      if (load !== state.loads.detail) return;
      state.workbook = XLSX.read(bytes, { type: 'array', cellDates: false, cellFormula: true });
      state.sheet = state.workbook.SheetNames[0];
      if (!state.sheet) throw new Error('文件中没有工作表');
      state.pending.detail = false; state.times.detail = savedTime || now();
      if (!savedTime) void cacheChange('detail', file, state.times.detail);
      analyzeSheet(); render();
    } catch (e) { if (load === state.loads.detail) { state.pending.detail = false; state.times.detail = ''; state.workbook = null; state.analysis = null; render(); status(e.message, true); } }
  }
  async function uploadTemplate(file, savedTime) {
    const load = beginLoad('template');
    if (!savedTime) void cacheChange('template', null, '', true);
    state.template = null;
    state.templateName = file ? file.name : '内置付款申请模板';
    render();
    status('正在检查模板结构…');
    try {
      if (file && !/\.docx$/i.test(file.name)) throw new Error('模板必须为 .docx 文件');
      const template = file ? await D.readTemplate(await file.arrayBuffer()) : await builtin();
      if (load !== state.loads.template) return;
      state.template = template; state.pending.template = false; state.times.template = savedTime || now();
      if (!savedTime) void cacheChange('template', file, state.times.template);
      render(); status('模板结构检查通过');
    } catch (e) { if (load === state.loads.template) { state.pending.template = false; render(); status(e.message + '。请重新选择模板或恢复内置模板。', true); } }
  }
  function refreshPayments() {
    const previous = new Map((state.analysis?.groups || []).map(g => [g.subject, g]));
    if (state.workbook) {
      analyzeSheet();
      for (const group of state.analysis.groups) {
        const old = previous.get(group.subject);
        if (old) { group.info = { ...old.info }; }
      }
    }
  }
  async function uploadBudget(file, savedTime, savedSheet = '') {
    if (!file) return;
    const load = beginLoad('budget');
    if (!savedTime) void cacheChange('budget', null, '', true);
    state.budgetName = file.name; state.budgetFile = null; state.budgetWorkbook = null;
    state.budgetSheet = ''; state.budget = null; state.budgetFailure = '';
    refreshPayments(); render();
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name) || file.size > 30 * 1024 * 1024) throw new Error('请选择 30 MB 以内的月份预算 Excel 文件');
      const bytes = await file.arrayBuffer();
      if (load !== state.loads.budget) return;
      const book = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true });
      const budget = C.parseBudget(book, file.name, XLSX, savedSheet);
      if (!budget.sheetNames.length) throw new Error(budget.errors.join('；'));
      state.budgetWorkbook = book; state.budgetFile = file; state.budget = budget; state.budgetSheet = budget.sheetName;
      state.pending.budget = false; state.times.budget = savedTime || now();
      if (!savedTime) void cacheChange('budget', file, state.times.budget, false, { sheet: state.budgetSheet });
      refreshPayments(); render();
    } catch (e) {
      if (load !== state.loads.budget) return;
      state.pending.budget = false; state.budgetWorkbook = null; state.budget = null; state.budgetFile = null; state.budgetFailure = e.message;
      refreshPayments(); render(); status(e.message, true);
    }
  }
  function renderBudgetSettings() {
    const host = root.querySelector('#paymentBudgetSettings');
    if (!state.budgetName) { host.replaceChildren(); return; }
    const budget = state.budget;
    host.innerHTML = `<section class="payment-group"><h3>月份预算表</h3><p>${esc(state.budgetName)}</p>
      ${budget ? `<label>预算工作表 <select id="paymentBudgetSheet"><option value="">请选择预算工作表</option>${budget.sheetNames.map(name => `<option value="${esc(name)}" ${name === state.budgetSheet ? 'selected' : ''}>${esc(name)}</option>`).join('')}</select></label>` : ''}
      <button id="paymentClearBudget" class="secondary-button">移除月份预算表</button>
      ${budget?.sheetName && !budget.errors.length ? `<p>已读取 ${budget.rows.length} 条预算，使用当前工作表全部记录核对。</p>` : ''}
      <details><summary>查看预算核对规则</summary><p class="payment-hint">预算月份和文件名不固定，以本次所选工作表为准。有OA号时优先按编号核对供应商和付款主体；任一侧缺号时，名称和主体唯一对应才匹配。金额逐笔完全一致才通过，不累计、不扣减预算。编号冲突、多候选或金额差异须逐条确认后继续；确认不修改原数据。核对信息不写入打印单。</p></details>
      </section>`;
    host.querySelector('#paymentBudgetSheet')?.addEventListener('change', e => {
      invalidate(); state.budgetConfirmed.clear();
      state.budget = C.parseBudget(state.budgetWorkbook, state.budgetName, XLSX, e.target.value);
      state.budgetSheet = state.budget.sheetName;
      void cacheChange('budget', state.budgetFile, state.times.budget, false, { sheet: state.budgetSheet });
      refreshPayments(); render();
    });
    host.querySelector('#paymentClearBudget').addEventListener('click', () => {
      invalidate(); state.loads.budget++; state.pending.budget = false; state.times.budget = ''; state.budgetConfirmed.clear();
      state.budgetWorkbook = null; state.budgetFile = null; state.budget = null; state.budgetName = ''; state.budgetSheet = ''; state.budgetFailure = '';
      void cacheChange('budget', null, '', true); refreshPayments(); render();
    });
  }
  function budgetProblems() {
    if (!state.budgetName) return [];
    if (state.budgetFailure) return [state.budgetFailure];
    if (!state.budget) return ['月份预算表尚未成功读取'];
    return state.budget.errors;
  }
  async function uploadOA(file) {
    if (!file) return;
    const load = beginLoad('oa');
    state.oaName = file.name; state.oaWorkbook = null; state.oa = null; state.oaSheet = ''; state.oaFailure = '';
    refreshPayments(); render();
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name) || file.size > 30 * 1024 * 1024) throw new Error('请选择 30 MB 以内的对公付款OA Excel 文件');
      const bytes = await file.arrayBuffer();
      if (load !== state.loads.oa) return;
      const book = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true });
      const oa = C.parsePaymentOA(book, file.name, XLSX);
      if (!oa.sheetNames.length) throw new Error(oa.errors.join('；'));
      state.oaWorkbook = book; state.oa = oa; state.oaSheet = oa.sheetName;
      state.pending.oa = false; state.times.oa = now();
      refreshPayments(); render();
    } catch (e) {
      if (load !== state.loads.oa) return;
      state.pending.oa = false; state.oaWorkbook = null; state.oa = null; state.oaFailure = e.message;
      refreshPayments(); render(); status(e.message, true);
    }
  }
  function renderOASettings() {
    const host = root.querySelector('#paymentOASettings');
    if (!state.oaName) { host.replaceChildren(); return; }
    host.innerHTML = `<section class="payment-group"><h3>对公付款OA</h3><p>${esc(state.oaName)}</p>
      ${state.oa ? `<label>OA工作表 <select id="paymentOASheet"><option value="">请选择OA工作表</option>${state.oa.sheetNames.map(name => `<option value="${esc(name)}" ${name === state.oaSheet ? 'selected' : ''}>${esc(name)}</option>`).join('')}</select></label>` : ''}
      <button id="paymentClearOA" class="secondary-button">移除对公付款OA</button>
      ${state.oa?.sheetName && !state.oa.errors.length ? `<p>已读取 ${state.oa.rows.length} 条OA记录。本文件不缓存，刷新或重新打开页面需再次上传。</p>` : ''}
      ${state.oa?.warnings?.length ? `<details><summary>查看 ${state.oa.warnings.length} 条来源提醒</summary><ul>${state.oa.warnings.map(message => `<li>${esc(message)}</li>`).join('')}</ul></details>` : ''}
      <details><summary>查看OA核对规则</summary><p class="payment-hint">付款明细“合同编号”作为付款OA号，精确匹配流程编码（兼容流程编号），逐笔核对付款金额并显示未操作者。同号重复行的金额、未操作者一致时读取一次；不一致、金额差异或未匹配时逐笔确认，不修改来源数据。未上传OA文件时提示未核对，仍允许生成。OA核对信息不写入Word/PDF。</p></details></section>`;
    host.querySelector('#paymentOASheet')?.addEventListener('change', e => {
      invalidate(); state.budgetConfirmed.clear();
      state.oa = C.parsePaymentOA(state.oaWorkbook, state.oaName, XLSX, e.target.value);
      state.oaSheet = state.oa.sheetName;
      refreshPayments(); render();
    });
    host.querySelector('#paymentClearOA').addEventListener('click', () => {
      invalidate(); state.budgetConfirmed.clear(); state.loads.oa++; state.pending.oa = false; state.times.oa = '';
      state.oaWorkbook = null; state.oa = null; state.oaName = ''; state.oaSheet = ''; state.oaFailure = '';
      refreshPayments(); render();
    });
  }
  function oaProblems() {
    if (!state.oaName) return [];
    if (state.oaFailure) return [state.oaFailure];
    if (!state.oa) return ['对公付款OA尚未成功读取'];
    return state.oa.errors;
  }
  function renderBudgetChecks() {
    const host = root.querySelector('#paymentBudgetChecks');
    const notices = `${state.budgetName ? '' : '<p class="payment-hint">未进行预算核对（未上传月份预算表）。</p>'}${state.oaName ? '' : '<p class="payment-hint">未进行OA核对（未上传对公付款OA文件），可继续生成。</p>'}`;
    if (!state.budgetName && !state.oaName) { host.innerHTML = notices; return; }
    if (state.pending.budget || state.pending.oa) { host.innerHTML = notices + '<p>正在读取核对文件…</p>'; return; }
    const errors = [...budgetProblems(), ...oaProblems()];
    if (errors.length) { host.innerHTML = `${notices}<div class="payment-error">${errors.map(esc).join('<br>')}<br>请修正或移除对应核对文件后重新生成。</div>`; return; }
    const rows = (state.analysis?.groups || []).flatMap(group => group.rows).filter(row => row._budgetCheck || row._oaCheck);
    if (!rows.length) { host.innerHTML = notices + '<p class="payment-hint">上传有效付款明细后逐笔核对预算及OA。</p>'; return; }
    const matched = rows.filter(row => row._budgetCheck?.status !== 'review' && row._oaCheck?.status !== 'review').length;
    const source = value => `${value.filename} / ${value.sheet} 第 ${value.row} 行`;
    const title = state.budgetName ? '月份预算核对' : '付款OA核对';
    host.innerHTML = `${notices}<section class="payment-group"><h3>${title} <small>${state.oaName ? '核对通过' : '预算内'} ${matched} 笔，待核对 ${rows.length - matched} 笔</small></h3>
      <div class="payment-budget-table-scroll" tabindex="0" role="region" aria-label="${title}表，可横向滚动"><table class="payment-budget-table ${state.oaName ? 'has-oa' : ''}" aria-label="${title}">
      <thead><tr>${['核对结果', '付款申请', '申请金额', '预算信息', ...(state.oaName ? ['OA付款金额', '未操作者'] : []), '核对说明', '确认'].map(title => `<th scope="col">${title}</th>`).join('')}</tr></thead><tbody>
      ${rows.map(row => {
        const check = row._budgetCheck, oa = row._oaCheck, passed = check?.status !== 'review' && oa?.status !== 'review';
        const oaGroups = new Map();
        for (const candidate of oa?.candidates || []) {
          const key = JSON.stringify([candidate.amountText, candidate.unoperated]);
          if (!oaGroups.has(key)) oaGroups.set(key, { ...candidate, sources: [] });
          oaGroups.get(key).sources.push(candidate.source);
        }
        const oaValues = [...oaGroups.values()];
        return `<tr class="payment-budget-check" data-budget-row="${row._row}">
          <td><span class="payment-budget-result ${passed ? 'is-matched' : 'is-review'}">${passed ? (state.oaName ? '核对通过' : '预算内') : '待核对'}</span></td>
          <td><strong>${esc(row['供应商全称'] || '（供应商空白）')}</strong><br>付款主体：${esc(row['付款主体'] || '（空白）')}<br>OA号：${esc(row['合同编号'] || '（空白）')}<small class="payment-budget-source">申请来源：${esc(source({filename:state.fileName, sheet:state.sheet, row:row._row}))}</small></td>
          <td class="payment-budget-amount">￥${esc(C.money(row._cents))}</td>
          <td>${check?.candidates.length ? check.candidates.map((candidate, index) => `<div class="payment-budget-candidate">${check.candidates.length > 1 ? `<strong>候选 ${index + 1}</strong><br>` : ''}供应商：${esc(candidate.supplier || '（空白）')}<br>付款主体：${esc(candidate.subject || '（空白）')}<br>OA号：${esc(candidate.oa || '（空白）')}<br>预算金额：<span class="payment-budget-amount">￥${esc(C.money(candidate.amount))}</span><small class="payment-budget-source">预算来源：${esc(source(candidate.source))}</small></div>`).join('') : check ? '未找到对应预算' : '未进行预算核对'}</td>
          ${state.oaName ? `<td class="payment-oa-amount">${oaValues.length ? oaValues.map((candidate, index) => `<div class="payment-budget-candidate">${oaValues.length > 1 ? `记录 ${index + 1}：` : ''}<span class="payment-budget-amount">￥${esc(candidate.amountText)}</span><small class="payment-budget-source">流程编码：${esc(candidate.oa)}<br>OA来源：${candidate.sources.map(value => esc(source(value))).join('<br>')}</small></div>`).join('') : '未找到对应OA'}</td><td class="payment-oa-unoperated">${oaValues.length ? oaValues.map((candidate, index) => `${oaValues.length > 1 ? `记录 ${index + 1}：` : ''}${esc(candidate.unoperated || '（空白）')}`).join('<br>') : '—'}</td>` : ''}
          <td>${[check ? `${state.oaName ? '预算：' : ''}${check.reason}` : '', oa ? `OA：${oa.reason}` : ''].filter(Boolean).map(esc).join('<br>')}</td>
          <td>${passed ? '无需确认' : `<label><input type="checkbox" data-budget-confirm ${state.budgetConfirmed.has(String(row._row)) ? 'checked' : ''} />已核实，确认继续生成</label>`}</td>
          </tr>`;
      }).join('')}</tbody></table></div></section>`;
    host.querySelectorAll('[data-budget-confirm]').forEach(input => input.addEventListener('change', () => {
      const row = input.closest('[data-budget-row]').dataset.budgetRow;
      if (input.checked) state.budgetConfirmed.add(row); else state.budgetConfirmed.delete(row);
      invalidate(); status('预算核对确认已更新，请重新生成预览');
    }));
  }
  async function uploadDirectory(file, savedTime) {
    if (!file) return;
    const load = beginLoad('directory');
    if (!savedTime) void cacheChange('directory', null, '', true);
    state.directoryName = file.name; state.directoryWorkbook = null; state.directoryRows = []; state.directoryHeaders = []; state.directoryErrors = []; state.directoryFailure = '';
    refreshPayments(); render();
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name) || file.size > 30 * 1024 * 1024) throw new Error('请选择 30 MB 以内的供应商名录 Excel 文件');
      const bytes = await file.arrayBuffer();
      if (load !== state.loads.directory) return;
      state.directoryWorkbook = XLSX.read(bytes, { type: 'array', cellFormula: true });
      state.directorySheet = state.directoryWorkbook.SheetNames[0];
      if (!state.directorySheet) throw new Error('名录中没有工作表');
      readDirectorySheet(); state.pending.directory = false;
      state.times.directory = savedTime || now();
      if (!savedTime) void cacheChange('directory', file, state.times.directory);
      refreshPayments(); render();
    } catch (e) {
      if (load !== state.loads.directory) return;
      state.pending.directory = false; state.directoryWorkbook = null; state.directoryFailure = e.message;
      refreshPayments(); render(); status(e.message, true);
    }
  }
  function readDirectorySheet() {
    state.directoryErrors = []; state.directoryRows = []; state.directoryHeaders = [];
    const sheet = state.directoryWorkbook.Sheets[state.directorySheet];
    if (!sheet?.['!ref']) { state.directoryErrors.push('供应商名录工作表为空'); return; }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (range.e.r > 100000 || range.e.c > 200) { state.directoryErrors.push('名录工作表范围过大'); return; }
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: true });
    const first = matrix.findIndex(row => row.some(v => C.text(v).trim()));
    if (first < 0) { state.directoryErrors.push('供应商名录没有数据'); return; }
    const headers = matrix[first].map(v => C.text(v).trim());
    const nonempty = headers.filter(Boolean);
    if (new Set(nonempty).size !== nonempty.length) { state.directoryErrors.push('名录存在重复列名，请修改后重新上传'); return; }
    state.directoryHeaders = nonempty;
    for (let i = first + 1; i < matrix.length; i++) {
      if (!matrix[i].some(v => C.text(v).trim())) continue;
      const record = { _row: range.s.r + i + 1, _errors: {} };
      headers.forEach((h, j) => {
        if (!h) return;
        const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r + i, c: range.s.c + j })];
        let value = matrix[i][j] ?? '';
        if (typeof value === 'number' && cell?.w && /^[\d ]+$/.test(cell.w)) value = cell.w;
        record[h] = value;
        if (cell?.t === 'e' || (cell?.f && cell.v == null)) record._errors[h] = '公式错误或缺少计算结果';
        else if (typeof cell?.v === 'number' && Math.abs(cell.v) >= 1e15) record._errors[h] = '长编号以数字保存可能丢失精度，请使用文本并核实账号';
      });
      state.directoryRows.push(record);
    }
    const pick = names => names.find(n => nonempty.includes(n)) || '';
    state.directoryMap = { supplier: pick(['供应商全称', '供应商名称', '供应商']) };
  }
  function currentDirectory() {
    if (!state.directoryName) return null;
    if (state.directoryFailure) return { errors: [state.directoryFailure], index: new Map() };
    const errors = [...state.directoryErrors], mapping = state.directoryMap;
    if (!state.directoryWorkbook) errors.push('供应商名录尚未成功读取');
    if (!mapping.supplier) errors.push('请选择名录的供应商全称列');
    if (!state.directoryHeaders.includes('收款账号')) errors.push('名录缺少收款账号列（银行账号不作为卡号来源）');
    if (!state.directoryHeaders.some(h => ['收款银行', '收款银行名称'].includes(h))) errors.push('名录需包含收款银行或收款银行名称列');
    if (errors.length) return { errors, index: new Map() };
    return C.buildDirectory(state.directoryRows.map(r => {
      const bankField = C.text(r['收款银行']).trim() ? '收款银行' : '收款银行名称';
      return { supplier: r[mapping.supplier], account: r['收款账号'], bankName: r[bankField], row: r._row,
        error: [mapping.supplier, '收款账号', bankField].map(k => r._errors[k] ? `${k}：${r._errors[k]}` : '').filter(Boolean).join('；') };
    }));
  }
  function renderDirectorySettings() {
    const host = root.querySelector('#paymentDirectorySettings');
    if (!state.directoryName) { host.replaceChildren(); return; }
    const directory = currentDirectory();
    const options = (selected, empty) => `<option value="">${empty}</option>` + state.directoryHeaders.map(h => `<option value="${esc(h)}" ${h === selected ? 'selected' : ''}>${esc(h)}</option>`).join('');
    host.innerHTML = `<section class="payment-group"><h3>供应商名录匹配设置</h3><div class="payment-info-grid">
      <label>名录工作表<select id="paymentDirectorySheet">${(state.directoryWorkbook?.SheetNames || []).map(n => `<option ${n === state.directorySheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label>
      <label>供应商全称列<select data-directory-field="supplier">${options(state.directoryMap.supplier, '请选择')}</select></label>
      <p>银行名称：收款银行，空白时取收款银行名称<br>卡号：只取收款账号<br>合并填入“开户银行及账号”列</p>
      </div><label>引用规则 <select id="paymentBankPolicy"><option value="directory" ${state.bankPolicy === 'directory' ? 'selected' : ''}>优先供应商名录，差异需核实</option><option value="fill" ${state.bankPolicy === 'fill' ? 'selected' : ''}>仅补全付款明细中的空白银行信息</option></select></label>
      <button id="paymentClearDirectory" class="secondary-button">移除名录</button>
      <p class="payment-hint">按供应商全称精确匹配；只获取银行信息，不校验名录中的付款主体。浙江迈德斯特医疗器械科技有限公司支持简称“迈德斯特”。同名不同账号会拦截，原始明细保留不变。引用时间为本次成功读取时间。</p>
      ${directory.errors.length ? `<p class="payment-error">${directory.errors.map(esc).join('<br>')}</p>` : `<p>已读取 ${state.directoryRows.length} 条名录记录</p>`}</section>`;
    host.querySelector('#paymentDirectorySheet').addEventListener('change', e => { invalidate(); state.directorySheet = e.target.value; readDirectorySheet(); refreshPayments(); render(); });
    host.querySelectorAll('[data-directory-field]').forEach(input => input.addEventListener('change', () => { invalidate(); state.directoryMap[input.dataset.directoryField] = input.value; refreshPayments(); render(); }));
    host.querySelector('#paymentBankPolicy').addEventListener('change', e => { invalidate(); state.bankPolicy = e.target.value; refreshPayments(); render(); });
    host.querySelector('#paymentClearDirectory').addEventListener('click', () => {
      void cacheChange('directory', null, '', true);
      invalidate(); state.loads.directory++; state.pending.directory = false; state.directoryWorkbook = null; state.directoryName = ''; state.times.directory = ''; state.directoryFailure = ''; state.directoryErrors = []; state.directoryRows = []; state.directoryHeaders = [];
      refreshPayments(); render();
    });
  }
  function analyzeSheet() {
    const sheet = state.workbook.Sheets[state.sheet];
    if (!sheet || !sheet['!ref']) { state.analysis = { errors: ['工作表没有数据'], groups: [] }; return; }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (range.e.r > 100000 || range.e.c > 200) { state.analysis = { errors: ['工作表范围过大，请移除无关区域后上传'], groups: [] }; return; }
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: true });
    const first = matrix.findIndex(row => row.some(v => C.text(v).trim()));
    if (first < 0) { state.analysis = { errors: ['工作表没有付款明细'], groups: [] }; return; }
    const headers = matrix[first].map(v => C.text(v).trim()).map(h => h === '付款总金额' ? '' : h), records = [], errors = [];
    for (let i = first + 1; i < matrix.length; i++) {
      if (!matrix[i].some((v, j) => headers[j] && C.text(v).trim())) continue;
      const record = { _row: range.s.r + i + 1 };
      headers.forEach((h, j) => {
        if (!h) return;
        let value = matrix[i][j] ?? '';
        const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r + i, c: range.s.c + j })];
        if (cell?.t === 'e' || (cell?.f && cell.v == null)) errors.push(`第 ${record._row} 行 ${h}：公式错误或缺少计算结果，请在 Excel 中重算保存`);
        if (['申请日期', '应付时间', '应付款时间'].includes(h) && typeof value === 'number') {
          const d = XLSX.SSF.parse_date_code(value, { date1904: !!state.workbook.Workbook?.WBProps?.date1904 });
          value = d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : value;
        }
        if (['开户银行及账号', '合同编号'].includes(h) && typeof value === 'number' && Math.abs(value) >= 1e15) errors.push(`第 ${record._row} 行 ${h}：长编号保存为数字可能已丢失精度，请改为文本并核对`);
        record[h] = value;
      });
      records.push(record);
    }
    if (state.directoryWorkbook && !headers.includes('开户银行及账号')) headers.push('开户银行及账号');
    state.analysis = C.analyze(headers, records);
    state.analysis.errors.unshift(...errors);
    const directory = currentDirectory();
    if (directory?.errors.length) state.analysis.errors.push(...directory.errors);
    for (const group of state.analysis.groups) {
      group.bankChecks = [];
      for (const row of group.rows) {
        const resolved = C.resolveBank(row, directory, state.bankPolicy);
        row._bankReference = resolved; row._resolvedBank = resolved.value;
        if (resolved.error) state.analysis.errors.push(`付款明细第 ${row._row} 行（${row['供应商全称']}）：${resolved.error}`);
        if (resolved.conflict) group.bankChecks.push(`第 ${row._row} 行 ${row['供应商全称']}：名录与明细的收款单位或账号不一致`);
      }
    }
    if (!state.analysis.errors.length && state.budgetName && !budgetProblems().length) {
      for (const group of state.analysis.groups) for (const row of group.rows) row._budgetCheck = C.matchBudget(row, state.budget);
    }
    if (!state.analysis.errors.length && state.oaName && !oaProblems().length) {
      for (const group of state.analysis.groups) for (const row of group.rows) row._oaCheck = C.matchPaymentOA(row, state.oa);
    }
    state.error = '';
  }
  function renderGroups() {
    const host = root.querySelector('#paymentGroups'), a = state.analysis;
    renderBudgetChecks();
    if (!a) { host.replaceChildren(); return; }
    if (a.errors.length) { host.innerHTML = `<div class="payment-error"><strong>请先修正以下问题后重新上传：</strong><ul>${a.errors.map(e => `<li>${esc(e)}</li>`).join('')}</ul></div>`; return; }
    host.innerHTML = a.groups.map((g, i) => `<section class="payment-group" data-group="${i}">
      <h3>${g.subject === '待核对' ? '核对付款主体' : esc(g.subject)} <small>${g.rows.length} 笔 · 合计 ￥${C.money(g.total)}</small></h3>
      ${g.subject === '待核对' ? `<p class="payment-warning">请核对原表第 ${g.rows.map(r => esc(r._row)).join('、')} 行的“付款主体”，填写实际付款公司后重新上传。</p>` : ''}
      <p>大写：${C.upper(g.total)}</p>
      <div class="payment-info-grid">${[...C.infoFields, '计划付款日期'].map(f => `<label>${f}<input data-field="${f}" type="${f.includes('日期') ? 'date' : 'text'}" value="${esc(g.info[f])}" ${f !== '计划付款日期' ? 'required' : ''} />${g.choices[f]?.length > 1 ? `<span class="payment-error">源表不一致：${g.choices[f].map(v => esc(v || '空白')).join('、')}。请填写统一值。</span>` : ''}</label>`).join('')}</div>
      ${g.bankChecks?.length ? `<div class="payment-warning">${g.bankChecks.map(esc).join('<br>')}<label><input type="checkbox" data-bank-confirm ${g.bankConfirmed ? 'checked' : ''} />已对照下方原始明细和引用银行信息，确认使用供应商名录</label></div>` : ''}
      ${g.rows.some(r => r._bankReference?.notice) ? `<p class="payment-warning">${g.rows.filter(r => r._bankReference?.notice).map(r => `第 ${r._row} 行：${esc(r._bankReference.notice)}`).join('<br>')}</p>` : ''}
      <details ${g.bankChecks?.length ? 'open' : ''}><summary>核对 ${g.rows.length} 条原始明细及引用信息</summary><div class="payment-detail-scroll"><table><thead><tr><th>源行</th>${C.fields.map(f => `<th>${f}</th>`).join('')}<th>银行信息来源</th><th>申请单引用银行信息</th></tr></thead><tbody>${g.rows.map(r => `<tr><td>${r._row}</td>${C.fields.map(f => `<td>${esc(r[f])}</td>`).join('')}<td>${esc(r._bankReference?.source)}</td><td>${esc(C.bankSummary(r._resolvedBank, r['供应商全称']).value)}</td></tr>`).join('')}</tbody></table></div></details>
    </section>`).join('');
    host.querySelectorAll('input').forEach(input => input.addEventListener('input', () => {
      const group = a.groups[Number(input.closest('[data-group]').dataset.group)];
      if (input.hasAttribute('data-bank-confirm')) group.bankConfirmed = input.checked;
      else group.info[input.dataset.field] = input.value;
      invalidate(); status('信息已更新，请重新生成预览');
    }));
    status(`已读取 ${a.groups.length} 个付款主体，共 ${a.groups.reduce((s, g) => s + g.rows.length, 0)} 条明细`);
  }
  async function generate() {
    if (state.busy) return;
    invalidate(); const revision = state.revision;
    try {
      if (state.restoring) throw new Error('正在恢复文件缓存，请稍候');
      const a = state.analysis;
      if (Object.values(state.pending).some(Boolean)) throw new Error('文件正在读取，请稍候再生成');
      if (!a || a.errors.length || !a.groups.length) throw new Error('请先上传并修正付款明细');
      const budgetErrors = [...budgetProblems(), ...oaProblems()];
      if (budgetErrors.length) throw new Error(budgetErrors.join('；'));
      const unconfirmed = a.groups.flatMap(group => group.rows).filter(row => (row._budgetCheck?.status === 'review' || row._oaCheck?.status === 'review') && !state.budgetConfirmed.has(String(row._row)));
      if (unconfirmed.length) throw new Error(`请逐条确认预算/OA核对：付款明细第 ${unconfirmed.map(row => row._row).join('、')} 行。`);
      if (!state.template) throw new Error('请先选择有效模板');
      for (const g of a.groups) {
        for (const f of C.infoFields) if (!g.info[f].trim()) throw new Error(`${g.subject}：请填写${f}，源数据冲突时需指定统一值`);
        C.date(g.info['申请日期']); C.date(g.info['计划付款日期']);
        if (g.bankChecks?.length && !g.bankConfirmed) throw new Error(`${g.subject}：请核实名录与明细的银行信息差异并勾选确认`);
      }
      state.busy = true;
      const button = root.querySelector('#paymentGenerate'); button.disabled = true; button.textContent = '正在排版…';
      root.querySelector('#paymentDownloadWord').disabled = true;
      status('正在排版，优先一页并保持付款明细清晰…');
      const groups = a.groups.map(g => ({ ...g, info: { ...g.info } }));
      const template = state.template;
      const layouts = [];
      for (const g of groups) { layouts.push(await D.layout(template, g)); await new Promise(resolve => setTimeout(resolve, 0)); if (revision !== state.revision) return; }
      const files = [];
      for (const pages of layouts) {
        files.push({ docx: await D.docx(template, pages), pdf: await D.pdf(pages) });
        if (revision !== state.revision) return;
      }
      const combinedWord = groups.length === 1 ? files[0].docx : await D.docx(template, layouts.flat());
      if (revision !== state.revision) return;
      state.outputs = { files, layouts, groups, combinedWord };
      showOutputs();
    } catch (e) { if (revision === state.revision) status(e.message, true); }
    finally {
      state.busy = false;
      const button = root.querySelector('#paymentGenerate'); if (button) { button.disabled = state.restoring; button.textContent = '核对并生成预览'; }
      const download = root.querySelector('#paymentDownloadWord'); if (download) download.disabled = state.restoring;
    }
  }
  async function downloadAllWord() {
    if (state.busy || state.restoring) return;
    try {
      if (!state.outputs) await generate();
      const outputs = state.outputs;
      if (!outputs) return;
      downloadBlob(outputs.combinedWord, C.printDownloadName(outputs.groups));
      status(`已开始下载合并 Word，共 ${outputs.groups.length} 个付款主体、${outputs.layouts.flat().length} 页。`);
    } catch (e) { status(e.message, true); }
  }
  function showOutputs() {
    const { layouts, groups } = state.outputs;
    const pages = root.querySelector('#paymentPages'); pages.replaceChildren();
    layouts.forEach((subjectPages, i) => {
      const card = document.createElement('section'); card.className = 'payment-preview';
      subjectPages.forEach((l, j) => {
        const label = document.createElement('p');
        label.textContent = `${groups[i].subject} · 第 ${j + 1} 页 / 共 ${subjectPages.length} 页 · 付款明细 ${l.detail.toFixed(1)} pt${subjectPages.length > 1 ? '（明细较多，已自动分页）' : ''}`;
        const canvas = D.canvas(l); canvas.setAttribute('aria-label', `${groups[i].subject}申请单第 ${j + 1} 页预览`);
        card.append(label, canvas);
      });
      if (groups.length > 1) {
        const actions = document.createElement('div'); actions.className = 'payment-actions';
        for (const kind of ['docx', 'pdf']) {
          const button = document.createElement('button'); button.className = 'secondary-button';
          button.dataset.paymentDownload = kind; button.dataset.paymentGroup = String(i);
          button.textContent = `下载${groups[i].subject} ${kind === 'docx' ? 'Word' : 'PDF'}`;
          actions.append(button);
        }
        card.append(actions);
      }
      pages.append(card);
    });
    root.querySelectorAll('[data-payment-download]').forEach(b => {
      const combined = !b.hasAttribute('data-payment-group');
      b.hidden = combined && groups.length > 1;
      b.disabled = b.hidden;
    });
    status(`已生成 ${layouts.flat().length} 页，共 ${groups.length} 个付款主体；优先一页，超出时自动分页。可一键下载合并 Word，或按付款主体分别下载 Word 和 PDF。`);
  }
  window.PaymentTool = { mount };
})();
