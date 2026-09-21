(function () {
  'use strict';
  const C = PaymentCore, D = PaymentDocument;
  const state = { revision: 0, workbook: null, sheet: '', fileName: '', template: null, templateName: '内置公用模板', analysis: null, outputs: null, error: '', busy: false };
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
  function mount(element) {
    root = element; render();
    if (!state.template) {
      const revision = state.revision;
      builtin().then(t => { if (revision === state.revision && !state.template) { state.template = t; status('模板已就绪，请上传付款明细'); } })
        .catch(e => { if (revision === state.revision && !state.template) status(e.message, true); });
    }
  }
  function render() {
    root.innerHTML = `<div class="payment-tool">
      <p class="payment-intro">按付款主体整理，每个主体一页 A4 横向。文件仅在本机浏览器处理。</p>
      <div class="payment-upload-grid">
        <label class="payment-upload">付款明细（Excel）<input id="paymentExcel" type="file" accept=".xlsx,.xls" /><span>${esc(state.fileName || '请选择付款明细文件')}</span></label>
        <label class="payment-upload">公用模板（Word，可选）<input id="paymentTemplate" type="file" accept=".docx" /><span>${esc(state.templateName)}</span></label>
      </div>
      <div class="payment-toolbar"><label>工作表 <select id="paymentSheet" ${state.workbook ? '' : 'disabled'}>${(state.workbook?.SheetNames || []).map(n => `<option ${n === state.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label><button id="paymentResetTemplate" class="secondary-button">恢复内置模板</button></div>
      <p class="payment-fixed">付款类型：<strong>采购货款</strong>　付款方式：<strong>对公转账</strong>（固定）</p>
      <div id="paymentStatus" role="status" aria-live="polite"></div>
      <div id="paymentGroups"></div>
      <div class="payment-actions"><button id="paymentGenerate" class="primary-button" ${state.busy ? 'disabled' : ''}>${state.busy ? '正在排版…' : '核对并生成预览'}</button><button class="secondary-button" data-payment-download="docx" disabled>下载 Word</button><button class="secondary-button" data-payment-download="pdf" disabled>下载 PDF 打印版</button></div>
      <p class="payment-hint">PDF 已按 85% 缩放排入纸张，打印请选择“实际大小 / 100%”。Word 可编辑，不同 Word/WPS 的字体替换可能影响分页；固定打印请使用 PDF。</p>
      <div id="paymentPages"></div>
    </div>`;
    root.querySelector('#paymentExcel').addEventListener('change', e => uploadExcel(e.target.files[0]));
    root.querySelector('#paymentTemplate').addEventListener('change', e => uploadTemplate(e.target.files[0]));
    root.querySelector('#paymentResetTemplate').addEventListener('click', () => uploadTemplate(null));
    root.querySelector('#paymentSheet').addEventListener('change', e => { invalidate(); state.sheet = e.target.value; analyzeSheet(); renderGroups(); });
    root.querySelector('#paymentGenerate').addEventListener('click', generate);
    root.querySelectorAll('[data-payment-download]').forEach(b => b.addEventListener('click', () => {
      const kind = b.dataset.paymentDownload, blob = state.outputs?.[kind];
      if (!blob) return;
      downloadBlob(blob, `付款申请单整理结果.${kind}`);
    }));
    renderGroups();
    if (state.outputs) showOutputs();
    else if (state.error) status(state.error, true);
  }
  async function uploadExcel(file) {
    if (!file) return;
    invalidate(); const revision = state.revision;
    state.workbook = null; state.analysis = null; state.fileName = file.name; render();
    try {
      if (!/\.(xlsx|xls)$/i.test(file.name) || file.size > 30 * 1024 * 1024) throw new Error('请选择 30 MB 以内的 Excel 文件');
      const bytes = await file.arrayBuffer();
      if (revision !== state.revision) return;
      state.workbook = XLSX.read(bytes, { type: 'array', cellDates: false, cellFormula: true });
      state.sheet = state.workbook.SheetNames[0];
      if (!state.sheet) throw new Error('文件中没有工作表');
      // A first upload can finish before the initial template fetch.
      const fallbackTemplate = !state.template ? await builtin() : null;
      if (revision !== state.revision) return;
      if (!state.template) state.template = fallbackTemplate;
      analyzeSheet(); render();
    } catch (e) { if (revision === state.revision) { state.workbook = null; state.analysis = null; render(); status(e.message, true); } }
  }
  async function uploadTemplate(file) {
    invalidate(); const revision = state.revision;
    state.template = null;
    status('正在检查模板结构…');
    try {
      if (file && !/\.docx$/i.test(file.name)) throw new Error('模板必须为 .docx 文件');
      const template = file ? await D.readTemplate(await file.arrayBuffer()) : await builtin();
      if (revision !== state.revision) return;
      state.template = template; state.templateName = file ? file.name : '内置公用模板';
      render(); status('模板结构检查通过');
    } catch (e) { if (revision === state.revision) status(e.message + '。请重新选择模板或恢复内置模板。', true); }
  }
  function analyzeSheet() {
    const sheet = state.workbook.Sheets[state.sheet];
    if (!sheet || !sheet['!ref']) { state.analysis = { errors: ['工作表没有数据'], groups: [] }; return; }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (range.e.r > 100000 || range.e.c > 200) { state.analysis = { errors: ['工作表范围过大，请移除无关区域后上传'], groups: [] }; return; }
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: true });
    const first = matrix.findIndex(row => row.some(v => C.text(v).trim()));
    if (first < 0) { state.analysis = { errors: ['工作表没有付款明细'], groups: [] }; return; }
    const headers = matrix[first].map(v => C.text(v).trim()), records = [], errors = [];
    for (let i = first + 1; i < matrix.length; i++) {
      if (!matrix[i].some(v => C.text(v).trim())) continue;
      const record = { _row: range.s.r + i + 1 };
      headers.forEach((h, j) => {
        let value = matrix[i][j] ?? '';
        const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r + i, c: range.s.c + j })];
        if (cell?.t === 'e' || (cell?.f && cell.v == null)) errors.push(`第 ${record._row} 行 ${h}：公式错误或缺少计算结果，请在 Excel 中重算保存`);
        if (h === '申请日期' && typeof value === 'number') {
          const d = XLSX.SSF.parse_date_code(value, { date1904: !!state.workbook.Workbook?.WBProps?.date1904 });
          value = d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : value;
        }
        if (['开户银行及账号', '合同编号'].includes(h) && typeof value === 'number' && Math.abs(value) >= 1e15) errors.push(`第 ${record._row} 行 ${h}：长编号保存为数字可能已丢失精度，请改为文本并核对`);
        record[h] = value;
      });
      records.push(record);
    }
    state.analysis = C.analyze(headers, records);
    state.analysis.errors.unshift(...errors);
    state.error = '';
  }
  function renderGroups() {
    const host = root.querySelector('#paymentGroups'), a = state.analysis;
    if (!a) { host.replaceChildren(); return; }
    if (a.errors.length) { host.innerHTML = `<div class="payment-error"><strong>请先修正以下问题后重新上传：</strong><ul>${a.errors.map(e => `<li>${esc(e)}</li>`).join('')}</ul></div>`; return; }
    host.innerHTML = a.groups.map((g, i) => `<section class="payment-group" data-group="${i}">
      <h3>${esc(g.subject)} <small>${g.rows.length} 笔 · 合计 ￥${C.money(g.total)}</small></h3>
      <p>大写：${C.upper(g.total)}</p>
      <div class="payment-info-grid">${[...C.infoFields, '计划付款日期'].map(f => `<label>${f}<input data-field="${f}" type="${f.includes('日期') ? 'date' : 'text'}" value="${esc(g.info[f])}" ${f !== '计划付款日期' ? 'required' : ''} />${g.choices[f]?.length > 1 ? `<span class="payment-error">源表不一致：${g.choices[f].map(v => esc(v || '空白')).join('、')}。请填写统一值。</span>` : ''}</label>`).join('')}</div>
      ${g.checks.length ? `<div class="payment-warning">${g.checks.map(esc).join('<br>')}<label><input type="checkbox" data-confirm ${g.confirmed ? 'checked' : ''} />已核实差异，仍以本次申请付款金额合计为准</label></div>` : ''}
      <details><summary>核对 ${g.rows.length} 条原始明细</summary><div class="payment-detail-scroll"><table><thead><tr><th>源行</th>${C.fields.map(f => `<th>${f}</th>`).join('')}</tr></thead><tbody>${g.rows.map(r => `<tr><td>${r._row}</td>${C.fields.map(f => `<td>${esc(r[f])}</td>`).join('')}</tr>`).join('')}</tbody></table></div></details>
    </section>`).join('');
    host.querySelectorAll('input').forEach(input => input.addEventListener('input', () => {
      const group = a.groups[Number(input.closest('[data-group]').dataset.group)];
      if (input.hasAttribute('data-confirm')) group.confirmed = input.checked;
      else group.info[input.dataset.field] = input.value;
      invalidate(); status('信息已更新，请重新生成预览');
    }));
    status(`已读取 ${a.groups.length} 个付款主体，共 ${a.groups.reduce((s, g) => s + g.rows.length, 0)} 条明细`);
  }
  async function generate() {
    if (state.busy) return;
    invalidate(); const revision = state.revision;
    try {
      const a = state.analysis;
      if (!a || a.errors.length || !a.groups.length) throw new Error('请先上传并修正付款明细');
      if (!state.template) throw new Error('请先选择有效模板');
      for (const g of a.groups) {
        for (const f of C.infoFields) if (!g.info[f].trim()) throw new Error(`${g.subject}：请填写${f}，源数据冲突时需指定统一值`);
        C.date(g.info['申请日期']); C.date(g.info['计划付款日期']);
        if (g.checks.length && !g.confirmed) throw new Error(`${g.subject}：请先核实付款总金额差异并勾选确认`);
      }
      state.busy = true;
      const button = root.querySelector('#paymentGenerate'); button.disabled = true; button.textContent = '正在排版…';
      status('正在测量内容并压缩到每主体一页…');
      const groups = a.groups.map(g => ({ ...g, info: { ...g.info } }));
      const template = state.template;
      const layouts = [];
      for (const g of groups) { layouts.push(await D.layout(template, g)); await new Promise(resolve => setTimeout(resolve, 0)); if (revision !== state.revision) return; }
      const docx = await D.docx(template, layouts), pdf = await D.pdf(layouts);
      if (revision !== state.revision) return;
      state.outputs = { docx, pdf, layouts, groups };
      showOutputs();
    } catch (e) { if (revision === state.revision) status(e.message, true); }
    finally {
      state.busy = false;
      const button = root.querySelector('#paymentGenerate'); if (button) { button.disabled = false; button.textContent = '核对并生成预览'; }
    }
  }
  function showOutputs() {
    const { layouts, groups } = state.outputs;
    const pages = root.querySelector('#paymentPages'); pages.replaceChildren();
    layouts.forEach((l, i) => {
      const card = document.createElement('section'); card.className = 'payment-preview';
      const label = document.createElement('p');
      label.textContent = `${groups[i].subject} · 第 ${i + 1} 页 / 共 ${layouts.length} 页 · 正文 ${(Math.floor(l.body * 2) / 2).toFixed(1)} pt，表格 ${(Math.floor(l.table * 2) / 2).toFixed(1)} pt${l.table < 6.5 ? '（明细较多，已继续压缩，请放大核对）' : ''}`;
      const canvas = D.canvas(l); canvas.setAttribute('aria-label', `${groups[i].subject}申请单预览`);
      card.append(label, canvas); pages.append(card);
    });
    root.querySelectorAll('[data-payment-download]').forEach(b => b.disabled = false);
    status(`已生成 ${layouts.length} 页，每个付款主体一页；Word 与 PDF 均可下载。`);
  }
  window.PaymentTool = { mount };
})();
