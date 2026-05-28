const tools = [
  { id: 'match', icon: '⇄', title: '智能匹配', desc: '根据关键列把源表字段填入目标表。' },
  { id: 'convert', icon: '↔', title: '格式转换', desc: 'Excel、CSV、JSON 互相转换。' },
  { id: 'clean', icon: '✓', title: '数据清洗', desc: '去空格、去空行、去重、规整文本。' },
  { id: 'merge', icon: '▦', title: '报表合并', desc: '多个文件或多个 Sheet 纵向合并。' },
  { id: 'diff', icon: '≠', title: '数据比对', desc: '按主键识别新增、删除、修改。' },
  { id: 'mask', icon: '◩', title: '敏感脱敏', desc: '手机号、身份证、姓名、邮箱快速遮蔽。' },
  { id: 'split', icon: '÷', title: '数据拆分', desc: '按列值或固定行数拆成多个文件。' },
  { id: 'formula', icon: 'ƒ', title: '派生列', desc: '用公式生成新字段，例如 [单价] * [数量]。' },
  { id: 'pivot', icon: '∑', title: '分组汇总', desc: '按维度做求和、计数、平均、最大、最小。' },
  { id: 'compare', icon: '⌁', title: '多列对比', desc: '选择多列数据，按行或按列找最大值、最小值。' },
  { id: 'chart', icon: '▥', title: '图表生成', desc: '基于字段自动聚合并绘制图表。' },
];

const app = {
  activeTool: 'match',
  files: {},
  result: null,
  resultName: '处理结果.xlsx',
  resultType: 'xlsx',
  chart: null,
};

const $ = (id) => document.getElementById(id);

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  renderToolList();
  switchTool('match');
  $('downloadResult').addEventListener('click', downloadResult);
  $('downloadChart').addEventListener('click', downloadChart);
});

function initTheme() {
  const saved = localStorage.getItem('dpc-theme');
  if (saved === 'dark') document.documentElement.dataset.theme = 'dark';
  $('themeToggle').addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? '' : 'dark';
    if (next) document.documentElement.dataset.theme = next;
    else delete document.documentElement.dataset.theme;
    localStorage.setItem('dpc-theme', next || 'light');
  });
}

function renderToolList() {
  $('toolList').innerHTML = tools.map(tool => `
    <button class="tool-tab" data-tool="${tool.id}">
      <span class="tab-icon">${tool.icon}</span>
      <span>${tool.title}</span>
    </button>
  `).join('');
  document.querySelectorAll('.tool-tab').forEach(button => {
    button.addEventListener('click', () => switchTool(button.dataset.tool));
  });
}

function switchTool(id) {
  app.activeTool = id;
  app.files = {};
  clearResult();
  const tool = tools.find(item => item.id === id);
  $('toolTitle').textContent = tool.title;
  $('toolDesc').textContent = tool.desc;
  document.querySelectorAll('.tool-tab').forEach(btn => btn.classList.toggle('active', btn.dataset.tool === id));
  $('chartCard').classList.toggle('hidden', id !== 'chart');
  document.querySelector('.result-layout').classList.toggle('single', id !== 'chart');
  renderToolPanel(id);
  updateStatus();
}

function renderToolPanel(id) {
  const renderers = {
    match: renderMatchPanel,
    convert: renderConvertPanel,
    clean: renderCleanPanel,
    merge: renderMergePanel,
    diff: renderDiffPanel,
    mask: renderMaskPanel,
    split: renderSplitPanel,
    formula: renderFormulaPanel,
    pivot: renderPivotPanel,
    compare: renderComparePanel,
    chart: renderChartPanel,
  };
  $('toolPanel').innerHTML = `<div class="tool-grid">${renderers[id]()}</div>`;
  bindUploads();
  bindActions(id);
}

function fileDropHTML(key, label, accept = '.xlsx,.xls,.csv,.json', multiple = false) {
  return `
    <label class="file-drop" data-upload="${key}">
      <span>
        <strong>${label}</strong>
        拖拽文件到这里，或 <em>点击选择</em>
        <span class="file-name" id="${key}Name">未选择文件</span>
      </span>
      <input type="file" data-input="${key}" accept="${accept}" ${multiple ? 'multiple' : ''} />
    </label>
  `;
}

function renderMatchPanel() {
  return `
    ${fileDropHTML('source', '源数据表')}
    ${fileDropHTML('target', '目标数据表')}
    ${sheetField('sourceSheet', '源表 Sheet')}
    ${sheetField('targetSheet', '目标表 Sheet')}
    ${selectField('sourceKey', '源表匹配列')}
    ${selectField('targetKey', '目标表匹配列')}
    ${selectField('sourceFill', '源表填充列')}
    ${selectField('targetFill', '目标表写入列')}
    <div class="button-row"><button class="primary-button" id="runAction">执行匹配</button></div>
  `;
}

function renderConvertPanel() {
  return `
    ${fileDropHTML('main', '上传待转换文件', '.xlsx,.xls,.csv,.json')}
    ${sheetField('mainSheet', '选择 Sheet')}
    <label class="field"><span>目标格式</span><select id="targetFormat"><option value="xlsx">Excel (.xlsx)</option><option value="csv">CSV (.csv)</option><option value="json">JSON (.json)</option></select></label>
    <div class="button-row"><button class="primary-button" id="runAction">转换并下载</button></div>
  `;
}

function renderCleanPanel() {
  return `
    ${fileDropHTML('main', '上传待清洗文件', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    <label class="field"><span>清洗规则</span><select id="cleanRules" multiple size="5"><option value="trim" selected>去除首尾空格</option><option value="blankRows" selected>删除空行</option><option value="duplicateRows">删除重复行</option><option value="spaces">合并多余空格</option><option value="newlines">去除换行符</option></select></label>
    <div class="button-row"><button class="primary-button" id="runAction">执行清洗</button></div>
  `;
}

function renderMergePanel() {
  return `
    <div class="merge-upload-row">
      ${mergeUploadHTML('merge1', '上传表格 1')}
      ${mergeUploadHTML('merge2', '上传表格 2')}
      ${mergeUploadHTML('merge3', '上传表格 3')}
    </div>
    <div class="merge-configs" id="mergeConfigs"></div>
    <div class="button-row merge-actions">
      <button class="secondary-button" id="previewMerge" type="button">数据预览</button>
      <button class="primary-button" id="runAction">生成合并结果</button>
    </div>
  `;
}

function mergeUploadHTML(key, label) {
  return `
    <label class="file-drop merge-file-drop" data-upload="${key}">
      <span>
        <strong>${label}</strong>
        每个框上传 1 个文件
        <span class="file-name" id="${key}Name">未选择文件</span>
      </span>
      <input type="file" data-input="${key}" accept=".xlsx,.xls,.csv" />
    </label>
  `;
}

function renderDiffPanel() {
  return `
    ${fileDropHTML('old', '旧版数据')}
    ${fileDropHTML('fresh', '新版数据')}
    ${sheetField('oldSheet', '旧版 Sheet')}
    ${sheetField('freshSheet', '新版 Sheet')}
    ${selectField('diffKey', '主键列')}
    <div class="button-row"><button class="primary-button" id="runAction">执行比对</button></div>
  `;
}

function renderMaskPanel() {
  return `
    ${fileDropHTML('main', '上传待脱敏文件', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    ${selectField('maskColumn', '脱敏字段')}
    <label class="field"><span>脱敏类型</span><select id="maskType"><option value="phone">手机号</option><option value="idcard">身份证</option><option value="name">姓名</option><option value="email">邮箱</option><option value="keepFirst">仅保留首字符</option></select></label>
    <div class="button-row"><button class="primary-button" id="runAction">一键脱敏</button></div>
  `;
}

function renderSplitPanel() {
  return `
    ${fileDropHTML('main', '上传待拆分文件', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    <label class="field"><span>拆分方式</span><select id="splitMode"><option value="column">按列值拆分</option><option value="rows">按固定行数拆分</option></select></label>
    ${selectField('splitColumn', '拆分依据列')}
    <label class="field"><span>每个文件最大行数</span><input id="splitRows" type="number" min="1" value="1000" /></label>
    <div class="button-row"><button class="primary-button" id="runAction">拆分并打包</button></div>
  `;
}

function renderFormulaPanel() {
  return `
    ${fileDropHTML('main', '上传待计算文件', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    <label class="field"><span>新列名称</span><input id="formulaName" value="新派生列" /></label>
    <label class="field wide"><span>计算公式</span><input id="formulaExpr" placeholder="例如：[单价] * [数量]，或 [省份] + [城市]" /></label>
    <div class="button-row"><button class="primary-button" id="runAction">生成派生列</button></div>
  `;
}

function renderPivotPanel() {
  return `
    ${fileDropHTML('main', '上传待汇总文件', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    ${selectField('groupColumn', '分组列')}
    ${selectField('valueColumn', '数值列')}
    <label class="field"><span>聚合方式</span><select id="aggType"><option value="sum">求和</option><option value="count">计数</option><option value="avg">平均值</option><option value="max">最大值</option><option value="min">最小值</option></select></label>
    <div class="button-row"><button class="primary-button" id="runAction">生成汇总表</button></div>
  `;
}

function renderComparePanel() {
  return `
    <label class="file-drop compare-upload" data-upload="main">
      <span>
        <strong>上传待对比文件</strong>
        拖拽文件到这里，或 <em>点击选择</em>
        <span class="file-name" id="mainName">未选择文件</span>
      </span>
      <input type="file" data-input="main" accept=".xlsx,.xls,.csv" />
    </label>
    <div class="compare-settings-row">
      <label class="field control-card compare-control"><span>选择 Sheet</span><select id="mainSheet"></select></label>
      <label class="field control-card compare-control"><span>表头所在行</span><input id="compareHeaderRow" type="number" min="1" value="1" /></label>
      <label class="field control-card compare-control"><span>对比方式</span><select id="compareMode"><option value="both">同时输出最大值和最小值</option><option value="max">只输出最大值</option><option value="min">只输出最小值</option></select></label>
      <div class="field control-card compare-control compare-apply">
        <span>表头设置</span>
        <button class="ghost-button" id="applyHeaderRow" type="button">应用表头行</button>
      </div>
    </div>
    <div class="field compare-columns">
      <span>选择参与对比的多列</span>
      <div class="compare-column-grid" id="compareColumns"></div>
    </div>
    <div class="button-row compare-actions">
      <button class="secondary-button" id="previewCompare" type="button">数据预览</button>
      <button class="secondary-button" id="clearCompareColumns" type="button">清除选择列</button>
      <button class="primary-button" id="runAction">生成对比结果</button>
    </div>
  `;
}

function renderChartPanel() {
  return `
    ${fileDropHTML('main', '上传图表数据', '.xlsx,.xls,.csv')}
    ${sheetField('mainSheet', '选择 Sheet')}
    <label class="field"><span>图表类型</span><select id="chartType"><option value="bar">柱状图</option><option value="line">折线图</option><option value="pie">饼图</option><option value="doughnut">环形图</option></select></label>
    ${selectField('xColumn', '分类列')}
    ${selectField('yColumn', '数值列')}
    <div class="button-row"><button class="primary-button" id="runAction">生成图表</button></div>
  `;
}

function sheetField(id, label) {
  return `<label class="field"><span>${label}</span><select id="${id}"></select></label>`;
}

function selectField(id, label) {
  return `<label class="field"><span>${label}</span><select id="${id}"></select></label>`;
}

function multiSelectField(id, label) {
  return `<label class="field wide"><span>${label}</span><select id="${id}" multiple size="8"></select></label>`;
}

function bindUploads() {
  document.querySelectorAll('[data-upload]').forEach(zone => {
    const key = zone.dataset.upload;
    const input = zone.querySelector('input');
    zone.addEventListener('dragover', event => {
      event.preventDefault();
      zone.classList.add('drag');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
    zone.addEventListener('drop', event => {
      event.preventDefault();
      zone.classList.remove('drag');
      if (event.dataTransfer.files.length) handleFiles(key, Array.from(event.dataTransfer.files));
    });
    input.addEventListener('change', () => {
      if (input.files.length) handleFiles(key, Array.from(input.files));
      input.value = '';
    });
  });
}

function bindActions(id) {
  const button = $('runAction');
  if (!button) return;
  const handlers = {
    match: runMatch,
    convert: runConvert,
    clean: runClean,
    merge: runMerge,
    diff: runDiff,
    mask: runMask,
    split: runSplit,
    formula: runFormula,
    pivot: runPivot,
    compare: runCompare,
    chart: runChart,
  };
  button.addEventListener('click', handlers[id]);
  if (id === 'merge') bindMergeControls();
  if (id === 'compare') bindCompareControls();
}

function bindMergeControls() {
  $('previewMerge')?.addEventListener('click', previewMergeResult);
  $('mergeConfigs')?.addEventListener('change', event => {
    if (event.target.matches('[data-merge-sheet]')) {
      updateMergeConfig();
    }
  });
}

function bindCompareControls() {
  $('applyHeaderRow')?.addEventListener('click', () => {
    refreshColumns();
    previewCompareData();
  });
  $('previewCompare')?.addEventListener('click', previewCompareData);
  $('clearCompareColumns')?.addEventListener('click', () => {
    const el = $('compareColumns');
    if (!el) return;
    el.querySelectorAll('input[type="checkbox"]').forEach(input => { input.checked = false; });
    toast('已清除参与对比列', 'success');
  });
  $('compareHeaderRow')?.addEventListener('change', () => {
    refreshColumns();
  });
}

async function handleFiles(key, files) {
  try {
    const loaded = [];
    for (const file of files) {
      loaded.push(await loadFile(file));
    }
    app.files[key] = key === 'many' ? loaded : loaded[0];
    $(key + 'Name').textContent = files.map(file => file.name).join('、');
    updateDependentFields();
    previewCurrentInput();
    updateStatus();
    toast(`已读取 ${files.length} 个文件`, 'success');
  } catch (error) {
    toast(`文件读取失败：${error.message}`, 'error');
  }
}

async function loadFile(file) {
  if (file.name.toLowerCase().endsWith('.json')) {
    const text = await file.text();
    const data = JSON.parse(text);
    const rows = Array.isArray(data) ? data : [data];
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'JSON');
    return { file, workbook: wb };
  }
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
  return { file, workbook };
}

function updateDependentFields() {
  const f = app.files;
  fillSheets('sourceSheet', f.source);
  fillSheets('targetSheet', f.target);
  fillSheets('mainSheet', f.main);
  fillSheets('oldSheet', f.old);
  fillSheets('freshSheet', f.fresh);
  if (app.activeTool === 'merge') updateMergeConfig();
  refreshColumns();
  ['sourceSheet', 'targetSheet', 'mainSheet', 'oldSheet', 'freshSheet'].forEach(id => {
    const el = $(id);
    if (el) el.onchange = () => {
      refreshColumns();
      previewCurrentInput();
    };
  });
}

function fillSheets(id, item) {
  const el = $(id);
  if (!el) return;
  const sheets = item ? item.workbook.SheetNames : [];
  el.innerHTML = sheets.map(name => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`).join('');
}

function refreshColumns() {
  const tool = app.activeTool;
  if (tool === 'match') {
    fillColumns('sourceKey', app.files.source, $('sourceSheet')?.value);
    fillColumns('sourceFill', app.files.source, $('sourceSheet')?.value);
    fillColumns('targetKey', app.files.target, $('targetSheet')?.value);
    fillColumns('targetFill', app.files.target, $('targetSheet')?.value);
  }
  if (tool === 'compare') {
    fillCompareColumns();
    return;
  }
  if (['convert', 'clean', 'mask', 'split', 'formula', 'pivot', 'chart'].includes(tool)) {
    const item = app.files.main;
    const sheet = $('mainSheet')?.value;
    ['maskColumn', 'splitColumn', 'groupColumn', 'valueColumn', 'xColumn', 'yColumn'].forEach(id => fillColumns(id, item, sheet));
  }
  if (tool === 'diff') {
    fillColumns('diffKey', app.files.fresh || app.files.old, $('freshSheet')?.value || $('oldSheet')?.value);
  }
}

function fillColumns(id, item, sheetName) {
  const el = $(id);
  if (!el) return;
  const headers = item && sheetName ? getHeaders(item.workbook, sheetName) : [];
  el.innerHTML = headers.map(name => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`).join('');
}

function fillCompareColumns() {
  const el = $('compareColumns');
  if (!el) return;
  const item = app.files.main;
  const sheet = $('mainSheet')?.value;
  const headers = item && sheet ? getCompareHeaders(item.workbook, sheet) : [];
  el.innerHTML = headers.map((name, index) => `
    <label class="compare-column-option" title="${escapeAttr(name)}">
      <input type="checkbox" value="${escapeAttr(name)}" checked />
      <span>${escapeHtml(name)}</span>
    </label>
  `).join('');
}

function updateMergeConfig() {
  const container = $('mergeConfigs');
  if (!container) return;
  const slots = getMergeSlots();
  container.innerHTML = slots.map(slot => renderMergeSlotConfig(slot)).join('');
}

function getMergeSlots() {
  return ['merge1', 'merge2', 'merge3'].map((key, index) => ({
    key,
    index: index + 1,
    item: app.files[key],
  }));
}

function renderMergeSlotConfig(slot) {
  if (!slot.item) {
    return `
      <section class="merge-slot-card empty">
        <h3>表格 ${slot.index}</h3>
        <p>上传文件后，可在这里选择 Sheet、列和映射列。</p>
      </section>
    `;
  }

  const sheets = slot.item.workbook.SheetNames;
  const selectedSheets = getMergeSelectedSheets(slot.key, sheets);
  const headerOptions = collectMergeHeaders(slot.item, selectedSheets);
  const selectedMap = document.querySelector(`[data-merge-map="${slot.key}"]`)?.value || '';
  const sheetBlocks = sheets.map(sheet => {
    const checked = selectedSheets.includes(sheet);
    const headers = getHeaders(slot.item.workbook, sheet);
    const selectedColumns = getMergeSelectedColumns(slot.key, sheet, headers);
    return `
      <div class="merge-sheet-block">
        <label class="merge-sheet-toggle">
          <input type="checkbox" data-merge-sheet="${slot.key}" value="${escapeAttr(sheet)}" ${checked ? 'checked' : ''} />
          <span>${escapeHtml(sheet)}</span>
        </label>
        ${checked ? `
          <div class="merge-column-grid">
            ${headers.map(header => `
              <label class="merge-column-option" title="${escapeAttr(header)}">
                <input type="checkbox" data-merge-column="${slot.key}" data-sheet="${escapeAttr(sheet)}" value="${escapeAttr(header)}" ${selectedColumns.includes(header) ? 'checked' : ''} />
                <span>${escapeHtml(header)}</span>
              </label>
            `).join('')}
          </div>
        ` : ''}
      </div>
    `;
  }).join('');

  return `
    <section class="merge-slot-card">
      <div class="merge-slot-head">
        <div>
          <h3>表格 ${slot.index}</h3>
          <span>${escapeHtml(slot.item.file.name)}</span>
        </div>
        <label class="merge-map-field">
          <span>映射列（表头内容）</span>
          <select data-merge-map="${slot.key}">
            <option value="">不设置映射列</option>
            ${headerOptions.map(header => `<option value="${escapeAttr(header)}" ${header === selectedMap ? 'selected' : ''}>${escapeHtml(header)}</option>`).join('')}
          </select>
        </label>
      </div>
      <div class="merge-sheet-list">${sheetBlocks}</div>
    </section>
  `;
}

function getMergeSelectedSheets(slotKey, fallbackSheets) {
  const checked = Array.from(document.querySelectorAll(`[data-merge-sheet="${slotKey}"]:checked`)).map(input => input.value);
  return checked.length ? checked : fallbackSheets;
}

function getMergeSelectedColumns(slotKey, sheetName, fallbackColumns) {
  const existing = Array.from(document.querySelectorAll(`[data-merge-column="${slotKey}"]`))
    .filter(input => input.dataset.sheet === sheetName);
  if (!existing.length) return fallbackColumns;
  return existing.filter(input => input.checked).map(input => input.value);
}

function collectMergeHeaders(item, sheetNames) {
  const set = new Set();
  sheetNames.forEach(sheet => getHeaders(item.workbook, sheet).forEach(header => set.add(header)));
  return Array.from(set);
}

function previewCurrentInput() {
  if (app.activeTool === 'merge') {
    previewMergeResult();
    return;
  }
  if (app.activeTool === 'compare') {
    previewCompareData();
    return;
  }
  let item = app.files.main || app.files.target || app.files.fresh || app.files.source || app.files.old;
  if (!item && Array.isArray(app.files.many)) item = app.files.many[0];
  if (!item) return renderTable([], []);
  const sheet = resolveSheet(item, $('mainSheet')?.value || $('targetSheet')?.value || $('freshSheet')?.value || $('sourceSheet')?.value || $('oldSheet')?.value);
  const data = getRows(item.workbook, sheet);
  renderTable(getHeaders(item.workbook, sheet), data);
  $('previewMeta').textContent = `${item.file.name} / ${sheet} / ${data.length} 行`;
  $('activeRowCount').textContent = `${data.length} 行数据`;
}

function previewCompareData() {
  const item = app.files.main;
  const sheet = $('mainSheet')?.value;
  if (!item || !sheet) return renderTable([], []);
  const { headers, rows } = getCompareTable(item.workbook, sheet);
  renderTable(headers, rows);
  $('previewMeta').textContent = `${item.file.name} / ${sheet} / 表头第 ${getCompareHeaderRow()} 行 / ${rows.length} 行`;
  $('activeRowCount').textContent = `${rows.length} 行数据`;
}

function requireFile(key, label) {
  const item = app.files[key];
  if (!item || (Array.isArray(item) && item.length === 0)) {
    toast(`请先上传${label}`, 'error');
    throw new Error('missing file');
  }
  return item;
}

function runMatch() {
  try {
    const source = requireFile('source', '源数据表');
    const target = requireFile('target', '目标数据表');
    const srcSheet = $('sourceSheet').value;
    const tgtSheet = $('targetSheet').value;
    const srcKey = $('sourceKey').value;
    const tgtKey = $('targetKey').value;
    const srcFill = $('sourceFill').value;
    const tgtFill = $('targetFill').value;
    const sourceRows = getRows(source.workbook, srcSheet);
    const targetRows = getRows(target.workbook, tgtSheet).map(row => ({ ...row }));
    const index = new Map(sourceRows.map(row => [String(row[srcKey] ?? '').trim(), row]));
    let matched = 0;
    targetRows.forEach(row => {
      const found = index.get(String(row[tgtKey] ?? '').trim());
      if (found) {
        row[tgtFill] = found[srcFill] ?? '';
        matched += 1;
      }
    });
    setWorkbookResult(targetRows, [...new Set([...getHeaders(target.workbook, tgtSheet), tgtFill])], `${baseName(target.file.name)}_匹配结果.xlsx`);
    toast(`匹配完成：${matched}/${targetRows.length} 行命中`, 'success');
  } catch (_) {}
}

function runConvert() {
  try {
    const item = requireFile('main', '待转换文件');
    const sheet = $('mainSheet').value;
    const format = $('targetFormat').value;
    const ws = item.workbook.Sheets[sheet];
    if (format === 'csv') {
      setRawResult('\ufeff' + XLSX.utils.sheet_to_csv(ws), `${baseName(item.file.name)}.csv`, 'csv');
    } else if (format === 'json') {
      setRawResult(JSON.stringify(XLSX.utils.sheet_to_json(ws, { defval: '' }), null, 2), `${baseName(item.file.name)}.json`, 'json');
    } else {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheet);
      setWorkbookObjectResult(wb, `${baseName(item.file.name)}.xlsx`);
    }
    renderTable(getHeaders(item.workbook, sheet), getRows(item.workbook, sheet));
    toast('转换结果已准备好', 'success');
  } catch (_) {}
}

function runClean() {
  try {
    const item = requireFile('main', '待清洗文件');
    const sheet = $('mainSheet').value;
    const headers = getHeaders(item.workbook, sheet);
    const rules = Array.from($('cleanRules').selectedOptions).map(opt => opt.value);
    let rows = getRows(item.workbook, sheet).map(row => {
      const next = {};
      headers.forEach(header => {
        let value = row[header];
        if (typeof value === 'string') {
          if (rules.includes('trim')) value = value.trim();
          if (rules.includes('spaces')) value = value.replace(/\s+/g, ' ');
          if (rules.includes('newlines')) value = value.replace(/[\r\n]+/g, ' ');
        }
        next[header] = value ?? '';
      });
      return next;
    });
    if (rules.includes('blankRows')) rows = rows.filter(row => headers.some(header => String(row[header] ?? '').trim() !== ''));
    if (rules.includes('duplicateRows')) {
      const seen = new Set();
      rows = rows.filter(row => {
        const key = JSON.stringify(headers.map(header => row[header]));
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }
    setWorkbookResult(rows, headers, `${baseName(item.file.name)}_清洗结果.xlsx`);
    toast(`清洗完成：剩余 ${rows.length} 行`, 'success');
  } catch (_) {}
}

async function runMerge() {
  try {
    const rows = buildMergeRows();
    if (!rows.length) {
      toast('请至少上传 1 个文件，并选择需要合并的 Sheet 和列', 'error');
      return;
    }
    const headers = collectHeaders(rows);
    setWorkbookResult(rows, headers, '报表合并结果.xlsx');
    toast(`合并完成：${rows.length} 行`, 'success');
  } catch (_) {}
}

function previewMergeResult() {
  try {
    const rows = buildMergeRows();
    if (!rows.length) {
      toast('暂无可预览数据，请先上传并选择 Sheet / 列', 'error');
      return;
    }
    const headers = collectHeaders(rows);
    renderTable(headers, rows);
    $('previewMeta').textContent = `最终合并表 / ${rows.length} 行 / ${headers.length} 列`;
    $('activeRowCount').textContent = `${rows.length} 行数据`;
  } catch (_) {}
}

function buildMergeRows() {
  const activeSlots = getMergeSlots().filter(slot => slot.item);
  if (!activeSlots.length) return [];

  const slotConfigs = [];
  for (const slot of activeSlots) {
    const mapColumn = document.querySelector(`[data-merge-map="${slot.key}"]`)?.value || '';
    if (!mapColumn) {
      toast(`请先设置表格 ${slot.index} 的映射列`, 'error');
      return [];
    }

    const selectedSheets = Array.from(document.querySelectorAll(`[data-merge-sheet="${slot.key}"]:checked`)).map(input => input.value);
    const sheetConfigs = selectedSheets.map(sheet => {
      const selectedColumns = Array.from(document.querySelectorAll(`[data-merge-column="${slot.key}"]:checked`))
        .filter(input => input.dataset.sheet === sheet)
        .map(input => input.value);
      return { sheet, selectedColumns };
    }).filter(config => config.selectedColumns.length > 0);

    if (!sheetConfigs.length) continue;

    const outputColumns = Array.from(new Set(sheetConfigs.flatMap(config => config.selectedColumns)));
    slotConfigs.push({ ...slot, mapColumn, sheetConfigs, outputColumns });
  }

  if (!slotConfigs.length) return [];

  const groups = new Map();
  slotConfigs.forEach(config => {
    config.sheetConfigs.forEach(sheetConfig => {
      getRows(config.item.workbook, sheetConfig.sheet).forEach(row => {
        const mapValue = String(row[config.mapColumn] ?? '').trim();
        if (!mapValue) return;
        if (!groups.has(mapValue)) groups.set(mapValue, {});
        const group = groups.get(mapValue);
        if (!group[config.key]) group[config.key] = [];
        group[config.key].push({ row, sheet: sheetConfig.sheet, columns: sheetConfig.selectedColumns });
      });
    });
  });

  const output = [];
  groups.forEach((slotRecords, mapValue) => {
    const maxRows = Math.max(1, ...slotConfigs.map(config => (slotRecords[config.key] || []).length));
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
      const next = { 映射值: mapValue };
      slotConfigs.forEach(config => {
        const records = slotRecords[config.key] || [];
        const record = records[rowIndex] || records[records.length - 1];
        config.outputColumns.forEach(column => {
          const value = record && record.columns.includes(column) ? record.row[column] : '';
          if (!(column in next) || isBlankCell(next[column])) {
            next[column] = isBlankCell(value) ? mapValue : value;
          }
        });
      });
      output.push(next);
    }
  });

  return output;
}

function runDiff() {
  try {
    const oldItem = requireFile('old', '旧版数据');
    const freshItem = requireFile('fresh', '新版数据');
    const oldSheet = $('oldSheet').value;
    const freshSheet = $('freshSheet').value;
    const key = $('diffKey').value;
    const oldRows = getRows(oldItem.workbook, oldSheet);
    const freshRows = getRows(freshItem.workbook, freshSheet);
    const headers = [...new Set([...getHeaders(oldItem.workbook, oldSheet), ...getHeaders(freshItem.workbook, freshSheet)])];
    const oldIndex = new Map(oldRows.map(row => [String(row[key] ?? ''), row]));
    const freshIndex = new Map(freshRows.map(row => [String(row[key] ?? ''), row]));
    const result = [];
    freshRows.forEach(row => {
      const oldRow = oldIndex.get(String(row[key] ?? ''));
      if (!oldRow) {
        result.push({ 差异类型: '新增', ...row });
        return;
      }
      const changed = headers.some(header => String(row[header] ?? '') !== String(oldRow[header] ?? ''));
      if (changed) result.push({ 差异类型: '修改', ...row });
    });
    oldRows.forEach(row => {
      if (!freshIndex.has(String(row[key] ?? ''))) result.push({ 差异类型: '删除', ...row });
    });
    setWorkbookResult(result, ['差异类型', ...headers], '差异报告.xlsx');
    renderTable(['差异类型', ...headers], result, row => {
      if (row.差异类型 === '新增') return 'row-added';
      if (row.差异类型 === '删除') return 'row-removed';
      if (row.差异类型 === '修改') return 'row-changed';
      return '';
    });
    toast(`比对完成：发现 ${result.length} 条差异`, 'success');
  } catch (_) {}
}

function runMask() {
  try {
    const item = requireFile('main', '待脱敏文件');
    const sheet = $('mainSheet').value;
    const col = $('maskColumn').value;
    const type = $('maskType').value;
    const headers = getHeaders(item.workbook, sheet);
    const rows = getRows(item.workbook, sheet).map(row => ({ ...row, [col]: maskValue(row[col], type) }));
    setWorkbookResult(rows, headers, `${baseName(item.file.name)}_脱敏结果.xlsx`);
    toast('脱敏完成', 'success');
  } catch (_) {}
}

async function runSplit() {
  try {
    const item = requireFile('main', '待拆分文件');
    if (typeof JSZip === 'undefined') {
      toast('JSZip 加载失败，无法打包', 'error');
      return;
    }
    const sheet = $('mainSheet').value;
    const headers = getHeaders(item.workbook, sheet);
    const rows = getRows(item.workbook, sheet);
    const zip = new JSZip();
    const chunks = {};
    if ($('splitMode').value === 'column') {
      const col = $('splitColumn').value;
      rows.forEach(row => {
        const key = sanitizeFileName(String(row[col] || '空值'));
        if (!chunks[key]) chunks[key] = [];
        chunks[key].push(row);
      });
    } else {
      const size = Math.max(1, Number($('splitRows').value) || 1000);
      rows.forEach((row, index) => {
        const key = `第${Math.floor(index / size) + 1}批`;
        if (!chunks[key]) chunks[key] = [];
        chunks[key].push(row);
      });
    }
    Object.entries(chunks).forEach(([name, chunkRows]) => {
      const wb = rowsToWorkbook(chunkRows, headers, '数据');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      zip.file(`${name}.xlsx`, out);
    });
    const blob = await zip.generateAsync({ type: 'blob' });
    app.result = blob;
    app.resultName = `${baseName(item.file.name)}_拆分结果.zip`;
    app.resultType = 'zip';
    $('downloadResult').disabled = false;
    renderTable(headers, rows);
    $('previewMeta').textContent = `已拆分为 ${Object.keys(chunks).length} 个文件`;
    toast(`拆分完成：${Object.keys(chunks).length} 个文件`, 'success');
  } catch (_) {}
}

function runFormula() {
  try {
    const item = requireFile('main', '待计算文件');
    const sheet = $('mainSheet').value;
    const headers = getHeaders(item.workbook, sheet);
    const name = $('formulaName').value.trim() || '新派生列';
    const expr = $('formulaExpr').value.trim();
    if (!expr) {
      toast('请输入计算公式', 'error');
      return;
    }
    const rows = getRows(item.workbook, sheet).map(row => {
      const next = { ...row };
      next[name] = evaluateFormula(row, headers, expr);
      return next;
    });
    setWorkbookResult(rows, [...new Set([...headers, name])], `${baseName(item.file.name)}_派生列.xlsx`);
    toast('派生列生成完成', 'success');
  } catch (_) {}
}

function runPivot() {
  try {
    const item = requireFile('main', '待汇总文件');
    const sheet = $('mainSheet').value;
    const groupCol = $('groupColumn').value;
    const valueCol = $('valueColumn').value;
    const agg = $('aggType').value;
    const groups = {};
    getRows(item.workbook, sheet).forEach(row => {
      const key = String(row[groupCol] || '空值').trim();
      if (!groups[key]) groups[key] = [];
      groups[key].push(row);
    });
    const metric = `${valueCol}_${aggLabel(agg)}`;
    const rows = Object.entries(groups).map(([key, groupRows]) => {
      const nums = groupRows.map(row => Number(row[valueCol])).filter(num => !Number.isNaN(num));
      return { [groupCol]: key, [metric]: aggregate(nums, groupRows.length, agg) };
    }).sort((a, b) => String(a[groupCol]).localeCompare(String(b[groupCol]), 'zh-CN'));
    setWorkbookResult(rows, [groupCol, metric], `${baseName(item.file.name)}_分组汇总.xlsx`);
    toast(`汇总完成：${rows.length} 个分组`, 'success');
  } catch (_) {}
}

function runCompare() {
  try {
    const item = requireFile('main', '待对比文件');
    const sheet = $('mainSheet').value;
    const { headers, rows: sourceRows } = getCompareTable(item.workbook, sheet);
    const selectedColumns = getSelectedValues('compareColumns');
    const mode = $('compareMode').value;
    if (selectedColumns.length < 2) {
      toast('请至少选择 2 列参与对比', 'error');
      return;
    }

    const rows = sourceRows.map(row => {
      const values = selectedColumns
        .map(column => ({ column, value: normalizeNumber(row[column]) }))
        .filter(item => item.value !== null);
      const maxItem = values.reduce((best, current) => !best || current.value > best.value ? current : best, null);
      const minItem = values.reduce((best, current) => !best || current.value < best.value ? current : best, null);
      return {
        ...row,
        ...(mode !== 'min' ? { 最大值: maxItem ? maxItem.value : '', 最大值来源列: maxItem ? maxItem.column : '' } : {}),
        ...(mode !== 'max' ? { 最小值: minItem ? minItem.value : '', 最小值来源列: minItem ? minItem.column : '' } : {}),
        参与对比列数: values.length,
      };
    });
    const resultHeaders = [
      ...headers,
      ...(mode !== 'min' ? ['最大值', '最大值来源列'] : []),
      ...(mode !== 'max' ? ['最小值', '最小值来源列'] : []),
      '参与对比列数'
    ];
    setWorkbookResult(rows, resultHeaders, `${baseName(item.file.name)}_多列对比结果.xlsx`);
    toast(`按行对比完成：${rows.length} 行，${selectedColumns.length} 列`, 'success');
  } catch (_) {}
}

function runChart() {
  try {
    const item = requireFile('main', '图表数据');
    if (typeof Chart === 'undefined') {
      toast('Chart.js 加载失败，无法生成图表', 'error');
      return;
    }
    const sheet = $('mainSheet').value;
    const xCol = $('xColumn').value;
    const yCol = $('yColumn').value;
    const chartType = $('chartType').value;
    const summary = {};
    getRows(item.workbook, sheet).forEach(row => {
      const label = String(row[xCol] || '未分类').trim();
      const value = Number(row[yCol]);
      summary[label] = (summary[label] || 0) + (Number.isNaN(value) ? 0 : value);
    });
    const top = Object.entries(summary)
      .map(([label, value]) => ({ label, value: Number(value.toFixed(2)) }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 30);
    if (app.chart) app.chart.destroy();
    app.chart = new Chart($('chartCanvas'), {
      type: chartType,
      data: {
        labels: top.map(item => item.label),
        datasets: [{
          label: `${yCol} 汇总`,
          data: top.map(item => item.value),
          backgroundColor: chartType === 'bar' || chartType === 'line' ? 'rgba(37, 99, 235, .72)' : palette(top.length),
          borderColor: chartType === 'line' ? '#2563eb' : '#fff',
          borderWidth: 2,
          tension: .25,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: chartType === 'pie' || chartType === 'doughnut' } },
      },
    });
    $('downloadChart').disabled = false;
    renderTable([xCol, yCol], top.map(item => ({ [xCol]: item.label, [yCol]: item.value })));
    $('previewMeta').textContent = `图表数据：${top.length} 个分类`;
    toast('图表已生成', 'success');
  } catch (_) {}
}

function setWorkbookResult(rows, headers, filename) {
  setWorkbookObjectResult(rowsToWorkbook(rows, headers, '结果'), filename);
  renderTable(headers, rows);
  $('previewMeta').textContent = `${rows.length} 行 / ${headers.length} 列`;
  $('activeRowCount').textContent = `${rows.length} 行数据`;
}

function setWorkbookObjectResult(wb, filename) {
  app.result = wb;
  app.resultName = filename;
  app.resultType = 'xlsx';
  $('downloadResult').disabled = false;
}

function setRawResult(content, filename, type) {
  app.result = content;
  app.resultName = filename;
  app.resultType = type;
  $('downloadResult').disabled = false;
}

function clearResult() {
  app.result = null;
  app.resultName = '处理结果.xlsx';
  app.resultType = 'xlsx';
  $('downloadResult').disabled = true;
  $('downloadChart').disabled = true;
  $('previewTable').innerHTML = '';
  $('previewMeta').textContent = '等待上传文件';
  if (app.chart) {
    app.chart.destroy();
    app.chart = null;
  }
}

function downloadResult() {
  if (!app.result) return;
  if (app.resultType === 'xlsx') {
    XLSX.writeFile(app.result, app.resultName);
    return;
  }
  const blob = app.result instanceof Blob
    ? app.result
    : new Blob([app.result], { type: app.resultType === 'json' ? 'application/json' : 'text/plain;charset=utf-8' });
  downloadBlob(blob, app.resultName);
}

function downloadChart() {
  const canvas = $('chartCanvas');
  downloadBlob(dataURLToBlob(canvas.toDataURL('image/png')), '数据图表.png');
}

function renderTable(headers, rows, rowClassFn) {
  const visibleRows = rows.slice(0, 300);
  if (!headers.length) {
    $('previewTable').innerHTML = '<tbody><tr><td>暂无数据</td></tr></tbody>';
    return;
  }
  const head = `<thead><tr><th>#</th>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>`;
  const body = visibleRows.map((row, index) => `
    <tr class="${rowClassFn ? rowClassFn(row) : ''}">
      <td>${index + 1}</td>
      ${headers.map(header => `<td title="${escapeAttr(String(row[header] ?? ''))}">${escapeHtml(String(row[header] ?? ''))}</td>`).join('')}
    </tr>
  `).join('');
  $('previewTable').innerHTML = `${head}<tbody>${body || '<tr><td colspan="99">暂无数据</td></tr>'}</tbody>`;
}

function getRows(wb, sheetName) {
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
}

function getCompareHeaderRow() {
  const value = Number($('compareHeaderRow')?.value || 1);
  return Math.max(1, Math.floor(value || 1));
}

function getCompareTable(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const headerIndex = Math.min(getCompareHeaderRow() - 1, Math.max(0, rawRows.length - 1));
  const headers = normalizeHeaders(rawRows[headerIndex] || []);
  const rows = rawRows.slice(headerIndex + 1).map(raw => {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = raw[index] ?? '';
    });
    return row;
  });
  return { headers, rows };
}

function getCompareHeaders(wb, sheetName) {
  return getCompareTable(wb, sheetName).headers;
}

function getHeaders(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
    headers.push(cell ? String(cell.v) : `列${c + 1}`);
  }
  return headers;
}

function normalizeHeaders(rawHeaders) {
  const used = new Map();
  return rawHeaders.map((value, index) => {
    const base = String(value ?? '').trim() || `列${index + 1}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function rowsToWorkbook(rows, headers, sheetName) {
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}

function resolveSheet(item, sheet) {
  return sheet && item.workbook.SheetNames.includes(sheet) ? sheet : item.workbook.SheetNames[0];
}

function updateStatus() {
  const count = Object.values(app.files).reduce((sum, item) => sum + (Array.isArray(item) ? item.length : item ? 1 : 0), 0);
  $('activeFileCount').textContent = `${count} 个文件`;
  if (!count) $('activeRowCount').textContent = '0 行数据';
}

function collectHeaders(rows) {
  const set = new Set();
  rows.forEach(row => Object.keys(row).forEach(key => set.add(key)));
  return Array.from(set);
}

function getSelectedValues(id) {
  const el = $(id);
  if (!el) return [];
  if (!('selectedOptions' in el)) {
    return Array.from(el.querySelectorAll('input[type="checkbox"]:checked')).map(input => input.value);
  }
  return Array.from(el.selectedOptions).map(option => option.value);
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Number.isNaN(value) ? null : value;
  const cleaned = String(value).replace(/,/g, '').trim();
  if (!cleaned) return null;
  const number = Number(cleaned);
  return Number.isNaN(number) ? null : number;
}

function isBlankCell(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function evaluateFormula(row, headers, expr) {
  let code = expr;
  headers.forEach(header => {
    const value = row[header] ?? '';
    const replacement = typeof value === 'number' ? String(value) : JSON.stringify(String(value));
    code = code.split(`[${header}]`).join(replacement);
  });
  try {
    return Function(`"use strict"; return (${code});`)();
  } catch (_) {
    return '计算出错';
  }
}

function maskValue(value, type) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (type === 'phone') return text.replace(/(\d{3})\d{4}(\d{4})/, '$1****$2');
  if (type === 'idcard') return text.length >= 15 ? `${text.slice(0, 6)}********${text.slice(-4)}` : '****';
  if (type === 'name') return text.length <= 1 ? '*' : `${text[0]}${'*'.repeat(Math.max(1, text.length - 1))}`;
  if (type === 'email') return text.replace(/^(.).+(@.+)$/, '$1***$2');
  return `${text[0]}${'*'.repeat(Math.max(1, text.length - 1))}`;
}

function aggregate(nums, rowCount, agg) {
  if (agg === 'count') return rowCount;
  if (!nums.length) return 0;
  if (agg === 'sum') return round(nums.reduce((a, b) => a + b, 0));
  if (agg === 'avg') return round(nums.reduce((a, b) => a + b, 0) / nums.length);
  if (agg === 'max') return Math.max(...nums);
  if (agg === 'min') return Math.min(...nums);
  return 0;
}

function aggLabel(agg) {
  return { sum: '求和', count: '计数', avg: '平均值', max: '最大值', min: '最小值' }[agg];
}

function palette(count) {
  const colors = ['#2563eb', '#0f9f6e', '#d97706', '#0891b2', '#7c3aed', '#db2777', '#dc2626', '#475569'];
  return Array.from({ length: count }, (_, index) => colors[index % colors.length]);
}

function round(value) {
  return Number(value.toFixed(2));
}

function baseName(name) {
  return name.replace(/\.[^.]+$/, '');
}

function sanitizeFileName(name) {
  return name.trim().replace(/[\\/:*?"<>|\s]+/g, '_') || '空值';
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

function dataURLToBlob(dataURL) {
  const [meta, data] = dataURL.split(',');
  const mime = meta.match(/:(.*?);/)[1];
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function toast(message, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = message;
  $('toastArea').appendChild(el);
  setTimeout(() => el.remove(), 3200);
}
