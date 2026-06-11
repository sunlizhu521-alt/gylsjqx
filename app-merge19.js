const tools = [
  { id: 'merge', icon: '▦', title: '报表合并', status: '可用-陈燕', desc: '多个文件按映射列合并，保留明细行。' },
  { id: 'compare', icon: '⌁', title: '多列对比', status: '可用-陈燕', desc: '选择多列数据，按行找最大值、最小值。' },
  { id: 'seal', icon: '章', title: '检验报告单-电子章', status: '可用-品质', desc: '检验报告单图片旋正后加盖电子章，并一键保存图片。' },
  { id: 'ocr', icon: '文', title: '图片转文字', status: '仅限简单图片', desc: '上传图片后识别文字，并导出 TXT、Excel、Word。' },
  { id: 'match', icon: '⇄', title: '智能匹配', status: '不可用仅是举例', desc: '根据关键列把源表字段填入目标表。' },
  { id: 'convert', icon: '↔', title: '格式转换', status: '不可用仅是举例', desc: 'Excel、CSV、JSON 互相转换。' },
  { id: 'clean', icon: '✓', title: '数据清洗', status: '不可用仅是举例', desc: '去空格、去空行、去重、规整文本。' },
  { id: 'diff', icon: '≠', title: '数据比对', status: '不可用仅是举例', desc: '按主键识别新增、删除、修改。' },
  { id: 'mask', icon: '◩', title: '敏感脱敏', status: '不可用仅是举例', desc: '手机号、身份证、姓名、邮箱快速遮蔽。' },
  { id: 'split', icon: '÷', title: '数据拆分', status: '不可用仅是举例', desc: '按列值或固定行数拆成多个文件。' },
  { id: 'formula', icon: 'ƒ', title: '派生列', status: '不可用仅是举例', desc: '用公式生成新字段，例如 [单价] * [数量]。' },
  { id: 'pivot', icon: '∑', title: '分组汇总', status: '不可用仅是举例', desc: '按维度做求和、计数、平均、最大、最小。' },
  { id: 'chart', icon: '▥', title: '图表生成', status: '不可用仅是举例', desc: '基于字段自动聚合并绘制图表。' },
];

const DEFAULT_SEAL_IMAGE = 'assets/quality-seal.png';

const app = {
  activeTool: '',
  toolStates: {},
  files: {},
  result: null,
  resultName: '处理结果.xlsx',
  resultType: 'xlsx',
  chart: null,
  ocrText: '',
  ocrImageUrl: '',
  ocrRotation: 0,
  sealRotation: 0,
  sealResultUrl: '',
};

const $ = (id) => document.getElementById(id);

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  renderToolList();
  switchTool('merge');
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
      <span class="tab-text">${tool.title}（${tool.status}）</span>
    </button>
  `).join('');
  document.querySelectorAll('.tool-tab').forEach(button => {
    button.addEventListener('click', () => {
      if (button.dataset.tool === app.activeTool) return;
      switchTool(button.dataset.tool);
    });
  });
}

function switchTool(id) {
  if (id === app.activeTool) return;
  saveToolState(app.activeTool);
  restoreToolState(id);
  app.activeTool = id;
  const tool = tools.find(item => item.id === id);
  $('toolTitle').textContent = tool.title;
  $('toolDesc').textContent = tool.desc;
  document.querySelectorAll('.tool-tab').forEach(btn => btn.classList.toggle('active', btn.dataset.tool === id));
  $('chartCard').classList.toggle('hidden', id !== 'chart');
  document.querySelector('.result-layout').classList.toggle('single', id !== 'chart');
  renderToolPanel(id);
  updateDependentFields();
  updateFileLabels();
  previewCurrentInput();
  updateStatus();
}

function saveToolState(id) {
  if (!id) return;
  app.toolStates[id] = {
    files: { ...app.files },
    result: app.result,
    resultName: app.resultName,
    resultType: app.resultType,
    ocrText: app.ocrText,
    ocrImageUrl: app.ocrImageUrl,
    ocrRotation: app.ocrRotation,
    sealRotation: app.sealRotation,
    sealResultUrl: app.sealResultUrl,
  };
}

function restoreToolState(id) {
  const state = app.toolStates[id] || {};
  app.files = { ...(state.files || {}) };
  app.result = state.result || null;
  app.resultName = state.resultName || '处理结果.xlsx';
  app.resultType = state.resultType || 'xlsx';
  app.ocrText = state.ocrText || '';
  app.ocrImageUrl = state.ocrImageUrl || '';
  app.ocrRotation = state.ocrRotation || 0;
  app.sealRotation = state.sealRotation || 0;
  app.sealResultUrl = state.sealResultUrl || '';
}

function updateFileLabels() {
  Object.entries(app.files).forEach(([key, item]) => {
    const label = $(key + 'Name');
    if (!label) return;
    if (Array.isArray(item)) label.textContent = item.map(entry => entry.file.name).join('、');
    else label.textContent = item?.file?.name || '未选择文件';
  });
  $('downloadResult').disabled = !app.result;
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
    seal: renderSealPanel,
    ocr: renderOcrPanel,
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
    <div class="merge-upload-card">
      <label class="file-drop merge-file-drop" data-upload="${key}">
        <span>
        <strong>${label}</strong>
        每个框上传 1 个文件
        <span class="file-name" id="${key}Name">未选择文件</span>
        </span>
        <input type="file" data-input="${key}" accept=".xlsx,.xls,.csv" />
      </label>
      <button class="ghost-button merge-clear-button" data-clear-merge="${key}" type="button">删除后重新上传</button>
    </div>
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

function renderSealPanel() {
  return `
    ${fileDropHTML('sealReport', '上传检验报告单图片', '.png,.jpg,.jpeg,.webp,.bmp')}
    <div class="ocr-preview-card seal-preview-card">
      <div class="ocr-image-panel">
        <div class="ocr-image-stage">
          <img id="sealImagePreview" alt="检验报告单预览" />
        </div>
        <div class="ocr-rotate-row">
          <button class="ghost-button" id="rotateSealLeft" type="button">左转 90°</button>
          <button class="ghost-button" id="rotateSealRight" type="button">右转 90°</button>
          <button class="ghost-button" id="resetSealRotation" type="button">重置旋转</button>
          <button class="ghost-button" id="clearSealImages" type="button">清空图片</button>
        </div>
      </div>
      <div class="seal-summary">
        <h3>检验报告单-电子章</h3>
        <p>把刚才常用的两个动作合并成一个工具：先旋正检验报告单，再在右下角加盖电子质检章。</p>
        <ul>
          <li>支持左转、右转、重置旋转，确保文字是正的。</li>
          <li>电子章已内置，不需要重复上传；按报告宽度约 18% 缩放后放在右下角。</li>
          <li>所有处理都在浏览器本地完成，不上传服务器。</li>
          <li>点击“一键保存到桌面”会下载盖章后的 JPG；浏览器下载目录设为桌面时会直接保存到桌面。</li>
        </ul>
      </div>
    </div>
    <div class="button-row ocr-actions">
      <button class="primary-button" id="runAction">旋正并加章</button>
      <button class="secondary-button" id="saveSealDesktop" type="button">一键保存到桌面</button>
    </div>
  `;
}

function renderOcrPanel() {
  return `
    ${fileDropHTML('image', '上传待识别图片', '.png,.jpg,.jpeg,.webp,.bmp')}
    <label class="field control-card"><span>识别语言</span><select id="ocrLang"><option value="chi_sim+eng">中文 + 英文</option><option value="chi_sim">中文</option><option value="eng">英文</option></select></label>
    <label class="field control-card"><span>识别模式</span><select id="ocrMode"><option value="document" selected>开票资料 / 清晰文档</option><option value="table">表格 / 多列内容</option><option value="strong">强增强 / 模糊图片</option><option value="off">原图识别</option></select></label>
    <div class="ocr-preview-card">
      <div class="ocr-image-panel">
        <div class="ocr-image-stage">
          <img id="ocrImagePreview" alt="图片预览" />
        </div>
        <div class="ocr-rotate-row">
          <button class="ghost-button" id="rotateOcrLeft" type="button">左转 90°</button>
          <button class="ghost-button" id="rotateOcrRight" type="button">右转 90°</button>
          <button class="ghost-button" id="resetOcrRotation" type="button">重置旋转</button>
          <button class="ghost-button" id="clearOcrImage" type="button">删除图片</button>
        </div>
      </div>
      <textarea id="ocrTextOutput" placeholder="识别出的文字会显示在这里"></textarea>
    </div>
    <div class="button-row ocr-actions">
      <button class="primary-button" id="runAction">识别文字</button>
      <button class="secondary-button" id="copyOcrText" type="button">复制文字</button>
      <button class="secondary-button" id="downloadOcrTxt" type="button">下载 TXT</button>
      <button class="secondary-button" id="exportOcrWord" type="button">转 Word</button>
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
    seal: runSealStamp,
    ocr: runOcr,
    chart: runChart,
  };
  button.addEventListener('click', handlers[id]);
  if (id === 'merge') bindMergeControls();
  if (id === 'compare') bindCompareControls();
  if (id === 'seal') bindSealControls();
  if (id === 'ocr') bindOcrControls();
}

function bindMergeControls() {
  $('previewMerge')?.addEventListener('click', previewMergeResult);
  document.querySelectorAll('[data-clear-merge]').forEach(button => {
    button.addEventListener('click', () => clearMergeFile(button.dataset.clearMerge));
  });
  $('mergeConfigs')?.addEventListener('change', event => {
    if (event.target.matches('[data-merge-sheet]')) {
      updateMergeConfig();
    }
  });
}

function clearMergeFile(key) {
  if (!app.files[key]) {
    toast('当前上传框还没有文件', 'info');
    return;
  }
  delete app.files[key];
  const name = $(key + 'Name');
  if (name) name.textContent = '未选择文件';
  updateMergeConfig();
  clearResult();
  updateStatus();
  toast('已删除该表格，可重新上传', 'success');
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

function bindOcrControls() {
  $('copyOcrText')?.addEventListener('click', copyOcrText);
  $('downloadOcrTxt')?.addEventListener('click', downloadOcrTxt);
  $('exportOcrWord')?.addEventListener('click', exportOcrWord);
  $('rotateOcrLeft')?.addEventListener('click', () => rotateOcrImage(-90));
  $('rotateOcrRight')?.addEventListener('click', () => rotateOcrImage(90));
  $('resetOcrRotation')?.addEventListener('click', () => resetOcrRotation());
  $('clearOcrImage')?.addEventListener('click', clearOcrImage);
  $('ocrTextOutput')?.addEventListener('input', event => {
    app.ocrText = event.target.value;
    renderOcrTextPreview();
  });
}

function bindSealControls() {
  $('rotateSealLeft')?.addEventListener('click', () => rotateSealImage(-90));
  $('rotateSealRight')?.addEventListener('click', () => rotateSealImage(90));
  $('resetSealRotation')?.addEventListener('click', resetSealRotation);
  $('clearSealImages')?.addEventListener('click', clearSealImages);
  $('saveSealDesktop')?.addEventListener('click', saveSealToDesktop);
}

async function handleFiles(key, files) {
  try {
    if (key === 'sealReport') {
      app.files[key] = await loadToolImageFile(files[0]);
      $(key + 'Name').textContent = files[0].name;
      if (key === 'sealReport') app.sealRotation = 0;
      clearResult();
      previewCurrentInput();
      updateStatus();
      toast('已读取图片', 'success');
      return;
    }
    if (key === 'image') {
      app.files[key] = await loadImageFile(files[0]);
      $(key + 'Name').textContent = files[0].name;
      previewCurrentInput();
      updateStatus();
      toast('已读取图片', 'success');
      return;
    }
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

async function loadImageFile(file) {
  if (!file || !file.type.startsWith('image/')) {
    throw new Error('请选择图片文件');
  }
  if (app.ocrImageUrl) URL.revokeObjectURL(app.ocrImageUrl);
  app.ocrImageUrl = URL.createObjectURL(file);
  app.ocrText = '';
  app.ocrRotation = 0;
  return { file, url: app.ocrImageUrl };
}

async function loadToolImageFile(file) {
  if (!file || !file.type.startsWith('image/')) {
    throw new Error('请选择图片文件');
  }
  return { file, url: URL.createObjectURL(file) };
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
  if (app.activeTool === 'seal') {
    previewSealImage();
    return;
  }
  if (app.activeTool === 'ocr') {
    previewOcrImage();
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

function previewSealImage() {
  const item = app.files.sealReport;
  const img = $('sealImagePreview');
  if (img) {
    img.src = app.sealResultUrl || item?.url || '';
    img.style.transform = app.sealResultUrl ? 'none' : `rotate(${app.sealRotation}deg)`;
    img.classList.toggle('hidden', !(app.sealResultUrl || item?.url));
  }
  const rows = [
    { 项目: '报告图片', 状态: item ? item.file.name : '未上传' },
    { 项目: '电子章', 状态: '已内置质检合格章' },
    { 项目: '旋转角度', 状态: `${app.sealRotation}°` },
    { 项目: '保存方式', 状态: '生成 JPG 后一键下载；浏览器下载目录设为桌面即可保存到桌面' },
  ];
  renderTable(['项目', '状态'], rows, null, { showIndex: false });
  $('previewMeta').textContent = app.result ? '已生成盖章图片' : (item ? '等待加章' : '等待上传检验报告单');
  $('activeRowCount').textContent = `${rows.length} 项`;
}

function rotateSealImage(degrees) {
  if (!app.files.sealReport) {
    toast('请先上传检验报告单图片', 'error');
    return;
  }
  app.sealRotation = normalizeRotation(app.sealRotation + degrees);
  clearSealResultOnly();
  previewSealImage();
}

function resetSealRotation() {
  if (!app.files.sealReport) {
    toast('请先上传检验报告单图片', 'error');
    return;
  }
  app.sealRotation = 0;
  clearSealResultOnly();
  previewSealImage();
}

function clearSealImages() {
  ['sealReport'].forEach(key => {
    if (app.files[key]?.url) URL.revokeObjectURL(app.files[key].url);
    delete app.files[key];
    const label = $(key + 'Name');
    if (label) label.textContent = '未选择文件';
  });
  app.sealRotation = 0;
  clearSealResultOnly();
  clearResult();
  previewSealImage();
  updateStatus();
  toast('已清空图片，可重新上传', 'success');
}

function clearSealResultOnly() {
  if (app.sealResultUrl) URL.revokeObjectURL(app.sealResultUrl);
  app.sealResultUrl = '';
  app.result = null;
  app.resultName = '处理结果.xlsx';
  app.resultType = 'xlsx';
  $('downloadResult').disabled = true;
}

function previewOcrImage() {
  const item = app.files.image;
  const img = $('ocrImagePreview');
  const output = $('ocrTextOutput');
  if (img) {
    img.src = item?.url || '';
    img.style.transform = `rotate(${app.ocrRotation}deg)`;
    img.classList.toggle('hidden', !item?.url);
  }
  if (output) output.value = app.ocrText || '';
  renderOcrTextPreview();
  $('previewMeta').textContent = item ? `${item.file.name} / 等待识别` : '等待上传图片';
}

function renderOcrTextPreview() {
  const lines = getOcrLines().map(text => ({ 识别文字: text }));
  renderTable(['识别文字'], lines, null, { showIndex: false });
  $('previewMeta').textContent = app.ocrText ? `${lines.length} 行文字` : '等待识别文字';
  $('activeRowCount').textContent = `${lines.length} 行文字`;
}

function rotateOcrImage(degrees) {
  if (!app.files.image) {
    toast('请先上传图片', 'error');
    return;
  }
  app.ocrRotation = normalizeRotation(app.ocrRotation + degrees);
  previewOcrImage();
}

function resetOcrRotation() {
  if (!app.files.image) {
    toast('请先上传图片', 'error');
    return;
  }
  app.ocrRotation = 0;
  previewOcrImage();
}

function clearOcrImage() {
  if (app.ocrImageUrl) URL.revokeObjectURL(app.ocrImageUrl);
  delete app.files.image;
  app.ocrImageUrl = '';
  app.ocrRotation = 0;
  app.ocrText = '';
  const label = $('imageName');
  if (label) label.textContent = '未选择文件';
  const output = $('ocrTextOutput');
  if (output) output.value = '';
  clearResult();
  previewOcrImage();
  updateStatus();
  toast('已删除图片，可重新上传', 'success');
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

async function runSealStamp(downloadAfter = false) {
  try {
    const report = requireFile('sealReport', '检验报告单图片');
    const reportImage = await loadHtmlImage(report.url);
    const stampImage = await loadHtmlImage(DEFAULT_SEAL_IMAGE);
    const rotation = normalizeRotation(app.sealRotation);
    const swapSize = rotation === 90 || rotation === 270;
    const canvas = document.createElement('canvas');
    canvas.width = swapSize ? reportImage.naturalHeight : reportImage.naturalWidth;
    canvas.height = swapSize ? reportImage.naturalWidth : reportImage.naturalHeight;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(rotation * Math.PI / 180);
    ctx.drawImage(reportImage, -reportImage.naturalWidth / 2, -reportImage.naturalHeight / 2);
    ctx.setTransform(1, 0, 0, 1, 0, 0);

    const sealWidth = Math.round(canvas.width * 0.18);
    const sealHeight = Math.round(stampImage.naturalHeight * sealWidth / stampImage.naturalWidth);
    const sealCanvas = document.createElement('canvas');
    const angle = -5 * Math.PI / 180;
    const extra = Math.ceil(Math.max(sealWidth, sealHeight) * 0.18);
    sealCanvas.width = sealWidth + extra * 2;
    sealCanvas.height = sealHeight + extra * 2;
    const sealCtx = sealCanvas.getContext('2d');
    sealCtx.translate(sealCanvas.width / 2, sealCanvas.height / 2);
    sealCtx.rotate(angle);
    sealCtx.drawImage(stampImage, -sealWidth / 2, -sealHeight / 2, sealWidth, sealHeight);

    const x = canvas.width - sealCanvas.width - Math.round(canvas.width * 0.055);
    const y = canvas.height - sealCanvas.height - Math.round(canvas.height * 0.045);
    ctx.drawImage(sealCanvas, x, y);

    const blob = await canvasToBlob(canvas, 'image/jpeg', 0.95);
    clearSealResultOnly();
    app.result = blob;
    app.resultName = `${baseName(report.file.name)}_旋正已盖章.jpg`;
    app.resultType = 'jpg';
    app.sealResultUrl = URL.createObjectURL(blob);
    $('downloadResult').disabled = false;
    previewSealImage();
    toast('已生成旋正盖章图片', 'success');
    if (downloadAfter) downloadResult();
  } catch (error) {
    toast(`处理失败：${error.message}`, 'error');
  }
}

async function saveSealToDesktop() {
  await runSealStamp(true);
}

async function runOcr() {
  try {
    const item = requireFile('image', '待识别图片');
    if (typeof Tesseract === 'undefined') {
      toast('OCR 识别库加载失败，请刷新页面后重试', 'error');
      return;
    }
    const output = $('ocrTextOutput');
    const lang = $('ocrLang')?.value || 'chi_sim+eng';
    const mode = $('ocrMode')?.value || 'document';
    $('previewMeta').textContent = '正在识别文字...';
    const imageForOcr = await getPreparedOcrSource(item);
    const result = await Tesseract.recognize(imageForOcr, lang, {
      tessedit_pageseg_mode: mode === 'table' ? '6' : '4',
      preserve_interword_spaces: '1',
      logger: info => {
        if (info.status && typeof info.progress === 'number') {
          $('previewMeta').textContent = `${info.status} / ${Math.round(info.progress * 100)}%`;
        }
      },
    });
    app.ocrText = cleanOcrText(result.data.text.trim(), mode);
    if (output) output.value = app.ocrText;
    renderOcrTextPreview();
    toast('图片文字识别完成', 'success');
  } catch (error) {
    toast(`识别失败：${error.message}`, 'error');
  }
}

async function copyOcrText() {
  const text = getOcrText();
  if (!text) {
    toast('暂无可复制文字', 'error');
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    toast('已复制文字', 'success');
  } catch (_) {
    const output = $('ocrTextOutput');
    output?.select();
    document.execCommand('copy');
    toast('已复制文字', 'success');
  }
}

function downloadOcrTxt() {
  const text = getOcrText();
  if (!text) {
    toast('暂无可下载文字', 'error');
    return;
  }
  downloadBlob(new Blob(['\ufeff' + text], { type: 'text/plain;charset=utf-8' }), getOcrBaseName() + '_识别文字.txt');
}

function exportOcrWord() {
  const text = getOcrText();
  if (!text) {
    toast('暂无可导出文字', 'error');
    return;
  }
  const body = text.split(/\r?\n/).map(line => `<p>${escapeHtml(line) || '&nbsp;'}</p>`).join('');
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>识别文字</title></head><body>${body}</body></html>`;
  downloadBlob(new Blob(['\ufeff' + html], { type: 'application/msword;charset=utf-8' }), getOcrBaseName() + '_识别文字.doc');
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

function loadHtmlImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('图片加载失败'));
    image.src = url;
  });
}

function canvasToBlob(canvas, type = 'image/png', quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(blob => {
      if (blob) resolve(blob);
      else reject(new Error('图片生成失败'));
    }, type, quality);
  });
}

function renderTable(headers, rows, rowClassFn, options = {}) {
  const showIndex = options.showIndex !== false;
  const visibleRows = rows.slice(0, 300);
  if (!headers.length) {
    $('previewTable').innerHTML = '<tbody><tr><td>暂无数据</td></tr></tbody>';
    return;
  }
  const indexHead = showIndex ? '<th>#</th>' : '';
  const head = `<thead><tr>${indexHead}${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>`;
  const body = visibleRows.map((row, index) => `
    <tr class="${rowClassFn ? rowClassFn(row) : ''}">
      ${showIndex ? `<td>${index + 1}</td>` : ''}
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

function getOcrText() {
  return ($('ocrTextOutput')?.value || app.ocrText || '').trim();
}

function getOcrLines() {
  return getOcrText().split(/\r?\n/).map(line => line.trim()).filter(Boolean);
}

function getOcrBaseName() {
  return baseName(app.files.image?.file?.name || '图片转文字');
}

function cleanOcrText(text, mode) {
  if (!text) return '';
  let next = text
    .replace(/\r/g, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g, '$1$2')
    .replace(/\s*[:：]\s*/g, '：')
    .replace(/纳税\s*人?\s*识别\s*号/g, '纳税人识别号')
    .replace(/地址\s*电话/g, '地址电话')
    .replace(/开户\s*行\s*及\s*账\s*号/g, '开户行及账号')
    .replace(/行\s*号/g, '行号')
    .replace(/单\s*位\s*名\s*称/g, '单位名称')
    .replace(/[|｜]/g, '')
    .replace(/[“”]/g, '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .join('\n');

  if (mode === 'document') {
    next = next
      .replace(/开\s*票\s*资\s*料/g, '开票资料')
      .replace(/(纳税人识别号：)\s*([0-9A-Z]+)/g, '$1$2')
      .replace(/(行号：)\s*([0-9]+)/g, '$1$2')
      .replace(/([0-9]{3,4})\s*-\s*([0-9]{6,8})/g, '$1-$2')
      .replace(/([0-9]{8,})\s+([0-9]{8,})/g, '$1\n$2');
  }
  return next;
}

function normalizeRotation(value) {
  return ((value % 360) + 360) % 360;
}

function getPreparedOcrSource(item) {
  const mode = $('ocrMode')?.value || 'document';
  const enhance = mode !== 'off';
  if (!app.ocrRotation && !enhance) return Promise.resolve(item.file);
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const rotation = normalizeRotation(app.ocrRotation);
      const swapSize = rotation === 90 || rotation === 270;
      const baseWidth = swapSize ? image.naturalHeight : image.naturalWidth;
      const baseHeight = swapSize ? image.naturalWidth : image.naturalHeight;
      const scale = getOcrScale(mode, baseWidth, baseHeight);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(baseWidth * scale);
      canvas.height = Math.round(baseHeight * scale);
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      ctx.translate(baseWidth / 2, baseHeight / 2);
      ctx.rotate(rotation * Math.PI / 180);
      ctx.drawImage(image, -image.naturalWidth / 2, -image.naturalHeight / 2);
      if (enhance) enhanceOcrCanvas(ctx, canvas, mode);
      canvas.toBlob(blob => {
        if (blob) resolve(blob);
        else reject(new Error('旋转图片处理失败'));
      }, 'image/png');
    };
    image.onerror = () => reject(new Error('图片加载失败'));
    image.src = item.url;
  });
}

function getOcrScale(mode, width, height) {
  if (mode === 'off') return 1;
  const longest = Math.max(width, height);
  if (mode === 'strong') return Math.min(4, Math.max(1.8, 2600 / longest));
  if (mode === 'table') return Math.min(3.5, Math.max(1.5, 2200 / longest));
  return Math.min(3, Math.max(1.35, 1800 / longest));
}

function enhanceOcrCanvas(ctx, canvas, mode) {
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;
  for (let i = 0; i < data.length; i += 4) {
    const gray = data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114;
    let value;
    if (mode === 'strong') {
      value = (gray - 128) * 1.75 + 128;
      value = value > 198 ? 255 : value < 108 ? 0 : value;
    } else if (mode === 'table') {
      value = (gray - 128) * 1.42 + 128;
      value = value > 224 ? 255 : value;
    } else {
      value = (gray - 128) * 1.25 + 128;
    }
    value = Math.max(0, Math.min(255, value));
    data[i] = value;
    data[i + 1] = value;
    data[i + 2] = value;
  }
  ctx.putImageData(imageData, 0, 0);
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
