const tools = [
  { id: 'internal-statement', icon: '对', title: '内部交易对账单', status: '可用', desc: '按采购组织整理内部交易，填充内置模板，分别下载 Excel 和 PDF。', href: 'internal-statement.html' },
  { id: 'funds', icon: '账', title: '付款节奏跟进', status: '可用', desc: '整理谈判付款、正常付款及申请核对记录，补充 OA 审批进度。', href: 'funds.html' },
  { id: 'negotiated-payment-summary', icon: '谈', title: '谈判付款整理', status: '可用', desc: '按供应商汇总每月支付金额，分别展示谈判付款和正常付款。', href: 'negotiated-payment-summary.html' },
  { id: 'funds-ledger', icon: '台', title: '资金台账整理', status: '可用', desc: '汇总全部月份的正常付款，匹配 OA 流程进度，已结清记录也保留。', href: 'funds-ledger.html' },
  { id: 'quarter-budget', icon: '季', title: '季度预算整理', status: '可用', desc: '合并付款预算、核对分期付款状态，按月份和分类汇总并下载 Excel。', href: 'quarter-budget.html' },
  { id: 'funds-budget-summary', icon: '预', title: '资金预算汇总', status: '可用', desc: '按月份、分类和状态汇总付款预算，更新汇总并保留原工作表。', href: 'funds-budget-summary.html' },
  { id: 'receipt-sign', icon: '签', title: '采购入库单签字', status: '可用', desc: '上传采购入库单后自动签字，合并下载 A4 竖向 PDF。', href: 'receipt-sign.html' },
  { id: 'seal', icon: '章', title: '检验报告单-电子章', status: '可用-品质', desc: '旋正检验报告图片，加盖内置质检合格章，下载 JPG。', href: 'quality-seal.html' },
  { id: 'beihuo', icon: '备', title: '备货分配', status: '可用', desc: '上传备货需求和维度表，匹配采购分工、汇总数量并导出分配明细。', href: 'beihuo/index.html' },
  { id: 'fulfillment', icon: '履', title: '订单履约表整理', status: '可用', desc: '合并徐英、操思敏、邢芳芳、李奇的履约明细，校验六项数量并统一日期格式。', href: 'fulfillment.html' },
  { id: 'contract', icon: '合', title: '合同生成', status: '可用', desc: '上传订单与合同模板，确认字段映射后在浏览器本地生成合同。', href: 'contract.html' },
  { id: 'payment', icon: '付', title: '付款申请单整理', status: '可用', desc: '按付款主体整理明细，优先一页 A4，明细较多时清晰分页，导出 Word 和 PDF。' },
];

const LAST_TOOL_KEY = 'gylsjqx-last-tool';

let activeTool = '';

const $ = (id) => document.getElementById(id);

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  renderToolList();
  let lastTool;
  try { lastTool = localStorage.getItem(LAST_TOOL_KEY); } catch (_) {}
  switchTool(tools.some(tool => tool.id === lastTool) ? lastTool : 'funds');
});

function initTheme() {
  let saved;
  try { saved = localStorage.getItem('dpc-theme'); } catch (_) {}
  if (saved === 'dark') document.documentElement.dataset.theme = 'dark';
  $('themeToggle').addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? '' : 'dark';
    if (next) document.documentElement.dataset.theme = next;
    else delete document.documentElement.dataset.theme;
    try { localStorage.setItem('dpc-theme', next || 'light'); } catch (_) {}
  });
}

function renderToolList() {
  $('toolList').innerHTML = tools.map(tool => `
    <button class="tool-tab" data-tool="${tool.id}" ${tool.href ? `data-href="${tool.href}"` : ''}>
      <span class="tab-icon">${tool.icon}</span>
      <span class="tab-text">${tool.title}（${tool.status}）</span>
    </button>
  `).join('');
  document.querySelectorAll('.tool-tab').forEach(button => {
    button.addEventListener('click', () => {
      if (button.dataset.tool === activeTool) return;
      switchTool(button.dataset.tool);
    });
  });
}

function switchTool(id) {
  const tool = tools.find(item => item.id === id);
  if (!tool) return;
  if (id === activeTool) return;
  activeTool = id;
  try { localStorage.setItem(LAST_TOOL_KEY, id); } catch (_) {}
  $('toolTitle').textContent = tool.title;
  $('toolDesc').textContent = tool.desc;
  document.querySelectorAll('.tool-tab').forEach(btn => btn.classList.toggle('active', btn.dataset.tool === id));
  $('toolPanel').hidden = Boolean(tool.href);
  $('embeddedTools').hidden = !tool.href;
  document.querySelectorAll('#embeddedTools iframe').forEach(frame => { frame.hidden = frame.dataset.tool !== id; });
  if (tool.href) {
    mountEmbeddedTool(tool);
    $('activeFileCount').textContent = '本地处理';
    $('activeRowCount').textContent = '切换保留内容';
    return;
  }
  if (id === 'payment') {
    PaymentTool.mount($('toolPanel'));
    $('activeFileCount').textContent = '本地处理';
    $('activeRowCount').textContent = '优先一页 · 清晰排版';
    return;
  }
}

// Keep each document alive so switching tools does not discard file inputs or results.
function mountEmbeddedTool(tool) {
  let frame = document.querySelector(`#embeddedTools iframe[data-tool="${tool.id}"]`);
  if (frame) return;
  frame = document.createElement('iframe');
  frame.dataset.tool = tool.id;
  frame.title = tool.title;
  frame.style.cssText = 'width:100%;height:800px;border:0';
  frame.addEventListener('load', () => {
    const doc = frame.contentDocument;
    const style = doc.createElement('style');
    style.textContent = `
      html,body{min-height:0!important;background:transparent!important;margin:0!important}
      body>main{max-width:none!important;margin:0!important;padding:0!important}
      body>main>a[href="index.html"],body>main>h1,body>main>h1+p,
      body>.app-header,.contract-sidebar,.contract-title-row>div:first-child,
      .app-shell>.sidebar,.topbar .eyebrow,.topbar h1{display:none!important}
      .contract-shell,.app-shell{display:block!important;min-height:0!important;padding:0!important}
      .dashboard{min-width:0!important;padding:0!important}
    `;
    doc.head.append(style);
    const syncTheme = () => { doc.documentElement.dataset.theme = document.documentElement.dataset.theme || ''; };
    syncTheme();
    new MutationObserver(syncTheme).observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});
    const resize = () => {
      if (!frame.hidden) frame.style.height = Math.ceil(doc.body.getBoundingClientRect().height + 24) + 'px';
    };
    new ResizeObserver(resize).observe(doc.body);
    resize();
  });
  // Share the homepage script version so cached child HTML cannot outlive a release.
  const script = [...document.scripts].find(item => item.src.includes('/app-merge19.js'));
  const source = new URL(tool.href, location.href);
  source.searchParams.set('v', new URL(script.src).searchParams.get('v') || '1');
  frame.src = source.href;
  $('embeddedTools').append(frame);
}
