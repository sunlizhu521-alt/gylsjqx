(function (root) {
  'use strict';
  const fields = ['供应商全称', '开户银行及账号', '合同编号', '合同总金额', '已付金额', '本次申请付款金额', '票据状态', '付款事由', '备注'];
  const required = ['申请日期', '付款主体', '申请部门', '经办人', ...fields];
  const infoFields = ['申请日期', '申请部门', '经办人'];
  const MAX_CENTS = 99999999999999n;
  function cents(value) {
    if (typeof value !== 'string' && typeof value !== 'number') throw new Error('金额不能为空或非数字');
    let s = String(value).trim().replace(/^[￥¥]\s*/, '');
    if (/^\d{1,3}(,\d{3})+(\.\d{1,2})?$/.test(s)) s = s.replace(/,/g, '');
    if (!/^\d+(\.\d{1,2})?$/.test(s)) throw new Error('金额须为非负数，且最多两位小数');
    const [a, b = ''] = s.split('.');
    const n = BigInt(a) * 100n + BigInt(b.padEnd(2, '0'));
    if (n > MAX_CENTS) throw new Error('金额超出支持范围（最高 9999 亿元）');
    return n;
  }
  function money(n) {
    return (n / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + (n % 100n).toString().padStart(2, '0');
  }
  function upper(n) {
    if (n < 0n || n > MAX_CENTS) throw new Error('大写金额超出范围');
    const digits = '零壹贰叁肆伍陆柒捌玖';
    function four(v) {
      const s = String(v), units = ['', '拾', '佰', '仟'];
      let out = '', zero = false;
      [...s].forEach((d, i) => {
        if (d === '0') { if (out) zero = true; }
        else { out += (zero ? '零' : '') + digits[Number(d)] + units[s.length - i - 1]; zero = false; }
      });
      return out;
    }
    let yuan = n / 100n, out = '';
    const groups = [];
    while (yuan) { groups.unshift(Number(yuan % 10000n)); yuan /= 10000n; }
    let zero = false;
    groups.forEach((v, i) => {
      if (!v) { if (out) zero = true; return; }
      if (out && (zero || v < 1000)) out += '零';
      out += four(v) + ['', '万', '亿'][groups.length - i - 1];
      zero = false;
    });
    out = (out || '零') + '元';
    const j = Number(n % 100n / 10n), f = Number(n % 10n);
    if (!j && !f) return out + '整';
    if (j) out += digits[j] + '角';
    if (f) out += (!j && n >= 100n ? '零' : '') + digits[f] + '分';
    return out;
  }
  function text(v) { return v == null ? '' : String(v); }
  function subjectName(value) { const name = text(value).trim(); return name === '浙江迈德斯特医疗器械科技有限公司' ? '迈德斯特' : name; }
  function downloadName(groups, extension) {
    if (!groups?.length) throw new Error('没有可下载的付款申请单');
    if (groups.length !== 1) throw new Error('请按付款主体分别下载');
    const g = groups[0];
    if (!text(g.info['计划付款日期']).trim()) throw new Error(`${g.subject}：请填写计划付款日期并重新生成后下载`);
    const period = date(g.info['计划付款日期']);
    const subject = subjectName(g.subject).replace(/[<>:"/\\|?*\x00-\x1f]/g, '_');
    return `${period}${subject}${money(g.total).replace(/,/g, '')}.${extension}`;
  }
  function printDownloadName(groups) {
    if (!groups?.length) throw new Error('没有可下载的付款申请单');
    const dates = groups.map(g => {
      const planned = date(g.info['计划付款日期']);
      if (!planned) throw new Error(`${g.subject}：请填写计划付款日期并重新生成后下载`);
      return planned;
    }).sort();
    const [, month, day] = dates[0].split('-');
    return `${Number(month)}月${Number(day)}日付款申请单-打印.docx`;
  }

  function bankSummary(value, supplier = '') {
    const raw = text(value);
    if (!raw.trim() || raw.trim() === '/') return { value: raw, compact: true };
    if (/^银行名称:[^\n]+\n卡号:[\d\s-]+$/.test(raw)) return { value: raw, compact: true };
    const input = raw.replace(/&#(?:x20|32);/gi, ' ');
    const labels = /收款银行名称|收款银行行号|开户银行行号|银行行号|开户行行号|收款单位|收款户名|账户名称|收款账号|银行账号|账户号码|开户银行|银行名称|开户行|联行号|单位|户名|账号|行号|备注/g;
    const tokens = [...input.matchAll(labels)].filter(m => /[:：\s]|^$/.test(input.slice(m.index + m[0].length, m.index + m[0].length + 1)));
    const units = [], accounts = [];
    tokens.forEach((m, i) => {
      const v = input.slice(m.index + m[0].length, tokens[i + 1]?.index ?? input.length).replace(/^\s*[:：]?\s*/, '').trim();
      if (/^(收款单位|收款户名|账户名称|单位|户名)$/.test(m[0]) && v) units.push(v);
      if (/^(收款账号|银行账号|账户号码|账号)$/.test(m[0]) && v) accounts.push(v);
    });
    const uniqueUnits = [...new Set(units)], uniqueAccounts = [...new Set(accounts)];
    const unit = uniqueUnits[0] || text(supplier).trim();
    if (uniqueUnits.length > 1 || uniqueAccounts.length !== 1 || !unit || !/^[\d\s-]+$/.test(uniqueAccounts[0])) return { value: raw, compact: false };
    return { value: `单位:${unit.replace(/[\r\n]+/g, ' ')}\n账号:${uniqueAccounts[0].replace(/[\r\n]+/g, ' ')}`, compact: true };
  }
  function bankIdentity(value, supplier) {
    const raw = text(value).trim();
    const summary = bankSummary(raw, supplier);
    return (summary.compact ? summary.value : /^[\d\s-]+$/.test(raw) ? `单位:${supplier}\n账号:${raw}` : raw).replace(/\s/g, '');
  }
  function buildDirectory(rows) {
    const index = new Map(), errors = [];
    for (const row of rows) {
      const name = text(row.supplier).trim();
      if (!name) { errors.push(`名录第 ${row.row} 行：供应商全称为空`); continue; }
      const raw = text(row.bank).trim(), unit = text(row.unit).trim() || name;
      const formatted = /^[\d\s-]+$/.test(raw) ? { value: `单位:${unit}\n账号:${raw}`, compact: true } : bankSummary(raw, unit);
      if (text(row.unit).trim() && formatted.compact && raw && raw !== '/') formatted.value = formatted.value.replace(/^单位:[^\n]*\n/, `单位:${unit}\n`);
      if (row.account !== undefined) {
        const account = text(row.account).trim(), bankName = text(row.bankName).trim();
        formatted.value = `银行名称:${bankName}\n卡号:${account}`;
        formatted.compact = !!bankName && /^[\d\s-]+$/.test(account);
      }
      const entry = { row: row.row, subject: subjectName(row.subject), bank: formatted.value, error: row.error || (row.account !== undefined ? (!text(row.account).trim() ? '收款账号为空' : !text(row.bankName).trim() ? '银行名称为空' : !formatted.compact ? '收款账号格式异常' : '') : (!raw || raw === '/' ? '银行信息为空' : !formatted.compact ? '银行信息无法明确识别，请选择账号列或使用带单位、账号标签的内容' : '')) };
      if (!index.has(name)) index.set(name, []);
      index.get(name).push(entry);
    }
    if (!rows.length) errors.push('供应商名录没有数据');
    return { index, errors };
  }
  function resolveBank(record, directory, policy = 'directory') {
    const original = text(record['开户银行及账号']);
    const name = text(record['供应商全称']).trim();
    const result = { value: original, source: '付款明细', notice: '', error: '', conflict: false };
    if (!directory) return result;
    if (policy === 'fill' && original.trim() && original.trim() !== '/') return result;
    const matches = directory.index.get(name) || [];
    if (!matches.length) {
      result.notice = '名录未匹配，沿用付款明细';
      if (!original.trim() || original.trim() === '/') result.error = '名录未匹配且付款明细银行信息为空';
      return result;
    }
    const invalid = matches.filter(e => e.error);
    if (invalid.length) { result.error = invalid.map(e => `名录第 ${e.row} 行：${e.error}`).join('；'); return result; }
    const identities = new Set(matches.map(e => bankIdentity(e.bank, name)));
    if (identities.size !== 1) { result.error = `名录第 ${matches.map(e => e.row).join('、')} 行存在不同的收款单位或账号，请核实名录`; return result; }
    result.value = matches[0].bank;
    result.source = `供应商名录第 ${matches.map(e => e.row).join('、')} 行`;
    result.conflict = !!original.trim() && original.trim() !== '/' && bankIdentity(original, name) !== bankIdentity(result.value, name);
    return result;
  }
  const budgetFields = {
    supplier: ['供应商', '供应商名称', '供应商全称'],
    subject: ['付款主体'],
    oa: ['对公付款0A流程号', '对公付款OA流程号', '对公付款OA号', '付款OA号', '流程编码', '流程编号', '合同编号', '合同号'],
    amount: ['OA-申请付款金额', '申请付款金额', '本次申请付款金额'],
  };
  const budgetHeader = value => text(value).normalize('NFKC').replace(/\s/g, '').toUpperCase();
  function budgetOA(value) {
    const raw = text(value).normalize('NFKC').trim(), key = raw.replace(/\s/g, '').toUpperCase();
    if (['', '/', '-', '未提交OA', '未申请OA', '未交OA', '未申请', '未提交'].includes(key)) return { key: '', invalid: false };
    const invalid = !/^[A-Z0-9_-]+$/.test(key) || (key.match(/MDST-DGFK/g) || []).length > 1 || /[A-Z0-9_-]*\d[A-Z0-9_-]*\s+[A-Z]+[-_]?\d/i.test(raw);
    return { key, invalid };
  }
  function budgetName(value) {
    const name = text(value).trim();
    return !!name && !['/', '-', '—', '--', '待定', '待核对', '未知', '未填写', '无', '暂无', '不详', '合计', '总计'].includes(name) && !/^[-+.\d]+$/.test(name) && !/多家供应商|多个供应商|各供应商|全部供应商|供应商合计|供应商汇总|[\r\n、;；/]/.test(name);
  }
  function parseBudget(book, filename, X, selectedSheet = '') {
    const result = { sheetNames: [], sheetName: '', rows: [], errors: [] }, keys = Object.keys(budgetFields);
    const aliases = new Map(Object.entries(budgetFields).flatMap(([key, names]) => names.map(name => [budgetHeader(name), key])));
    const tables = new Map();
    for (const name of book?.SheetNames || []) {
      const sheet = book.Sheets?.[name], byRow = new Map();
      let lastRow = 0;
      for (const address of Object.keys(sheet || {})) {
        if (!/^[A-Z]+\d+$/.test(address)) continue;
        const cell = sheet[address];
        if (!cell?.f && !text(cell?.v).trim()) continue;
        const position = X.utils.decode_cell(address); lastRow = Math.max(lastRow, position.r);
        if (position.r >= 100) continue;
        const key = aliases.get(budgetHeader(cell?.v));
        if (key) { if (!byRow.has(position.r)) byRow.set(position.r, {}); (byRow.get(position.r)[key] ||= []).push(position.c); }
      }
      const headers = [...byRow].sort((a, b) => a[0] - b[0]);
      const complete = headers.find(([, map]) => keys.every(key => map[key]));
      const best = complete || headers.sort((a, b) => Object.keys(b[1]).length - Object.keys(a[1]).length)[0];
      tables.set(name, { sheet, header: best?.[0], map: best?.[1] || {}, lastRow });
      if (complete) result.sheetNames.push(name);
    }
    result.sheetName = selectedSheet || (result.sheetNames.length === 1 ? result.sheetNames[0] : '');
    if (!result.sheetName) {
      result.errors.push(result.sheetNames.length > 1 ? '请选择预算工作表' : `${filename}：未找到预算表头，必须包含供应商、付款主体、对公付款OA流程号、OA-申请付款金额`);
      return result;
    }
    const table = tables.get(result.sheetName), location = `${filename} / ${result.sheetName}`;
    if (!table) { result.errors.push(`${location}：工作表不存在，请重新选择`); return result; }
    const missing = keys.filter(key => !table.map[key]);
    if (missing.length) { result.errors.push(`${location}：缺少预算表头 ${missing.map(key => budgetFields[key][0]).join('、')}`); return result; }
    for (const key of keys) if (table.map[key].length !== 1) result.errors.push(`${location} 第 ${table.header + 1} 行：预算字段“${budgetFields[key][0]}”对应多个列，请保留唯一列`);
    if (table.lastRow > 100000) result.errors.push(`${location}：预算数据超过100000行，请检查源表范围`);
    if (result.errors.length) return result;
    for (let r = table.header + 1; r <= table.lastRow; r++) {
      const cells = Object.fromEntries(keys.map(key => [key, table.sheet[X.utils.encode_cell({ r, c: table.map[key][0] })]]));
      if (keys.every(key => !cells[key]?.f && !text(cells[key]?.v).trim())) continue;
      const values = {};
      for (const key of keys) {
        const cell = cells[key], address = X.utils.encode_cell({ r, c: table.map[key][0] }), label = `${location}!${address}（${budgetFields[key][0]}）`;
        if (cell?.t === 'e' || cell?.f && (cell.v == null || cell.t === 'z' || cell.t === 'n' && typeof cell.v !== 'number')) result.errors.push(`${label}：公式错误或缺少有效计算结果，请在Excel/WPS重新计算并保存`);
        if (typeof cell?.v === 'number' && !Number.isFinite(cell.v)) result.errors.push(`${label}：数值异常`);
        values[key] = text(cell?.v).trim();
        if (key === 'oa' && typeof cell?.v === 'number') {
          if (!Number.isSafeInteger(cell.v) || Math.abs(cell.v) >= 1e15) result.errors.push(`${label}：OA编号保存为数字可能丢失精度，请改为文本并核实`);
          else { const displayed = X.utils.format_cell({ ...cell }); values.oa = /^\d+$/.test(displayed) ? displayed : String(cell.v); }
        }
      }
      let amount;
      try { amount = cents(cells.amount?.v); }
      catch (error) { result.errors.push(`${location}!${X.utils.encode_cell({ r, c: table.map.amount[0] })}（${budgetFields.amount[0]}）：${error.message}`); continue; }
      result.rows.push({ supplier: values.supplier, subject: subjectName(values.subject), oa: values.oa, amount, amountText: money(amount), source: { filename, sheet: result.sheetName, row: r + 1 } });
    }
    if (result.errors.length) result.rows = [];
    else if (!result.rows.length) result.errors.push(`${location}：没有预算明细`);
    return result;
  }
  function matchBudget(record, budget) {
    const result = { status: 'review', reason: '', method: '', candidates: [] };
    if (!budget || budget.errors?.length) { result.reason = '月份预算表尚未有效读取，请检查预算文件'; return result; }
    const supplier = text(record['供应商全称']).trim(), subject = subjectName(record['付款主体']), oa = budgetOA(record['合同编号']);
    const rows = budget.rows || [], sameNames = rows.filter(row => row.supplier === supplier && row.subject === subject);
    if (!budgetName(supplier) || !budgetName(subject)) { result.reason = '付款明细供应商或付款主体为空、待核对或包含多家，不能确认预算归属'; return result; }
    if (oa.invalid) { result.reason = '付款明细OA号包含多个编号或格式异常，请核对单笔付款OA号'; result.candidates = sameNames; return result; }
    let amount;
    try { amount = cents(record['本次申请付款金额']); }
    catch (error) { result.reason = `付款明细申请金额异常：${error.message}`; return result; }
    if (oa.key) result.candidates = rows.filter(row => { const code = budgetOA(row.oa); return !code.invalid && code.key === oa.key; });
    if (result.candidates.length) result.method = 'oa';
    else { result.candidates = sameNames; result.method = 'supplier-subject'; }
    if (!result.candidates.length) { result.reason = '预算表中未找到对应付款，请核实OA号、供应商及付款主体'; return result; }
    if (result.candidates.length !== 1) { result.reason = `${result.method === 'oa' ? '同一OA号' : '同一供应商及付款主体'}对应多条预算，请核实具体明细，不按金额选择或合并预算`; return result; }
    const candidate = result.candidates[0], candidateOA = budgetOA(candidate.oa);
    if (!budgetName(candidate.supplier) || !budgetName(candidate.subject)) { result.reason = '预算记录供应商或付款主体为空、待核对或包含多家，不能分摊预算'; return result; }
    if (candidateOA.invalid) { result.reason = '预算记录OA号包含多个编号或格式异常，请核对单笔预算OA号'; return result; }
    if (supplier !== candidate.supplier || subject !== candidate.subject) { result.reason = `OA号相同，但供应商或付款主体不一致；预算为“${candidate.supplier}／${candidate.subject}”，请核实`; return result; }
    if (result.method !== 'oa' && oa.key && candidateOA.key) { result.reason = `付款明细OA号“${oa.key}”与预算OA号“${candidateOA.key}”不一致，不能按名称回退匹配`; return result; }
    if (amount !== candidate.amount) { result.reason = `申请金额 ${money(amount)} 与预算金额 ${candidate.amountText} 不一致，请核实分批申请或金额差异`; return result; }
    result.status = 'matched'; result.reason = result.method === 'oa' ? 'OA号、供应商、付款主体及金额一致' : '一方缺OA号，供应商、付款主体唯一且金额一致';
    return result;
  }
  function parsePaymentOA(book, filename, X, selectedSheet = '') {
    const result = { sheetNames: [], sheetName: '', rows: [], errors: [], warnings: [] }, tables = new Map();
    const aliases = new Map([['流程编码', 'code'], ['流程编号', 'legacy'], ['付款金额', 'amount'], ['未操作者', 'unoperated']]);
    for (const name of book?.SheetNames || []) {
      const sheet = book.Sheets?.[name], byRow = new Map(); let lastRow = 0;
      for (const address of Object.keys(sheet || {})) {
        if (!/^[A-Z]+\d+$/.test(address)) continue;
        const cell = sheet[address]; if (!cell?.f && !text(cell?.v).trim()) continue;
        const position = X.utils.decode_cell(address); lastRow = Math.max(lastRow, position.r);
        if (position.r >= 100) continue;
        const key = aliases.get(budgetHeader(cell?.v));
        if (key) { if (!byRow.has(position.r)) byRow.set(position.r, {}); (byRow.get(position.r)[key] ||= []).push(position.c); }
      }
      const headers = [...byRow].sort((a, b) => a[0] - b[0]).map(([row, map]) => [row, { oa: map.code || map.legacy, amount: map.amount, unoperated: map.unoperated }]);
      const complete = headers.find(([, map]) => map.oa && map.amount && map.unoperated);
      const best = complete || headers.sort((a, b) => Object.values(b[1]).filter(Boolean).length - Object.values(a[1]).filter(Boolean).length)[0];
      tables.set(name, { sheet, header: best?.[0], map: best?.[1] || {}, lastRow });
      if (complete) result.sheetNames.push(name);
    }
    result.sheetName = selectedSheet || (result.sheetNames.length === 1 ? result.sheetNames[0] : '');
    if (!result.sheetName) {
      result.errors.push(result.sheetNames.length > 1 ? '请选择对公付款OA工作表' : `${filename}：未找到OA表头，必须包含流程编码或流程编号、付款金额、未操作者`);
      return result;
    }
    const table = tables.get(result.sheetName), location = `${filename} / ${result.sheetName}`;
    if (!table) { result.errors.push(`${location}：工作表不存在，请重新选择`); return result; }
    const labels = { oa: '流程编码或流程编号', amount: '付款金额', unoperated: '未操作者' }, keys = Object.keys(labels);
    const missing = keys.filter(key => !table.map[key]);
    if (missing.length) { result.errors.push(`${location}：缺少OA表头 ${missing.map(key => labels[key]).join('、')}`); return result; }
    for (const key of keys) if (table.map[key].length !== 1) result.errors.push(`${location} 第 ${table.header + 1} 行：OA字段“${labels[key]}”对应多个列，请保留唯一列`);
    if (table.lastRow > 100000) result.errors.push(`${location}：OA数据超过100000行，请检查源表范围`);
    if (result.errors.length) return result;
    for (let r = table.header + 1; r <= table.lastRow; r++) {
      const cells = Object.fromEntries(keys.map(key => [key, table.sheet[X.utils.encode_cell({ r, c: table.map[key][0] })]]));
      if (keys.every(key => !cells[key]?.f && !text(cells[key]?.v).trim())) continue;
      const values = {}, previousErrors = result.errors.length;
      let missingOA = false;
      for (const key of keys) {
        const cell = cells[key], label = `${location}!${X.utils.encode_cell({ r, c: table.map[key][0] })}（${labels[key]}）`;
        if (cell?.t === 'e' || cell?.f && (cell.v == null || cell.t === 'z' || cell.t === 'n' && typeof cell.v !== 'number')) result.errors.push(`${label}：公式错误或缺少有效计算结果，请在Excel/WPS重新计算并保存`);
        if (typeof cell?.v === 'number' && !Number.isFinite(cell.v)) result.errors.push(`${label}：数值异常`);
        values[key] = text(cell?.v).trim();
        if (key === 'oa' && typeof cell?.v === 'number') {
          if (!Number.isSafeInteger(cell.v) || Math.abs(cell.v) >= 1e15) result.errors.push(`${label}：OA编号保存为数字可能丢失精度，请改为文本并核实`);
          else { const displayed = X.utils.format_cell({ ...cell }); values.oa = /^\d+$/.test(displayed) ? displayed : String(cell.v); }
        }
        if (key === 'oa' && !budgetOA(values.oa).key) {
          missingOA = true;
          if (result.errors.length === previousErrors) result.warnings.push(`${location} 第 ${r + 1} 行：流程号为空，未参与OA核对`);
          break;
        }
      }
      if (missingOA) continue;
      let amount;
      try { amount = cents(cells.amount?.v); }
      catch (error) { result.errors.push(`${location}!${X.utils.encode_cell({ r, c: table.map.amount[0] })}（付款金额）：${error.message}`); continue; }
      result.rows.push({ oa: values.oa, amount, amountText: money(amount), unoperated: values.unoperated, source: { filename, sheet: result.sheetName, row: r + 1 } });
    }
    if (result.errors.length) result.rows = [];
    else if (!result.rows.length) result.errors.push(`${location}：没有OA明细`);
    return result;
  }
  function matchPaymentOA(record, parsed) {
    const result = { status: 'review', reason: '', candidates: [], unoperated: '' };
    if (!parsed || parsed.errors?.length) { result.reason = '对公付款OA文件尚未有效读取，请检查OA文件'; return result; }
    const rawOA = record['合同编号'], oa = budgetOA(rawOA);
    if (typeof rawOA === 'number' && (!Number.isSafeInteger(rawOA) || Math.abs(rawOA) >= 1e15)) { result.reason = '付款明细OA编号保存为数字可能丢失精度，请改为文本并核实'; return result; }
    if (!oa.key) { result.reason = '付款明细缺少付款OA号，请核实合同编号列'; return result; }
    if (oa.invalid) { result.reason = '付款明细OA号包含多个编号或格式异常，请核对单笔付款OA号'; return result; }
    result.candidates = (parsed.rows || []).filter(row => { const code = budgetOA(row.oa); return !code.invalid && code.key === oa.key; });
    if (!result.candidates.length) { result.reason = `对公付款OA表未找到流程编号“${oa.key}”，请核实编号或重新导出OA文件`; return result; }
    const first = result.candidates[0];
    if (result.candidates.some(row => row.amount !== first.amount || row.unoperated !== first.unoperated)) { result.reason = `同一OA号“${oa.key}”的付款金额或未操作者不一致，请核实对应流程记录`; return result; }
    let amount;
    try { amount = cents(record['本次申请付款金额']); }
    catch (error) { result.reason = `付款明细申请金额异常：${error.message}`; return result; }
    if (amount !== first.amount) { result.reason = `本次申请付款金额 ${money(amount)} 与OA付款金额 ${first.amountText} 不一致，请核实分批申请或金额差异`; return result; }
    result.status = 'matched'; result.unoperated = first.unoperated; result.reason = '付款OA号及付款金额一致';
    return result;
  }
  function date(v) {
    if (v === '' || v == null) return '';
    let s = text(v).trim().replace(/[年月/.]/g, '-').replace(/日$/, '');
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]00:00:00)?$/);
    if (!m) throw new Error('日期应为 YYYY-MM-DD');
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (d.getUTCFullYear() !== +m[1] || d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) throw new Error('日期无效');
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }
  function weekFriday(now = new Date()) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(now).map(p => [p.type, p.value]));
    const friday = new Date(Date.UTC(+parts.year, +parts.month - 1, +parts.day));
    // Monday-Sunday is one week, so weekends use the Friday just passed.
    friday.setUTCDate(friday.getUTCDate() + 4 - (friday.getUTCDay() + 6) % 7);
    return friday.toISOString().slice(0, 10);
  }
  function analyze(headers, records) {
    const errors = [], groups = new Map();
    const missing = required.filter(f => !headers.includes(f));
    const duplicates = headers.filter((h, i) => h && headers.indexOf(h) !== i);
    if (missing.length) errors.push('缺少列：' + missing.join('、'));
    if (duplicates.length) errors.push('重复列名：' + [...new Set(duplicates)].join('、'));
    if (errors.length) return { errors, groups: [] };
    for (const record of records) {
      const subject = subjectName(record['付款主体']), source = record._row;
      // Skip only a valid, exact zero; invalid or missing amounts still need correction.
      try { if (cents(record['本次申请付款金额']) === 0n) continue; } catch (_) {}
      if (!subject) { errors.push(`第 ${source} 行：缺少付款主体`); continue; }
      if (!groups.has(subject)) groups.set(subject, { subject, rows: [], total: 0n, info: {}, choices: {}, checks: [] });
      const group = groups.get(subject);
      const row = { ...record };
      for (const field of ['合同总金额', '已付金额']) if (!text(row[field]).trim()) row[field] = '/';
      row['票据状态'] = '已开票';
      row['备注'] = [record['付款分类'], record['经办人'], record['备注']].map(v => text(v).trim()).filter(Boolean).join('-');
      try {
        const due = text(record['应付时间'] ?? record['应付款时间']).trim();
        const parsed = due && due !== '/' ? date(due.replace(/[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$/, '')) : '';
        const deadline = parsed ? `${Number(parsed.slice(5, 7))}月${Number(parsed.slice(8, 10))}日前付款` : '';
        row['付款事由'] = [text(record['付款事由']).trim(), deadline].filter(Boolean).join('，');
      } catch (e) { errors.push(`第 ${source} 行：应付时间${e.message}，请核对付款期限`); }
      try { row._cents = cents(record['本次申请付款金额']); group.total += row._cents; }
      catch (e) { errors.push(`第 ${source} 行：本次申请付款金额${e.message}`); }
      if (group.total > MAX_CENTS) errors.push(`${subject}：合计超出支持范围`);
      try { row['申请日期'] = date(record['申请日期']); }
      catch (e) { errors.push(`第 ${source} 行：申请日期${e.message}`); }
      group.rows.push(row);
    }
    const plannedDate = weekFriday();
    for (const group of groups.values()) {
      for (const field of infoFields) {
        const choices = [...new Set(group.rows.map(r => text(r[field]).trim()))];
        group.choices[field] = choices;
        group.info[field] = choices.length === 1 ? choices[0] : '';
      }
      group.info['计划付款日期'] = plannedDate;
    }
    if (!records.length) errors.push('工作表没有付款明细');
    else if (!groups.size && !errors.length) errors.push('本次申请付款金额均为 0，没有需要生成的付款明细');
    return { errors, groups: [...groups.values()] };
  }
  const api = { fields, required, infoFields, cents, money, upper, text, date, weekFriday, analyze, downloadName, printDownloadName, subjectName, bankSummary, buildDirectory, resolveBank, parseBudget, matchBudget, parsePaymentOA, matchPaymentOA };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PaymentCore = api;
})(typeof window === 'undefined' ? globalThis : window);
