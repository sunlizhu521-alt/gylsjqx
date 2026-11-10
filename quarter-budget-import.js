(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.QuarterBudgetImport = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const aliases = {
    month: ['预估付款月份', '申请付款月份', '预计付款月份'],
    supplier: ['供应商名称', '供应商', '供应商全称'],
    subject: ['付款主体'], oa: ['对公付款OA流程号', '对公付款OA号', 'OA流程号'],
    summary: ['摘要'], amount: ['申请付款金额', '本次申请付款金额', 'OA-申请付款金额', '申请金额'],
    status: ['是否已支付', '是否已付款', '是否已付', '付款状态'],
    category: ['分类'], owner: ['跟单员', '采购跟单员', '创跟单员'], excluded: ['是否剔除'],
    detail: ['谈判明细'], paid: ['已付金额'], settled: ['是否结清']
  };
  const noteHeaders = /备注|账期|付款条件|应付款?时间|当前节点|未操作者|流程状态|需要提交付款审单的日期|是否提交纸质付款审单/;
  const norm = value => String(value == null ? '' : value).replace(/\s/g, '').replace(/0A/gi, 'OA');
  const nonempty = value => value !== undefined && value !== null && value !== '';
  const text = value => value == null ? '' : String(value);
  const sourceId = source => source.fingerprint || source.filename || 'source';
  function columnName(column) {
    let name = '';
    for (let n = column + 1; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
    return name;
  }
  function address(row, col) { return columnName(col) + row; }
  function coordinates(key) {
    const m = /^([A-Z]+)([1-9]\d*)$/.exec(key);
    if (!m) return null;
    let col = 0;
    for (const c of m[1]) col = col * 26 + c.charCodeAt(0) - 64;
    return { row: Number(m[2]), col: col - 1 };
  }
  function raw(cell) {
    if (!cell) return null;
    if (cell.t === 'e') return cell.w || text(cell.v) || '#ERROR!';
    if (cell.v instanceof Date) return Number.isNaN(cell.v.getTime()) ? 'Invalid Date' : cell.v.toISOString();
    return cell.v == null ? null : cell.v;
  }
  function status(value) {
    const v = norm(value);
    if (/^(已支付|已付款|已付|付清|已结清)$/.test(v)) return '已支付';
    if (/^(未支付|未付款|未付|不支付)$/.test(v)) return '未支付';
    return '待核对';
  }
  function decimal(value) {
    if (typeof value !== 'number' && typeof value !== 'string') return null;
    if (typeof value === 'number' && !Number.isFinite(value)) return null;
    let str = String(value).trim().replace(/^[￥¥]\s*/, '').replace(/\s*元$/, '');
    if (str.includes(',') && !/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(str)) return null;
    str = str.replace(/,/g, '');
    const match = /^([+-]?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(str);
    if (!match) return null;
    const exponent = Number(match[4] || 0);
    if (!Number.isInteger(exponent) || Math.abs(exponent) > 100) return null;
    const digits = match[2] + (match[3] || '');
    const scale = (match[3] || '').length - exponent;
    let units = BigInt(digits), cents;
    if (scale <= 2) cents = units * 10n ** BigInt(2 - scale);
    else {
      const divisor = 10n ** BigInt(scale - 2);
      cents = units / divisor + (units % divisor * 2n >= divisor ? 1n : 0n);
    }
    if (match[1] === '-') cents = -cents;
    if (cents > BigInt(Number.MAX_SAFE_INTEGER) || cents < BigInt(Number.MIN_SAFE_INTEGER)) return null;
    return { cents: Number(cents), precision: scale > 2 };
  }
  function dated(value, date1904) {
    if (typeof value === 'number' && Number.isFinite(value) && value > 12 && value < 2958466) {
      // Excel's fictitious 1900-02-29 has no Gregorian equivalent.
      if (!date1904 && Math.floor(value) === 60) return null;
      const day = date1904 ? value + 1462 : value;
      const epoch = day < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
      const d = new Date(epoch + Math.floor(day) * 86400000);
      return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
    }
    const str = text(value).trim();
    const m = /^(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})(?:\s*(?:月|[-/.])\s*(\d{1,2}))?(?:月|日)?(?:T.*)?$/.exec(str);
    if (!m) return null;
    const year = Number(m[1]), month = Number(m[2]), day = Number(m[3] || 1);
    const d = new Date(Date.UTC(year, month - 1, day));
    return d.getUTCFullYear() === year && d.getUTCMonth() + 1 === month && d.getUTCDate() === day ? { year, month, day } : null;
  }
  function monthNumber(value) {
    if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 12) return value;
    const m = /^(\d{1,2})月(?:\d{1,2}日)?$/.exec(text(value).trim());
    return m && +m[1] >= 1 && +m[1] <= 12 ? +m[1] : null;
  }
  function keyFor(value, options, date1904) {
    const d = dated(value, date1904);
    const month = d ? d.month : monthNumber(value);
    const year = d ? d.year : Number.isInteger(options.budgetYear) ? options.budgetYear : null;
    return year && month ? `${year}-${String(month).padStart(2, '0')}` : '';
  }
  function isRed(cell) {
    const s = cell && cell.s;
    if (!s || typeof s !== 'object') return false;
    const fill = s.fill || s;
    if (fill.patternType && fill.patternType !== 'solid') return false;
    const color = fill.fgColor || {};
    if (color.indexed === 10) return true;
    return /^(?:FF)?(?:FF0000|CC0000|C00000|9C0006)$/i.test(color.rgb || '');
  }
  function installments(value) {
    const parts = [], re = /(?:(\d{4})年)?(\d{1,2})月(?:\d{1,2}日)?\s*[－—–:：-]\s*([￥¥]?\s*\d[\d,]*(?:\.\d+)?)(万)?(?:元)?/g;
    let match;
    while ((match = re.exec(text(value)))) {
      const month = Number(match[2]);
      if (month < 1 || month > 12) continue;
      const value = match[3].replace(/[￥¥,\s]/g, '');
      const amount = Number(value) * (match[4] ? 10000 : 1);
      parts.push({ month: match[1] ? `${match[1]}年${month}月` : `${month}月`, monthNumber: month, amount, fragment: match[0] });
    }
    return parts;
  }
  function inspect(sheet) {
    const rows = new Map();
    for (const name of Object.keys(sheet)) {
      const pos = coordinates(name), cell = sheet[name];
      if (!pos || !cell || !nonempty(cell.v) && !cell.f && cell.t !== 'e') continue;
      if (!rows.has(pos.row)) rows.set(pos.row, new Map());
      rows.get(pos.row).set(pos.col, { cell, name });
    }
    const orderedRows = [...rows.keys()].sort((a, b) => a - b);
    for (const row of orderedRows.slice(0, 30)) {
      const fields = {};
      for (const [col, entry] of rows.get(row)) {
        const label = norm(raw(entry.cell));
        for (const field of Object.keys(aliases)) if (aliases[field].some(a => norm(a) === label) && fields[field] == null) fields[field] = col;
      }
      if (fields.supplier != null && (fields.amount != null || fields.detail != null)) return { rows, orderedRows, fields, headerRow: row, headerCells: rows.get(row) };
    }
    return { rows, orderedRows, fields: {}, headerRow: null };
  }

  function parse(source, role, options) {
    options = options || {};
    if (!source || !source.book || !source.book.Sheets) throw new TypeError('缺少已读取的 Excel 工作簿');
    if (!['negotiated', 'ledger', 'budget', 'existing', 'evidence'].includes(role)) throw new TypeError('未知导入角色');
    const result = { records: [], warnings: [], errors: [], sheets: [], evidence: [] };
    const book = source.book, date1904 = !!(book.Workbook && book.Workbook.WBProps && book.Workbook.WBProps.date1904);
    const evidenceIds = new Set(), summaryEvidenceIds = new Set();
    const location = (sheet, row, part) => ({ role, file: source.filename || '', sheet, row, part: part || 1 });
    function issue(level, code, message, sheet, row, cell) {
      result[level].push({ code, message, source: location(sheet, row), cell: cell || '' });
    }
    function readAmount(cell, sheet, row, coordinate) {
      const amount = raw(cell);
      if (cell && (cell.t === 'e' || cell.f && !nonempty(cell.v))) {
        issue('errors', 'AMOUNT_FORMULA_ERROR', '金额公式无有效缓存或包含错误，不能按零计算', sheet, row, coordinate);
        return { amount, cents: null };
      }
      if (!nonempty(amount)) return { amount: null, cents: null };
      const d = decimal(amount);
      if (!d) issue('errors', 'INVALID_AMOUNT', '金额无效或超出安全整数分范围，保留原值', sheet, row, coordinate);
      else if (d.precision) issue('warnings', 'AMOUNT_PRECISION', '原金额超过两位小数，原值保留，统计按分四舍五入', sheet, row, coordinate);
      return { amount, cents: d ? d.cents : null };
    }
    function addEvidence(record, sheet, row, cell, cents, state, monthKey, part, kind) {
      if (cents == null || !['已支付', '未支付'].includes(state)) return null;
      const id = `${sourceId(source)}:${encodeURIComponent(sheet)}:${cell}`;
      if (!evidenceIds.has(id)) {
        evidenceIds.add(id);
        result.evidence.push({ id, oa: record.oa, supplier: record.supplier, subject: record.subject, summary: record.summary,
          cents, monthKey: monthKey || '', status: state, source: { ...location(sheet, row, part), cell }, part: part || 1, kind, strength: 'explicit' });
      }
      return id;
    }
    const names = options.sheetName ? [options.sheetName] : (book.SheetNames || Object.keys(book.Sheets));
    const inspected = new Map(names.filter(name => book.Sheets[name]).map(name => [name, inspect(book.Sheets[name])]));
    let selected = new Set(names);
    if (!options.sheetName && role !== 'evidence') {
      const candidates = names.filter(name => !/原始表|备份|旧表|backup/i.test(name) && inspected.get(name) && inspected.get(name).headerRow);
      let preferred = [];
      if (role === 'negotiated') {
        preferred = candidates.filter(name => name === '谈判付款');
        if (!preferred.length) preferred = candidates.filter(name => /谈判.*付款|付款.*谈判/.test(name));
        if (!preferred.length) preferred = candidates.filter(name => { const f = inspected.get(name).fields; return f.detail != null || f.month != null && f.amount != null; });
      } else if (role === 'ledger') {
        preferred = candidates.filter(name => name === '正常付款');
        if (!preferred.length) preferred = candidates.filter(name => /正常付款/.test(name));
        if (!preferred.length) preferred = candidates.filter(name => inspected.get(name).fields.detail == null);
      } else if (role === 'existing') {
        preferred = candidates.filter(name => name === '付款汇总明细');
        if (!preferred.length) preferred = candidates.filter(name => inspected.get(name).fields.month != null && inspected.get(name).fields.detail == null);
      } else {
        const budgetCandidates = candidates.filter(name => { const f = inspected.get(name).fields; return f.month != null && f.detail == null; });
        preferred = budgetCandidates.filter(name => /^(预估预算.*|未预估付款明细|预估付款明细|预算明细)$/.test(name));
        if (!preferred.length) {
          const standardized = budgetCandidates.filter(name => inspected.get(name).fields.category != null);
          preferred = standardized.length ? standardized : budgetCandidates;
        }
      }
      const rawBudgetSheets = role === 'budget' && preferred.every(name => inspected.get(name).fields.category == null && !/预算.*明细|预估预算/.test(name));
      if (preferred.length > 1 && !rawBudgetSheets) {
        issue('errors', 'AMBIGUOUS_SHEETS', `存在多个可用工作表：${preferred.join('、')}；请明确选择工作表`, '', 0);
        selected = new Set();
      } else selected = new Set(preferred);
    }
    for (const sheetName of names) {
      const sheet = book.Sheets[sheetName];
      if (!sheet) { issue('errors', 'MISSING_SHEET', '指定工作表不存在', sheetName, 0); continue; }
      if (!options.sheetName && /原始表|备份|旧表|backup/i.test(sheetName)) {
        result.sheets.push({ name: sheetName, skipped: true, reason: '旧备份不重复纳入', recordCount: 0 }); continue;
      }
      if (!selected.has(sheetName)) { result.sheets.push({ name: sheetName, skipped: true, reason: '不属于本次角色的明细工作表或需明确选择', recordCount: 0 }); continue; }
      // Iterate stored content, never !ref: formatted workbooks can claim a million rows.
      const { rows, orderedRows, headerRow, fields, headerCells } = inspected.get(sheetName);
      if (!headerRow) { result.sheets.push({ name: sheetName, skipped: true, reason: '未找到付款明细表头', recordCount: 0 }); continue; }
      const initialCount = result.records.length;
      const getCell = (row, field) => fields[field] == null ? null : sheet[address(row, fields[field])];
      const get = (row, field) => raw(getCell(row, field));
      const noteColumns = [...headerCells].filter(([, entry]) => noteHeaders.test(text(raw(entry.cell))));
      const details = fields.detail != null && (role === 'negotiated' || role === 'evidence');
      const scheduleColumns = details ? [...headerCells].filter(([col, entry]) => col > fields.detail && (dated(raw(entry.cell), date1904) || monthNumber(raw(entry.cell)))).map(([col]) => col).sort((a, b) => a - b) : [];
      const mergeByRow = new Map();
      for (const merge of sheet['!merges'] || []) {
        if (details && fields.detail >= merge.s.c && fields.detail <= merge.e.c && merge.e.r > merge.s.r) {
          const start = merge.s.r + 1, end = merge.e.r + 1;
          const activeRows = orderedRows.filter(r => r >= start && r <= end && nonempty(get(r, 'supplier')));
          const parts = installments(raw(sheet[address(start, fields.detail)]));
          for (const row of activeRows) mergeByRow.set(row, { start, end, rows: activeRows, parts });
        }
      }
      function base(row, part) {
        const originalNotes = noteColumns.filter(([col]) => nonempty(raw(sheet[address(row, col)]))).map(([col, entry]) => `${norm(raw(entry.cell)) === '备注' ? '' : text(raw(entry.cell)) + '：'}${text(raw(sheet[address(row, col)]))}`);
        const rec = { id: `${sourceId(source)}:${encodeURIComponent(sheetName)}:${row}:${part || 1}`, source: location(sheetName, row, part),
          month: get(row, 'month'), monthKey: keyFor(get(row, 'month'), options, date1904), monthDate1904: date1904, supplier: text(get(row, 'supplier')),
          subject: text(get(row, 'subject')), oa: text(get(row, 'oa')), summary: text(get(row, 'summary')),
          amount: null, cents: null, status: status(get(row, 'status')), notes: originalNotes.join('\n'),
          category: text(get(row, 'category')) || ({ negotiated: '谈判付款', ledger: '正常付款', budget: '预估付款' }[role] || ''),
          owner: text(get(row, 'owner')) || (role === 'budget' && fields.category == null ? sheetName : ''),
          excluded: norm(get(row, 'excluded')) === '是' ? '是' : '否', evidenceIds: [] };
        if (!rec.oa) {
          const oas = [...new Set((rec.notes.match(/MDST[-－]DGFK\d{11,}/gi) || []))];
          if (oas.length === 1) rec.oa = oas[0];
          else if (oas.length > 1) issue('warnings', 'MULTIPLE_OA', '备注含多个流程号，不能唯一提取', sheetName, row);
        }
        if (/原流程|历史流程/.test(rec.notes) && /未提交|未提交流程/.test(rec.notes)) rec.notes += '\n流程说明：原流程，本期未提交。';
        if (!rec.monthKey && nonempty(rec.month)) issue('warnings', 'UNKNOWN_MONTH_YEAR', '付款月份没有可唯一确定的年份，保留原值', sheetName, row);
        return rec;
      }
      for (const row of orderedRows.filter(r => r > headerRow)) {
        const supplier = get(row, 'supplier');
        if (!nonempty(supplier) || /^(合计|总计|小计|备注|说明)[:：]?$/.test(norm(supplier))) continue;
        if (!details && !['amount', 'summary', 'month', 'oa', 'subject'].some(field => nonempty(get(row, field)) || getCell(row, field) && getCell(row, field).f)) continue;
        if (!details) {
          const rec = base(row), cell = getCell(row, 'amount'), coord = fields.amount == null ? '' : address(row, fields.amount);
          Object.assign(rec, readAmount(cell, sheetName, row, coord));
          if (rec.cents == null && !nonempty(rec.amount)) issue('warnings', 'MISSING_AMOUNT', '申请付款金额为空，保留空值', sheetName, row, coord);
          if (fields.status == null) {
            const bareNotes = noteColumns.map(([col]) => text(raw(sheet[address(row, col)])).trim());
            if (bareNotes.some(n => /^(已付|已付款|已支付)[。！!]?$/u.test(n) || /^(本笔|本期)(款项)?已(支付|付款|付清)[。！!]?$/u.test(n) || /^(?:\d{4}[年/.-]\d{1,2}[月/.-]\d{1,2}[日号]?\s*[，,、和及]?\s*)+已(?:支付|付款|付)[。！!]?$/u.test(n))) rec.status = '已支付';
            const paid = decimal(get(row, 'paid'));
            if (rec.cents > 0 && paid && paid.cents === rec.cents && norm(get(row, 'settled')) === '是') rec.status = '已支付';
          }
          if (text(get(row, 'status')).includes('部分')) {
            rec.notes += `${rec.notes ? '\n' : ''}原付款状态：${get(row, 'status')}；已付组成及剩余金额待核对。`;
            issue('warnings', 'PARTIAL_PAYMENT', '部分支付不能视为全额已付', sheetName, row);
          }
          const ev = role === 'existing' ? null : addEvidence(rec, sheetName, row, coord, rec.cents, rec.status, rec.monthKey, 1, 'ledger');
          if (ev) rec.evidenceIds.push(ev);
          if (ev && fields.month == null && fields.summary == null) summaryEvidenceIds.add(ev);
          if (role !== 'evidence') result.records.push(rec);
          continue;
        }
        const recBase = base(row), merge = mergeByRow.get(row);
        const originalDetail = merge ? raw(sheet[address(merge.start, fields.detail)]) : get(row, 'detail');
        let parts = merge ? merge.parts : installments(originalDetail), ambiguousMerge = false;
        if (merge) {
          ambiguousMerge = merge.parts.length !== merge.rows.length;
          parts = ambiguousMerge ? [] : [{ ...merge.parts[merge.rows.indexOf(row)], originalPart: merge.rows.indexOf(row) + 1 }];
          if (ambiguousMerge) issue('warnings', 'MERGED_PARTS_AMBIGUOUS', '合并谈判明细期数与有效源行数不一致，不将同一明细复制到不同OA', sheetName, row, address(merge.start, fields.detail));
        }
        const schedule = [];
        for (const col of scheduleColumns) {
          const cell = sheet[address(row, col)];
          if (!cell || !nonempty(cell.v) && !cell.f && cell.t !== 'e') continue;
          const value = readAmount(cell, sheetName, row, address(row, col));
          const d = dated(raw(sheet[address(headerRow, col)]), date1904);
          const mo = d ? d.month : monthNumber(raw(sheet[address(headerRow, col)]));
          schedule.push({ ...value, cell, col, coordinate: address(row, col), month: mo, monthKey: d ? `${d.year}-${String(d.month).padStart(2, '0')}` : '', used: false });
        }
        if (!parts.length) {
          const rec = { ...recBase, month: null, monthKey: '', amount: null, cents: null, status: '待核对', evidenceIds: [] };
          rec.notes += `${rec.notes ? '\n' : ''}${ambiguousMerge ? '合并明细不能唯一分配' : nonempty(originalDetail) ? '谈判明细无法识别月份及金额' : '谈判明细为空'}${nonempty(originalDetail) ? `；原文：${originalDetail}` : ''}`;
          if (!ambiguousMerge) issue('warnings', 'MISSING_INSTALLMENTS', '缺少可解析的谈判分期，保留源行并待核对', sheetName, row);
          if (role !== 'evidence') result.records.push(rec);
        }
        const yearsByPaymentMonth = new Map();
        function rememberYear(month, key) {
          if (!key) return;
          if (!yearsByPaymentMonth.has(month)) yearsByPaymentMonth.set(month, new Set());
          yearsByPaymentMonth.get(month).add(key.slice(0, 4));
        }
        parts.forEach((part, index) => {
          rememberYear(part.monthNumber, keyFor(part.month, {}, date1904));
          const equal = schedule.filter(item => item.month === part.monthNumber && item.cents === decimal(part.amount)?.cents);
          const paired = schedule.length === parts.length ? schedule[index] : equal.length === 1 ? equal[0] : null;
          if (paired && paired.month === part.monthNumber) rememberYear(part.monthNumber, paired.monthKey);
        });
        for (let index = 0; index < parts.length; index++) {
          const part = parts[index], ordinal = part.originalPart || index + 1;
          const rec = { ...recBase, id: `${recBase.id.replace(/:\d+$/, '')}:${ordinal}`, source: location(sheetName, row, ordinal), month: part.month, amount: part.amount, cents: null, evidenceIds: [] };
          Object.assign(rec, readAmount({ v: part.amount }, sheetName, row, address(merge ? merge.start : row, fields.detail)));
          // Application-week headers can precede the payment month; intact sequences take precedence.
          let matched = schedule.length === parts.length ? schedule[index] : schedule.find(s => !s.used && s.cents === rec.cents && s.month === part.monthNumber);
          if (!matched) {
            const sameAmount = schedule.filter(s => !s.used && s.cents === rec.cents);
            if (sameAmount.length === 1) matched = sameAmount[0];
          }
          rec.monthKey = keyFor(rec.month, options, date1904);
          rec.notes += `${rec.notes ? '\n' : ''}谈判明细：${originalDetail}`;
          const knownYears = yearsByPaymentMonth.get(part.monthNumber);
          if (!rec.monthKey && knownYears && knownYears.size === 1) {
            const year = [...knownYears][0];
            rec.monthKey = `${year}-${String(part.monthNumber).padStart(2, '0')}`;
          }
          rec.status = '待核对';
          if (matched && !matched.used) {
            matched.used = true;
            if (!rec.monthKey && matched.month === part.monthNumber) rec.monthKey = matched.monthKey;
            if (matched.cents !== rec.cents) {
              rec.notes += `${rec.notes ? '\n' : ''}谈判金额${rec.amount}与排期${matched.coordinate}金额${matched.amount}不一致，保留谈判金额。`;
              issue('warnings', 'SCHEDULE_AMOUNT_CONFLICT', '谈判金额与对应排期金额不一致', sheetName, row, matched.coordinate);
            }
            if (matched.month !== part.monthNumber) issue('warnings', 'SCHEDULE_MONTH_CONFLICT', '付款月份与申请排期表头月份不同，年份不从冲突表头推定', sheetName, row, matched.coordinate);
            if (matched.month === part.monthNumber && matched.monthKey && rec.monthKey && matched.monthKey !== rec.monthKey) issue('warnings', 'SCHEDULE_YEAR_CONFLICT', '谈判明细年份与排期表头年份不一致，保留谈判年份且不向其他分期传递冲突年份', sheetName, row, matched.coordinate);
            if (matched.cents != null) rec.status = isRed(matched.cell) ? '已支付' : '未支付';
            const ev = addEvidence(rec, sheetName, row, matched.coordinate, matched.cents, rec.status, rec.monthKey, ordinal, 'schedule');
            if (ev) rec.evidenceIds.push(ev);
          } else {
            rec.notes += `${rec.notes ? '\n' : ''}无唯一对应排期，付款状态待核对。`;
            issue('warnings', 'UNMATCHED_INSTALLMENT', '谈判分期无唯一对应金额单元格', sheetName, row);
          }
          if (!rec.monthKey) issue('warnings', 'UNKNOWN_MONTH_YEAR', '分期月份没有可唯一确定的年份', sheetName, row);
          else {
            const [year, month] = rec.monthKey.split('-').map(Number);
            rec.month = Math.round((Date.UTC(year, month - 1, 1) - Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30)) / 86400000);
          }
          if (role !== 'evidence') result.records.push(rec);
        }
        for (const item of schedule.filter(s => !s.used)) {
          addEvidence(recBase, sheetName, row, item.coordinate, item.cents, isRed(item.cell) ? '已支付' : '未支付', item.monthKey, schedule.indexOf(item) + 1, 'schedule');
          issue('warnings', 'UNASSIGNED_SCHEDULE', '排期金额未唯一对应到谈判分期，证据独立保留', sheetName, row, item.coordinate);
        }
      }
      result.sheets.push({ name: sheetName, headerRow, recordCount: result.records.length - initialCount, skipped: false });
    }
    if (role === 'evidence') {
      const detailed = result.evidence.filter(e => !summaryEvidenceIds.has(e.id));
      result.evidence = result.evidence.filter(e => {
        const duplicate = summaryEvidenceIds.has(e.id) && norm(e.oa) && detailed.some(d => norm(d.oa) === norm(e.oa) && norm(d.supplier) === norm(e.supplier) && norm(d.subject) === norm(e.subject) && d.cents === e.cents && d.status === e.status);
        if (duplicate) issue('warnings', 'DUPLICATE_SUMMARY_EVIDENCE', '该汇总付款证据已有逐期或本笔明细支持，不重复分配到其他期次', e.source.sheet, e.source.row, e.source.cell);
        return !duplicate;
      });
    }
    if (!result.sheets.some(s => !s.skipped)) issue('errors', 'NO_PAYMENT_SHEET', '没有找到可导入的付款明细工作表', '', 0);
    return result;
  }
  return { parse };
});
