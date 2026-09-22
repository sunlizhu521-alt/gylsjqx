(function (root) {
  'use strict';
  const fields = ['供应商全称', '开户银行及账号', '合同编号', '合同总金额', '已付金额', '本次申请付款金额', '票据状态', '付款事由', '备注'];
  const required = ['申请日期', '付款主体', '申请部门', '经办人', '付款总金额', ...fields];
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
    const candidates = directory.index.get(name) || [];
    const scoped = candidates.filter(e => e.subject);
    const matches = scoped.length ? candidates.filter(e => e.subject === subjectName(record['付款主体'])) : candidates;
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
  function date(v) {
    if (v === '' || v == null) return '';
    let s = text(v).trim().replace(/[年月/.]/g, '-').replace(/日$/, '');
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]00:00:00)?$/);
    if (!m) throw new Error('日期应为 YYYY-MM-DD');
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (d.getUTCFullYear() !== +m[1] || d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) throw new Error('日期无效');
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
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
      if (!subject) { errors.push(`第 ${source} 行：缺少付款主体`); continue; }
      if (!groups.has(subject)) groups.set(subject, { subject, rows: [], total: 0n, info: {}, choices: {}, checks: [] });
      const group = groups.get(subject);
      const row = { ...record };
      try { row._cents = cents(record['本次申请付款金额']); group.total += row._cents; }
      catch (e) { errors.push(`第 ${source} 行：本次申请付款金额${e.message}`); }
      if (group.total > MAX_CENTS) errors.push(`${subject}：合计超出支持范围`);
      if (text(record['付款总金额']).trim()) {
        try { row._check = cents(record['付款总金额']); }
        catch (e) { errors.push(`第 ${source} 行：付款总金额${e.message}`); }
      }
      try { row['申请日期'] = date(record['申请日期']); }
      catch (e) { errors.push(`第 ${source} 行：申请日期${e.message}`); }
      group.rows.push(row);
    }
    for (const group of groups.values()) {
      for (const field of infoFields) {
        const choices = [...new Set(group.rows.map(r => text(r[field]).trim()))];
        group.choices[field] = choices;
        group.info[field] = choices.length === 1 ? choices[0] : '';
      }
      group.info['计划付款日期'] = '';
      group.checks = group.rows.filter(r => r._check != null && r._check !== group.total)
        .map(r => `第 ${r._row} 行付款总金额 ${money(r._check)}，明细合计 ${money(group.total)}`);
    }
    if (!records.length) errors.push('工作表没有付款明细');
    return { errors, groups: [...groups.values()] };
  }
  const api = { fields, required, infoFields, cents, money, upper, text, date, analyze, subjectName, bankSummary, buildDirectory, resolveBank };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PaymentCore = api;
})(typeof window === 'undefined' ? globalThis : window);
