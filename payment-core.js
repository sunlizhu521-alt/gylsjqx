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
      const subject = text(record['付款主体']).trim(), source = record._row;
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
  const api = { fields, required, infoFields, cents, money, upper, text, date, analyze };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PaymentCore = api;
})(typeof window === 'undefined' ? globalThis : window);
