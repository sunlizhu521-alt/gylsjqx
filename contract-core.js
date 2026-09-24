(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.ContractCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const INVALID_FILE_NAME = /[\\/:*?"<>|\x00-\x1f]/g;
  const MAX_AMOUNT_CENTS = 99999999999999n;

  function text(value) {
    return value === null || value === undefined ? '' : String(value);
  }

  function cleanHeader(value) {
    return text(value).replace(/\s+/g, ' ').trim();
  }

  function sanitizeFileName(value) {
    const cleaned = text(value).replace(INVALID_FILE_NAME, '_').replace(/\.+$/g, '').trim();
    return cleaned || '合同';
  }

  function findHeaderIndex(matrix) {
    const hints = /(?:序号|编码|名称|物料|商品|产品|规格|型号|SKU|单位|数量|单价|金额|合计|备注|交货|供应商|采购方|订单号|合同编号|日期|税率|付款)/i;
    let best = null;
    for (let index = 0; index < Math.min(matrix.length, 50); index += 1) {
      const row = Array.isArray(matrix[index]) ? matrix[index] : [];
      const values = row.map(cleanHeader).filter(Boolean);
      if (!values.length) continue;
      const unique = new Set(values);
      const hintCount = values.filter(value => hints.test(value)).length;
      const numericCount = values.filter(value => /^[-+]?\d+(?:[.,]\d+)?$/.test(value)).length;
      const longTextCount = values.filter(value => value.length > 40).length;
      const score = values.length * 10 + hintCount * 30 + unique.size - numericCount * 4 - longTextCount * 5;
      if (!best || score > best.score) best = { index, score };
    }
    return best?.index ?? -1;
  }

  function analyzeMatrix(matrix, maxRows = 1000) {
    const headerIndex = findHeaderIndex(matrix);
    if (headerIndex < 0) throw new Error('订单工作表没有数据');
    const headers = matrix[headerIndex].map(cleanHeader);
    const populated = headers.filter(Boolean);
    if (!populated.length) throw new Error('订单工作表没有表头');
    if (new Set(populated).size !== populated.length) throw new Error('订单表头存在重复字段，请修改后重新上传');

    const rows = [];
    for (let index = headerIndex + 1; index < matrix.length; index += 1) {
      const row = matrix[index] || [];
      if (!row.some(value => cleanHeader(value))) continue;
      const record = { _row: index + 1 };
      headers.forEach((header, column) => {
        if (header) record[header] = row[column] ?? '';
      });
      rows.push(record);
      if (rows.length > maxRows) throw new Error(`订单明细超过 ${maxRows} 行，请拆分后处理`);
    }
    if (!rows.length) throw new Error('订单工作表只有表头，没有明细数据');
    return { headerIndex, headers: populated, rows };
  }

  function normalizeLabel(value) {
    return cleanHeader(value).toUpperCase().replace(/[^A-Z0-9\u4e00-\u9fff]/g, '');
  }

  function extractAdjacentLabelValues(matrix, aliasesByKey, mergedRanges = [], excludedRowIndexes = []) {
    const definitions = Object.entries(aliasesByKey || {}).map(([key, aliases]) => ({
      key,
      aliases: (aliases || []).map(normalizeLabel).filter(Boolean),
    }));
    const allLabels = new Set(definitions.flatMap(definition => definition.aliases));
    const excludedRows = new Set(excludedRowIndexes);
    const values = {};
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
      if (excludedRows.has(rowIndex)) continue;
      const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        const label = normalizeLabel(row[columnIndex]);
        if (!label) continue;
        const definition = definitions.find(item => item.aliases.includes(label));
        if (!definition || values[definition.key] !== undefined) continue;
        const merged = mergedRanges.find(range => rowIndex >= range.s.r && rowIndex <= range.e.r && columnIndex >= range.s.c && columnIndex <= range.e.c);
        const startColumn = merged ? merged.e.c + 1 : columnIndex + 1;
        for (let valueColumn = startColumn; valueColumn < Math.min(row.length, startColumn + 4); valueColumn += 1) {
          const candidate = cleanHeader(row[valueColumn]);
          if (!candidate) continue;
          if (allLabels.has(normalizeLabel(candidate))) break;
          values[definition.key] = candidate;
          break;
        }
      }
    }
    return values;
  }

  function distinctValues(rows, field) {
    const seen = new Set();
    const values = [];
    for (const row of rows) {
      const value = text(row[field]).trim();
      if (!value || seen.has(value)) continue;
      seen.add(value);
      values.push(value);
    }
    return values;
  }

  function selectDetailRows(rows, headers = [], fieldSelections = {}) {
    const sourceRows = Array.isArray(rows) ? rows : [];
    const sequenceField = (headers || []).find(header => /^(?:序号|行号|明细序号)$/.test(normalizeLabel(header)));
    if (sequenceField) {
      const sequenced = sourceRows.filter(row => {
        const value = text(row[sequenceField]).trim().replace(/[、.．。]$/, '');
        return /^\d+(?:\.0+)?$/.test(value) && Number(value) >= 0;
      });
      if (sequenced.length) return sequenced;
    }

    const isIdentity = value => {
      const normalized = text(value).replace(/[\s：:]/g, '');
      return !!normalized && !/^(序号|合计|总计|小计|交货时间|交货日期|交期|付款方式|付款条件|备注|说明|签字|盖章)$/.test(normalized);
    };
    const identityFields = ['materialCode', 'materialName', 'sku'].map(key => fieldSelections[key]).filter(field => field && field !== '@sequence');
    if (identityFields.length) {
      const matched = sourceRows.filter(record => identityFields.some(field => isIdentity(record[field])));
      if (matched.length) return matched;
    }
    const fallbackFields = ['specification', 'quantity', 'taxUnitPrice', 'taxAmount'].map(key => fieldSelections[key]).filter(field => field && field !== '@sequence');
    if (fallbackFields.length) {
      const matched = sourceRows.filter(record => fallbackFields.some(field => text(record[field]).trim()));
      if (matched.length) return matched;
    }
    return sourceRows;
  }

  function parseNumber(value) {
    const normalized = text(value).trim().replace(/[￥¥,，\s]/g, '');
    if (!normalized || !/^-?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
    const number = Number(normalized);
    return Number.isFinite(number) ? number : null;
  }

  function amountCents(value) {
    if (typeof value !== 'string' && typeof value !== 'number') throw new Error('金额不能为空或非数字');
    let normalized = String(value).trim().replace(/^[￥¥]\s*/, '');
    if (/^\d{1,3}([,，]\d{3})+(\.\d{1,2})?$/.test(normalized)) normalized = normalized.replace(/[,，]/g, '');
    if (!/^\d+(?:\.\d{1,2})?$/.test(normalized)) throw new Error('金额须为非负数，且最多两位小数');
    const [yuan, decimal = ''] = normalized.split('.');
    const cents = BigInt(yuan) * 100n + BigInt(decimal.padEnd(2, '0'));
    if (cents > MAX_AMOUNT_CENTS) throw new Error('金额超出支持范围（最高 9999 亿元）');
    return cents;
  }

  function formatAmountCents(cents) {
    return `${cents / 100n}.${(cents % 100n).toString().padStart(2, '0')}`;
  }

  function amountUpper(cents) {
    if (cents < 0n || cents > MAX_AMOUNT_CENTS) throw new Error('大写金额超出范围');
    const digits = '零壹贰叁肆伍陆柒捌玖';
    const four = value => {
      const source = String(value), units = ['', '拾', '佰', '仟'];
      let output = '', pendingZero = false;
      [...source].forEach((digit, index) => {
        if (digit === '0') { if (output) pendingZero = true; return; }
        output += (pendingZero ? '零' : '') + digits[Number(digit)] + units[source.length - index - 1];
        pendingZero = false;
      });
      return output;
    };
    let yuan = cents / 100n, output = '';
    const groups = [];
    while (yuan) { groups.unshift(Number(yuan % 10000n)); yuan /= 10000n; }
    let pendingZero = false;
    groups.forEach((value, index) => {
      if (!value) { if (output) pendingZero = true; return; }
      if (output && (pendingZero || value < 1000)) output += '零';
      output += four(value) + ['', '万', '亿'][groups.length - index - 1];
      pendingZero = false;
    });
    output = (output || '零') + '元';
    const jiao = Number(cents % 100n / 10n), fen = Number(cents % 10n);
    if (!jiao && !fen) return output + '整';
    if (jiao) output += digits[jiao] + '角';
    if (fen) output += (!jiao && cents >= 100n ? '零' : '') + digits[fen] + '分';
    return output;
  }

  function sumAmountField(rows, field) {
    if (!field) throw new Error('请先为“含税运总金额（元）”选择订单明细列');
    let total = 0n;
    for (const row of rows || []) {
      try { total += amountCents(row[field]); }
      catch (error) { throw new Error(`${field} 第 ${row._row || '?'} 行${error.message}`); }
      if (total > MAX_AMOUNT_CENTS) throw new Error(`${field} 汇总金额超出支持范围（最高 9999 亿元）`);
    }
    return { cents: total, lower: formatAmountCents(total), upper: amountUpper(total) };
  }

  function formatNumber(value) {
    if (!Number.isFinite(value)) return '';
    return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(10)));
  }

  function resolveField(rows, field, strategy = 'first', manual = '') {
    const values = distinctValues(rows, field);
    if (strategy === 'manual') return text(manual);
    if (strategy === 'merge') return values.join('、');
    if (strategy === 'sum') {
      const numbers = rows.map(row => parseNumber(row[field]));
      if (numbers.some(value => value === null)) throw new Error(`${field} 包含非数字内容，不能求和`);
      return formatNumber(numbers.reduce((total, value) => total + value, 0));
    }
    return values[0] ?? '';
  }

  function mappingIssues(rows, mappings, detailRow) {
    const issues = [];
    const entries = Object.entries(mappings || {}).filter(([, mapping]) => mapping?.field);
    if (!entries.length) issues.push('至少设置一个字段映射');
    for (const [target, mapping] of entries) {
      if (mapping.mode === 'detail' && !detailRow) issues.push(`${target} 已设为明细字段，但尚未选择明细模板行`);
      if (mapping.mode !== 'detail') {
        const values = distinctValues(rows, mapping.field);
        if (values.length > 1 && !['first', 'merge', 'manual', 'sum'].includes(mapping.strategy)) {
          issues.push(`${mapping.field} 有多个不同值，需要选择处理方式`);
        }
        if (mapping.strategy === 'manual' && !text(mapping.manual).trim()) issues.push(`${mapping.field} 的手工值不能为空`);
        if (mapping.strategy === 'sum' && rows.some(row => parseNumber(row[mapping.field]) === null)) issues.push(`${mapping.field} 含非数字内容，不能求和`);
      }
    }
    return [...new Set(issues)];
  }

  return {
    text,
    cleanHeader,
    sanitizeFileName,
    analyzeMatrix,
    extractAdjacentLabelValues,
    distinctValues,
    selectDetailRows,
    parseNumber,
    amountCents,
    formatAmountCents,
    amountUpper,
    sumAmountField,
    resolveField,
    mappingIssues,
  };
});
