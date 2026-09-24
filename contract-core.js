(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.ContractCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const INVALID_FILE_NAME = /[\\/:*?"<>|\x00-\x1f]/g;

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
    resolveField,
    mappingIssues,
  };
});
