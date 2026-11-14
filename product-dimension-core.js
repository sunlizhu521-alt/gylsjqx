(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'),require('./product-dimension-enrichment.js'));
  else root.ProductDimensionCore = factory(root.XLSX,root.ProductDimensionEnrichment);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(XLSX, Enrichment) {
  'use strict';
  const MAPPING = [
    ['物料编码', '编码'], ['领星SKU', '领星SKU'], ['SKU', 'SKU'],
    ['金蝶名称', '名称'], ['金蝶创建时间', '创建日期'], ['物料分组', '物料分组']
  ];
  const value = cell => cell?.v ?? '';
  const text = value => String(value ?? '').trim();
  const date1904 = book => ['true', '1'].includes(String(book?.Workbook?.WBProps?.date1904).toLowerCase());
  const nonempty = cell => !!cell?.f || cell?.t === 'e' || value(cell) !== '';
  function signature(cells) {
    return JSON.stringify(cells.map(cell => value(cell) === '' ? ['', ''] : [cell?.t || typeof value(cell), value(cell)]));
  }
  function issue(issues, code, message, details = {}) {
    issues.push({severity:'error', code, message, ...details});
  }
  function locate(book, role, issues) {
    const label = role === 'dimension' ? '商品维度表' : '物料表';
    const required = [...MAPPING.map(pair => pair[role === 'dimension' ? 0 : 1]),...(role==='dimension'?Enrichment.FIELD_HEADERS:[])];
    const matches = [];
    if (!book?.SheetNames?.length) {
      issue(issues, 'missing-workbook', `请上传${label}`, {role});
      return null;
    }
    for (const name of book.SheetNames) {
      const sheet = book.Sheets[name];
      if (!sheet?.['!ref']) continue;
      const range = XLSX.utils.decode_range(sheet['!ref']);
      for (let r = range.s.r; r <= Math.min(range.e.r, 19); r++) {
        const headers = Array.from({length:range.e.c + 1}, (_, c) => text(value(sheet[XLSX.utils.encode_cell({r, c})])));
        if (required.every(header => headers.includes(header))) matches.push({name, sheet, range, headers, headerRow:r + 1});
      }
    }
    if (matches.length !== 1) {
      issue(issues, matches.length ? 'ambiguous-sheet' : 'missing-headers', matches.length
        ? `${label}有多个匹配的工作表或表头，请保留一个：${matches.map(m => `${m.name}第${m.headerRow}行`).join('、')}`
        : `${label}前20行未找到完整表头：${required.join('、')}`, {role});
      return null;
    }
    const match = matches[0];
    const duplicated = required.filter(header => match.headers.indexOf(header) !== match.headers.lastIndexOf(header));
    if (duplicated.length) {
      issue(issues, 'duplicate-header', `${label}表头重复：${duplicated.join('、')}`, {role, sheet:match.name, row:match.headerRow});
      return null;
    }
    return match;
  }
  function readRows(table, role, issues, existingCodes = new Map()) {
    const codeColumn = table.headers.indexOf(role === 'dimension' ? '物料编码' : '编码');
    const mappedColumns = new Set(MAPPING.map(pair => table.headers.indexOf(pair[1])));
    const groups = new Map();
    let total = 0, formulaCount = 0;
    for (let r = table.headerRow; r <= table.range.e.r; r++) {
      const cells = table.headers.map((_, c) => table.sheet[XLSX.utils.encode_cell({r, c})] || null);
      if (!cells.some(nonempty)) continue;
      total++;
      const codeCell = cells[codeColumn], rawCode = value(codeCell), code = text(rawCode);
      let valid = true;
      for (let c = 0; c < cells.length; c++) {
        if (role === 'material' && c !== codeColumn && (existingCodes.has(code) || !mappedColumns.has(c))) continue;
        const cell = cells[c];
        const address = XLSX.utils.encode_cell({r, c});
        const details = {role, sheet:table.name, row:r + 1, column:table.headers[c], address};
        if (cell?.f) formulaCount++;
        if (cell?.t === 'e') {
          issue(issues, 'cell-error', `${table.name}!${address}是错误值，请先核对`, details);
          valid = false;
        } else if (cell?.f && (cell.v === undefined || cell.v === null)) {
          issue(issues, 'missing-formula-cache', `${table.name}!${address}的公式没有计算结果，请先在Excel中计算并保存`, details);
          valid = false;
        } else if (typeof cell?.v === 'number' && !Number.isFinite(cell.v)) {
          issue(issues, 'invalid-number', `${table.name}!${address}不是有效数值`, details);
          valid = false;
        }
      }
      const details = {role, sheet:table.name, row:r + 1, column:table.headers[codeColumn], address:XLSX.utils.encode_cell({r, c:codeColumn})};
      if (!code) {
        issue(issues, 'empty-code', `${table.name}第${r + 1}行有数据但物料编码为空`, details);
        valid = false;
      } else if (codeCell?.t !== 'e' && !((typeof rawCode === 'string') || (typeof rawCode === 'number' && Number.isSafeInteger(rawCode)))) {
        issue(issues, 'invalid-code', `${table.name}第${r + 1}行物料编码不是文本或安全整数，请将原始编码保存为文本`, details);
        valid = false;
      }
      if (!valid) continue;
      const row = {rowNumber:r + 1, cells, values:cells.map(value)};
      if (!groups.has(code)) groups.set(code, {code, rows:[]});
      groups.get(code).rows.push(row);
    }
    return {groups, total, formulaCount};
  }
  function analyze(dimensionWorkbook, materialWorkbook, settings = {}) {
    const issues = [];
    const dimension = locate(dimensionWorkbook, 'dimension', issues);
    const material = locate(materialWorkbook, 'material', issues);
    const result = {
      dimensionSheet:dimension?.name || '', materialSheet:material?.name || '',
      headerRow:dimension?.headerRow || 0, materialHeaderRow:material?.headerRow || 0,
      headers:dimension?.headers || [], groups:[], added:[], counts:{}, issues
    };
    if (!dimension || !material) return result;
    const target1904 = date1904(dimensionWorkbook), source1904 = date1904(materialWorkbook);
    const existing = readRows(dimension, 'dimension', issues);
    const incoming = readRows(material, 'material', issues, existing.groups);
    result.groups = [...existing.groups.values()].map(group => ({
      ...group, conflict:group.rows.some(row => signature(row.cells) !== signature(group.rows[0].cells))
    }));
    for (const group of result.groups.filter(group => group.conflict)) {
      const keptRow = group.rows[0].rowNumber, droppedRows = group.rows.slice(1).map(row => row.rowNumber);
      issues.push({severity:'warning', code:'dimension-duplicate-conflict', role:'dimension', sheet:dimension.name,
        row:keptRow, keptRow, droppedRows,
        message:`编码${group.code}的重复行内容不同，保留原第${keptRow}行，去除原第${droppedRows.join('、')}行`});
    }
    for (const group of incoming.groups.values()) {
      if (existing.groups.has(group.code)) continue;
      const rows = group.rows.map(row => {
        const cells = dimension.headers.map(() => null);
        for (const [target, source] of MAPPING) {
          const sourceColumn = material.headers.indexOf(source);
          let sourceCell = row.cells[sourceColumn];
          if (target === '金蝶创建时间' && typeof sourceCell?.v === 'number' && source1904 !== target1904) {
            const serial = sourceCell.v + (source1904 ? 1462 : -1462);
            const before = XLSX.SSF.parse_date_code(sourceCell.v, {date1904:source1904});
            const after = XLSX.SSF.parse_date_code(serial, {date1904:target1904});
            if (!before || !after || ['y','m','d','H','M','S'].some(part => before[part] !== after[part])) {
              issue(issues, 'incompatible-date-system', `${material.name}第${row.rowNumber}行创建日期不能转换到商品维度表的日期系统，请先核对源文件`,
                {role:'material', sheet:material.name, row:row.rowNumber, column:source, address:XLSX.utils.encode_cell({r:row.rowNumber - 1, c:sourceColumn})});
            } else {
              sourceCell = {...sourceCell, v:serial};
              delete sourceCell.w;
            }
          }
          const c = dimension.headers.indexOf(target);
          cells[c] = target === '物料编码' ? {t:'s', v:group.code}
            : ['SKU', '领星SKU'].includes(target) ? {t:'s', v:String(value(sourceCell))}
            : sourceCell;
        }
        return {rowNumber:row.rowNumber, cells, values:cells.map(value)};
      });
      const conflict = rows.some(row => signature(row.cells) !== signature(rows[0].cells));
      result.added.push({code:group.code, rows, sourceRows:rows.map(row => row.rowNumber), values:rows[0].values, conflict});
      if (conflict) issue(issues, 'new-material-conflict', `新增编码${group.code}在物料表第${rows.map(row => row.rowNumber).join('、')}行的映射内容不同，请先核对源文件`,
        {role:'material', sheet:material.name, row:rows[0].rowNumber, sourceRows:rows.map(row => row.rowNumber)});
    }
    result.counts = {
      dimensionRows:existing.total, dimensionUnique:existing.groups.size,
      dimensionDuplicateRows:[...existing.groups.values()].reduce((n, group) => n + group.rows.length - 1, 0),
      dimensionDuplicateGroups:result.groups.filter(group => group.rows.length > 1).length,
      dimensionConflictGroups:result.groups.filter(group => group.conflict).length,
      materialRows:incoming.total, materialUnique:incoming.groups.size,
      materialDuplicateRows:[...incoming.groups.values()].reduce((n, group) => n + group.rows.length - 1, 0),
      materialConflictGroups:result.added.filter(group => group.conflict).length,
      addedUnique:result.added.length, finalUnique:existing.groups.size + result.added.length,
      formulaCount:existing.formulaCount
    };
    return Enrichment.prepare(result,{...settings,date1904:target1904});
  }
  function resolve(analysis, overrides = {}) {
    const rows = [];
    for (const [kind, groups] of [['dimension', analysis.groups], ['material', analysis.added]]) {
      for (const group of groups) {
        const chosen = group.rows[0];
        const values = chosen.values.slice(), cells = chosen.cells.slice(), codeColumn = analysis.headers.indexOf('物料编码');
        values[codeColumn] = group.code;
        cells[codeColumn] = {...cells[codeColumn], t:'s', v:group.code};
        delete cells[codeColumn].f;
        delete cells[codeColumn].F;
        delete cells[codeColumn].w;
        rows.push({code:group.code, sourceRow:chosen.rowNumber, values, cells, isNew:kind === 'material'});
      }
    }
    const blocked = analysis.issues.some(item => item.severity === 'error');
    const enriched = !blocked ? Enrichment.enrich(analysis,rows,overrides) : {rows:[],pendingFields:[],pendingRows:0};
    const result = {...enriched,unresolved:[],counts:{...analysis.counts,resolvedRows:enriched.rows.length,pendingRows:enriched.pendingRows},issues:analysis.issues,canExport:false};
    result.canExport=!blocked&&!enriched.pendingFields.length&&!Enrichment.validateResult(analysis,result).length;
    return result;
  }
  return {MAPPING, analyze, resolve,REQUIRED_HEADERS:Enrichment.REQUIRED_HEADERS,EDITABLE_HEADERS:Enrichment.EDITABLE_HEADERS,FIELD_HEADERS:Enrichment.FIELD_HEADERS,validateResult:Enrichment.validateResult};
});
