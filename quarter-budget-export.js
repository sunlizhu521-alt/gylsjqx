(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'));
  else root.QuarterBudgetExport = factory(root.XLSX, root.JSZip);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip) {
  'use strict';
  const DETAIL_HEADERS = ['预估付款月份','供应商名称','付款主体','对公付款0A流程号','摘要','申请付款金额','是否已支付','备注','分类','跟单员','是否剔除'];
  const CLASSIFIED_HEADERS = ['付款月份','分类','申请总金额','已付款金额','剩余付款金额','备注'];
  const MONTHLY_HEADERS = ['付款月份','申请总金额','已付款金额','剩余付款金额','备注'];
  const AMOUNT_FORMAT = '#,##0.00';
  const width = text => [...String(text)].reduce((n, c) => n + (c.codePointAt(0) > 255 ? 2 : 1), 0);

  function cell(value, format) {
    if (value === null || value === undefined) return {t:'s',v:''};
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) throw new Error('导出金额或日期包含无效数字，请先核对。');
      return {t:'n',v:value,...(format ? {z:format} : {})};
    }
    return {t:'s',v:String(value)};
  }

  function money(cents) {
    if (!Number.isSafeInteger(cents)) throw new Error('预算汇总金额无效，请重新整理后导出。');
    return cell(cents / 100, AMOUNT_FORMAT);
  }

  function month(key) {
    if (!key) return cell('月份待核对');
    const match = /^(\d{4})-(0[1-9]|1[0-2])$/.exec(key);
    if (!match) return cell(key);
    return {...cell(Number(match[1]) + '年' + Number(match[2]) + '月'),z:'@'};
  }

  function formatCell(value) {
    return value.t === 'n' && value.z ? X.SSF.format(value.z, value.v) : String(value.v ?? '');
  }

  function rowHeight(row, columns, sheetName, index) {
    const lines = Math.max(1, ...row.map((value, c) => formatCell(value).split(/\r\n|\r|\n/).reduce((n, line) => n + Math.max(1, Math.ceil(width(line) / (columns[c].wch - 2))), 0)));
    const height = Math.max(26, lines * 16 + 10);
    if (height > 409) throw new Error(`${sheetName}第${index + 1}行内容过长，超过Excel单行可显示范围，请缩短该行说明后导出。`);
    return height;
  }

  return async function(result) {
    if (!Array.isArray(result?.records) || !Array.isArray(result.summary?.classified) || !Array.isArray(result.summary?.monthly) || !result.summary?.totals) {
      throw new Error('缺少季度预算整理结果，请先整理数据。');
    }
    const detailRows = [DETAIL_HEADERS.map(v => cell(v)), ...result.records.map(r => [
      r.monthKey ? month(r.monthKey) : {...cell(r.month == null ? '' : String(r.month)),z:'@'}, cell(r.supplier), cell(r.subject), cell(r.oa), cell(r.summary), cell(r.amount, AMOUNT_FORMAT),
      cell(r.status), cell(r.notes), cell(r.category), cell(r.owner), cell(r.excluded)
    ])];
    const total = result.summary.totals;
    const summaryRows = [CLASSIFIED_HEADERS.map(v => cell(v)), ...result.summary.classified.map(r => [
      month(r.monthKey), cell(r.category), money(r.totalCents), money(r.paidCents), money(r.remainingCents), cell('')
    ])];
    const totalRows = new Set([summaryRows.length]);
    summaryRows.push([cell('合计'), cell(''), money(total.totalCents), money(total.paidCents), money(total.remainingCents), cell('')], []);
    const summaryHeaders = new Set([0, summaryRows.length]);
    summaryRows.push(MONTHLY_HEADERS.map(v => cell(v)), ...result.summary.monthly.map(r => [
      month(r.monthKey), money(r.totalCents), money(r.paidCents), money(r.remainingCents), cell('')
    ]));
    totalRows.add(summaryRows.length);
    summaryRows.push([cell('合计'), money(total.totalCents), money(total.paidCents), money(total.remainingCents), cell('')]);
    const sheets = [
      {name:'付款汇总明细', rows:detailRows, widths:[18,34,20,29,56,18,15,76,17,15,13], headers:new Set([0]), totals:new Set()},
      {name:'付款预算汇总', rows:summaryRows, widths:[19,20,20,20,20,58], headers:summaryHeaders, totals:totalRows}
    ];
    const book = X.utils.book_new();
    for (const info of sheets) {
      const sheet = X.utils.aoa_to_sheet(info.rows);
      sheet['!cols'] = info.widths.map(wch => ({wch}));
      sheet['!rows'] = info.rows.map((row, index) => ({hpt:row.length ? rowHeight(row, sheet['!cols'], info.name, index) : 12}));
      if (info.name === '付款汇总明细') sheet['!autofilter'] = {ref:sheet['!ref']};
      info.sheet = sheet;
      X.utils.book_append_sheet(book, sheet, info.name);
    }

    const zip = await Zip.loadAsync(X.write(book, {type:'array',bookType:'xlsx'}));
    let styles = await zip.file('xl/styles.xml').async('string');
    const original = styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
    const formats = original.map(s => s.match(/\bnumFmtId="(\d+)"/)?.[1] || '0'), count = formats.length;
    const fonts = '<fonts count="3"><font><sz val="11"/><color rgb="FF172033"/><name val="等线"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font><font><b/><sz val="11"/><color rgb="FF172033"/><name val="等线"/></font></fonts>';
    const fills = '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315B9D"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill></fills>';
    const borders = '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' + ['thin','hair'].map(style => '<border>' + ['left','right','top','bottom'].map(side => `<${side} style="${style}"><color rgb="FF404040"/></${side}>`).join('') + '<diagonal/></border>').join('') + '</borders>';
    const xf = (format, font, fill, blank = false) => `<xf numFmtId="${format}" fontId="${font}" fillId="${fill}" borderId="${font===1?1:blank?0:2}" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1" shrinkToFit="0"/></xf>`;
    const bodyStyles = blank => [[0,2],[0,4],[2,2]].map(([font,fill]) => formats.map(f => xf(f,font,fill,blank)).join('')).join('');
    styles = styles.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/, fonts).replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, fills).replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, borders)
      .replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${count * 6 + 1}">${bodyStyles(false)}${xf(0,1,3)}${bodyStyles(true)}</cellXfs>`);
    zip.file('xl/styles.xml', styles);
    for (let index = 0; index < sheets.length; index++) {
      const info = sheets[index], path = `xl/worksheets/sheet${index + 1}.xml`;
      const contentRows = new Set(), contentColumns = new Set();
      info.rows.forEach((row,r) => row.forEach((cell,c) => {
        if (cell.v == null || (typeof cell.v === 'string' && !cell.v.trim())) return;
        contentRows.add(r); contentColumns.add(c);
      }));
      let xml = await zip.file(path).async('string');
      xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (tag, attrs, end) => {
        const address = attrs.match(/\br="([A-Z]+)(\d+)"/);
        if (!address) return tag;
        const row = Number(address[2]) - 1, col = X.utils.decode_col(address[1]);
        const old = Number(attrs.match(/\bs="(\d+)"/)?.[1] || 0), record = index === 0 ? result.records[row - 1] : null;
        const red = record && (record.related === true || record.excluded === '是' || (record.status === '已支付' && (col === 5 || col === 6)));
        const blank = !contentRows.has(row) || !contentColumns.has(col);
        const style = info.headers.has(row) ? count * 3 : old + (red ? count : info.totals.has(row) ? count * 2 : 0) + (blank ? count * 3 + 1 : 0);
        return '<c' + attrs.replace(/\s+s="\d+"/, '') + ` s="${style}"${end}>`;
      });
      const views = '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
      xml = /<sheetViews\b/.test(xml) ? xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, views) : xml.replace(/<(?:sheetFormatPr|cols|sheetData)\b/, match => views + match);
      zip.file(path, xml);
    }
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  };
});
