(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), require('./vendor/jszip.min.js'), require('./table-compare-core.js'));
  else root.TableCompareExport = factory(root.XLSX, root.JSZip, root.TableCompareCore);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(X, Zip, Core) {
  'use strict';
  const wrapColumns = new Set([5, 10, 11, 12, 13, 14]);
  const width = value => [...String(value)].reduce((n, ch) => n + (ch.codePointAt(0) > 255 ? 2 : 1), 0);
  const displayWidth = value => Math.max(0, ...String(value).split(/\r?\n/).map(width));
  function valueCell(cell) {
    if (!cell || cell.v == null) return {t:'s', v:'', z:'@'};
    if (cell.t === 'd' || cell.t === 'n' && cell.z && X.SSF.is_date(cell.z)) return {t:'s', v:Core.formatCell(cell), z:'@'};
    return {t:cell.t, v:cell.v, z:cell.t === 'n' ? '#,##0.00' : '@'};
  }
  return async function(result) {
    const source = result.sheets[0], book = X.utils.book_new();
    const rows = source.rows.map(row => row.map(valueCell));
    const sheet = X.utils.aoa_to_sheet([source.headers, ...rows]);
    sheet['!cols'] = source.headers.map((header, c) => {
      const content = rows.reduce((max, row) => Math.max(max, displayWidth(Core.formatCell(row[c]))), 0);
      return {wch:Math.max(12, width(header) + 4, (wrapColumns.has(c) ? Math.min(64, content) : content) + 3)};
    });
    sheet['!autofilter'] = {ref:'A1:' + X.utils.encode_col(source.headers.length - 1) + (rows.length + 1)};
    X.utils.book_append_sheet(book, sheet, '比对汇总');
    const zip = await Zip.loadAsync(X.write(book, {type:'array', bookType:'xlsx'}));
    let styles = await zip.file('xl/styles.xml').async('string');
    const original = styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g);
    const count = original.length, formats = original.map(s => s.match(/\bnumFmtId="(\d+)"/)?.[1] || '0');
    const xf = (format, wrap, header = false) => `<xf numFmtId="${format}" fontId="${header ? 1 : 0}" fillId="${header ? 3 : 2}" borderId="${header ? 1 : 2}" xfId="0" applyNumberFormat="1" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="${wrap ? 1 : 0}" shrinkToFit="0"/></xf>`;
    const fonts = '<fonts count="2"><font><sz val="11"/><color rgb="FF172033"/><name val="等线"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font></fonts>';
    const fills = '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF315B9D"/><bgColor indexed="64"/></patternFill></fill></fills>';
    const borders = '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' + ['thin','hair'].map(style => '<border>' + ['left','right','top','bottom'].map(side => `<${side} style="${style}"><color rgb="FF404040"/></${side}>`).join('') + '<diagonal/></border>').join('') + '</borders>';
    styles = styles.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/, fonts).replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, fills).replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, borders).replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${count * 2 + 1}">${formats.map(f => xf(f, false)).join('')}${formats.map(f => xf(f, true)).join('')}${xf(0, false, true)}</cellXfs>`);
    zip.file('xl/styles.xml', styles);
    let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (tag, attrs, end) => {
      const address = attrs.match(/\br="([A-Z]+)(\d+)"/);
      if (!address) return tag;
      const old = Number(attrs.match(/\bs="(\d+)"/)?.[1] || 0);
      const style = address[2] === '1' ? count * 2 : old + (wrapColumns.has(X.utils.decode_col(address[1])) ? count : 0);
      return '<c' + attrs.replace(/\s+s="\d+"/, '') + ` s="${style}"${end}>`;
    });
    xml = xml.replace(/<row\b([^>]*)>/g, (tag, attrs) => {
      const r = Number(attrs.match(/\br="(\d+)"/)?.[1]);
      const lines = r === 1 ? 1 : Math.max(1, ...[...wrapColumns].map(c => Core.formatCell(rows[r - 2]?.[c]).split(/\r?\n/).reduce((n, line) => n + Math.max(1, Math.ceil(width(line) / Math.max(1, sheet['!cols'][c].wch - 2))), 0)));
      const height = r === 1 ? 28 : Math.min(409, Math.max(26, lines * 15 + 8));
      return '<row' + attrs.replace(/\s+(?:ht|customHeight)="[^"]*"/g, '') + ` ht="${height}" customHeight="1">`;
    });
    const views = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
    xml = /<sheetViews\b/.test(xml) ? xml.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, views) : xml.replace(/<sheetData\b/, views + '<sheetData');
    zip.file('xl/worksheets/sheet1.xml', xml);
    return zip.generateAsync({type:'uint8array', compression:'DEFLATE'});
  };
});
