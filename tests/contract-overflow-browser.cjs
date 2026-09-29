const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');
(async () => {
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage();
    await page.route('**/contract-tool.js*', async route => {
      const response = await route.fetch();
      const source = (await response.text()).replace(/\}\)\(\);\s*$/, `window.layoutTest = async bytes => {
        state.templateSheet = '合同';
        const zip = await JSZip.loadAsync(new Uint8Array(bytes));
        const context = await locateXlsxSheet(zip, '合同');
        await fitExcelRowHeights(zip, context);
        await fitExcelContractMainPage(zip, context);
        await forceWorkbookRecalculation(zip, context);
        zip.file(context.sheetPath, new XMLSerializer().serializeToString(context.sheetDoc));
        return Array.from(await zip.generateAsync({type:'uint8array'}));
      };})();`);
      await route.fulfill({ response, body: source });
    });
    await page.goto(process.env.QA_URL || 'http://127.0.0.1:8899/contract.html');
    await page.waitForFunction(() => window.layoutTest && window.XLSX);
    for (const mode of ['single', 'horizontal', 'vertical', 'multiple']) {
      const result = await page.evaluate(async mode => {
        const text = '完整保留条款：交货验收与付款约定，禁止丢失任何文字。\n'.repeat(['horizontal','vertical'].includes(mode) ? 32 : 90);
        const sheet = XLSX.utils.aoa_to_sheet([['虚构合同']]);
        sheet.C18 = { t: 's', v: text };
        sheet.A24 = { t: 's', v: '通用条款' };
        sheet.A25 = { t: 's', v: '附件一：供应方廉洁诚信承诺书' };
        sheet.C23 = { t: 's', v: '后续签署日期' };
        sheet.B23 = { t: 'n', f: 'SUM(A23:A24)', v: 0 };
        sheet['!ref'] = 'A1:F25'; sheet['!cols'] = Array.from({length:6},()=>({wch:18}));
        if (mode === 'horizontal') sheet['!merges'] = [XLSX.utils.decode_range('C18:F18')];
        if (mode === 'vertical') sheet['!merges'] = [XLSX.utils.decode_range('C18:F20')];
        if (mode === 'multiple') sheet.D18 = { t:'s', v: text };
        const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, '合同');
        const ref = XLSX.utils.aoa_to_sheet([['参考']]); ref.A2={t:'n',f:'合同!B23',v:0}; ref['!ref']='A1:A2'; XLSX.utils.book_append_sheet(book,ref,'参考');
        const output = await window.layoutTest(Array.from(new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'}))));
        const generated = XLSX.read(new Uint8Array(output),{type:'array'}), target=generated.Sheets['合同'];
        const texts = column => Object.entries(target).filter(([a,c])=>new RegExp('^'+column+'\\d+$').test(a)&&c.t==='s'&&c.v!=='后续签署日期').sort((a,b)=>Number(a[0].slice(1))-Number(b[0].slice(1))).map(([,c])=>c.v).join('');
        const signature=Object.keys(target).find(a=>target[a]?.v==='后续签署日期'), row=Number(signature.slice(1));
        const zip=await JSZip.loadAsync(new Uint8Array(output)), xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
        const doc=new DOMParser().parseFromString(xml,'application/xml'), ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
        return {output, text, chunks:Object.entries(target).filter(([a,c])=>/^C\d+$/.test(a)&&c.t==='s'&&c.v!=='后续签署日期').map(([,c])=>c.v), c:texts('C'), d:mode==='multiple'?texts('D'):null, row,
          formula:[...doc.getElementsByTagNameNS(ns,'c')].find(n=>n.getAttribute('r')==='B'+row)?.getElementsByTagNameNS(ns,'f')[0]?.textContent, reference:(await zip.file('xl/worksheets/sheet2.xml').async('string')).match(/<f>(.*?)<\/f>/)?.[1],
          heights:[...doc.getElementsByTagNameNS(ns,'row')].map(n=>Number(n.getAttribute('ht')||15)),
          merges:[...doc.getElementsByTagNameNS(ns,'mergeCell')].map(n=>XLSX.utils.decode_range(n.getAttribute('ref'))),
          breaks:[...doc.getElementsByTagNameNS(ns,'brk')].map(n=>Number(n.getAttribute('id')))};
      }, mode);
      assert.equal(result.c,result.text,mode+' retains every character');
      if(result.d) assert.equal(result.d,result.text);
      for(let i=0;i<result.merges.length;i++) for(let j=i+1;j<result.merges.length;j++) {
        const a=result.merges[i],b=result.merges[j];
        assert.ok(a.e.r<b.s.r || b.e.r<a.s.r || a.e.c<b.s.c || b.e.c<a.s.c, '续排合并区域不得重叠');
      }
      if (['horizontal','vertical'].includes(mode)) {
        assert.equal(result.row,23,'原合并区域不得拆分或插入续排行');
        assert.equal(result.merges.length,1);
        assert.equal(result.chunks.length,1);
      } else assert.ok(result.row>23);
      assert.equal(result.formula,`SUM(A${result.row}:A${result.row+1})`);
      assert.equal(result.reference,`合同!B${result.row}`);
      assert.ok(result.heights.every(h=>h<=(['horizontal','vertical'].includes(mode)?409.5:220)), '续排行高必须按实际内容计算，不保留固定大空白');
      assert.ok(result.chunks.every(t=>t.startsWith('完整保留条款：')&&t.endsWith('\n')), '可容纳的完整段落不得从中间截开');
      assert.ok(result.breaks.includes(result.row));
      assert.ok(result.breaks.includes(result.row+1));
      if(mode==='horizontal' && process.env.QA_OUTPUT) {await fs.mkdir(process.env.QA_OUTPUT,{recursive:true}); await fs.writeFile(path.join(process.env.QA_OUTPUT,'overflow.xlsx'),Buffer.from(result.output));}
    }
    console.log('PASS: C18 single, horizontal, vertical, multiple overflow; exact text, formulas, page breaks and row limits');
  } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
