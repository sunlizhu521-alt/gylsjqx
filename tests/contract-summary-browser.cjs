const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const XLSX = require('../vendor/xlsx.full.min.js');
const assert = require('node:assert/strict');
(async () => {
 const browser = await chromium.launch({ headless: true, executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' });
 try {
  const page = await browser.newPage();
  await page.goto(process.env.QA_URL || 'http://127.0.0.1:8940/contract.html');
  for (const [price, label] of [[29999.99,'￥29,999.99（＜3万）'],[30000,'￥30,000.00（＝3万）'],[30000.01,'￥30,000.01（＞3万）'],[0,'￥0.00（＜3万）'],['异常','合同总金额：待核对']]) {
   const book = XLSX.utils.book_new();
   XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['序号','物料编码','名称','数量','含税单价'],[1,'TEST001','测试产品',1,price]]),'Sheet1');
   await page.locator('#orderFile').setInputFiles({name:'虚构订单.xls',mimeType:'application/vnd.ms-excel',buffer:Buffer.from(XLSX.write(book,{type:'buffer',bookType:'biff8'}))});
   await page.waitForFunction(expected=>document.querySelector('#fileSummary').textContent.includes(expected),label);
   assert((await page.locator('#fileSummary').innerText()).includes('虚构订单.xls（1行）'));
  }
  console.log('PASS order summary total, below/equal/above 30000, zero and invalid amount');
 } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exitCode=1;});
