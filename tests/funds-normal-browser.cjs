const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const {chromium} = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X = require('../vendor/xlsx.full.min.js'), C = require('../funds-core.js');
const base = process.env.QA_URL || 'http://127.0.0.1:8963/';
const FLOW = ['当前节点','未操作者','流程状态','申请人','流程编号'];
const LH = ['登记月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单',...FLOW,'预估付款时间'];
const AH = ['申请日期','供应商名称','付款主体','流程编码','申请付款金额','付款分类'];
const OH = ['流程编号','收款单位','付款金额',...FLOW.slice(0,4)];
const encode = book => Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const flowCode = n => 'MDST-DGFK20261114' + String(n).padStart(3,'0');
function append(book,name,headers,records) { X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...records.map(row=>headers.map(header=>row[header] ?? ''))]),name); }
function row(name,amount,flag,oa,due,extra={}) {
  return {'登记月份':'2026-11-01','供应商名称':name,'付款主体':'合成主体','对公付款0A流程号':oa,'开票时间':'2026-11-01','开票金额':amount,'摘要':name,'申请付款金额':amount,'应付款时间':due,'单据号【财务应付单】':'','OA合同编号':'合同-'+name,'合同总金额':amount,'已付金额':0,'是否结清':amount===0?'是':'否','剩余金额':amount,'是否谈判付款':flag,'是否提交付款申请单':'旧申请状态',...Object.fromEntries(FLOW.map(key=>[key,'原'+key])),'预估付款时间':'2027-01-01',...extra};
}
function fixture() {
  const records = {
    李奇: [row('空OA唯一候选',100,'否','','2026-11-20',{'合同总金额':1000,'已付金额':300,'剩余金额':700,'是否结清':'否'}),row('空白分类多候选',101,'　\n','　','2026-11-21'),row('谈判无OA排除',201,'是','','2026-11-05'),row('谈判有OA排除',202,'是',flowCode(202),'2026-11-06')],
    徐英: [row('有OA已结清',102,'否',flowCode(102),'2026-11-22',{'已付金额':102,'剩余金额':0,'是否结清':'是'}),row('零额缺日期',0,'','','')],
    邢芳芳: [row('台账否命中谈判表',103,'否','','月底协商')],
    操思敏: [row('台账空申请分类谈判',104,'',flowCode(104),'2026-11-24')],
    孙立柱: [row('别名表头本行日期',105,'否','','2026-11-25')]
  };
  const ledger=X.utils.book_new();
  for(const name of C.PEOPLE){
    const aliases=name==='孙立柱'?LH.map(h=>h==='应付款时间'?'应付时间':h==='对公付款0A流程号'?'对公付款OA流程号':h):LH;
    const sheet=X.utils.aoa_to_sheet([aliases,...records[name].map(record=>LH.map(header=>record[header]??''))]);
    X.utils.book_append_sheet(ledger,sheet,'对账明细-'+name);
  }
  const applications=X.utils.book_new();append(applications,'付款申请汇总',AH,[
    {'申请日期':'2026-10-10','供应商名称':'空OA唯一候选','付款主体':'合成主体','流程编码':flowCode(100),'申请付款金额':100,'付款分类':'正常付款'},
    {'申请日期':'2026-10-12','供应商名称':'台账空申请分类谈判','付款主体':'合成主体','流程编码':flowCode(104),'申请付款金额':104,'付款分类':'谈判付款'}
  ]);
  const oa=X.utils.book_new();append(oa,'Sheet0',OH,[[100,'空OA唯一候选',100],[101,'空白分类多候选',101],[111,'空白分类多候选',101],[102,'有OA已结清',102],[103,'台账否命中谈判表',103],[104,'台账空申请分类谈判',104],[105,'别名表头本行日期',105],[202,'谈判有OA排除',202]].map(([id,name,amount])=>({'流程编号':flowCode(id),'收款单位':name,'付款金额':amount,'当前节点':'财务审核','未操作者':'合成人员','流程状态':'未归档','申请人':'合成申请人'})));
  const negotiated=X.utils.book_new();append(negotiated,'谈判付款',LH,[row('台账否命中谈判表',103,'是',flowCode(103),'2026-12-31',{'是否提交付款申请单':'未提交'})]);
  return {ledger,applications,oa,negotiated,records};
}
(async()=>{
  const data=fixture(),browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'});
  try {
    const page=await browser.newPage({acceptDownloads:true}),errors=[];
    page.on('pageerror',error=>errors.push(error.message));
    await page.route('**/*',route=>new URL(route.request().url()).origin===new URL(base).origin?route.continue():route.abort());
    await page.goto(base,{waitUntil:'networkidle'});await page.locator('.tool-tab[data-tool="funds"]').click();
    const frame=page.frameLocator('iframe[data-tool="funds"]');
    async function upload(role,book=data[role]) { await frame.locator('#'+role+'File:not([disabled])').setInputFiles({name:'合成-'+role+'.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:encode(book)}); }
    async function analyze() { await frame.locator('#analyze:not([disabled])').click();await frame.locator('#normalPreview').waitFor(); }
    async function download() { await frame.locator('#download:not([disabled])').waitFor({timeout:60000});const wait=page.waitForEvent('download');await frame.locator('#download').click();const file=await wait;assert.match(file.suggestedFilename(),/^付款节奏跟进-\d{4}-\d{2}-\d{2}-V\d+\.xlsx$/);return X.read(await fs.readFile(await file.path()),{cellStyles:true,cellNF:true}); }
    for(const role of ['ledger','applications','oa','negotiated'])await upload(role);
    await analyze();const book=await download(),sheet=book.Sheets.正常付款,grid=X.utils.sheet_to_json(sheet,{header:1,defval:''}),headers=grid[0],rows=grid.slice(1),col=key=>headers.indexOf(key),find=name=>rows.find(row=>row[col('供应商名称')]===name);
    assert.deepEqual(headers,C.NORMAL_HEADERS);assert.equal(rows.length,7);
    assert(!find('谈判无OA排除'));assert(!find('谈判有OA排除'));
    const expectedDates={'空OA唯一候选':'2026年11月20日','空白分类多候选':'2026年11月21日','有OA已结清':'2026年11月22日','零额缺日期':'','台账否命中谈判表':'月底协商','台账空申请分类谈判':'2026年11月24日','别名表头本行日期':'2026年11月25日'};
    const originalAmounts=Object.fromEntries(Object.values(data.records).flat().map(record=>[record['供应商名称'],record['申请付款金额']]));
    const blankNames=['空OA唯一候选','空白分类多候选','零额缺日期','台账否命中谈判表','别名表头本行日期'];
    for(const [name,expected] of Object.entries(expectedDates)){
      const record=find(name);assert(record,name+'应纳入');assert.equal(record[col('应付时间')],expected,name+'应付时间取台账本行');assert.equal(record[col('本次申请付款金额')],originalAmounts[name]);
      const excelRow=rows.indexOf(record)+2;
      if(record[col('跟单员')]==='李奇')assert.equal(sheet['O'+excelRow].f,undefined,'李奇保留原余额值，不使用申请金额重算');
      else assert.equal(sheet['O'+excelRow].f,`G${excelRow}-M${excelRow}`,'其他人余额公式跟随实际行');
      if(blankNames.includes(name)){assert.equal(record[col('对公付款0A流程号')],'未提交OA');assert.equal(record[col('是否对账')],'未对账');for(const field of FLOW)assert.equal(record[col(field)],'',name+'不填'+field);}
    }
    assert.deepEqual(['合同总金额','已付金额','剩余金额','是否结清'].map(h=>find('空OA唯一候选')[col(h)]),[1000,300,700,'否']);
    assert.equal(find('有OA已结清')[col('对公付款0A流程号')],flowCode(102));assert.equal(find('有OA已结清')[col('流程编号')],flowCode(102));assert.equal(find('有OA已结清')[col('已付金额')],102);assert.equal(find('有OA已结清')[col('剩余金额')],0);
    assert.equal(find('零额缺日期')[col('本次申请付款金额')],0);assert.equal(find('零额缺日期')[col('是否结清')],'是');assert.equal(find('零额缺日期')[col('申请付款的日期')],'日期待确认');
    assert.equal(find('空OA唯一候选')[col('是否提交付款申请单')],'已提交','纸质申请与OA号独立');
    assert.equal(find('台账空申请分类谈判')[col('是否提交付款申请单')],'未申请','申请分类不改变台账纳入，只影响申请归属');
    const preview=await frame.locator('#normalPreview tbody tr').evaluateAll(nodes=>nodes.map(row=>[...row.cells].map(cell=>cell.textContent)));
    for(const record of rows){const shown=preview.find(row=>row[col('供应商名称')]===record[col('供应商名称')]);for(const key of ['对公付款0A流程号','应付时间',...FLOW])assert.equal(shown[col(key)],String(record[col(key)]));}
    const redRow=rows.indexOf(find('空OA唯一候选'))+2;
    assert(/^FF0000$/.test(sheet['A'+redRow].s.fgColor.rgb),'已匹配纸质申请行标红，排序后跟随记录');
    assert.equal(book.Sheets.谈判付款.I2.v,'2026-12-31','谈判原应付日期保留');
    assert.equal(await frame.locator('#pendingCount').innerText(),'0');
    assert.doesNotMatch(await frame.locator('#classificationSummary').innerText(),/无对公付款 OA 号未纳入/);

    const bad=X.read(encode(data.ledger),{cellNF:true});
    bad.Sheets['对账明细-李奇'].P2={t:'s',v:'待定'};
    bad.Sheets['对账明细-李奇'].P3={t:'e',v:23,w:'#REF!'};
    delete bad.Sheets['对账明细-孙立柱'].P1;
    await upload('ledger',bad);await analyze();
    await frame.locator('[data-issue-id^="classification:"]').first().waitFor();
    assert(await frame.locator('#download').isDisabled());
    const badNames=await frame.locator('#normalPreview tbody tr td:nth-child(3)').allTextContents();
    for(const name of ['空OA唯一候选','空白分类多候选','别名表头本行日期'])assert(!badNames.includes(name),'异常/缺列不能按空白纳入');
    assert.equal(await frame.locator('[data-issue-id^="classification-header:"]').count(),1);
    const corrected=X.read(encode(data.ledger),{cellNF:true});corrected.Sheets['对账明细-李奇'].I2={t:'s',v:'2026-12-01'};
    await upload('ledger',corrected);await analyze();
    const refreshed=await download(),latest=X.utils.sheet_to_json(refreshed.Sheets.正常付款,{defval:''});
    assert.equal(latest.length,7);assert.equal(latest.find(row=>row.供应商名称==='空OA唯一候选').应付时间,'2026年12月1日','替换台账采用最新本行日期');
    assert.deepEqual(errors,[]);
    console.log(JSON.stringify({status:'PASS',base,normalRows:7,classification:'ledger 是 excluded; 否/blank included despite other sources',emptyOA:'fixed 未提交OA and five blank flow fields even with candidates',due:'exact source row, empty/text retained, latest upload used',financial:'李奇合同口径原值保留/no false overpayment; others zero/settled/formulas/highlights',invalidClassification:'error/value/missing column excluded with review'}));
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
