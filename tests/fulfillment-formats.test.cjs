const {test}=require('node:test'),assert=require('node:assert/strict');
const X=require('../vendor/xlsx.full.min.js'),C=require('../fulfillment-core.js'),F=require('../fulfillment-formats.js');
const headers=['事业部','物料编码','下单月份','下单日期',...C.QUANTITY_FIELDS,'合同约定交期','SKU','OA备货单号','说明'];
function fixture(name='徐英',columns=headers){const sheet=X.utils.aoa_to_sheet([columns]);return {SheetNames:[name],Sheets:{[name]:sheet}};}
function set(book,name,values,row=2){const sheet=book.Sheets[name];for(const [key,cell]of Object.entries(values)){const keys=Object.keys(sheet).filter(k=>/^[A-Z]+1$/.test(k)),header=keys.find(k=>sheet[k].v===key);sheet[header.replace(/1$/,row)]=typeof cell==='object'&&cell!==null?cell:{t:typeof cell==='number'?'n':'s',v:cell};}sheet['!ref']=sheet['!ref'].replace(/\d+$/,row);}
function get(book,name,key,row=2){const sheet=book.Sheets[name],header=Object.keys(sheet).find(k=>/^[A-Z]+1$/.test(k)&&sheet[k].v===key);return sheet[header.replace(/1$/,row)];}
test('四个人员按各自表头转换，月份和日期为真实日期，其他内容为文本，原工作簿不改',()=>{
 const book={SheetNames:[],Sheets:{}};
 for(const [index,person]of C.SHEETS.entries()){
  const name=person+(index===2?'1017':''),columns=[...headers.slice(index),...headers.slice(0,index)],part=fixture(name,columns);book.Sheets[name]=part.Sheets[name];book.SheetNames.push(name);
  set(book,name,{'事业部':'海外事业一部','物料编码':{t:'n',v:123,z:'000000',w:'000123'},'下单月份':'2025-11','下单日期':'2026年10月27日','未交付数量':'约1,245套','已下单未备料未生产':'2.5','已备料未生产':0,'生产中产品':1e-7,'完工未发产品':'','已发货数量':3,'合同约定交期':{t:'n',v:46322,z:C.DATE_FORMAT},'SKU':'0012','OA备货单号':'00004','说明':{t:'n',v:1e21,w:'1E+21'}});
 }
 book.Sheets.孙立柱={A1:{t:'n',v:123,f:'1+122'},'!ref':'A1'};const before=JSON.stringify(book),output=F.normalize(book);
 assert.equal(JSON.stringify(book),before);assert.equal(output.normalizedSheetCount,4);assert.deepEqual(output.issues,[]);assert.deepEqual(output.warnings,[]);assert.strictEqual(output.book.Sheets.孙立柱,book.Sheets.孙立柱);
 for(const name of book.SheetNames){
  assert.equal(get(output.book,name,'下单月份').z,'yyyy"年"m"月"');assert.equal(X.SSF.format('yyyy-mm-dd',get(output.book,name,'下单月份').v),'2025-11-01');
  assert.equal(get(output.book,name,'下单日期').z,C.DATE_FORMAT);assert.equal(X.SSF.format('yyyy-mm-dd',get(output.book,name,'下单日期').v),'2026-10-27');
  for(const [field,value,z]of [['未交付数量',1245,'0'],['已下单未备料未生产',2.5,'0.0'],['已备料未生产',0,'0'],['生产中产品',1e-7,'0.0000000'],['完工未发产品','','0']]){const cell=get(output.book,name,field);assert.equal(cell.v,value);assert.equal(cell.z,z);assert.equal(cell.t,value===''?'s':'n');}
  assert.equal(get(output.book,name,'物料编码').v,'000123');assert.equal(get(output.book,name,'SKU').v,'0012');assert.equal(get(output.book,name,'OA备货单号').v,'00004');
  assert.equal(get(output.book,name,'说明').v,'1000000000000000000000');assert.equal(get(output.book,name,'合同约定交期').v,'2026年10月27日');
  for(const field of ['物料编码','SKU','OA备货单号','合同约定交期','说明']){assert.equal(get(output.book,name,field).t,'s');assert.equal(get(output.book,name,field).z,'@');}
 }
});
test('1904 日期体系、原日期数值、历史订单和月斜杠默认不产生四年偏移',()=>{
 const book=fixture();book.Workbook={WBProps:{date1904:true}};
 set(book,'徐英',{'下单月份':{t:'n',v:44860,z:'yyyy-mm'},'下单日期':{t:'n',v:44860,z:C.DATE_FORMAT},'合同约定交期':{t:'n',v:44860,z:C.DATE_FORMAT},'说明':{t:'n',v:44860,z:'yyyy-mm'}});
 const out=F.normalize(book);assert.equal(get(out.book,'徐英','下单日期').v,44860);assert.equal(get(out.book,'徐英','下单月份').v,44860);assert.equal(get(out.book,'徐英','合同约定交期').v,'2026年10月27日');assert.equal(get(out.book,'徐英','说明').v,'2026年10月');
 set(book,'徐英',{'下单月份':'/','下单日期':'历史订单'});const defaults=F.normalize(book);assert.equal(get(defaults.book,'徐英','下单月份').v,C.dateCell({v:'2026-04-01'}).v-1462);assert.equal(get(defaults.book,'徐英','下单日期').v,C.dateCell({v:'2025-01-01'}).v-1462);
 set(book,'徐英',{'下单月份':'历史订单','下单日期':{t:'d',v:new Date('2026-10-27T00:00:00Z')}});const dated=F.normalize(book);assert.equal(get(dated.book,'徐英','下单月份').v,C.dateCell({v:'2025-01-01'}).v-1462);assert.equal(get(dated.book,'徐英','下单日期').v,44860);
});
test('数字标识符即使误设日期格式也保持编码文本，不转换成年月日',()=>{
 const book=fixture();for(const key of ['物料编码','SKU','OA备货单号'])set(book,'徐英',{[key]:{t:'n',v:1234,z:'yyyy-mm-dd',w:'1903-05-18'}});
 const out=F.normalize(book);for(const key of ['物料编码','SKU','OA备货单号']){const cell=get(out.book,'徐英',key);assert.equal(cell.v,'1234');assert.equal(cell.t,'s');assert.equal(cell.z,'@');}
});
test('纵向合并数量只第一行提取，后续补0，合并和样式不变',()=>{
 const book=fixture('邢芳芳');set(book,'邢芳芳',{'生产中产品':{t:'s',v:'物料通用 474套左右',s:{fill:{fgColor:{rgb:'FF0000'}}}}});const sheet=book.Sheets.邢芳芳;
 const c=headers.indexOf('生产中产品');sheet['!merges']=[{s:{r:1,c},e:{r:3,c}}];sheet['!ref']='A1:N4';const before=JSON.stringify(book),out=F.normalize(book);
 assert.deepEqual([2,3,4].map(row=>get(out.book,'邢芳芳','生产中产品',row).v),[474,0,0]);assert.equal(JSON.stringify(book),before);assert.deepEqual(out.book.Sheets.邢芳芳['!merges'],sheet['!merges']);assert.deepEqual(get(out.book,'邢芳芳','生产中产品').s,get(book,'邢芳芳','生产中产品').s);
 assert.equal(out.updates.filter(u=>u.cell.startsWith(X.utils.encode_col(c))&&[474,0].includes(u.value)).length,3);
});
test('未知日期保留文字和警告，空日期空数量不改0，残余公式和异常数量留原值并提示',()=>{
 const book=fixture();set(book,'徐英',{'下单月份':'下月待定','下单日期':'','未交付数量':'10-20套','已下单未备料未生产':{t:'n',f:'1+1',v:0},'已备料未生产':'','生产中产品':false});
 const out=F.normalize(book);assert.equal(get(out.book,'徐英','下单月份').v,'下月待定');assert.equal(get(out.book,'徐英','下单日期').v,'');assert.equal(get(out.book,'徐英','已备料未生产').v,'');assert.equal(get(out.book,'徐英','未交付数量').v,'10-20套');assert.equal(get(out.book,'徐英','已下单未备料未生产').f,'1+1');assert.equal(out.warnings.length,1);assert.equal(out.issues.length,3);
 assert.ok(out.issues.some(i=>i.message.includes('公式尚未成功重算')));assert.ok(out.issues.some(i=>i.message.includes('多个数字')));
});
test('无表头或非指定工作表不乱转换，支持表头前标题和八位日期后缀',()=>{
 const book=fixture('李奇20261027');const source=book.Sheets.李奇20261027;
 for(const key of Object.keys(source).filter(k=>/^[A-Z]+1$/.test(k))){source[key.replace(/1$/,'3')]=source[key];delete source[key];}
 source.A1={t:'n',v:123};source['!ref']='A1:N3';book.Sheets.徐英={A1:{t:'n',v:23},'!ref':'A1'};book.Sheets.李奇旧版={A1:{t:'n',v:23},'!ref':'A1'};
 const out=F.normalize(book);assert.equal(out.normalizedSheetCount,1);assert.equal(out.issues.length,1);assert.equal(out.book.Sheets.李奇20261027.A1.t,'s');assert.strictEqual(out.book.Sheets.徐英,book.Sheets.徐英);assert.strictEqual(out.book.Sheets.李奇旧版,book.Sheets.李奇旧版);
});
test('规范结果写出重读后再次处理不更新，不因读取时生成的w推进缓存',()=>{
 const book=fixture();set(book,'徐英',{'事业部':'国内事业部','物料编码':'00123','下单月份':'2026-10','下单日期':'2026-10-27','未交付数量':'1.25套','生产中产品':0,'合同约定交期':{t:'n',v:46322,z:C.DATE_FORMAT}});
 const first=F.normalize(book);assert.equal(first.normalizedSheetCount,1);const reread=X.read(X.write(first.book,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellNF:true,cellStyles:true});
 assert.ok(reread.Sheets.徐英.A1.w);const repeated=F.normalize(reread);assert.deepEqual(repeated.updates,[]);assert.equal(repeated.normalizedSheetCount,0);assert.deepEqual(repeated.issues,[]);assert.strictEqual(repeated.book.Sheets.徐英,reread.Sheets.徐英);
});
