const X=require('../vendor/xlsx.full.min.js');
const Zip=require('../vendor/jszip.min.js');
const PEOPLE=['李奇','徐英','邢芳芳','操思敏','孙立柱'];
const HEADERS=['月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付款金额','是否结清','剩余付款金额','是否谈判付款','是否提交付款申请单'];
const NORMAL_HEADERS=['登记月份','跟单员','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单','提交时间汇总','当前节点','未操作者','流程状态','申请人','流程编号','是否对账','付款自然周'];
const REVIEW_HEADERS=['付款类别','供应商名称','付款主体','申请付款金额','是否提交付款申请单','核对说明','来源文件','来源工作表','来源行号','跟单员'];
const NEGOTIATED_HEADERS=['供应商','付款主体','对公付款0A流程号','合同总金额','已付金额','摘要','申请付款金额','单据号【财务应付单】','创建人','是否已确认付款','是否结清','应付款时间','账期','剩余付款金额','备注',46325,46329,'提交时间汇总'];
const write=book=>Buffer.from(X.write(book,{type:'buffer',bookType:'xlsx'}));
const sheetBook=(name,rows)=>{const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet(rows),name);return book;};
async function fixtures(){
 const ledger=X.utils.book_new(),applications=[['申请日期','供应商全称','付款主体','流程编码','合同金额','已付金额','付款事由','备注','对接人','本次申请付款金额']],oa=[['流程编号','当前节点','未操作者','流程状态','申请人','收款单位','付款金额']];
 // Old flags intentionally disagree: classification must use the selected negotiated table.
 const records=PEOPLE.map((person,i)=>({person,supplier:'正常供应商'+(i+1),amount:100+i,paid:0,id:'MDST-DGFK20261024'+String(i+1).padStart(3,'0'),ap:'AP00030'+String(i+1).padStart(3,'0'),contract:'HT-NORMAL-'+(i+1),reason:'10月货款',negotiated:i===0?'是':i===1?'':'否'}));
 records.push({person:'李奇',supplier:'谈判甲公司',amount:300,paid:100,id:'MDST-DGFK20261024010',ap:'AP00030010',contract:'HT-NEG-10',reason:'9月尾款',negotiated:'否'});
 records.push({person:'徐英',supplier:'谈判乙公司',amount:0,paid:0,id:'MDST-DGFK20261024011',ap:'AP00030011',contract:'HT-NEG-11',reason:'零余额对账',negotiated:''});
 for(const person of PEOPLE){
  const rows=[HEADERS];
  for(const r of records.filter(r=>r.person===person))rows.push([46300,r.supplier,'迈德斯特',r.id,46300,r.amount,r.reason,r.amount,46329,r.ap,r.contract,r.amount,r.paid,r.amount===r.paid?'是':'否',r.amount-r.paid,r.negotiated,r.supplier==='正常供应商5'?'':'是']);
  const sheet=X.utils.aoa_to_sheet(rows);for(let row=2;row<=rows.length;row++)for(const col of ['A','E','I'])sheet[col+row].z='yyyy年m月d日';
  X.utils.book_append_sheet(ledger,sheet,'对账明细-'+person);
 }
 // Missing and partial applications are exported for review without blocking downloads.
 for(const r of records){
  const application=(date,amount)=>[date,r.supplier,'迈德斯特',r.id,r.amount,r.paid,r.reason,r.ap,r.person,amount];
  if(r.supplier==='谈判甲公司')applications.push(application('2026-10-24',80),application('2026-10-23',100));
  else if(r.supplier!=='正常供应商4')applications.push(application('2026-10-24',r.supplier==='正常供应商5'?103:r.amount));
  const flow=[r.id,'财务审核','测试审核人','未归档',r.person,r.supplier,r.amount];oa.push(flow,flow.slice());
 }
 const negotiated=X.utils.book_new(),sheet=X.utils.aoa_to_sheet([NEGOTIATED_HEADERS,
 ['谈判甲公司','','',300,0,'9月尾款',300,'','李奇','是','否',46329,'票到30天',300,'必须保留的人工备注',80,220],
 ['谈判乙公司','迈德斯特','MDST-DGFK20261024011',0,0,'零余额对账',0,'AP00030011','徐英','是','是',46329,'约定账期',0,'',0,0],
 ['合计','','','','','','','','','','','','','','','','']]);
 sheet.N2={t:'n',f:'G2-E2',v:300};sheet.Q2={t:'n',f:'N2-P2',v:220};sheet.N3={t:'n',f:'G3-E3',v:0};sheet.R2={t:'s',v:'2026-09-01-999.00'};sheet.L2.z=sheet.L3.z='yyyy年m月d日';sheet.O2.c=[{a:'原作者',t:'不能丢失的原批注'}];
 sheet['!merges']=[{s:{r:1,c:14},e:{r:2,c:14}}];sheet['!cols']=NEGOTIATED_HEADERS.map((_,i)=>({wch:i===14?36:20}));sheet['!rows']=[{hpt:30},{hpt:36},{hpt:36},{hpt:22}];
 X.utils.book_append_sheet(negotiated,sheet,'10月更新谈判付款计划');
 X.utils.book_append_sheet(negotiated,X.utils.aoa_to_sheet([['不应导出的其他工作表'],['私有备注']]),'辅助说明');
 X.utils.book_append_sheet(negotiated,X.utils.aoa_to_sheet([NEGOTIATED_HEADERS,['不应默认选取的反馈版本']]),'谈判付款-反馈');
 // Give the source a distinctive existing cell style to detect destructive round-trips.
 const zip=await Zip.loadAsync(write(negotiated));let styles=await zip.file('xl/styles.xml').async('string');
 styles=styles.replace('<fonts count="1">','<fonts count="2">').replace('</fonts>','<font><sz val="14"/><color rgb="FF923456"/><name val="Microsoft YaHei"/></font></fonts>');
 const xf=styles.match(/<cellXfs count="(\d+)">/);const newStyle=Number(xf[1]);styles=styles.replace(xf[0],`<cellXfs count="${newStyle+1}">`).replace('</cellXfs>','<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>');
 zip.file('xl/styles.xml',styles);let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');xml=xml.replace(/<c r="O2"([^>]*)>/,'<c r="O2" s="'+newStyle+'"$1>');zip.file('xl/worksheets/sheet1.xml',xml);
 const noMatch=write(sheetBook('普通汇总',[['供应商','申请付款金额'],['测试',1]]));
 const singleBook=X.utils.book_new();X.utils.book_append_sheet(singleBook,sheet,'其他月份谈判付款');
 const oaZip=await Zip.loadAsync(write(sheetBook('Sheet0',oa)));oaZip.file('xl/worksheets/sheet1.xml',(await oaZip.file('xl/worksheets/sheet1.xml').async('string')).replace(/<dimension ref="[^"]+"\/>/,'<dimension ref="A1"/>')); 
 return {ledger:write(ledger),applications:write(sheetBook('付款申请汇总',applications)),negotiated:await zip.generateAsync({type:'nodebuffer'}),oa:await oaZip.generateAsync({type:'nodebuffer'}),noMatch,single:write(singleBook),sheetName:'10月更新谈判付款计划',records};
}
module.exports={fixtures,PEOPLE,NORMAL_HEADERS,NEGOTIATED_HEADERS,REVIEW_HEADERS};
