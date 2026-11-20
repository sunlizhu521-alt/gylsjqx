const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const {fixture,book,ID}=require('./funds-progress-fixtures.cjs');
const HEADERS=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','已付金额','剩余付款金额','是否结清','是否提交付款申请单','谈判明细','分批总金额','核对','2026-10-26','2026-11-16','2026-11-30'];
const APP=['申请日期','供应商名称','付款主体','流程编码','本次申请付款金额','付款分类','是否OA备注','是否已付款','应付时间'];
const prefixes=['批次（付款节奏跟进）：','申请日期（付款节奏跟进）：','OA备注时间（付款节奏跟进）：','付款时间（付款节奏跟进）：'];
function sample(oldAmount=15230,newAmount=49990,extra=false){
 const input=fixture(),rows=[
  ['批次供应商','国源',ID(1),oldAmount,oldAmount,0,'是','已经提交全部','10月30日－'+oldAmount,oldAmount,'一致',oldAmount,'',''],
  ['批次供应商','国源',ID(2),newAmount,0,newAmount,'否','未提交','11月25日－'+newAmount,newAmount,'一致','',newAmount,'']
 ];
 if(extra)rows.push(['批次供应商','国源','未提交OA',newAmount,0,newAmount,'否','未提交','12月25日－'+newAmount,newAmount,'一致','','',newAmount]);
 input.negotiated={name:'批次原表.xlsx',book:book('谈判付款',[HEADERS,...rows])};
 const sheet=input.negotiated.book.Sheets.谈判付款;
 sheet.L2.c=[{a:'人工',t:['原付款说明保留',prefixes[0]+'共1批，第1批',prefixes[1]+'2026年10月21日',prefixes[2]+'已备注，时间未填',prefixes[3]+'2026年10月21日'].join('\n')}];
 sheet.L2.s={patternType:'solid',fgColor:{rgb:'FF0000'}};
 for(let r=2;r<=rows.length+1;r++){sheet['F'+r].f='D'+r+'-E'+r;sheet['J'+r].f='SUM(L'+r+':N'+r+')';}
 input.applications={name:'申请.xlsx',book:book('付款申请汇总',[APP,
  ['2026-10-21','批次供应商','国源',ID(1),oldAmount,'谈判付款','是','2026-10-21','2026-10-30'],
  ['2026-11-19','批次供应商','国源',ID(1),newAmount,'谈判付款-谈判付款表','2026-11-19','','2026-11-25']])};
 input.oa={name:'OA.xlsx',book:book('Sheet0',[['流程编号','收款单位','付款金额',...C.FLOW_FIELDS],
  [ID(1),'批次供应商',oldAmount,'归档抄送','甲','归档','李奇'],
  [ID(2),'批次供应商',newAmount,'审批','乙','未归档','李奇']])};
 return input;
}
module.exports={sample,HEADERS,APP,ID,prefixes};
