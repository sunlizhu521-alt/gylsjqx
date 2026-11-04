const X=require('../vendor/xlsx.full.min.js'),C=require('../funds-core.js');
const LH=['登记月份','供应商名称','付款主体','对公付款0A流程号','开票时间','开票金额','摘要','申请付款金额','应付款时间','单据号【财务应付单】','OA合同编号','合同总金额','已付金额','是否结清','剩余金额','是否谈判付款','是否提交付款申请单'];
const NH=['供应商名称','付款主体','对公付款0A流程号','申请付款金额','已付金额','剩余付款金额','是否结清','应付款时间','备注','谈判明细','分批总金额','核对',46328,46335];
const AH=['申请日期','供应商全称','付款主体','流程编码','本次申请付款金额','付款分类'];
const OH=['流程编号','收款单位','付款金额','当前节点','未操作者','流程状态','申请人','签字意见'];
const PAID40='[出纳执行]\n部分支付：40元\n合成出纳 2026-11-01 15:30:00';
const ID=n=>'MDST-DGFK20261101'+String(n).padStart(3,'0');
const book=(name,rows)=>{const b=X.utils.book_new();X.utils.book_append_sheet(b,X.utils.aoa_to_sheet(rows),name);return b;};
function fixture(){
 const ledger=X.utils.book_new();
 for(const person of C.PEOPLE){
  const row=(n,amount,neg,due)=>[46300,'合成供应商'+n,'迈德斯特',ID(n),46300,amount,'合成货款',amount,due,'AP2026110100'+n,'HT-TEST-'+n,amount,0,'否',amount,neg?'是':'否','未申请'];
  const rows=person==='李奇'?[row(1,100,true,46329),row(3,300,false,46342)]:person==='徐英'?[row(2,200,false,46335)]:[];
  const s=X.utils.aoa_to_sheet([LH,...rows]);for(let r=2;r<=rows.length+1;r++)for(const c of ['A','E','I'])s[c+r].z=C.DATE_FORMAT;
  X.utils.book_append_sheet(ledger,s,'对账明细-'+person);
 }
 const negotiated=book('十一月谈判付款主表',[NH,['合成供应商1','迈德斯特',ID(1),100,0,100,'否',46332,'保留备注','1、11月2日－40\n2、11月9日－60',100,'一致',40,60]]);
 const ns=negotiated.Sheets[negotiated.SheetNames[0]];
 ns.F2={t:'n',v:100,f:'D2-E2',z:'#,##0.00'};ns.K2={t:'n',v:100,f:'SUM(M2:N2)'};ns.H2.z=ns.M1.z=ns.N1.z=C.DATE_FORMAT;ns.I2.c=[{a:'人工',t:'保留原批注'}];
 const applications=book('付款申请汇总',[AH,
  ['2026-11-01','合成供应商1','浙江迈德斯特医疗器械科技有限公司',ID(1),40,'谈判付款'],
  ['2026-11-02','合成供应商2','迈德斯特',ID(2),200,'正常付款'],
  ['2026-11-03','合成供应商1','迈德斯特',ID(1),40,'谈判付款'],
  ['2026-11-04','未进入原表供应商','迈德斯特',ID(4),0,'正常付款']]);
 const oa=book('Sheet0',[OH,...[1,2,3,4].map(n=>[ID(n),n===4?'未进入原表供应商':'合成供应商'+n,n===4?0:n*100,'财务审批','合成审核人',n===1?'归档':'未归档','OA申请人',n===1?PAID40:'']),[ID(1),'合成供应商1',100,'财务审批','合成审核人','归档','OA申请人',PAID40]]);
 return Object.fromEntries(Object.entries({ledger,applications,negotiated,oa}).map(([k,book])=>[k,{book,name:k+'.xlsx'}]));
}
function set(i,role,header,row,value,name){const b=i[role].book,s=b.Sheets[name||b.SheetNames[0]],headers=X.utils.sheet_to_json(s,{header:1})[0],c=headers.indexOf(header);if(c<0)throw Error('missing '+header);s[X.utils.encode_cell({r:row-1,c})]=value&&typeof value==='object'?value:{t:typeof value==='number'?'n':'s',v:value};}
function append(i,role,values,name){const b=i[role].book;X.utils.sheet_add_aoa(b.Sheets[name||b.SheetNames[0]],[values],{origin:-1});}
module.exports={fixture,set,append,ID,LH,NH,AH,OH,book,PAID40};
