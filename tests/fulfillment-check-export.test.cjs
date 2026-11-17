const test=require('node:test'),assert=require('node:assert/strict'),X=require('../vendor/xlsx.full.min.js'),Z=require('../vendor/jszip.min.js'),E=require('../fulfillment-check-export.js'),C=require('../fulfillment-check-core.js');
function book(name){const b=X.utils.book_new(),s=X.utils.aoa_to_sheet([['编号','数量'],['001',0.5],['002',0],['003',2]]);s.B4={t:'n',f:'B2+B3+1.5',v:2};s.A2.c=[{a:'人工',t:'原批注'}];s['!merges']=[{s:{r:4,c:0},e:{r:4,c:1}}];s['!ref']='A1:B5';s['!cols']=[{wch:24},{wch:16,hidden:true}];s['!rows']=[{hpt:30},{hpt:26,hidden:true}];s['!autofilter']={ref:'A1:B4'};X.utils.book_append_sheet(b,s,name);return b;}
const bytes=b=>X.write(b,{type:'buffer',bookType:'xlsx',bookSST:true});
const result={canExport:true,summary:[['甲',0.3,0.5,0.2,0,1,0,0,0,0]],detail:[['数量差异','甲','甲','001','','','CGDD1','CGDD1','CGDD1：0.3',0.3,0.5,0.2,'未关闭','采购.xlsx / 采购 第 2 行（D2）','履约.xlsx / 徐英 第 2 行（D2）','名称一致','核实两边数量']]};
const section=(xml,n)=>(xml.match(new RegExp('<'+n+'\\b[^>]*>([\\s\\S]*?)</'+n+'>'))||[])[1]||'';
test('两张结果表先于全部原表；原母包非元数据部分逐字保留',async()=>{
 const a=book('采购'),b=book('徐英');X.utils.book_append_sheet(b,X.utils.aoa_to_sheet([['原表其他'],[123]]),'隐藏表');b.Workbook={Sheets:[{Hidden:0},{Hidden:2}]};
 const p=bytes(a),f=bytes(b),out=await E(p,f,result),z=await Z.loadAsync(out),source=await Z.loadAsync(f),read=X.read(out,{type:'array',cellStyles:true,cellNF:true});
 assert.deepEqual(read.SheetNames,['汇总','差异明细','采购','徐英','隐藏表']);assert.equal(read.Workbook.Sheets[4].Hidden,2);
 const changed=new Set(['xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/styles.xml','[Content_Types].xml','docProps/app.xml']);
 for(const n of Object.keys(source.files).filter(n=>!source.files[n].dir&&!changed.has(n)))assert.equal(await z.file(n).async('base64'),await source.file(n).async('base64'),n);
 for(const name of ['采购','徐英']){assert.equal(read.Sheets[name].B4.f,'B2+B3+1.5');assert.equal(read.Sheets[name].B4.v,2);assert.match(read.Sheets[name].A2.c[0].t,/原批注/);assert.equal(read.Sheets[name]['!autofilter'].ref,'A1:B4');assert.deepEqual(read.Sheets[name]['!merges'],a.Sheets.采购['!merges']);}
 assert.deepEqual(X.utils.sheet_to_json(read.Sheets.汇总,{header:1})[0],C.SUMMARY);assert.deepEqual(X.utils.sheet_to_json(read.Sheets.差异明细,{header:1})[0],C.DETAIL);assert.equal(read.Sheets.汇总.B2.t,'n');assert.equal(read.Sheets.差异明细.D2.t,'s');assert.equal(read.Sheets.差异明细.D2.v,'001');assert.equal(read.Sheets.汇总.A1.s.fgColor.rgb,'315B9D');
 assert.equal(read.Sheets.差异明细['!autofilter'].ref,'A1:Q2');assert.equal(read.Sheets.汇总['!autofilter'].ref,'A1:J2');
 const styles=await z.file('xl/styles.xml').async('string'),original=await source.file('xl/styles.xml').async('string');for(const name of ['fonts','fills','borders','cellXfs'])assert(section(styles,name).startsWith(section(original,name)),name);
 const types=await z.file('[Content_Types].xml').async('string'),parts=[...types.matchAll(/PartName="([^"]+)"/g)].map(m=>m[1]);assert.equal(new Set(parts).size,parts.length,'no duplicate content types');
 const xml=await z.file('xl/fulfillment-check-import/result0.xml').async('string');assert.match(xml,/state="frozen"/);assert.match(styles,/style="hair"/);
});
test('名称冲突不覆盖源表；无公式与名称的采购表可安全改名',async()=>{
 const p=book('同名'),b=book('同名');delete p.Sheets.同名.B4.f;delete p.Sheets.同名['!autofilter'];const out=await E(bytes(p),bytes(b),result);assert.deepEqual(X.read(out,{type:'array'}).SheetNames,['汇总','差异明细','采购-同名-1','同名']);
 await assert.rejects(E(bytes(b),bytes(b),result),/重名且含公式/);
 await assert.rejects(E(bytes(p),bytes(book('汇总')),result),/已有/);
});
test('原始的空白和数值零不混淆；特殊采购结构明确停止，不静默丢失',async()=>{
 const p=book('采购'),b=book('徐英');const r={...result,detail:[result.detail[0].map((v,i)=>[9,11].includes(i)?null:v)]};
 const out=X.read(await E(bytes(p),bytes(b),r),{type:'array'});assert.equal(out.Sheets.差异明细.J2?.v,undefined);assert.equal(out.Sheets.差异明细.K2.v,0.5);
 const z=await Z.loadAsync(bytes(p));z.file('xl/worksheets/sheet1.xml',(await z.file('xl/worksheets/sheet1.xml').async('string')).replace('</worksheet>','<tableParts count="1"><tablePart r:id="rId2"/></tableParts></worksheet>'));
 await assert.rejects(E(await z.generateAsync({type:'uint8array'}),bytes(b),result),/结构化表/);
});
