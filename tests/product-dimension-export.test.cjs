const test = require('node:test');
const assert = require('node:assert/strict');
const X = require('../vendor/xlsx.full.min.js');
const Zip = require('../vendor/jszip.min.js');
const Core = require('../product-dimension-core.js');
const Export = require('../product-dimension-export.js');
const D = ['物料编码','领星SKU','SKU','金蝶名称','金蝶创建时间','物料分组','人工分类'];
const M = ['编码','领星SKU','SKU','名称','创建日期','物料分组'];
const MIME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
function makeBook(headers,rows,name='维度原表') {
  const book=X.utils.book_new();X.utils.book_append_sheet(book,X.utils.aoa_to_sheet([headers,...rows]),name);return book;
}
async function fixture() {
  const book=makeBook(D,[[' 001 ','保留SKU',0,'原商品','原日期','原分组',2],['001','','','重复行','','',99],['002','','','人工商品','','','手工填写'],[0,'0',0,'零编码','','','文字缓存']]);
  const sheet=book.Sheets.维度原表;
  sheet.G2={t:'n',v:2,f:'1+1'};sheet.G5={t:'s',v:'文字缓存',f:'"文字缓存"'};
  sheet.G1.c=[{a:'作者',t:'保留表头批注'}];sheet.G3.c=[{a:'作者',t:'重复行批注去除'}];sheet.G4.c=[{a:'作者',t:'保留原第四行批注'}];
  sheet.D3.l={Target:'https://example.com/dropped'};sheet.D4.l={Target:'https://example.com/kept'};
  sheet['!rows']=[{}, {hidden:true}, {hidden:true}, {hidden:true}];sheet['!cols']=[{wch:22},{hidden:true}];sheet['!autofilter']={ref:'A1:G5'};
  const aux=X.utils.aoa_to_sheet([['辅助原表',2]]);aux.B1.f='1+1';aux.A1.c=[{a:'人工',t:'辅助不变'}];X.utils.book_append_sheet(book,aux,'辅助原表');
  const zip=await Zip.loadAsync(X.write(book,{type:'buffer',bookType:'xlsx'}));
  let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
  xml=xml.replace(/<autoFilter[^>]*\/>/,'<autoFilter ref="A1:G5"><filterColumn colId="0"><filters><filter val="001"/></filters></filterColumn></autoFilter><conditionalFormatting sqref="A2:A5"><cfRule type="duplicateValues" priority="1"/></conditionalFormatting>');
  xml=xml.replace('<f>1+1</f>','<f t="shared" si="0" ref="G2:G3">1+1</f>');
  xml=xml.replace(/(<c r="G3"[^>]*>)/,'$1<f t="shared" si="0"/>');
  zip.file('xl/worksheets/sheet1.xml',xml);zip.file('custom/keep.bin',new Uint8Array([0,255,1,99]));
  const source=await zip.generateAsync({type:'uint8array'});
  const material=makeBook(M,[['001','不得覆盖',10,'新名称','新日期','新分组'],['0003',0,'0000','新增名称',46340,'新增分组']],'物料');material.Sheets.物料.E3.z='yyyy-mm-dd';
  const analysis=Core.analyze(X.read(source,{cellNF:true,cellStyles:true,cellDates:false}),material),result=Core.resolve(analysis);
  assert.equal(result.canExport,true);return {source,analysis,result};
}
async function patch(source,path,fn) {
  const zip=await Zip.loadAsync(source);zip.file(path,fn(await zip.file(path).async('string')));return zip.generateAsync({type:'uint8array'});
}
const MAIN='xl/worksheets/sheet1.xml';
test('keeps the first original row, appends unique codes, stores cached values and preserves native workbook parts',async()=>{
  const {source,analysis,result}=await fixture(),bytes=await Export(source,analysis,result);
  const before=await Zip.loadAsync(source),after=await Zip.loadAsync(bytes),book=X.read(bytes,{cellNF:true,cellStyles:true}),sheet=book.Sheets.维度原表,xml=await after.file(MAIN).async('string');
  assert.deepEqual(book.SheetNames,['维度原表','辅助原表']);
  assert.deepEqual([2,3,4,5].map(row=>sheet['A'+row].v),['001','002','0','0003']);
  assert([2,3,4,5].every(row=>sheet['A'+row].t==='s'));
  assert.equal(sheet.B2.v,'保留SKU');assert.equal(sheet.C2.v,0);assert.equal(sheet.G2.v,2);assert.equal(sheet.G3.v,'手工填写');assert.equal(sheet.G4.v,'文字缓存');
  assert.equal(sheet.B3?.v ?? '','');assert.equal(sheet.G5?.v ?? '','');assert.equal(sheet.B5.v,'0');assert.equal(sheet.C5.v,'0000');assert.equal(sheet.E5.v,46340);assert.equal(sheet.E5.z,'yyyy-mm-dd');
  assert(!/<f\b/.test(xml));assert(!/<row\b[^>]*hidden=/.test(xml));assert.match(xml,/<col\b[^>]*hidden="(?:1|true)"/);assert.match(xml,/<autoFilter ref="A1:G5"\/>/);assert(!xml.includes('<filterColumn'));
  assert.match(xml,/<conditionalFormatting sqref="A2:A4">/);
  assert.equal(sheet.G1.c[0].t,'保留表头批注');assert.equal(sheet.G3.c[0].t,'保留原第四行批注');assert.equal(sheet.G4.c,undefined);
  assert.equal(sheet.D3.l.Target,'https://example.com/kept');assert.equal(sheet.D2.l,undefined);
  const relations=await after.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');assert(!relations.includes('example.com/dropped'));
  const vmlFile=Object.keys(after.files).find(path=>/vmlDrawing1\.vml$/.test(path)),vml=await after.file(vmlFile).async('string');assert.match(vml,/<x:Row>2<\/x:Row>/);assert(!vml.includes('<x:Row>3</x:Row>'));
  for(const name of Object.keys(before.files).filter(name=>!before.files[name].dir)){
    if([MAIN,'xl/styles.xml','xl/workbook.xml','xl/worksheets/_rels/sheet1.xml.rels','xl/comments1.xml',vmlFile].includes(name))continue;
    assert(after.file(name),name);assert.deepEqual(await after.file(name).async('uint8array'),await before.file(name).async('uint8array'),name+' must remain unchanged');
  }
  const oldStyles=await before.file('xl/styles.xml').async('string'),newStyles=await after.file('xl/styles.xml').async('string');
  for(const name of ['fonts','fills','borders','cellXfs']){const re=new RegExp('<'+name+'[^>]*>([\\s\\S]*?)</'+name+'>');assert(newStyles.match(re)[1].startsWith(oldStyles.match(re)[1]),name+' preserves original entries');}
  assert.match(newStyles,/<alignment horizontal="center" vertical="center" wrapText="1"\/>/);
  assert.equal(book.Sheets.辅助原表.B1.f,'1+1');
});

test('raw formula validation detects missing caches and cached errors even when analysis was previously valid',async()=>{
  const {source,analysis,result}=await fixture();
  for(const replacement of ['<c r="G2" t="str"><f>""</f></c>','<c r="G2"><f>1+1</f></c>','<c r="G2" t="e"><f>1/0</f><v>#DIV/0!</v></c>','<c r="G2"><f>1+1</f><v>NaN</v></c>']){
    const input=await patch(source,MAIN,xml=>xml.replace(/<c r="G2"[^>]*>[\s\S]*?<\/c>/,replacement));
    await assert.rejects(()=>Export(input,analysis,result),/G2.*缺少有效的已保存结果/);
  }
  const empty=await patch(source,MAIN,xml=>xml.replace(/<c r="G2"[^>]*>[\s\S]*?<\/c>/,'<c r="G2" t="str"><f>""</f><v/></c>'));
  const output=await Zip.loadAsync(await Export(empty,analysis,result));assert.match(await output.file(MAIN).async('string'),/<c r="G2" t="inlineStr"><is><t xml:space="preserve"><\/t><\/is><\/c>/);
});

test('numeric zero and boolean false formula caches remain values',async()=>{
  const {source,analysis,result}=await fixture();
  const input=await patch(source,MAIN,xml=>xml.replace(/<c r="G2"[^>]*>[\s\S]*?<\/c>/,'<c r="G2"><f>0</f><v>0</v></c>').replace(/<c r="G5"[^>]*>[\s\S]*?<\/c>/,'<c r="G5" t="b"><f>FALSE()</f><v>0</v></c>'));
  const output=X.read(await Export(input,analysis,result));assert.equal(output.Sheets.维度原表.G2.v,0);assert.equal(output.Sheets.维度原表.G4.v,false);
});

test('complex row-dependent structures fail clearly rather than silently corrupting the source',async()=>{
  const {source,analysis,result}=await fixture();
  for(const [insert,message] of [['<mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>',/mergeCells/],['<tableParts count="1"><tablePart r:id="table"/></tableParts>',/tableParts/],['<conditionalFormatting sqref="G2:G5"><cfRule type="expression" priority="9"><formula>G2&gt;0</formula></cfRule></conditionalFormatting>',/公式条件格式/]]){
    const input=await patch(source,MAIN,xml=>xml.replace('</worksheet>',insert+'</worksheet>'));
    await assert.rejects(()=>Export(input,analysis,result),message);
  }
  const array=await patch(source,MAIN,xml=>xml.replace(/<f t="shared" si="0" ref="G2:G3">1\+1<\/f>/,'<f t="array" ref="G2:G3">1+1</f>'));
  await assert.rejects(()=>Export(array,analysis,result),/数组或数据表公式/);
  const cross=await patch(source,'xl/worksheets/sheet2.xml',xml=>xml.replace('<f>1+1</f>',"<f>'维度原表'!G4</f>"));
  await assert.rejects(()=>Export(cross,analysis,result),/其他工作表公式引用/);
});

test('updates filter named range, preserves auxiliary calculation chain entries and removes main sheet entries',async()=>{
  const {source,analysis,result}=await fixture(),zip=await Zip.loadAsync(source);
  zip.file('xl/calcChain.xml','<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="G2" i="1"/><c r="G5"/><c r="B1" i="2"/></calcChain>');
  zip.file('xl/_rels/workbook.xml.rels',(await zip.file('xl/_rels/workbook.xml.rels').async('string')).replace('</Relationships>',`<Relationship Id="rIdCalcChain" Type="${MIME}calcChain" Target="calcChain.xml"/></Relationships>`));
  zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml').async('string')).replace('</Types>','<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
  const output=await Zip.loadAsync(await Export(await zip.generateAsync({type:'uint8array'}),analysis,result)),chain=await output.file('xl/calcChain.xml').async('string');
  assert.match(chain,/<c r="B1" i="2"\/>/);assert(!chain.includes('G2'));assert(!chain.includes('G5'));
  const workbook=await output.file('xl/workbook.xml').async('string');assert.match(workbook,/_xlnm\._FilterDatabase/);assert(workbook.includes('&apos;维度原表&apos;!$A$1:$G$5'));
});

test('export rejects invalid resolution and duplicate keys',async()=>{
  const {source,analysis,result}=await fixture();
  await assert.rejects(()=>Export(source,analysis,{...result,canExport:false}),/待处理错误/);
  await assert.rejects(()=>Export(source,analysis,{...result,rows:[result.rows[0],result.rows[0]]}),/重复物料编码/);
  await assert.rejects(()=>Export(source,analysis,{...result,rows:[result.rows[1],result.rows[0]]}),/原商品行顺序/);
});
