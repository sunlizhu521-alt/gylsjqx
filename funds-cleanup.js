/* Remove trailing blank formatting without rebuilding values or native workbook parts. */
window.FundsCleanup = async function(source) {
  'use strict';
  const zip=source?.files?source:await JSZip.loadAsync(source);
  const parser=new DOMParser(),serializer=new XMLSerializer();
  const all=(doc,name)=>[...doc.getElementsByTagNameNS('*',name)];
  const first=(doc,name)=>all(doc,name)[0];
  const parse=async path=>{
    const doc=parser.parseFromString(await zip.file(path).async('string'),'application/xml');
    if(first(doc,'parsererror'))throw Error('清理空白格式时无法读取 '+path);
    return doc;
  };
  const resolve=(base,target)=>new URL(target,'https://xlsx.invalid/'+base).pathname.slice(1);
  const relPath=path=>path.replace(/([^/]+)$/,'_rels/$1.rels');
  const workbook=await parse('xl/workbook.xml'),relations=await parse('xl/_rels/workbook.xml.rels');
  const sharedRelation=all(relations,'Relationship').find(r=>r.getAttribute('Type').endsWith('/sharedStrings'));
  const shared=sharedRelation?all(await parse(resolve('xl/workbook.xml',sharedRelation.getAttribute('Target'))),'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[];
  const sheetRelations=new Map(all(relations,'Relationship').filter(r=>r.getAttribute('Type').endsWith('/worksheet')).map(r=>[r.getAttribute('Id'),resolve('xl/workbook.xml',r.getAttribute('Target'))]));
  let workbookChanged=false;
  for(const [sheetIndex,entry] of all(workbook,'sheet').entries()){
    const path=sheetRelations.get(entry.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
    if(!path)continue;
    const doc=await parse(path),data=first(doc,'sheetData');if(!data)continue;
    let changed=false,end={r:0,c:0},start={r:Infinity,c:Infinity};
    const set=(node,key,value)=>{if(node.getAttribute(key)!==String(value)){node.setAttribute(key,String(value));changed=true;}};
    const remove=node=>{node.remove();changed=true;};
    const range=ref=>{
      if(!/^\$?[A-Z]+\$?[1-9]\d*(?::\$?[A-Z]+\$?[1-9]\d*)?$/.test(ref||''))return null;
      return XLSX.utils.decode_range(ref);
    };
    const include=ref=>{const r=range(ref);if(r){start.r=Math.min(start.r,r.s.r);start.c=Math.min(start.c,r.s.c);end.r=Math.max(end.r,r.e.r);end.c=Math.max(end.c,r.e.c);}};
    const hasContent=cell=>[...cell.children].some(e=>!['v','is'].includes(e.localName)||e.localName==='v'&&e.textContent!==''&&(cell.getAttribute('t')!=='s'||shared[Number(e.textContent)]!=='')||e.localName==='is'&&all(e,'t').some(t=>t.textContent!==''))||[...cell.attributes].some(a=>a.namespaceURI!=='http://www.w3.org/2000/xmlns/'&&!['r','s','t'].includes(a.name));
    for(const cell of all(data,'c'))if(hasContent(cell)){
      include(cell.getAttribute('r'));
      for(const formula of all(cell,'f'))include(formula.getAttribute('ref'));
    }
    for(const row of all(data,'row'))if([...row.children].some(c=>c.localName!=='c'))end.r=Math.max(end.r,Number(row.getAttribute('r'))-1);
    for(const link of all(doc,'hyperlink'))include(link.getAttribute('ref'));
    for(const validation of all(doc,'dataValidation'))for(const ref of (validation.getAttribute('sqref')||first(validation,'sqref')?.textContent||'').split(/\s+/))include(ref);
    if(zip.file(relPath(path))){
      const links=await parse(relPath(path));
      for(const link of all(links,'Relationship')){
        if(link.getAttribute('TargetMode')==='External')continue;
        const type=link.getAttribute('Type'),target=resolve(path,link.getAttribute('Target'));
        if(!zip.file(target))continue;
        if(/\/(?:comments|threadedComment|threadedComments|table)$/.test(type)){
          const related=await parse(target);
          for(const tag of ['comment','threadedComment','table'])for(const anchor of all(related,tag))include(anchor.getAttribute('ref'));
        }
        if(type.endsWith('/drawing')){
          const drawing=await parse(target);
          for(const tag of ['from','to'])for(const anchor of all(drawing,tag)){
            const c=Number(first(anchor,'col')?.textContent),r=Number(first(anchor,'row')?.textContent);
            if(Number.isInteger(c)&&Number.isInteger(r)&&c>=0&&r>=0)include(XLSX.utils.encode_cell({c,r}));
          }
        }
        if(type.endsWith('/vmlDrawing'))for(const shape of all(await parse(target),'shape')){
          const anchor=first(shape,'ClientData');if(!anchor||anchor.getAttribute('ObjectType')==='Note')continue;
          const offsets=first(anchor,'Anchor')?.textContent.split(',').map(Number);
          if(offsets?.length===8)for(const i of [0,4])if(Number.isInteger(offsets[i])&&Number.isInteger(offsets[i+2]))include(XLSX.utils.encode_cell({c:offsets[i],r:offsets[i+2]}));
        }
      }
    }
    // Keep merges connected to the real table, including intentional blank input cells.
    const merges=all(doc,'mergeCell');let expanded=true;
    while(expanded){expanded=false;for(const merge of merges){const r=range(merge.getAttribute('ref'));if(r&&r.s.r<=end.r&&r.s.c<=end.c){const before=end.r+':'+end.c;include(merge.getAttribute('ref'));if(before!==end.r+':'+end.c)expanded=true;}}}
    for(const merge of merges){const r=range(merge.getAttribute('ref'));if(r&&(r.s.r>end.r||r.s.c>end.c))remove(merge);}
    const mergeList=first(doc,'mergeCells');if(mergeList){if(!mergeList.children.length)remove(mergeList);else set(mergeList,'count',mergeList.children.length);}
    for(const cell of all(data,'c')){const p=XLSX.utils.decode_cell(cell.getAttribute('r'));if((p.r>end.r||p.c>end.c)&&!hasContent(cell))remove(cell);}
    for(const row of all(data,'row')){
      if(Number(row.getAttribute('r'))>end.r+1&&!row.children.length){remove(row);continue;}
      if(row.hasAttribute('spans')){
        const columns=[...row.children].filter(c=>c.localName==='c').map(c=>XLSX.utils.decode_cell(c.getAttribute('r')).c+1);
        if(columns.length)set(row,'spans',Math.min(...columns)+':'+Math.max(...columns));
        else{row.removeAttribute('spans');changed=true;}
      }
    }
    for(const col of all(doc,'col')){
      if(Number(col.getAttribute('min'))>end.c+1)remove(col);
      else if(Number(col.getAttribute('max'))>end.c+1)set(col,'max',end.c+1);
    }
    const clip=ref=>{const r=range(ref);if(!r)return ref;if(r.s.r>end.r||r.s.c>end.c)return '';return XLSX.utils.encode_range({s:r.s,e:{r:Math.min(r.e.r,end.r),c:Math.min(r.e.c,end.c)}});};
    const filter=first(doc,'autoFilter');
    if(filter){
      const ref=clip(filter.getAttribute('ref'));
      if(!ref)remove(filter);
      else{set(filter,'ref',ref);const bounds=range(ref);for(const col of all(filter,'filterColumn'))if(Number(col.getAttribute('colId'))>bounds.e.c-bounds.s.c)remove(col);}
    }
    for(const cf of all(doc,'conditionalFormatting')){
      const sqref=cf.hasAttribute('sqref')?cf.getAttribute('sqref'):first(cf,'sqref')?.textContent;
      if(sqref==null)continue;
      const refs=sqref.split(/\s+/).filter(Boolean).map(clip).filter(Boolean).join(' ');
      if(!refs)remove(cf);
      else if(refs!==sqref){if(cf.hasAttribute('sqref'))set(cf,'sqref',refs);else{first(cf,'sqref').textContent=refs;changed=true;}}
    }
    let dimension=first(doc,'dimension');
    if(!dimension){dimension=doc.createElementNS(doc.documentElement.namespaceURI,'dimension');doc.documentElement.insertBefore(dimension,doc.documentElement.firstChild);changed=true;}
    const previous=range(dimension.getAttribute('ref'));
    set(dimension,'ref',XLSX.utils.encode_range({s:{r:Math.min(previous?.s.r??0,start.r,end.r),c:Math.min(previous?.s.c??0,start.c,end.c)},e:end}));
    for(const name of all(workbook,'definedName'))if(name.getAttribute('localSheetId')===String(sheetIndex)&&name.getAttribute('name')==='_xlnm._FilterDatabase'){
      const current=first(doc,'autoFilter');
      if(!current){name.remove();workbookChanged=true;continue;}
      const r=range(current.getAttribute('ref')),address=p=>'$'+XLSX.utils.encode_col(p.c)+'$'+(p.r+1);
      const value="'"+entry.getAttribute('name').replace(/'/g,"''")+"'!"+address(r.s)+':'+address(r.e);
      const unquoted=entry.getAttribute('name')+'!'+address(r.s)+':'+address(r.e);
      if(name.textContent!==value&&name.textContent!==unquoted){name.textContent=value;workbookChanged=true;}
    }
    if(changed)zip.file(path,serializer.serializeToString(doc));
  }
  if(workbookChanged)zip.file('xl/workbook.xml',serializer.serializeToString(workbook));
  return zip.generateAsync({type:'uint8array'});
};
