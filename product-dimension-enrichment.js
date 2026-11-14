(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'));
  else root.ProductDimensionEnrichment = factory(root.XLSX);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(XLSX) {
  'use strict';
  const REQUIRED_CLASS = ['销售产品线','销售系列','销售产品分类','场景分组','一级品类','二级分类','三级分类','销售区域'];
  const OPTIONAL = ['型号','品牌','国家','是否正常备货'];
  const PRICES = ['财务加权平均价','不含税结算价','结算价'];
  const DATES = ['当前日期','上架时间（月）','生命周期'];
  const FIELD_HEADERS = ['销售产品线','销售系列','销售产品分类',...PRICES,'场景分组','一级品类','二级分类','三级分类',...OPTIONAL.slice(0,3),'销售区域',...DATES,'是否正常备货'];
  const REQUIRED_HEADERS = [...REQUIRED_CLASS,...DATES];
  const EDITABLE_HEADERS = ['金蝶创建时间',...REQUIRED_CLASS,...OPTIONAL,...PRICES];
  const own = (o,k) => Object.prototype.hasOwnProperty.call(o || {},k);
  const text = v => String(v ?? '').trim();
  const missing = v => !text(v) || /^(?:待确认|待补|待补全|待补充|待核对|未匹配|无法区分|未知|无|无资料|N\/A|NA|null|undefined|\/)$/i.test(text(v)) || /^[-—–―－﹣−]+$/.test(text(v));
  const numericCategory = v => typeof v==='number' || /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text(v));
  const sku = v => missing(v) || text(v) === '0' ? '' : text(v);
  const key = values => JSON.stringify(values);
  const unique = values => [...new Set(values)];
  const today = () => new Date(Date.now() + 8 * 3600000).toISOString().slice(0,10);
  const dateLabel = d => `${String(d.y).padStart(4,'0')}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  function dateValue(value, date1904) {
    let parts;
    if (typeof value === 'number' && Number.isFinite(value)) parts = XLSX.SSF.parse_date_code(value,{date1904});
    else if (typeof value === 'string') {
      const match = text(value).match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
      if (!match || Number(match[4] || 0)>23 || Number(match[5] || 0)>59 || Number(match[6] || 0)>59) return null;
      parts = {y:Number(match[1]),m:Number(match[2]),d:Number(match[3])};
    }
    if (!parts || parts.y < 1900 || parts.y > 9999) return null;
    const utc = Date.UTC(parts.y,parts.m-1,parts.d), d = new Date(utc);
    if (d.getUTCFullYear()!==parts.y || d.getUTCMonth()+1!==parts.m || d.getUTCDate()!==parts.d) return null;
    let serial = (utc-Date.UTC(1899,11,31))/86400000 + (utc>=Date.UTC(1900,2,1)?1:0) - (date1904?1462:0);
    if (serial < 0 || !Number.isFinite(serial)) return null;
    return {y:parts.y,m:parts.m,d:parts.d,serial,label:dateLabel(parts)};
  }
  function age(created,asOf) {
    const months=(asOf.y-created.y)*12+asOf.m-created.m-(asOf.d<created.d?1:0);
    return {months,age:`${months}个月`,stage:months<=6?'新品期':months<=24?'稳定期':'衰弱期'};
  }
  function rawValue(record,field,headers) { return record.values[headers.indexOf(field)]; }
  function learnedValue(record,field,headers) {
    const v=rawValue(record,field,headers);
    if (PRICES.includes(field)) return !missing(v) && (typeof v==='number'||typeof v==='string') && Number.isFinite(Number(v)) ? Number(v) : null;
    if (missing(v)) return null;
    if (field!=='型号' && numericCategory(v)) return null;
    return text(v);
  }
  function options(records,field,headers) {
    const map=new Map();
    for(const record of records) {
      const v=learnedValue(record,field,headers);
      if(v===null)continue;
      const id=key([typeof v,v]);
      if(!map.has(id))map.set(id,{value:v,sourceRows:[]});
      map.get(id).sourceRows.push(record.rowNumber);
    }
    return [...map.values()].map(item=>({...item,sourceRows:unique(item.sourceRows)}));
  }
  const detail = (method,reason,candidates=[],sourceRows=[]) => ({method,reason,sourceRows:unique(sourceRows.length?sourceRows:candidates.flatMap(c=>c.sourceRows)),candidates});
  function choose(records,field,headers,method,reason) {
    const candidates=options(records,field,headers);
    if(!candidates.length)return null;
    return {value:candidates.length===1?candidates[0].value:'待确认',field:detail(candidates.length===1?method:'pending',candidates.length===1?reason:reason+'存在不同值，需确认',candidates),conflict:candidates.length>1};
  }
  function states(values,headers) {
    const ss=unique(['领星SKU','SKU'].map(h=>sku(values[headers.indexOf(h)]).match(/^(RE|K1)-/i)?.[1]?.toUpperCase()).filter(Boolean));
    const name=text(values[headers.indexOf('金蝶名称')]);
    const ns=/^全新换包装/.test(name)?'RE':/^二手/.test(name)?'K1':'';
    return ss.length>1 || (ss[0]&&ns&&ss[0]!==ns) ? 'conflict' : ss[0]||ns||'';
  }
  function model(values,headers) {
    const valuesToTry=['领星SKU','SKU'].map(h=>sku(values[headers.indexOf(h)])).filter(Boolean);
    const roots=unique(valuesToTry.map(v=>v.replace(/^(RE|K1)-/i,'').match(/^([A-Z]{1,8}\d{1,6})(?:-|$)/i)?.[1]?.toUpperCase()).filter(Boolean));
    if(roots.length===1)return roots[0];
    if(roots.length>1)return '';
    return text(values[headers.indexOf('金蝶名称')]).replace(/^(全新换包装|二手)/,'').match(/^([A-Z]{1,8}\d{1,6})(?!\d)/i)?.[1]?.toUpperCase() || '';
  }
  function prepare(analysis,settings={}) {
    analysis.date1904=!!settings.date1904;
    analysis.asOfDate=settings.asOfDate || today();
    const asOf=dateValue(analysis.asOfDate,analysis.date1904);
    if(!asOf)analysis.issues.push({severity:'error',code:'invalid-as-of-date',message:'处理日期无效，请使用YYYY-MM-DD'});
    const history=analysis.groups.map(group=>({code:group.code,...group.rows[0]})), headers=analysis.headers;
    const index={sku:new Map(),name:new Map(),model:new Map(),group:new Map(),gh:new Map(),g:new Map(),state:new Map()};
    const add=(map,k,r)=>{if(!k)return;if(!map.has(k))map.set(k,[]);map.get(k).push(r);};
    for(const r of history) {
      const get=h=>text(rawValue(r,h,headers)),group=get('物料分组'),name=get('金蝶名称'),base=model(r.values,headers),g=get('销售产品线'),h=get('销售系列');
      for(const s of unique(['领星SKU','SKU'].map(f=>sku(rawValue(r,f,headers))).filter(Boolean)))add(index.sku,s,r);
      if(group&&name)add(index.name,key([group,name]),r);
      if(group&&base)add(index.model,key([group,base]),r);
      add(index.group,group,r);add(index.g,g,r);
      if(g&&h){add(index.gh,key([g,h]),r);const state=states(r.values,headers);if(state&&state!=='conflict')add(index.state,key([g,h,state]),r);}
    }
    analysis.learning={history,index,asOf};
    return analysis;
  }
  function matches(index,keys) { return [...new Map(keys.flatMap(k=>index.get(k)||[]).map(r=>[r.code,r])).values()]; }
  function baseRecords(analysis,values) {
    const keys=['领星SKU','SKU'].map(h=>sku(values[analysis.headers.indexOf(h)])).filter(s=>/^(RE|K1)-/i.test(s)).map(s=>s.replace(/^(RE|K1)-/i,''));
    return matches(analysis.learning.index.sku,keys).filter(r=>!states(r.values,analysis.headers));
  }
  function tiers(analysis,values) {
    const hs=analysis.headers,ix=analysis.learning.index,get=h=>text(values[hs.indexOf(h)]),group=get('物料分组'),name=get('金蝶名称'),base=model(values,hs);
    return [
      ['history-sku','按相同SKU的已维护记录',matches(ix.sku,unique(['领星SKU','SKU'].map(h=>sku(values[hs.indexOf(h)])).filter(Boolean)))],
      ['history-name','按同物料分组、相同商品名称',group&&name?ix.name.get(key([group,name]))||[]:[]],
      ['history-model','按同物料分组、基础型号'+base,group&&base?ix.model.get(key([group,base]))||[]:[]],
      ['history-group','按相同完整物料分组',group?ix.group.get(group)||[]:[]]
    ];
  }
  function lookup(analysis,values,field,limit=4) {
    for(const [method,reason,records]of tiers(analysis,values).slice(0,limit)) {
      const selected=choose(sameState(analysis,values,field,records),field,analysis.headers,method,reason);
      if(selected)return selected;
    }
    return null;
  }
  function sameState(analysis,values,field,records) {
    if(!['销售产品分类','三级分类','型号','是否正常备货'].includes(field))return records;
    const state=states(values,analysis.headers);
    return state==='conflict'?[]:records.filter(r=>states(r.values,analysis.headers)===state);
  }
  function countryFacts(values,headers) {
    let name=text(values[headers.indexOf('金蝶名称')]);
    const skus=['领星SKU','SKU'].map(h=>sku(values[headers.indexOf(h)])).join(' ');
    const countries=[];
    const names=['印度尼西亚','马来西亚','澳大利亚','意大利','西班牙','加拿大','新加坡','俄罗斯','法国','德国','沙特','印度','越南','韩国','智利','波兰','日本','美国','英国','中国','国内','印尼'];
    for(const word of names.sort((a,b)=>b.length-a.length))if(name.includes(word)) {
      countries.push(word==='国内'?'中国':word==='印尼'?'印度尼西亚':word);
      name=name.split(word).join(' ');
    }
    if(/\bUSA?\b/i.test(name)||/(?:^|[- ])USA?(?:[- ]|$)/i.test(skus))countries.push('美国');
    if(/\bUK\b/i.test(name)||/(?:^|[- ])UK(?:[- ]|$)/i.test(skus))countries.push('英国');
    if(/\bCHINA\b/i.test(name)||/(?:^|[- ])CN(?:[- ]|$)/i.test(skus))countries.push('中国');
    return unique(countries);
  }
  function countryValue(analysis,values,field) {
    const countries=countryFacts(values,analysis.headers);
    const name=text(values[analysis.headers.indexOf('金蝶名称')]);
    if(countries.length>1)return {value:'待确认',conflict:true,field:detail('pending','名称或SKU包含多个国家（'+countries.join('、')+'），请确认',field==='国家'?countries.map(value=>({value,sourceRows:[]})):[])};
    if(field==='国家')return countries.length?{value:countries[0],field:detail('derived','名称或SKU明确标注国家',[{value:countries[0],sourceRows:[]}])}:null;
    const regionWords=['欧洲','欧盟'].filter(word=>name.includes(word));
    if(!countries.length&&!regionWords.length)return null;
    const country=countries[0],fixed={美国:'美国',英国:'欧洲',中国:'中国'};
    if(country&&fixed[country])return {value:fixed[country],field:detail('derived','根据明确国家填写销售区域；英国对应欧洲',[{value:fixed[country],sourceRows:[]}])};
    const exact=lookup(analysis,values,field,2);if(exact)return exact;
    const records=analysis.learning.history.filter(r=>{
      const found=countryFacts(r.values,analysis.headers);
      return country?found.length===1&&found[0]===country:regionWords.some(word=>text(rawValue(r,'金蝶名称',analysis.headers)).includes(word));
    });
    return choose(records,field,analysis.headers,'derived','按明确国家或地区对应的唯一历史区域')
      || {value:'待确认',conflict:true,field:detail('pending','已识别国家或地区，但没有唯一的历史销售区域，请确认')};
  }
  function dependencies(analysis,values,field) {
    const get=h=>text(values[analysis.headers.indexOf(h)]),g=get('销售产品线'),h=get('销售系列'),ix=analysis.learning.index;
    if(!missing(g)&&!missing(h)) {
      const picked=choose(sameState(analysis,values,field,ix.gh.get(key([g,h]))||[]),field,analysis.headers,'derived','按销售产品线、销售系列的已维护值');
      if(picked)return picked;
    }
    return field==='场景分组'&&!missing(g)?choose(ix.g.get(g)||[],field,analysis.headers,'derived','按销售产品线的已维护场景'):null;
  }
  function returnValue(analysis,values,field) {
    const headers=analysis.headers,state=states(values,headers);
    if(!state)return null;
    if(state==='conflict')return {value:'待确认',conflict:true,field:detail('pending','名称与SKU的退回状态冲突')};
    const bases=baseRecords(analysis,values),get=h=>text(values[headers.indexOf(h)]),g=get('销售产品线'),h=get('销售系列');
    if(['销售产品线','销售系列','场景分组','一级品类','二级分类','销售区域'].includes(field)) {
      const bg=options(bases,'销售产品线',headers),bh=options(bases,'销售系列',headers);
      if((!missing(g)&&bg.length===1&&g!==bg[0].value)||(!missing(h)&&bh.length===1&&h!==bh[0].value))return null;
      return choose(bases,field,headers,'derived','去掉RE-/K1-后，SKU精确对应原成品');
    }
    if(['销售产品分类','三级分类'].includes(field))return (!missing(g)&&!missing(h)?choose(analysis.learning.index.state.get(key([g,h,state]))||[],field,headers,'derived','按同系列、同退回状态的已维护记录'):null)
      || {value:'待确认',conflict:true,field:detail('pending','缺少同系列、同退回状态的分类记录，请确认')};
    if(field!=='型号'||missing(g)||missing(h))return null;
    const qs=options(bases,'型号',headers);
    if(qs.length!==1||/售后配件|待确认|定制|--/.test(String(qs[0]?.value||'')))return {value:'待确认',conflict:true,field:detail('pending','无法唯一确认退回商品对应的原成品型号',qs)};
    const pairs=[];
    for(const record of analysis.learning.index.state.get(key([g,h,state]))||[]) {
      const actual=learnedValue(record,'型号',headers);
      if(actual===null||/--/.test(String(actual)))continue;
      const oldBase=baseRecords(analysis,record.values),oldQ=options(oldBase,'型号',headers);
      if(oldQ.length!==1||/售后配件|待确认|定制|--/.test(String(oldQ[0].value)))continue;
      pairs.push({record,valid:actual===state+'-'+oldQ[0].value,sourceRows:[record.rowNumber,...oldQ[0].sourceRows]});
    }
    if(pairs.length<2||pairs.some(pair=>!pair.valid))return {value:'待确认',conflict:true,field:detail('pending',pairs.length<2?'同系列同状态不足两组成对型号记录':'同系列退回型号转换存在反例，需确认',qs,pairs.flatMap(p=>p.sourceRows))};
    return {value:state+'-'+qs[0].value,field:detail('derived','原成品型号增加'+state+'-；已由至少两组同系列同状态记录验证',[],[...qs[0].sourceRows,...pairs.flatMap(p=>p.sourceRows)])};
  }
  function automatic(analysis,values,field,manual) {
    const headers=analysis.headers,name=text(values[headers.indexOf('金蝶名称')]);
    if(['销售产品分类','三级分类','型号','是否正常备货'].includes(field)&&states(values,headers)==='conflict')return {value:'待确认',conflict:true,field:detail('pending','名称与SKU的退回状态冲突')};
    if(field==='品牌') {
      const brands=unique(analysis.learning.history.map(r=>learnedValue(r,field,headers)).filter(v=>v!==null&&name.includes(v))).map(value=>({value,sourceRows:analysis.learning.history.filter(r=>learnedValue(r,field,headers)===value).map(r=>r.rowNumber)}));
      if(brands.length)return {value:brands.length===1?brands[0].value:'待确认',field:detail(brands.length===1?'derived':'pending',brands.length===1?'商品名称明确包含品牌':'名称包含多个品牌，需确认',brands)};
      return lookup(analysis,values,field,2);
    }
    if(field==='是否正常备货') {
      if(/(?:不再|不|非|暂停|停止|暂不|取消|无需|未|待确认)[^，。；;\n]{0,12}正常备货|正常备货[^，。；;\n]{0,6}(?:暂停|停止|取消)/.test(name))return {value:'待确认',field:detail('pending','名称包含否定或暂停正常备货的表述，请确认')};
      if(/正常备货/.test(name))return {value:'正常备货',field:detail('derived','商品名称明确标注正常备货')};
      return lookup(analysis,values,field,2);
    }
    if(PRICES.includes(field))return null;
    if(['国家','销售区域'].includes(field)) {
      const named=countryValue(analysis,values,field);if(named)return named;
      const exact=lookup(analysis,values,field,2);if(exact)return exact;
      if(field==='国家')return null;
    }
    if(['场景分组','一级品类','二级分类','三级分类'].includes(field)&&(own(manual,'销售产品线')||own(manual,'销售系列'))) {
      const exactRows=tiers(analysis,values)[0][2];
      const changed=['销售产品线','销售系列'].some(h=>own(manual,h)&&(()=>{const known=options(exactRows,h,headers);return known.length!==1||text(manual[h])!==known[0].value;})());
      if(changed)return dependencies(analysis,values,field)||{value:'待确认',conflict:true,field:detail('pending','人工产品线或系列已变更，尚无对应的分类依据')};
    }
    const direct=lookup(analysis,values,field,2);if(direct)return direct;
    const returned=returnValue(analysis,values,field);if(returned)return returned;
    for(const [method,reason,records]of tiers(analysis,values).slice(2)) {
      const chosen=choose(sameState(analysis,values,field,records),field,headers,method,reason);if(chosen)return chosen;
    }
    return ['场景分组','一级品类','二级分类','三级分类'].includes(field)?dependencies(analysis,values,field):null;
  }
  function enrich(analysis,rows,overrides={}) {
    const headers=analysis.headers,newCodes=new Set(analysis.added.map(r=>r.code)),asOf=analysis.learning?.asOf;
    const pendingFields=[],addPending=(code,field,reason,candidates=[])=>{if(!pendingFields.some(p=>p.code===code&&p.field===field))pendingFields.push({code,field,reason,candidates});};
    for(const row of rows) {
      if(!newCodes.has(row.code))continue;
      const manual=own(overrides,row.code)?overrides[row.code]||{}:{},fields={};row.fields=fields;
      const set=(field,v,metadata)=>{const c=headers.indexOf(field);if(c<0)return;row.values[c]=v;row.cells[c]={t:typeof v==='number'?'n':'s',v};fields[field]=metadata;};
      for(const field of ['销售产品线','销售系列','场景分组','销售产品分类','一级品类','二级分类','三级分类','型号','销售区域','品牌','国家','是否正常备货',...PRICES]) {
        if(own(manual,field)) {
          let v=manual[field]??'';
          if(PRICES.includes(field)&&text(v)==='')v='';
          else if(PRICES.includes(field)) {
            if(!['number','string'].includes(typeof v)||!Number.isFinite(Number(v))){set(field,v,detail('manual','人工金额不是有效有限数值'));addPending(row.code,field,'请输入有效数值，金额允许0；也可清空');continue;}
            v=Number(v);
          } else if(!PRICES.includes(field))v=text(v);
          set(field,v,detail('manual',text(v)===''?'已人工清空，不自动覆盖':'人工填写优先'));
        } else {
          const selected=automatic(analysis,row.values,field,manual);
          if(selected&&(!PRICES.includes(field)||!selected.conflict))set(field,selected.value,selected.field);
          else set(field,PRICES.includes(field)?'':'待确认',selected?.field||detail(PRICES.includes(field)?'empty':'pending',PRICES.includes(field)?'新增编码没有可靠价格来源，留空可下载，可人工填写':'没有无冲突的历史依据，需确认'));
        }
      }
      const rawCreated=own(manual,'金蝶创建时间')?manual['金蝶创建时间']:row.values[headers.indexOf('金蝶创建时间')];
      const created=dateValue(rawCreated,analysis.date1904),creationMethod=own(manual,'金蝶创建时间')?'manual':'source';
      if(!created||!asOf||created.serial>asOf.serial) {
        const reason=!created?'创建日期缺失或无法识别，请填写有效日期':'创建日期晚于处理日期，请核对';
        set('金蝶创建时间',rawCreated??'',detail(creationMethod,reason));addPending(row.code,'金蝶创建时间',reason);
        set('上架时间（月）','待确认',detail('pending','需先补全有效创建日期'));
        set('生命周期','待确认',detail('pending','需先补全有效创建日期'));
      } else {
        set('金蝶创建时间',created.serial,detail(creationMethod,'按目标工作簿日期系统保存日期'));
        row.cells[headers.indexOf('金蝶创建时间')].z='yyyy-mm-dd';
        const computed=age(created,asOf);
        set('上架时间（月）',computed.age,detail('derived','按创建日期至处理日期计算完整月数'));
        set('生命周期',computed.stage,detail('derived','完整月数≤6新品期，≤24稳定期，其余衰弱期'));
      }
      if(asOf){set('当前日期',asOf.serial,detail('derived','北京时间处理日期 '+analysis.asOfDate));row.cells[headers.indexOf('当前日期')].z='yyyy-mm-dd';}
      for(const field of REQUIRED_CLASS)if(missing(row.values[headers.indexOf(field)]))addPending(row.code,field,fields[field]?.reason||'必填字段待补',fields[field]?.candidates||[]);
      for(const field of [...REQUIRED_CLASS,'品牌','国家','是否正常备货'])if(numericCategory(row.values[headers.indexOf(field)])) {
        fields[field].reason='分类字段不能填写纯数值，请填写实际分类';
        addPending(row.code,field,fields[field].reason);
      }
    }
    return {rows,pendingFields,pendingRows:new Set(pendingFields.map(p=>p.code)).size};
  }
  function validateResult(analysis,result) {
    const errors=[],push=(code,field,reason)=>errors.push({code,field,reason,message:reason}),headers=analysis.headers;
    if(!Array.isArray(result?.rows)){push('','','缺少输出数据');return errors;}
    const newCodes=new Set(analysis.added.map(r=>r.code)),expected=new Map([...analysis.groups.map(g=>[g.code,g]),...analysis.added.map(g=>[g.code,g])]),seen=new Set();
    const asOf=dateValue(analysis.asOfDate,analysis.date1904);
    for(const row of result.rows) {
      if(seen.has(row.code)){push(row.code,'物料编码','物料编码重复');continue;}seen.add(row.code);
      if(!expected.has(row.code)){push(row.code,'物料编码','输出包含来源之外的物料编码');continue;}
      if(!Array.isArray(row.values)||row.values.length!==headers.length){push(row.code,'','输出列数不一致');continue;}
      const get=h=>row.values[headers.indexOf(h)];
      if(get('物料编码')!==row.code){push(row.code,'物料编码','输出编码与原始分组不一致');}
      if(!newCodes.has(row.code))continue;
      for(const field of REQUIRED_CLASS)if(typeof get(field)!=='string'||missing(get(field))||numericCategory(get(field)))push(row.code,field,'必填字段为空、待确认或为无效纯数值');
      for(const field of ['品牌','国家','是否正常备货'])if(numericCategory(get(field)))push(row.code,field,'分类字段不能填写纯数值');
      for(const field of PRICES)if(get(field)!==''&&get(field)!=null&&(typeof get(field)!=='number'||!Number.isFinite(get(field))))push(row.code,field,'价格必须是有限数值或空白');
      const created=typeof get('金蝶创建时间')==='number'?dateValue(get('金蝶创建时间'),analysis.date1904):null;
      if(!created||!asOf||created.serial>asOf.serial||created.serial!==get('金蝶创建时间'))push(row.code,'金蝶创建时间','创建日期必须为有效、非未来的Excel日期数值');
      if(!asOf||typeof get('当前日期')!=='number'||get('当前日期')!==asOf.serial)push(row.code,'当前日期','当前日期与处理日期不一致');
      if(created&&asOf&&created.serial<=asOf.serial){const computed=age(created,asOf);if(get('上架时间（月）')!==computed.age)push(row.code,'上架时间（月）','上架月龄与完整月数不一致');if(get('生命周期')!==computed.stage)push(row.code,'生命周期','生命周期与完整月数不一致');}
      else {if(missing(get('上架时间（月）')))push(row.code,'上架时间（月）','缺少有效月龄');if(missing(get('生命周期')))push(row.code,'生命周期','缺少有效生命周期');}
    }
    for(const code of expected.keys())if(!seen.has(code))push(code,'物料编码','输出缺少来源记录');
    for(const problem of analysis.issues.filter(i=>i.severity==='error'))push('','',problem.message);
    return errors;
  }
  return {REQUIRED_HEADERS,EDITABLE_HEADERS,FIELD_HEADERS,prepare,enrich,validateResult,dateValue};
});
