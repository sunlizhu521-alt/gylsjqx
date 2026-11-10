(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./quarter-budget-import.js'));
  else root.QuarterBudgetCore = factory(root.QuarterBudgetImport);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(Import) {
  'use strict';
  const HEADERS = ['预估付款月份','供应商名称','付款主体','对公付款0A流程号','摘要','申请付款金额','是否已支付','备注','分类','跟单员','是否剔除'];
  const PRIORITY = {negotiated:0, ledger:1, budget:2};
  const CATEGORY = {negotiated:'谈判付款',ledger:'正常付款',budget:'预估付款'};
  const text = v => String(v ?? '').trim();
  const norm = v => text(v).normalize('NFKC').replace(/\s+/g, '');
  const entity = v => ({'浙江迈德斯特医疗器械科技有限公司':'迈德斯特','杭州国源养老科技有限公司':'国源'}[norm(v)] || norm(v));
  const usable = v => v && !['待核对','未提交OA','不支付'].includes(norm(v));
  function oaKey(value) {
    const key = text(value).normalize('NFKC').toUpperCase();
    return /^[A-Z][A-Z0-9_-]*\d[A-Z0-9_-]*$/.test(key) ? key : '';
  }
  const sourceLabel = s => `${s?.file || '来源'} / ${s?.sheet || '工作表'} 第${s?.row || '?'}行${s?.part ? ' 第'+s.part+'期' : ''}`;
  const append = (record, note) => { if (note && !String(record.notes || '').includes(note)) record.notes = [record.notes, note].filter(Boolean).join('\n'); };
  const formatMoney = cents => cents == null ? '' : (cents / 100).toLocaleString('zh-CN', {minimumFractionDigits:2,maximumFractionDigits:2});
  const monthLabel = record => /^\d{4}-\d{2}$/.test(record.monthKey || '') ? `${record.monthKey.slice(0,4)}年${Number(record.monthKey.slice(5))}月` : '月份待核对';
  const toRow = r => [r.month ?? null,r.supplier,r.subject,r.oa,r.summary,r.amount ?? null,r.status,r.notes,r.category,r.owner,r.excluded];
  const message = w => typeof w === 'string' ? w : w?.message || String(w);
  function sameParty(a,b) {
    if (usable(a.supplier) && usable(b.supplier) && norm(a.supplier) !== norm(b.supplier)) return false;
    if (usable(a.subject) && usable(b.subject) && entity(a.subject) !== entity(b.subject)) return false;
    return true;
  }
  function matchingIdentity(record, evidence) {
    if (!sameParty(record, evidence)) return false;
    const a=oaKey(record.oa), b=oaKey(evidence.oa);
    if (a && b) return a===b;
    return !a && !b && usable(record.supplier) && norm(record.supplier)===norm(evidence.supplier) && usable(record.subject) && entity(record.subject)===entity(evidence.subject) && norm(record.summary) && norm(record.summary)===norm(evidence.summary);
  }
  function moneyInNotes(record, cents) {
    if (!/差额|排期/.test(record.notes || '')) return false;
    return [...String(record.notes).matchAll(/\d[\d,]*(?:\.\d{1,2})?/g)].some(m=>Math.round(Number(m[0].replace(/,/g,''))*100)===cents);
  }
  const evidenceOrder = e => [e.monthKey || '9999-99',String(e.source?.row || 0).padStart(8,'0'),String(e.order ?? e.part ?? e.source?.part ?? 0).padStart(8,'0'),e.id].join(':');
  function summarize(records) {
    const groups=new Map(), months=new Map();
    function add(map,key,r) {
      if (!map.has(key)) map.set(key,{monthKey:r.monthKey || '',category:r.category,total:0n,paid:0n});
      if (r.cents===null || r.cents===undefined) return;
      if (!Number.isSafeInteger(r.cents)) throw Error(sourceLabel(r.source)+'：金额超过可安全统计范围');
      const t=map.get(key);t.total+=BigInt(r.cents);if(r.status==='已支付')t.paid+=BigInt(r.cents);
    }
    for (const r of records.filter(r=>r.excluded!=='是')) { add(groups,JSON.stringify([r.monthKey || '',r.category]),r);add(months,r.monthKey || '',r); }
    function finish(t,category=true) {
      if ([t.total,t.paid,t.total-t.paid].some(v=>v>BigInt(Number.MAX_SAFE_INTEGER)||v< -BigInt(Number.MAX_SAFE_INTEGER))) throw Error('汇总金额超过可安全统计范围');
      return {monthKey:t.monthKey,...(category?{category:t.category}:{}),totalCents:Number(t.total),paidCents:Number(t.paid),remainingCents:Number(t.total-t.paid)};
    }
    const order=['谈判付款','正常付款','预估付款','暂停付款'];
    const byMonth=(a,b)=>(a.monthKey || '9999-99').localeCompare(b.monthKey || '9999-99');
    const classified=[...groups.values()].map(t=>finish(t)).sort((a,b)=>byMonth(a,b)||(order.includes(a.category)?order.indexOf(a.category):99)-(order.includes(b.category)?order.indexOf(b.category):99)||a.category.localeCompare(b.category,'zh-CN'));
    const monthly=[...months.values()].map(t=>finish(t,false)).sort(byMonth);
    const all={total:0n,paid:0n};for(const t of months.values()){all.total+=t.total;all.paid+=t.paid;}
    const totals=finish(all,false);delete totals.monthKey;
    return {classified,monthly,totals};
  }
  function build(inputs, options={}) {
    const mode=options.mode || 'new', decisions=options.decisions || {}, matches=options.matches || {};
    if(!['new','existing'].includes(mode))throw Error('请选择新建整理或继续已有明细');
    for(const role of mode==='existing'?['existing']:['negotiated','ledger','budget'])if(!inputs[role]?.book)throw Error('请先上传'+({negotiated:'谈判付款',ledger:'资金台账整理',budget:'预估预算',existing:'已有付款汇总明细'}[role]));
    const parsed={},warnings=[],issues=[],errors=[];
    for(const role of ['negotiated','ledger','budget','evidence','existing'])if(inputs[role]?.book){
      parsed[role]=Import.parse(inputs[role],role,options);
      warnings.push(...(parsed[role].warnings || []).map(message));
      errors.push(...(parsed[role].errors || []).map(message));
    }
    let records=(mode==='existing'?parsed.existing.records:['negotiated','ledger','budget'].flatMap(role=>parsed[role].records.map(r=>({...r,category:CATEGORY[role]})))).map((r,index)=>({...r,source:{...r.source},notes:r.notes || '',status:r.status || '待核对',excluded:r.excluded==='是'?'是':'否',index,related:false}));
    const deleted=[];
    if(mode==='new'){
      const groups=new Map();
      for(const r of records){const oa=oaKey(r.oa);if(!oa||r.cents==null)continue;const key=oa+'|'+r.cents;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);}
      const remove=new Set();
      for(const group of groups.values()){
        const ranks=group.map(r=>PRIORITY[r.source.role]);
        if(new Set(ranks).size<2)continue;
        const best=Math.min(...ranks),kept=group.filter(r=>PRIORITY[r.source.role]===best);
        for(const r of group)if(PRIORITY[r.source.role]!==best){remove.add(r.id);deleted.push({record:r,kept,reason:'跨表OA号和金额完全一致，按谈判付款、资金台账、预估预算的优先级保留。'});}
      }
      records=records.filter(r=>!remove.has(r.id));
      const parents=new Map(records.map(r=>[r.id,r.id]));
      function find(id){let p=parents.get(id);if(p!==id){p=find(p);parents.set(id,p);}return p;}
      for(let i=0;i<records.length;i++)for(let j=i+1;j<records.length;j++){
        const a=records[i],b=records[j];if(a.source.role===b.source.role)continue;
        const sameOA=oaKey(a.oa)&&oaKey(a.oa)===oaKey(b.oa);
        const sameSupplier=norm(a.supplier)&&norm(a.supplier)===norm(b.supplier)&&usable(a.subject)&&entity(a.subject)===entity(b.subject);
        const sameAmount=a.cents!=null&&a.cents===b.cents;
        const sameBusiness=a.monthKey&&a.monthKey===b.monthKey&&norm(a.summary)&&norm(a.summary)===norm(b.summary);
        if(sameOA || (sameSupplier&&(sameAmount||sameBusiness)))parents.set(find(b.id),find(a.id));
      }
      const related=new Map();for(const r of records){const k=find(r.id);if(!related.has(k))related.set(k,[]);related.get(k).push(r);}
      records=[...related.values()].flatMap(group=>{
        if(group.length>1){
          const label=group.map(r=>sourceLabel(r.source)).join('；');
          issues.push({id:'duplicate:'+group[0].id,type:'duplicate',message:'疑似重复，默认保留。请核对业务、金额和分期后确认是否剔除。'+label,recordIds:group.map(r=>r.id)});
          for(const r of group){r.related=true;append(r,'疑似重复待核对：'+label);}
        }
        return group;
      });
    }
    for(const r of records){
      if(decisions[r.id] && ['是','否'].includes(decisions[r.id].excluded)){
        r.excluded=decisions[r.id].excluded;
        if(r.excluded==='是')append(r,'人工确认剔除：'+(text(decisions[r.id].reason)||'已核对为重复款项'));
        else if(text(decisions[r.id].reason))append(r,'人工确认保留：'+text(decisions[r.id].reason));
      }
      if(!r.monthKey)warnings.push(sourceLabel(r.source)+'：付款年月未确定，计入“月份待核对”');
      if(r.cents==null)warnings.push(sourceLabel(r.source)+'：金额为空或无法识别，保留明细，未计入金额汇总');
    }
    // A dedicated schedule replaces the older schedule as evidence, not as row data.
    const evidenceSources=[parsed.evidence || parsed.negotiated,parsed.ledger].filter(Boolean);
    let evidence=[...new Map(evidenceSources.flatMap(p=>p.evidence || []).filter(e=>e.id&&Number.isSafeInteger(e.cents)).map(e=>[e.id,e])).values()];
    const conflictingProofs=new Set();
    // A ledger's aggregate state does not create another installment payment beside its schedule cell.
    evidence=evidence.filter(e=>{
      if(e.kind==='schedule')return true;
      const schedule=evidence.filter(s=>s.kind==='schedule'&&matchingIdentity(e,s)&&e.cents===s.cents&&(!e.monthKey||e.monthKey===s.monthKey)&&(!norm(e.summary)||!norm(s.summary)||norm(e.summary)===norm(s.summary)));
      if(!schedule.length)return true;
      if(e.monthKey&&['已支付','未支付'].includes(e.status)&&!schedule.some(s=>s.status===e.status))schedule.forEach(s=>conflictingProofs.add(s.id));
      return false;
    });
    const active=records.filter(r=>r.excluded!=='是'), candidates=new Map();
    for(const r of active){
      if(/原流程.*本期未提交|历史流程/.test(r.notes)){candidates.set(r.id,[]);continue;}
      let possible=evidence.filter(e=>matchingIdentity(r,e)&&(r.cents===e.cents||(r.evidenceIds || []).includes(e.id)||(r.status==='已支付'&&moneyInNotes(r,e.cents))));
      const otherMonths=possible.filter(e=>r.monthKey&&e.monthKey&&r.monthKey!==e.monthKey);
      possible=possible.filter(e=>!r.monthKey||!e.monthKey||r.monthKey===e.monthKey);
      if(!possible.length&&otherMonths.length){
        append(r,'付款依据月份不同：'+otherMonths.map(e=>monthLabel(e)+'（'+sourceLabel(e.source)+'）').join('；')+'；未用于本笔付款状态。');
        issues.push({id:'month:'+r.id,type:'evidence',message:'同OA或业务的同额依据属于其他月份，未认定本笔已付。',recordIds:[r.id]});
      }
      const own=possible.filter(e=>(r.evidenceIds || []).includes(e.id));
      if(own.length)possible=own;
      const sameMonth=possible.filter(e=>r.monthKey&&e.monthKey===r.monthKey);
      if(sameMonth.length)possible=sameMonth;
      const sameSummary=possible.filter(e=>norm(r.summary)&&norm(e.summary)===norm(r.summary));
      if(sameSummary.length)possible=sameSummary;
      candidates.set(r.id,possible.filter(e=>e.status==='已支付'||e.status==='未支付').sort((a,b)=>evidenceOrder(a).localeCompare(evidenceOrder(b))));
    }
    const used=new Map(),assigned=new Map(),conflicted=new Set();
    function assign(r,e){
      if(used.has(e.id)&&used.get(e.id)!==r.id){conflicted.add(r.id);conflicted.add(used.get(e.id));return false;}
      used.set(e.id,r.id);assigned.set(r.id,e);return true;
    }
    for(const r of active)if(matches[r.id]){
      const e=(candidates.get(r.id)||[]).find(e=>e.id===matches[r.id]);
      if(e)assign(r,e);else{conflicted.add(r.id);warnings.push(sourceLabel(r.source)+'：先前选择的付款依据已失效，请重新核对');}
    }
    const sets=new Map();
    for(const r of active.filter(r=>!assigned.has(r.id)&&!conflicted.has(r.id))){
      const list=candidates.get(r.id)||[];if(!list.length)continue;
      const key=list.map(e=>e.id).sort().join('|');if(!sets.has(key))sets.set(key,[]);sets.get(key).push(r);
    }
    for(const group of [...sets.values()].sort((a,b)=>candidates.get(a[0].id).length-candidates.get(b[0].id).length)){
      const list=candidates.get(group[0].id).filter(e=>!used.has(e.id));
      if(group.length===1&&list.length===1){assign(group[0],list[0]);continue;}
      if(group.length===list.length&&list.length>1){
        if(new Set(list.map(e=>e.status)).size===1){group.forEach((r,i)=>assign(r,list[i]));continue;}
        // Existing explicit states may locate otherwise identical installments; edited row order cannot.
        if(['已支付','未支付'].every(status=>group.filter(r=>r.status===status).length===list.filter(e=>e.status===status).length)){
          for(const r of group)assign(r,list.find(e=>!used.has(e.id)&&e.status===r.status));
          continue;
        }
      }
      const paid=group.filter(r=>r.status==='已支付');
      if(list.length===1&&paid.length===1&&list[0].status==='已支付')assign(paid[0],list[0]);
      for(const r of group)if(!assigned.has(r.id))conflicted.add(r.id);
    }
    for(const r of active){
      const e=assigned.get(r.id);
      if(e&&conflictingProofs.has(e.id)){
        r.status='待核对';append(r,'付款状态待核对：台账明确状态与同月分期付款依据冲突。');
        issues.push({id:'proof-conflict:'+r.id,type:'evidence',message:'同月台账状态与分期付款依据冲突，请核对。',recordIds:[r.id]});
      }else if(conflicted.has(r.id)){
        r.status='待核对';append(r,'付款状态待核对：同一付款依据对应多笔记录，或分期对应关系存在冲突，未重复分配已付款。');
        issues.push({id:'evidence:'+r.id,type:'evidence',message:'付款依据需要确认：'+sourceLabel(r.source),recordIds:[r.id],candidates:(candidates.get(r.id)||[]).map(e=>({id:e.id,label:`${sourceLabel(e.source)}｜${monthLabel(e)}｜${e.summary || ''}｜${formatMoney(e.cents)}｜${e.status}`}))});
      }else if(e){
        if(r.status==='已支付'&&e.status==='未支付'){
          r.status='待核对';append(r,'付款状态待核对：原记录为已支付，对应来源金额仍标未支付（'+sourceLabel(e.source)+'）。');
          issues.push({id:'status:'+r.id,type:'evidence',message:'原付款状态与本笔来源证据冲突，需核对。',recordIds:[r.id]});
        }else r.status=e.status;
      }
      // Combined rows cannot reuse one paid component already represented elsewhere.
      const parts=[...String(r.notes || '').matchAll(/((?:\d[\d,]*(?:\.\d{1,2})?)[ \t]*[+＋][ \t]*(?:\d[\d,]*(?:\.\d{1,2})?))/g)].flatMap(m=>m[1].split(/[+＋]/).map(v=>Math.round(Number(v.replace(/[,\s]/g,''))*100)));
      if(parts.length>=2&&parts.reduce((a,b)=>a+b,0)===r.cents){
        const components=parts.map(c=>evidence.filter(e=>e.cents===c&&sameParty(r,e)&&norm(r.supplier)===norm(e.supplier)));
        const paid=components.flat().filter(e=>e.status==='已支付');
        if(paid.length&&components.some(es=>es.some(e=>e.status==='未支付'))){
          r.status='待核对';append(r,'付款状态待核对：合并金额包含已付和未付部分，已付部分可能与其他明细共用付款依据；请核对是否重复计入及所属月份。');
          issues.push({id:'components:'+r.id,type:'evidence',message:'合并金额包含不同付款状态，不能认定整笔已支付。',recordIds:[r.id]});
        }
      }
    }
    let summary;
    try{summary=summarize(records);}catch(e){errors.push(e.message);summary={classified:[],monthly:[],totals:{totalCents:null,paidCents:null,remainingCents:null}};}
    issues.push(...errors.map((message,i)=>({id:'error:'+i,type:'error',message,recordIds:[]})));
    return {records,deleted,issues,warnings:[...new Set([...warnings,...errors])],summary,sources:Object.fromEntries(Object.entries(parsed).map(([role,p])=>[role,{filename:inputs[role].filename,sheets:p.sheets}])),canExport:errors.length===0};
  }
  return {HEADERS,build,summarize,toRow,monthLabel,formatMoney,oaKey};
});
