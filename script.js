/* ── PALETTE ── */
const CC  = ['#0B3D2E','#1F7A4C','#7BC67B','#2FA85F','#C99A2E','#3D7A5E'];
const CC2 = ['#0B3D2E','#1F7A4C','#7BC67B','#2FA85F','#C99A2E','#3D7A5E'];
const CI  = {
  'Fertilizante/ha':'#0B3D2E','Defensivos/ha':'#1A6B40','Sementes/ha':'#2FA85F',
  'Combustivel/ha':'#7BC67B','Mão de obra/ha':'#A8D8A8','Manutenção/ha':'#C4EAC4'
};
const charts = {};

/* ── STATE ── */
let raw = [], filtered = [], empresas = [];

/* ── CHART DEFAULTS ── */
Chart.defaults.font.family = 'Trebuchet MS,Trebuchet,Arial,sans-serif';
const cd = {
  maintainAspectRatio:false, animation:{duration:500},
  plugins:{legend:{display:false},tooltip:{bodyFont:{family:'Trebuchet MS'},titleFont:{family:'Trebuchet MS',weight:'700'}}},
  scales:{
    x:{grid:{display:false},ticks:{font:{family:'Trebuchet MS',size:10},color:'#5A6E5A'},border:{display:false}},
    y:{grid:{color:'#EDF3ED'},ticks:{font:{family:'Trebuchet MS',size:10},color:'#5A6E5A'},border:{display:false}}
  }
};

/* ══════════════════════════════════════════════════════
   FILE HANDLING
══════════════════════════════════════════════════════ */
function handleDrop(e){
  e.preventDefault();
  document.getElementById('upCard').classList.remove('drag-over');
  const f=e.dataTransfer.files[0]; if(f) processFile(f);
}
function handleFile(e){const f=e.target.files[0];if(f)processFile(f);}
function processFile(file){
  if(!file.name.match(/\.(xlsx|xls)$/i)){alert('Selecione um arquivo .xlsx');return;}
  showLoad(true);
  const rd=new FileReader();
  rd.onload=e=>{
    try{parseWB(XLSX.read(e.target.result,{type:'array'}),file.name);}
    catch(err){showLoad(false);alert('Erro: '+err.message);}
  };
  rd.readAsArrayBuffer(file);
}
function parseWB(wb,fname){
  const baseSheet=wb.Sheets['BASE']||wb.Sheets[wb.SheetNames[0]];
  if(!baseSheet){showLoad(false);alert('Aba BASE não encontrada.');return;}
  const empSheet=wb.Sheets['EMPRESAS'];
  if(empSheet) empresas=XLSX.utils.sheet_to_json(empSheet);

  raw=XLSX.utils.sheet_to_json(baseSheet).map(r=>{
    const n={}; Object.keys(r).forEach(k=>n[k.trim()]=r[k]);
    let v=n['Valor']; if(typeof v==='string') v=parseFloat(v.replace(',','.'))||null;
    return {...n,Valor:v};
  });
  populateFilters();
  showLoad(false);
  setOK(fname,raw.length);
  applyFilters();
}

/* ══════════════════════════════════════════════════════
   FILTERS
══════════════════════════════════════════════════════ */
function populateFilters(){
  const safras=[...new Set(raw.map(r=>r['Safra']).filter(Boolean))].sort().reverse();
  const culturas=[...new Set(raw.map(r=>r['Cultura']).filter(Boolean))].sort();
  const emps=[...new Set(raw.map(r=>r['Empresa']).filter(Boolean))].sort();
  fillSel('fSafra',safras,'Todas as Safras');
  fillSel('fCultura',culturas,'Todas as Culturas');
  fillSel('fEmpresa',emps,'Todas as Empresas');
  if(safras.length) document.getElementById('fSafra').value=safras[0];
  // build culture tabs for ranking
  buildCultureTabs(culturas);
}
function fillSel(id,items,ph){
  const s=document.getElementById(id);
  s.innerHTML=`<option value="">${ph}</option>`+items.map(i=>`<option value="${i}">${i}</option>`).join('');
}
function buildCultureTabs(culturas){
  const tb=document.getElementById('tabsRkCult');
  tb.innerHTML='<button class="tab on" onclick="filterRankByCulture(\'\',this)">Todas</button>'
    +culturas.map(c=>`<button class="tab" onclick="filterRankByCulture('${c}',this)">${c}</button>`).join('');
}
function switchRkInd(wrapId,btn){
  ['rkProdWrap','rkCustoWrap','rkMargemWrap'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.style.display=id===wrapId?'block':'none';
  });
  document.querySelectorAll('#tabsRkInd .tab').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
}
let currentRankCulture='';
function filterRankByCulture(cult,btn){
  document.querySelectorAll('#tabsRkCult .tab').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  currentRankCulture=cult;
  renderRankings(getEmps(filtered));
}
function applyFilters(){
  const sf=document.getElementById('fSafra').value;
  const cu=document.getElementById('fCultura').value;
  const em=document.getElementById('fEmpresa').value;
  filtered=raw.filter(r=>{
    if(sf&&r['Safra']!==sf)return false;
    if(cu&&r['Cultura']!==cu)return false;
    if(em&&r['Empresa']!==em)return false;
    return true;
  });
  document.getElementById('rCount').textContent=filtered.length;
  document.getElementById('secSafra').textContent='· '+(sf||'Todas as Safras');
  if(raw.length>0) renderDashboard();
}

/* ══════════════════════════════════════════════════════
   DATA HELPERS
══════════════════════════════════════════════════════ */
function getVal(data,empresa,ind,avg=false){
  let rows=data.filter(r=>r['Indicador']===ind);
  if(empresa) rows=rows.filter(r=>r['Empresa']===empresa);
  const vs=rows.map(r=>r['Valor']).filter(v=>v!=null&&!isNaN(v)&&isFinite(v));
  if(!vs.length) return null;
  return avg?vs.reduce((a,b)=>a+b,0)/vs.length:vs.reduce((a,b)=>a+b,0);
}
function getEmps(data){return[...new Set(data.map(r=>r['Empresa']).filter(Boolean))].sort();}
function getSafras(){return[...new Set(raw.map(r=>r['Safra']).filter(Boolean))].sort();}
function getCultures(data){return[...new Set(data.map(r=>r['Cultura']).filter(Boolean))].sort();}
function fmtBRL(v,dec=0){if(v==null||isNaN(v))return'—';return'R$ '+v.toLocaleString('pt-BR',{minimumFractionDigits:dec,maximumFractionDigits:dec});}
function fmtK(v){if(v==null||isNaN(v))return'—';if(Math.abs(v)>=1e6)return'R$ '+(v/1e6).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'M';if(Math.abs(v)>=1e3)return'R$ '+(v/1e3).toFixed(0)+'K';return fmtBRL(v);}
function fmtN(v,d=1){if(v==null||isNaN(v))return'—';return v.toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d});}
function fmtP(v,d=1){if(v==null||isNaN(v))return'—';return(v*100).toFixed(d)+'%';}
function destr(id){if(charts[id]){charts[id].destroy();delete charts[id];}}
function buildLeg(emps,id){
  const el=document.getElementById(id);if(!el)return;
  el.innerHTML=emps.map((e,i)=>`<div class="leg-i"><div class="leg-d" style="background:${CC[i%CC.length]}"></div>${e}</div>`).join('');
}

/* ══════════════════════════════════════════════════════
   AGRO PERFORMANCE SCORE
══════════════════════════════════════════════════════ */
function calcScore(data, empresa){
  const emps=getEmps(data);
  // Indicators: higher is better unless marked ↓
  const indicators=[
    {ind:'Produtividade sc/ha', higher:true,  weight:25},
    {ind:'Margem EBITDA',       higher:true,  weight:25},
    {ind:'EBITDA/ha',          higher:true,  weight:20},
    {ind:'Custo por hectare',  higher:false, weight:15},
    {ind:'Resultado por hectare', higher:true, weight:15},
  ];
  let totalScore=0, totalWeight=0;
  indicators.forEach(cfg=>{
    const vals=emps.map(e=>getVal(data,e,cfg.ind,true)).filter(v=>v!=null);
    if(!vals.length) return;
    const empVal=getVal(data,empresa,cfg.ind,true);
    if(empVal==null) return;
    const mn=Math.min(...vals), mx=Math.max(...vals);
    const range=mx-mn; if(range===0) return;
    let norm=(empVal-mn)/range; // 0–1
    if(!cfg.higher) norm=1-norm;
    totalScore+=norm*cfg.weight;
    totalWeight+=cfg.weight;
  });
  if(!totalWeight) return {score:0,grade:'N/A'};
  const score=Math.round((totalScore/totalWeight)*100);
  let grade='F';
  if(score>=85) grade='A';
  else if(score>=70) grade='B';
  else if(score>=55) grade='C';
  else if(score>=40) grade='D';
  return {score,grade};
}
function scoreClass(g){return{A:'sc-A',B:'sc-B',C:'sc-C',D:'sc-D',F:'sc-F'}[g]||'sc-F';}

/* ══════════════════════════════════════════════════════
   RENDER DASHBOARD
══════════════════════════════════════════════════════ */
function renderDashboard(){
  document.getElementById('emptyState').style.display='none';
  document.getElementById('dashContent').style.display='block';
  const emps=getEmps(filtered);
  renderKPIs(emps);
  renderReceita(emps);
  renderCustoHa(emps);
  renderProdRanking(emps);
  renderScatter(emps);
  renderHeatmap(emps);
  renderRadar(emps);
  renderCustoTbl(emps);
  renderWaterfall(emps);
  renderCustoSaca(emps);
  renderEbitdaHa(emps);
  renderArea(emps);
  renderResultado(emps);
  renderEvolucao();
  renderRankings(emps);
  renderScoreChart(emps);
  renderExecTable(emps);
  renderExecHighlights(emps);
  renderIndicadoresTable(emps);
}

/* ══════════════════════════════════════════════════════
   KPIs
══════════════════════════════════════════════════════ */
function renderKPIs(emps){
  const cfgs=[
    {lbl:'Receita Líquida Total', ind:'Receita Líquida', fmtFn:fmtK,  unit:'', ico:'💵', sum:true},
    {lbl:'EBITDA Agro Total',     ind:'EBITDA Agro',     fmtFn:fmtK,  unit:'', ico:'📈', sum:true},
    {lbl:'Margem EBITDA Média',   ind:'Margem EBITDA',   fmtFn:v=>fmtP(v), unit:'', ico:'📊', avg:true},
    {lbl:'Produtividade Média',   ind:'Produtividade sc/ha', fmtFn:v=>fmtN(v)+' sc/ha', unit:'', ico:'🌾', avg:true},
    {lbl:'Custo/ha Médio',        ind:'Custo por hectare',fmtFn:v=>fmtBRL(v,0)+'/ha', unit:'', ico:'💸', avg:true, lowerBetter:true},
    {lbl:'Área Colhida Total',    ind:'Área Colhida (ha)',fmtFn:v=>fmtN(v,0)+' ha', unit:'', ico:'🗺️', sum:true},
    {lbl:'EBITDA/ha Médio',       ind:'EBITDA/ha',        fmtFn:v=>fmtBRL(v,0)+'/ha', unit:'', ico:'💎', avg:true},
    {lbl:'Resultado/ha Médio',    ind:'Resultado por hectare', fmtFn:v=>fmtBRL(v,0)+'/ha', unit:'', ico:'🎯', avg:true},
  ];
  const g=document.getElementById('kpiGrid');
  g.innerHTML='';
  cfgs.forEach((c,i)=>{
    let val;
    if (c.ind === 'Margem EBITDA') {
      const totalReceita = getVal(filtered, null, 'Receita Líquida', false);
      const totalEbitda = getVal(filtered, null, 'EBITDA Agro', false);
      val = (totalReceita && totalReceita !== 0) ? (totalEbitda / totalReceita) : 0;
    } else {
      val = c.avg ? getVal(filtered, null, c.ind, true) : getVal(filtered, null, c.ind, false);
    }
    // benchmark = avg of all companies for this indicator
    const bVals=emps.map(e=>getVal(filtered,e,c.ind,true)).filter(v=>v!=null);
    const bench=bVals.length?bVals.reduce((a,b)=>a+b,0)/bVals.length:null;
    // variation vs benchmark (only for avg)
    let varHtml='';
    if(val!=null && bench!=null && !c.sum && bVals.length>1){
      const diff=val-bench;
      const pct=bench!==0?Math.abs(diff/bench)*100:0;
      const better=c.lowerBetter?(diff<=0):(diff>=0);
      const arrow=better?'▲':'▼';
      varHtml=`<span class="badge ${better?'bp':'bn'}">${arrow} ${pct.toFixed(1)}% vs Bm</span>`;
      // Per-company badges
      emps.forEach((e,ei)=>{
        const ev=getVal(filtered,e,c.ind,true);
        if(ev==null||bench==null) return;
        const ed=ev-bench;
        const ep=bench!==0?Math.abs(ed/bench)*100:0;
        const eb=c.lowerBetter?(ed<=0):(ed>=0);
      });
    }
    // Best / worst company
    let benchNote='';
    if(bVals.length>1){
      const sorted=[...emps.map(e=>({e,v:getVal(filtered,e,c.ind,true)})).filter(r=>r.v!=null)];
      sorted.sort((a,b)=>c.lowerBetter?a.v-b.v:b.v-a.v);
      if(sorted.length) benchNote=`Benchmark: ${c.fmtFn(bench||0)}`;
    }
    const d=document.createElement('div');
    d.className='kpi-c'; d.style.animationDelay=(i*0.04)+'s';
    d.innerHTML=`<div class="kpi-lbl">${c.ico} ${c.lbl}</div>
      <div class="kpi-val">${val!=null?c.fmtFn(val):'—'}</div>
      <div class="kpi-row">${varHtml||'<span class="badge bx">'+emps.length+' empresa'+( emps.length!==1?'s':'')+'</span>'}</div>
      <div class="kpi-bench">${benchNote}</div>`;
    g.appendChild(d);
  });
}

/* ══════════════════════════════════════════════════════
   RECEITA LÍQUIDA
══════════════════════════════════════════════════════ */
function renderReceita(emps){
  destr('cReceita'); buildLeg(emps,'legReceita');
  // always use `filtered` so all active filters (cultura, empresa, safra) are respected
  const safras=[...new Set(filtered.map(r=>r['Safra']).filter(Boolean))].sort().reverse();
  const ds=emps.map((e,i)=>({
    label:e,
    data:safras.map(s=>{
      const rows=filtered.filter(r=>r['Empresa']===e&&r['Safra']===s&&r['Indicador']==='Receita Líquida');
      return rows.reduce((a,r)=>a+(r['Valor']||0),0)/1e6;
    }),
    backgroundColor:CC[i%CC.length]+'DD', borderRadius:5, borderSkipped:false
  }));
  const tot=ds.flatMap(d=>d.data).reduce((a,b)=>a+b,0);
  document.getElementById('bdReceita').textContent='Total: R$'+tot.toFixed(1)+'M';
  charts['cReceita']=new Chart(document.getElementById('cReceita'),{
    type:'bar', data:{labels:safras,datasets:ds},
    options:{...cd, plugins:{...cd.plugins,legend:{display:true,position:'top',
      labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,pointStyle:'circle',padding:10}}},
      scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v.toFixed(1)+'M'}}}}
  });
}

/* ══════════════════════════════════════════════════════
   CUSTO/HA
══════════════════════════════════════════════════════ */
function renderCustoHa(emps){
  destr('cCusto');
  const vals=emps.map(e=>getVal(filtered,e,'Custo por hectare',true));
  const avg=vals.filter(v=>v).reduce((a,b)=>a+b,0)/vals.filter(v=>v).length;
  document.getElementById('bdCusto').textContent='Média: R$'+Math.round(avg).toLocaleString('pt-BR');
  charts['cCusto']=new Chart(document.getElementById('cCusto'),{
    type:'bar',
    data:{labels:emps, datasets:[
      {data:vals, backgroundColor:vals.map(v=>v&&v<=avg?CC[1]+'CC':CC[3]+'CC'), borderRadius:7, borderSkipped:false},
      {type:'line',data:Array(emps.length).fill(avg), borderColor:CC[0], borderDash:[5,3], borderWidth:2, pointRadius:0}
    ]},
    options:{...cd, scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v.toLocaleString('pt-BR')}}}}
  });
}

/* ══════════════════════════════════════════════════════
   RANKING PRODUTIVIDADE
══════════════════════════════════════════════════════ */
function renderProdRanking(emps){
  destr('cProd');
  const cultFilter=document.getElementById('fCultura').value;
  const cults=getCultures(filtered);

  if(cultFilter){
    // Single culture: simple horizontal bar sorted by value
    const vals=emps.map(e=>getVal(filtered,e,'Produtividade sc/ha',true));
    const avg=vals.filter(v=>v).reduce((a,b)=>a+b,0)/vals.filter(v=>v).length;
    document.getElementById('bdProd').textContent='Média: '+avg.toFixed(1)+' sc/ha';
    const sorted=[...emps.map((e,i)=>({e,v:vals[i],c:CC[i%CC.length]}))].sort((a,b)=>(b.v||0)-(a.v||0));
    charts['cProd']=new Chart(document.getElementById('cProd'),{
      type:'bar',
      data:{
        labels:sorted.map(r=>r.e),
        datasets:[
          {data:sorted.map(r=>r.v), backgroundColor:sorted.map(r=>r.c+'CC'), borderRadius:7, borderSkipped:false, label:cultFilter},
          {type:'line',data:Array(emps.length).fill(avg), borderColor:CC[0], borderDash:[5,3], borderWidth:2, pointRadius:0, label:'Média'}
        ]
      },
      options:{...cd,indexAxis:'y',
        plugins:{...cd.plugins,legend:{display:false}},
        scales:{
          x:{...cd.scales.x,ticks:{...cd.scales.x.ticks,callback:v=>v+' sc/ha'}},
          y:{grid:{display:false},ticks:{font:{family:'Trebuchet MS',size:10},color:'#5A6E5A'},border:{display:false}}
        }
      }
    });
  } else {
    // Multiple cultures: grouped bars — empresa × cultura so comparison is fair
    const ds=cults.map((c,ci)=>({
      label:c,
      data:emps.map(e=>{
        const rows=filtered.filter(r=>r['Empresa']===e&&r['Cultura']===c&&r['Indicador']==='Produtividade sc/ha');
        const vs=rows.map(r=>r['Valor']).filter(v=>v!=null&&!isNaN(v));
        return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:null;
      }),
      backgroundColor:CC[ci%CC.length]+'CC', borderRadius:5, borderSkipped:false
    }));
    document.getElementById('bdProd').textContent=cults.join(' · ')+' separados';
    charts['cProd']=new Chart(document.getElementById('cProd'),{
      type:'bar',
      data:{labels:emps, datasets:ds},
      options:{...cd,
        plugins:{...cd.plugins,legend:{display:true,position:'top',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,pointStyle:'circle',padding:8}}},
        scales:{
          x:{...cd.scales.x},
          y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>v+' sc/ha'}}
        }
      }
    });
  }
}

/* ══════════════════════════════════════════════════════
   EFICIÊNCIA OPERACIONAL (SCATTER)
══════════════════════════════════════════════════════ */
let scatterCultureFilter='';
function renderScatter(emps){
  destr('cScatter');
  const cultFilter=document.getElementById('fCultura').value;
  const cults=getCultures(filtered);
  const scatterCard=document.getElementById('cScatter').closest('.card');

  // Show/hide culture filter buttons
  let ctrlEl=document.getElementById('scatterCultCtrl');
  if(!cultFilter && cults.length>1){
    if(!ctrlEl){
      ctrlEl=document.createElement('div');
      ctrlEl.id='scatterCultCtrl';
      ctrlEl.style.cssText='display:flex;gap:4px;flex-wrap:wrap;margin-bottom:10px';
      document.getElementById('cScatter').parentElement.insertAdjacentElement('beforebegin',ctrlEl);
    }
    // set default if not set or not in list
    if(!scatterCultureFilter||!cults.includes(scatterCultureFilter)) scatterCultureFilter=cults[0];
    ctrlEl.innerHTML=cults.map(c=>`<button onclick="setScatterCulture('${c}',this)" style="padding:4px 10px;border-radius:6px;border:1.5px solid #DEE8DE;background:${c===scatterCultureFilter?'#0B3D2E':'#fff'};color:${c===scatterCultureFilter?'#fff':'#5A6E5A'};font-family:var(--font);font-size:11px;font-weight:700;cursor:pointer">${c}</button>`).join('');
  } else {
    if(ctrlEl) ctrlEl.remove();
    scatterCultureFilter='';
  }

  const activeCult=cultFilter||(cults.length>1?scatterCultureFilter:cults[0])||'';
  const dataToPlot=activeCult?filtered.filter(r=>r['Cultura']===activeCult):filtered;

  const ds=emps.map((e,i)=>({
    label:e,
    data:[{
      x:getVal(dataToPlot.filter(r=>r['Empresa']===e),[null],'Custo por hectare',true)||getVal(dataToPlot,''+e,'Custo por hectare',true)||0,
      y:getVal(dataToPlot.filter(r=>r['Empresa']===e),[null],'Produtividade sc/ha',true)||getVal(dataToPlot,''+e,'Produtividade sc/ha',true)||0
    }],
    backgroundColor:CC[i%CC.length]+'CC',
    pointRadius:12, pointHoverRadius:15
  }));
  // fix: use helper correctly
  const ds2=emps.map((e,i)=>{
    const d=dataToPlot.filter(r=>r['Empresa']===e);
    return {
      label:e,
      data:[{x:getVal(d,e,'Custo por hectare',true)||0, y:getVal(d,e,'Produtividade sc/ha',true)||0}],
      backgroundColor:CC[i%CC.length]+'CC',
      pointRadius:12, pointHoverRadius:15
    };
  });

  charts['cScatter']=new Chart(document.getElementById('cScatter'),{
    type:'scatter', data:{datasets:ds2},
    options:{
      maintainAspectRatio:false, animation:{duration:500},
      plugins:{
        legend:{display:true,position:'top',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,pointStyle:'circle',padding:8}},
        tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: Custo R$${ctx.parsed.x.toLocaleString('pt-BR')} | Prod. ${ctx.parsed.y.toFixed(1)} sc/ha`}}
      },
      scales:{
        x:{title:{display:true,text:`Custo/ha (R$)${activeCult?' — '+activeCult:''}`,font:{family:'Trebuchet MS',size:10}},
          ticks:{font:{family:'Trebuchet MS',size:9},callback:v=>'R$'+v.toLocaleString('pt-BR')},
          grid:{color:'#EDF3ED'},border:{display:false}},
        y:{title:{display:true,text:'Produtividade (sc/ha)',font:{family:'Trebuchet MS',size:10}},
          ticks:{font:{family:'Trebuchet MS',size:9},callback:v=>v+' sc'},
          grid:{color:'#EDF3ED'},border:{display:false}}
      }
    }
  });
}
function setScatterCulture(cult,btn){
  scatterCultureFilter=cult;
  const ctrl=document.getElementById('scatterCultCtrl');
  if(ctrl) ctrl.querySelectorAll('button').forEach(b=>{
    const isSel=b.textContent===cult;
    b.style.background=isSel?'#0B3D2E':'#fff';
    b.style.color=isSel?'#fff':'#5A6E5A';
  });
  renderScatter(getEmps(filtered));
}

/* ══════════════════════════════════════════════════════
   HEATMAP
══════════════════════════════════════════════════════ */
function renderHeatmap(emps){
  const indicators=[
    {ind:'Produtividade sc/ha', lbl:'Produtividade', higher:true},
    {ind:'Margem EBITDA',       lbl:'Margem EBITDA', higher:true},
    {ind:'Custo por hectare',   lbl:'Custo/ha',      higher:false},
    {ind:'EBITDA/ha',          lbl:'EBITDA/ha',     higher:true},
    {ind:'Resultado por hectare',lbl:'Resultado/ha', higher:true},
    {ind:'Custo por saca',      lbl:'Custo/sc',      higher:false},
    {ind:'Preço Médio Vendido', lbl:'Preço Vendido', higher:true},
  ];
  const wrap=document.getElementById('heatmapWrap');
  let html='<div style="overflow-x:auto"><table class="hmap"><thead><tr><th>Empresa</th>';
  indicators.forEach(c=>html+=`<th>${c.lbl}</th>`);
  html+='<th>Score</th></tr></thead><tbody>';

  // compute min/max per indicator
  const ranges=indicators.map(c=>{
    const vs=emps.map(e=>getVal(filtered,e,c.ind,true)).filter(v=>v!=null);
    return{mn:Math.min(...vs),mx:Math.max(...vs)};
  });

  emps.forEach(e=>{
    const {score,grade}=calcScore(filtered,e);
    html+=`<tr><td><strong>${e}</strong></td>`;
    indicators.forEach((c,i)=>{
      const v=getVal(filtered,e,c.ind,true);
      const r=ranges[i]; let norm=0.5;
      if(v!=null&&r.mx!==r.mn) norm=(v-r.mn)/(r.mx-r.mn);
      if(!c.higher) norm=1-norm;
      const bg=heatColor(norm);
      const txt=v!=null?(c.ind==='Margem EBITDA'?fmtP(v):c.ind.includes('Custo')||c.ind==='EBITDA/ha'||c.ind==='Resultado por hectare'?fmtBRL(v,0):fmtN(v,1)):'—';
      html+=`<td style="background:${bg};color:${norm>0.6?'#0A3020':norm<0.35?'#6B1515':'#333'}">${txt}</td>`;
    });
    html+=`<td><span class="score-pill ${scoreClass(grade)}">${grade} (${score})</span></td></tr>`;
  });
  html+='</tbody></table></div>';
  wrap.innerHTML=html;
}
function heatColor(norm){
  // green=good, red=bad, interpolate
  const r=Math.round(220-norm*150);
  const g=Math.round(140+norm*100);
  const b=Math.round(140-norm*100);
  return `rgba(${r},${g},${b},0.55)`;
}

/* ══════════════════════════════════════════════════════
   RADAR CUSTOS
══════════════════════════════════════════════════════ */
function renderRadar(emps){
  destr('cRadar'); buildLeg(emps,'legRadar');
  const axes=['Fertilizante/ha','Defensivos/ha','Sementes/ha','Combustivel/ha','Mão de obra/ha','Manutenção/ha'];
  const ds=emps.map((e,i)=>({
    label:e,
    data:axes.map(a=>getVal(filtered,e,a,true)||0),
    borderColor:CC[i%CC.length], backgroundColor:CC[i%CC.length]+'22',
    borderWidth:2, pointRadius:3, pointBackgroundColor:CC[i%CC.length]
  }));
  charts['cRadar']=new Chart(document.getElementById('cRadar'),{
    type:'radar', data:{labels:axes.map(a=>a.replace('/ha','')), datasets:ds},
    options:{
      maintainAspectRatio:false, animation:{duration:500},
      plugins:{legend:{display:true,position:'bottom',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,padding:10}}},
      scales:{r:{ticks:{display:false,backdropColor:'transparent'},
        grid:{color:'#DEE8DE'},pointLabels:{font:{family:'Trebuchet MS',size:10},color:'#5A6E5A'}}}
    }
  });
}

/* ══════════════════════════════════════════════════════
   CUSTO TABLE
══════════════════════════════════════════════════════ */
function renderCustoTbl(emps){
  const axes=['Fertilizante/ha','Defensivos/ha','Sementes/ha','Combustivel/ha','Mão de obra/ha','Manutenção/ha'];
  const tbl=document.getElementById('tblCusto');
  let h=`<thead><tr><th>Indicador</th>${emps.map(e=>`<th>${e}</th>`).join('')}<th>Benchmark</th></tr></thead><tbody>`;
  axes.forEach(a=>{
    const vals=emps.map(e=>getVal(filtered,e,a,true));
    const bench=vals.filter(v=>v).reduce((s,v)=>s+v,0)/vals.filter(v=>v).length;
    h+=`<tr><td style="font-weight:600;color:#0B3D2E">${a.replace('/ha','')}</td>`;
    vals.forEach(v=>{
      const diff=v&&bench?(v-bench)/Math.abs(bench)*100:null;
      const cls=diff==null?'':diff<0?'var-pos':'var-neg';
      h+=`<td><span class="tnum">${v!=null?fmtBRL(v,0):'—'}</span>${diff!=null?`<br><span class="${cls}">${diff>0?'+':''}${diff.toFixed(1)}%</span>`:''}`;
      h+=`</td>`;
    });
    h+=`<td class="tnum">${fmtBRL(bench,0)}</td></tr>`;
  });
  h+='</tbody>';
  tbl.innerHTML=h;
}

/* ══════════════════════════════════════════════════════
   WATERFALL FINANCEIRO — valores em R$/ha
══════════════════════════════════════════════════════ */
function renderWaterfall(emps){
  destr('cWaterfall');
  const empsToUse=emps.length?emps:getEmps(filtered);
  function avgHa(ind){
    const vs=empsToUse.map(e=>getVal(filtered,e,ind,true)).filter(v=>v!=null&&!isNaN(v));
    return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:0;
  }
  // Receita líquida/ha = Receita Total / Área Colhida
  const receitaTotal=empsToUse.map(e=>getVal(filtered,e,'Receita Líquida',true)).filter(v=>v).reduce((a,b)=>a+b,0)/empsToUse.length;
  const areaTotal=empsToUse.map(e=>getVal(filtered,e,'Área Colhida (ha)',true)).filter(v=>v).reduce((a,b)=>a+b,0)/empsToUse.length;
  const receitaHa = areaTotal>0 ? receitaTotal/areaTotal : avgHa('Receita Líquida');

  // Cost items per ha
  const fertHa   = -avgHa('Fertilizante/ha');
  const defHa    = -avgHa('Defensivos/ha');
  const semHa    = -avgHa('Sementes/ha');
  const combHa   = -avgHa('Combustivel/ha');
  const mdoHa    = -avgHa('Mão de obra/ha');
  const manutHa  = -avgHa('Manutenção/ha');
  const custoHaTot= avgHa('Custo por hectare');
  // "Outros custos" = total custo/ha minus known items
  const knownCosts=Math.abs(fertHa)+Math.abs(defHa)+Math.abs(semHa)+Math.abs(combHa)+Math.abs(mdoHa)+Math.abs(manutHa);
  const outrosHa = custoHaTot>knownCosts ? -(custoHaTot-knownCosts) : 0;
  const ebitdaHa  = avgHa('EBITDA/ha');
  const lucroHa   = avgHa('Resultado por hectare');

  const labels=['Receita\nLíq./ha','Fertiliz.','Defensivos','Sementes','Combustív.','M.O.','Manutenção','Outros\nCustos','EBITDA/ha','Lucro/ha'];
  const values=[receitaHa, fertHa, defHa, semHa, combHa, mdoHa, manutHa, outrosHa, ebitdaHa, lucroHa];

  // Floating waterfall bars
  let running=receitaHa;
  const floatData=values.map((v,i)=>{
    if(i===0) return [0, v];           // Receita: starts from 0
    if(i===8){                         // EBITDA: totals from 0
      running=0; return [0, ebitdaHa];
    }
    if(i===9) return [0, lucroHa];     // Lucro: from 0
    const base=running;
    running+=v;
    return v>=0?[base,running]:[running,base];
  });
  const colors=values.map((v,i)=>{
    if(i===0) return '#0B3D2E';
    if(i===8) return '#1F7A4C';
    if(i===9) return '#0B3D2E';
    return v>=0?'#2FA85F':'#C0392B';
  });

  charts['cWaterfall']=new Chart(document.getElementById('cWaterfall'),{
    type:'bar',
    data:{labels,datasets:[{
      data:floatData,
      backgroundColor:colors, borderRadius:4, borderSkipped:false
    }]},
    options:{...cd,
      plugins:{...cd.plugins,
        tooltip:{callbacks:{label:ctx=>{
          const v=values[ctx.dataIndex];
          return ` ${labels[ctx.dataIndex].replace('\n',' ')}: R$${Math.abs(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}/ha`;
        }}}
      },
      scales:{...cd.scales,
        x:{...cd.scales.x,ticks:{...cd.scales.x.ticks,maxRotation:30,font:{family:'Trebuchet MS',size:9}}},
        y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v.toLocaleString('pt-BR',{maximumFractionDigits:0})}}
      }
    }
  });
}

/* ══════════════════════════════════════════════════════
   CUSTO SACA
══════════════════════════════════════════════════════ */
function renderCustoSaca(emps){
  destr('cCustoSaca');
  const cults=getCultures(filtered);
  const ds=emps.map((e,i)=>({
    label:e, data:cults.map(c=>{
      const rows=filtered.filter(r=>r['Empresa']===e&&r['Cultura']===c&&r['Indicador']==='Custo por saca');
      const vs=rows.map(r=>r['Valor']).filter(v=>v!=null&&!isNaN(v));
      return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:null;
    }),
    backgroundColor:CC[i%CC.length]+'CC', borderRadius:5
  }));
  charts['cCustoSaca']=new Chart(document.getElementById('cCustoSaca'),{
    type:'bar', data:{labels:cults,datasets:ds},
    options:{...cd, plugins:{...cd.plugins,legend:{display:true,position:'top',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,padding:8}}},
      scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v}}}}
  });
}

/* ══════════════════════════════════════════════════════
   EBITDA/HA
══════════════════════════════════════════════════════ */
function renderEbitdaHa(emps){
  destr('cEbitdaHa');
  const vals=emps.map(e=>getVal(filtered,e,'EBITDA/ha',true));
  charts['cEbitdaHa']=new Chart(document.getElementById('cEbitdaHa'),{
    type:'bar', data:{labels:emps,datasets:[{
      data:vals, backgroundColor:vals.map(v=>v&&v>0?CC[1]+'CC':'#C0392B88'), borderRadius:7, borderSkipped:false
    }]},
    options:{...cd, scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v.toLocaleString('pt-BR')}}}}
  });
}

/* ══════════════════════════════════════════════════════
   ÁREA
══════════════════════════════════════════════════════ */
function renderArea(emps){
  destr('cArea');
  const cults=getCultures(filtered);
  const ds=emps.map((e,i)=>({
    label:e, data:cults.map(c=>filtered.filter(r=>r['Empresa']===e&&r['Cultura']===c&&r['Indicador']==='Área Colhida (ha)').reduce((a,r)=>a+(r['Valor']||0),0)),
    backgroundColor:CC[i%CC.length]+'CC', borderRadius:5
  }));
  charts['cArea']=new Chart(document.getElementById('cArea'),{
    type:'bar', data:{labels:cults,datasets:ds},
    options:{...cd, plugins:{...cd.plugins,legend:{display:true,position:'top',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,padding:8}}},
      scales:{...cd.scales,x:{...cd.scales.x,stacked:true},y:{...cd.scales.y,stacked:true,ticks:{...cd.scales.y.ticks,callback:v=>v.toLocaleString('pt-BR')+' ha'}}}}
  });
}

/* ══════════════════════════════════════════════════════
   RESULTADO/HA
══════════════════════════════════════════════════════ */
function renderResultado(emps){
  destr('cResultado');
  const vals=emps.map(e=>getVal(filtered,e,'Resultado por hectare',true));
  charts['cResultado']=new Chart(document.getElementById('cResultado'),{
    type:'bar', data:{labels:emps,datasets:[{
      data:vals, backgroundColor:vals.map(v=>v&&v>0?CC[1]+'CC':'#C0392B88'), borderRadius:7
    }]},
    options:{...cd, scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>'R$'+v.toLocaleString('pt-BR')}}}}
  });
}

/* ══════════════════════════════════════════════════════
   EVOLUÇÃO
══════════════════════════════════════════════════════ */
function renderEvolucao(){
  destr('cEvolucao');
  const emps=getEmps(raw);
  const safras=[...new Set(raw.map(r=>r['Safra']).filter(Boolean))].sort();
  const cult=document.getElementById('fCultura').value||'SOJA';
  const ds=emps.map((e,i)=>({
    label:e,
    data:safras.map(s=>{
      const rows=raw.filter(r=>r['Empresa']===e&&r['Safra']===s&&r['Indicador']==='Produtividade sc/ha'&&(cult?r['Cultura']===cult:true));
      const vs=rows.map(r=>r['Valor']).filter(v=>v!=null&&!isNaN(v));
      return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:null;
    }),
    borderColor:CC[i%CC.length], backgroundColor:CC[i%CC.length]+'22',
    borderWidth:2.5, fill:false, tension:.4,
    pointBackgroundColor:CC[i%CC.length], pointRadius:5, pointHoverRadius:7
  }));
  charts['cEvolucao']=new Chart(document.getElementById('cEvolucao'),{
    type:'line', data:{labels:safras,datasets:ds},
    options:{...cd, plugins:{...cd.plugins,legend:{display:true,position:'top',labels:{font:{family:'Trebuchet MS',size:10},usePointStyle:true,pointStyle:'circle',padding:12}}},
      scales:{...cd.scales,y:{...cd.scales.y,ticks:{...cd.scales.y.ticks,callback:v=>v.toFixed(1)+' sc/ha'}}}}
  });
}

/* ══════════════════════════════════════════════════════
   RANKINGS
══════════════════════════════════════════════════════ */
function renderRankings(emps){
  const cult=currentRankCulture;
  let dataToUse=filtered;
  if(cult) dataToUse=filtered.filter(r=>r['Cultura']===cult);
  buildRankTbl('rkProd',emps,'Produtividade sc/ha',true,'sc/ha',v=>v.toFixed(1),dataToUse);
  buildRankTbl('rkCusto',emps,'Custo por hectare',false,'R$/ha',v=>'R$'+v.toLocaleString('pt-BR',{maximumFractionDigits:0}),dataToUse);
  buildRankTbl('rkMargem',emps,'Margem EBITDA',true,'%',v=>(v*100).toFixed(1)+'%',dataToUse);
}
function buildRankTbl(id,emps,ind,higherBetter,unit,fmtFn,data){
  const rows=emps.map(e=>({e,v:getVal(data,e,ind,true)})).filter(r=>r.v!=null);
  rows.sort((a,b)=>higherBetter?b.v-a.v:a.v-b.v);
  const mx=Math.max(...rows.map(r=>r.v));
  const tbl=document.getElementById(id);
  tbl.innerHTML=`<thead><tr><th>#</th><th>Empresa</th><th>${unit}</th><th style="min-width:90px">vs Grupo</th></tr></thead>
  <tbody>${rows.map((r,i)=>`<tr>
    <td><div class="rk ${['rk1','rk2','rk3','rkx'][Math.min(i,3)]}">${i+1}</div></td>
    <td><strong>${r.e}</strong></td>
    <td class="tnum">${fmtFn(r.v)}</td>
    <td><div style="display:flex;align-items:center;gap:6px"><div class="bar-bg"><div class="bar-fill" style="width:${(r.v/mx*100).toFixed(0)}%"></div></div></div></td>
  </tr>`).join('')}</tbody>`;
}

/* ══════════════════════════════════════════════════════
   SCORE CHART
══════════════════════════════════════════════════════ */
function renderScoreChart(emps){
  destr('cScore');
  const scores=emps.map(e=>calcScore(filtered,e));
  const sc=document.getElementById('scoreCards');
  sc.innerHTML=emps.map((e,i)=>{
    const {score,grade}=scores[i];
    return `<span class="score-pill ${scoreClass(grade)}">${e}: ${grade} (${score})</span>`;
  }).join('');
  charts['cScore']=new Chart(document.getElementById('cScore'),{
    type:'bar',
    data:{labels:emps, datasets:[{
      data:scores.map(s=>s.score),
      backgroundColor:scores.map(s=>({A:CC[0],B:CC[1],C:'#C99A2E',D:'#E07B2A',F:'#C0392B'}[s.grade]||CC[3])+'CC'),
      borderRadius:8, borderSkipped:false
    }]},
    options:{...cd,
      plugins:{...cd.plugins,tooltip:{callbacks:{label:ctx=>{const {score,grade}=scores[ctx.dataIndex];return` Score ${score}/100 — Grau ${grade}`;}}}},
      scales:{...cd.scales,y:{...cd.scales.y,min:0,max:100,ticks:{...cd.scales.y.ticks,callback:v=>v+' pts'}}}
    }
  });
}

/* ══════════════════════════════════════════════════════
   TABELA EXECUTIVA (com Score)
══════════════════════════════════════════════════════ */
function renderExecTable(emps){
  const safra=document.getElementById('fSafra').value;
  const allRows=[];
  const empAll=getEmps(raw);
  const safras=[...new Set(raw.map(r=>r['Safra']).filter(Boolean))].sort().reverse();
  const cults=[...new Set(raw.map(r=>r['Cultura']).filter(Boolean))].sort();
  (safra?[safra]:safras).forEach(s=>{
    cults.forEach(c=>{
      empAll.forEach(e=>{
        const d=raw.filter(r=>r['Empresa']===e&&r['Safra']===s&&r['Cultura']===c);
        if(!d.length) return;
        const g=(ind,avg=true)=>getVal(d,e,ind,avg);
        const {score,grade}=calcScore(d,e);
        const rec=g('Receita Líquida'); const ebi=g('EBITDA Agro');
        const marg=g('Margem EBITDA',true); const prod=g('Produtividade sc/ha',true);
        const custo=g('Custo por hectare',true); const ebitdaha=g('EBITDA/ha',true);
        if(rec==null&&prod==null) return;
        allRows.push({e,s,c,rec,ebi,marg,prod,custo,ebitdaha,score,grade});
      });
    });
  });
  const tbl=document.getElementById('tblExec');
  tbl.innerHTML=`<thead><tr>
    <th>Empresa</th><th>Safra</th><th>Cultura</th>
    <th>Receita Líq.</th><th>EBITDA</th><th>Marg.EBITDA</th>
    <th>Produt.sc/ha</th><th>Custo/ha</th><th>EBITDA/ha</th>
    <th>Score</th>
  </tr></thead><tbody>${allRows.map(r=>`<tr>
    <td><strong>${r.e}</strong></td>
    <td>${r.s}</td><td>${r.c}</td>
    <td class="tnum">${fmtK(r.rec)}</td>
    <td class="tnum">${fmtK(r.ebi)}</td>
    <td class="tnum">${r.marg!=null?fmtP(r.marg):'—'}</td>
    <td class="tnum">${r.prod!=null?fmtN(r.prod,1)+' sc/ha':'—'}</td>
    <td class="tnum">${r.custo!=null?fmtBRL(r.custo,0):'—'}</td>
    <td class="tnum">${r.ebitdaha!=null?fmtBRL(r.ebitdaha,0):'—'}</td>
    <td><span class="score-pill ${scoreClass(r.grade)}">${r.grade} (${r.score})</span></td>
  </tr>`).join('')}</tbody>`;
}

/* ══════════════════════════════════════════════════════
   DESTAQUES EXECUTIVOS
══════════════════════════════════════════════════════ */
function renderExecHighlights(emps){
  const g=(ind,avg=true)=>getVal(filtered,null,ind,avg);
  const best=(ind,higher)=>{
    const sorted=emps.map(e=>({e,v:getVal(filtered,e,ind,true)})).filter(r=>r.v!=null);
    sorted.sort((a,b)=>higher?b.v-a.v:a.v-b.v);
    return sorted[0]||{e:'—',v:null};
  };
    const destaques = [
    {
      lbl: 'Maior Receita/ha',
      ico: '💰',
      ...(function() {
        let bestEmp = '—', maxVal = -Infinity;
        emps.forEach(e => {
          const rec = getVal(filtered, e, 'Receita Líquida', false);
          const area = getVal(filtered, e, 'Área Colhida (ha)', false);
          const val = (area && area > 0) ? (rec / area) : 0;
          if (val > maxVal) { maxVal = val; bestEmp = e; }
        });
        return { e: bestEmp, v: maxVal === -Infinity ? null : maxVal };
      })(),
      fmtFn: v => fmtBRL(v, 0) + '/ha'
    },
    {lbl:'Maior Produtividade', ico:'🌾', ...best('Produtividade sc/ha',true), fmtFn:v=>fmtN(v)+' sc/ha'},
    {lbl:'Menor Custo/ha',      ico:'💡', ...best('Custo por hectare',false), fmtFn:v=>fmtBRL(v,0)+'/ha'},
    {lbl:'Melhor Margem EBITDA',ico:'📊', ...best('Margem EBITDA',true),      fmtFn:fmtP},
  ];
  document.getElementById('execHighlights').innerHTML=destaques.map(d=>`
    <div class="exec-item">
      <div>
        <div class="ei-lbl">${d.lbl}</div>
        <div class="ei-val">${d.v!=null?d.fmtFn(d.v):'—'}</div>
        <div class="ei-name">${d.e}</div>
      </div>
      <div class="ei-ico">${d.ico}</div>
    </div>`).join('');
}

/* ══════════════════════════════════════════════════════
   TABELA COMPLETA INDICADORES
══════════════════════════════════════════════════════ */
function renderIndicadoresTable(emps){
  const indicators=[
    'Receita Líquida','EBITDA Agro','Lucro Operacional','Margem EBITDA',
    'Resultado por hectare','Área Colhida (ha)','Produção Total','Produtividade sc/ha',
    'Custo por hectare','Custo por saca','Fertilizante/ha','Defensivos/ha','Sementes/ha',
    'Combustivel/ha','Mão de obra/ha','Manutenção/ha','EBITDA/ha','Lucro/sc','Preço Médio Vendido'
  ];
  const tbl=document.getElementById('tblIndicadores');
  let h=`<thead><tr><th>Indicador</th>${emps.map(e=>`<th>${e}</th>`).join('')}<th>Benchmark</th></tr></thead><tbody>`;
  indicators.forEach(ind=>{
    const vals=emps.map(e=>getVal(filtered,e,ind,true));
    const bVals=vals.filter(v=>v!=null);
    const bench=bVals.length?bVals.reduce((a,b)=>a+b,0)/bVals.length:null;
    const isPct=ind.includes('Margem');
    // Always 2 decimal places for all values
    const fFn=v=>{
      if(v==null) return'—';
      if(isPct) return fmtP(v,2);
      if(Math.abs(v)>=1e6) return'R$ '+(v/1e6).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'M';
      return v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
    };
    h+=`<tr><td style="font-weight:600;color:#0B3D2E;white-space:nowrap">${ind}</td>`;
    vals.forEach(v=>{
      const diff=v!=null&&bench?((v-bench)/Math.abs(bench))*100:null;
      const lB=['Custo por hectare','Custo por saca','Fertilizante/ha','Defensivos/ha','Sementes/ha','Combustivel/ha','Mão de obra/ha','Manutenção/ha'].includes(ind);
      const pos=diff==null?false:lB?diff<0:diff>0;
      const neg=diff==null?false:lB?diff>0:diff<0;
      h+=`<td><span class="tnum" style="font-size:12px">${fFn(v)}</span>`;
      if(diff!=null) h+=`<br><span class="${pos?'var-pos':neg?'var-neg':'var-neu'}">${diff>0?'+':''}${diff.toFixed(1)}%</span>`;
      h+=`</td>`;
    });
    h+=`<td class="tnum">${fFn(bench)}</td></tr>`;
  });
  h+='</tbody>';
  tbl.innerHTML=h;
}

/* ══════════════════════════════════════════════════════
   EXPORT EXCEL
══════════════════════════════════════════════════════ */
function exportExcel(){
  const tbl=document.getElementById('tblExec');
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.table_to_sheet(tbl);
  XLSX.utils.book_append_sheet(wb,ws,'Ranking Executivo');
  XLSX.writeFile(wb,'benchmarking-agro-ranking.xlsx');
}
function exportIndicadores(){
  const tbl=document.getElementById('tblIndicadores');
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.table_to_sheet(tbl);
  XLSX.utils.book_append_sheet(wb,ws,'Indicadores');
  // Add highlights sheet
  const tbl2=document.getElementById('tblExec');
  const ws2=XLSX.utils.table_to_sheet(tbl2);
  XLSX.utils.book_append_sheet(wb,ws2,'Ranking Score');
  XLSX.writeFile(wb,'benchmarking-agro-indicadores.xlsx');
}

/* ══════════════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════════════ */
function showLoad(s){document.getElementById('lov').classList.toggle('show',s);}
function setOK(fname,count){
  const c=document.getElementById('upCard'); c.classList.add('ok');
  document.getElementById('upOkTxt').textContent=`${fname} · ${count} registros`;
  document.getElementById('upOk').style.display='flex';
  document.getElementById('dataPill').style.display='flex';
  document.getElementById('pillTxt').textContent=`${count} registros`;
  document.getElementById('upIco').textContent='✅';
  document.getElementById('upH3').textContent='Dados carregados com sucesso';
  document.getElementById('upP').textContent='Clique para atualizar com um novo arquivo';
}
function switchTab(groupId,paneId,btn){
  const grp=document.getElementById(groupId);
  grp.querySelectorAll('.tab').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  // Only hide direct sibling panes listed by this tab group (data-panes attribute), not every .tp in card
  const panes=grp.dataset.panes?grp.dataset.panes.split(','):[];
  panes.forEach(id=>{const el=document.getElementById(id);if(el)el.classList.remove('on');});
  const target=document.getElementById(paneId);
  if(target) target.classList.add('on');
}

/* ══════════════════════════════════════════════════════
   DEMO DATA
══════════════════════════════════════════════════════ */
(function(){
  const emps=['EMPRESA 01','EMPRESA 02','EMPRESA 03','EMPRESA 04'];
  const safras=['2022/2023','2023/2024','2024/2025'];
  const culturas=['SOJA','MILHO'];
  const inds=[
    {n:'Receita Líquida',cat:'Executivo',avg:28e6,std:10e6},
    {n:'EBITDA Agro',cat:'Executivo',avg:7e6,std:2.5e6},
    {n:'Lucro Operacional',cat:'Executivo',avg:4.5e6,std:2e6},
    {n:'Margem EBITDA',cat:'Executivo',avg:0.24,std:0.08},
    {n:'Resultado por hectare',cat:'Executivo',avg:900,std:400},
    {n:'Área Colhida (ha)',cat:'Agrícola',avg:3200,std:1400},
    {n:'Produção Total',cat:'Agrícola',avg:210000,std:90000},
    {n:'Produtividade sc/ha',cat:'Agrícola',avg:68,std:13},
    {n:'Custo por hectare',cat:'Custos',avg:4900,std:750},
    {n:'Custo por saca',cat:'Custos',avg:48,std:14},
    {n:'Fertilizante/ha',cat:'Custos',avg:880,std:190},
    {n:'Defensivos/ha',cat:'Custos',avg:790,std:180},
    {n:'Sementes/ha',cat:'Custos',avg:490,std:90},
    {n:'Combustivel/ha',cat:'Custos',avg:270,std:55},
    {n:'Mão de obra/ha',cat:'Custos',avg:390,std:90},
    {n:'Manutenção/ha',cat:'Custos',avg:340,std:75},
    {n:'EBITDA/ha',cat:'Financeiro',avg:1450,std:550},
    {n:'Lucro/sc',cat:'Financeiro',avg:19,std:9},
    {n:'Preço Médio Vendido',cat:'Comercialização',avg:112,std:9},
  ];
  const demo=[];
  emps.forEach((e,ei)=>{
    safras.forEach((s,si)=>{
      culturas.forEach(c=>{
        const m=c==='MILHO'?0.32:1;
        const trend=1+(si-1)*0.04;
        inds.forEach(ind=>{
          const base=ind.avg*m*trend*(0.75+(ei+1)*0.12+(Math.random()-.5)*.14);
          demo.push({Indicador:ind.n,Categoria:ind.cat,Empresa:e,Safra:s,Cultura:c,Valor:base});
        });
      });
    });
  });
  raw=demo;
  populateFilters();
  applyFilters();
})();
/* ══════════════════════════════════════════════════════
   AUTO-LOAD DATA FROM REPOSITORY
   Tenta carregar 'dados.xlsx' automaticamente ao iniciar
══════════════════════════════════════════════════════ */
window.addEventListener('DOMContentLoaded', () => {
  const defaultFile = 'dados.xlsx';
  fetch(defaultFile)
    .then(response => {
      if (response.ok) return response.arrayBuffer();
      throw new Error('Arquivo dados.xlsx não encontrado no repositório.');
    })
    .then(data => {
      showLoad(true);
      parseWB(XLSX.read(data, { type: 'array' }), defaultFile);
      console.log('Dados carregados automaticamente de:', defaultFile);
    })
    .catch(err => {
      console.log(err.message);
      // Se falhar, apenas mantém o estado inicial (aguardando upload manual)
    });
});
