/* ═══════════════════════════════════════════════════════════
   HR Intelligence Dashboard — Application Logic
   Reads Excel via SheetJS · Charts via Chart.js
═══════════════════════════════════════════════════════════ */

/* ── Chart.js global defaults ─────────────────────────────── */
Chart.defaults.color          = '#94a3b8';
Chart.defaults.borderColor    = '#1f2d45';
Chart.defaults.font.family    = 'Inter, system-ui, sans-serif';
Chart.defaults.plugins.legend.labels.boxWidth = 10;
Chart.defaults.plugins.legend.labels.padding  = 14;

const PALETTE = {
  accent : '#06b6d4',
  accent2: '#0ea5e9',
  green  : '#22c55e',
  red    : '#ef4444',
  amber  : '#f59e0b',
  purple : '#a855f7',
  pink   : '#ec4899',
  teal   : '#14b8a6',
  indigo : '#6366f1',
  orange : '#f97316',
};

const DEPT_COLORS = [
  '#06b6d4','#22c55e','#a855f7','#f59e0b',
  '#ef4444','#6366f1','#14b8a6','#f97316',
];

/* ── Application State ─────────────────────────────────────── */
const STATE = {
  raw: null,          // parsed Excel sheets
  cfg: null,          // Business_Config map
  kpi: null,          // KPI_Summary rows
  emp: null,          // Employees rows
  rec: null,          // Recruitment rows
  att: null,          // Attendance rows
  trn: null,          // Training rows
  activeSection: 'overview',
  dept: 'all',
  liveMode: false,
  liveTimer: null,
  charts: {},
};

/* ════════════════════════════════════════════════════════════
   EMBEDDED SAMPLE DATA (used before Excel is uploaded)
════════════════════════════════════════════════════════════ */
function buildSampleState() {
  const months = [
    '2024-01','2024-02','2024-03','2024-04','2024-05','2024-06',
    '2024-07','2024-08','2024-09','2024-10','2024-11','2024-12',
  ];

  const kpi = months.map((m, i) => {
    const t = i / 11;
    return {
      Month:               m,
      Headcount:           Math.round(465 + t*15 + jitter(3)),
      TurnoverRate:        +(13.2  - t*2.1  + jitter(.3)).toFixed(1),
      TimeToHire_Days:     +(31    - t*5    + jitter(1)).toFixed(1),
      CostPerHire:         Math.round(4800  - t*400  + jitter(100)),
      EngagementScore:     +(68    + t*8    + jitter(.5)).toFixed(1),
      AbsenteeismRate:     +(3.8   - t*.6   + jitter(.1)).toFixed(1),
      RevenuePerEmployee:  Math.round(174000 + t*8000 + jitter(1000)),
      PerformanceAvg:      +(3.4   + t*.3   + jitter(.05)).toFixed(2),
      TrainingHoursPerEmp: +(24    + t*12   + jitter(1)).toFixed(1),
      TrainingROI:         Math.round(220 + t*80 + jitter(10)),
      OfferAcceptRate:     +(78    + t*8    + jitter(1)).toFixed(1),
      QualityOfHire:       +(72    + t*6    + jitter(.5)).toFixed(1),
      RetentionRate:       +(86.8  + t*2.1  + jitter(.3)).toFixed(1),
      eNPS:                Math.round(22 + t*18 + jitter(2)),
    };
  });

  const depts = ['Engineering','Sales','Operations','Marketing','HR','Finance','Customer Success','Product'];
  const deptCounts = { Engineering:120, Sales:80, Operations:60, Marketing:40, HR:30, Finance:35, 'Customer Success':55, Product:60 };

  const emp = [];
  const genders  = ['Male','Female','Non-Binary'];
  const ages     = ['Under 25','25-34','35-44','45-54','55+'];
  const levels   = ['Individual Contributor','Senior','Lead/Principal','Manager','Director','VP+'];
  const statuses = ['Active','Active','Active','Active','Active','Resigned','Terminated'];
  let id = 1001;
  for (const d of depts) {
    const n = deptCounts[d];
    for (let i = 0; i < n; i++) {
      emp.push({
        EmpID: `EMP${id++}`,
        Department: d,
        Gender:     genders[wRandom([52,45,3])],
        AgeGroup:   ages[wRandom([12,35,30,16,7])],
        Level:      levels[wRandom([30,25,18,14,9,4])],
        Status:     statuses[wRandom([70,10,5,5,3,5,2])],
        PerformanceScore: +clamp(gauss(3.5,.65),1,5).toFixed(1),
        EngagementScore:  +clamp(gauss(72,12),30,100).toFixed(1),
        Salary: Math.round(55000 + Math.random()*120000),
        Tenure_Years: +(Math.random()*14+.3).toFixed(1),
      });
    }
  }

  const rec = [];
  for (const m of months) {
    for (const d of depts) {
      const h = Math.floor(Math.random()*4);
      if (h > 0) rec.push({
        Month: m, Department: d, NewHires: h,
        TimeToHire_Days: +clamp(gauss(28,8),7,90).toFixed(1),
        CostPerHire: Math.round(clamp(gauss(4500,800),1000,12000)),
        OfferAcceptRate: +clamp(gauss(80,10),50,100).toFixed(1),
        QualityScore: +clamp(gauss(73,10),40,100).toFixed(1),
      });
    }
  }

  const att = [];
  for (let i = 0; i < months.length; i++) {
    for (const d of depts) {
      const wdays = 22, hc = deptCounts[d] + jitter(3);
      const tot = hc * wdays;
      const abs = Math.round(tot * (0.025 + Math.random()*.02));
      att.push({ Month: months[i], Department: d, Headcount: hc, AbsenceDays: abs, TotalWorkdays: wdays, AbsenteeismRate: +(abs/tot*100).toFixed(2) });
    }
  }

  const programs = ['Leadership Dev','Technical Skills','Compliance','Sales Excellence','DEI Training','Project Mgmt','Comm Skills'];
  const trn = [];
  for (const m of months) {
    for (const d of depts) {
      const p = Math.floor(Math.random()*25)+4;
      const h = +(p * (4+Math.random()*12)).toFixed(1);
      const pre  = +clamp(gauss(65,10),40,85).toFixed(1);
      const post = +clamp(pre+gauss(13,4),50,100).toFixed(1);
      trn.push({ Month: m, Department: d, Program: programs[Math.floor(Math.random()*programs.length)], Participants: p, TotalHours: h, TotalCost: Math.round(h*(50+Math.random()*100)), CompletionRate: +clamp(gauss(88,7),60,100).toFixed(1), PreScore: pre, PostScore: post });
    }
  }

  const cfg = {
    AnnualRevenue: 85000000, AverageSalary: 85000, CurrentHeadcount: 480,
    IndustryTurnoverBenchmark: 14.0, IndustryTimeToHireBenchmark: 36,
    IndustryCostPerHireBenchmark: 4700, IndustryEngagementBenchmark: 66,
    IndustryAbsenteeismBenchmark: 3.5, IndustryRevenuePerEmpBenchmark: 160000,
    TurnoverReplacementCostMultiplier: 1.5, AbsenteeismIndirectCostMultiplier: 1.5,
    TargetTurnoverRate: 10, TargetEngagementScore: 78, TargetTimeToHire: 25,
    TargetCostPerHire: 4000, TargetRevenuePerEmployee: 185000,
  };

  return { kpi, emp, rec, att, trn, cfg };
}

/* ── Math helpers ─────────────────────────────────────────── */
function jitter(n) { return (Math.random()*2-1)*n; }
function gauss(mean, sd) { let u=0,v=0; while(!u) u=Math.random(); while(!v) v=Math.random(); return mean + sd * Math.sqrt(-2*Math.log(u)) * Math.cos(2*Math.PI*v); }
function clamp(v,lo,hi) { return Math.min(hi, Math.max(lo, v)); }
function wRandom(weights) { const t=weights.reduce((a,b)=>a+b,0); let r=Math.random()*t; for(let i=0;i<weights.length;i++){r-=weights[i]; if(r<=0) return i;} return weights.length-1; }
function fmt$(n) { if(Math.abs(n)>=1e6) return '$'+(n/1e6).toFixed(1)+'M'; if(Math.abs(n)>=1e3) return '$'+(n/1e3).toFixed(0)+'K'; return '$'+Math.round(n); }
function fmtN(n, decimals=1) { return n==null?'—':typeof n==='number'?n.toFixed(decimals):n; }
function pct(n) { return (n>=0?'+':'')+n.toFixed(1)+'%'; }

/* ════════════════════════════════════════════════════════════
   EXCEL PARSING
════════════════════════════════════════════════════════════ */
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const get = name => {
          const ws = wb.Sheets[name];
          return ws ? XLSX.utils.sheet_to_json(ws) : [];
        };
        const cfg = {};
        get('Business_Config').forEach(r => { cfg[r.Metric] = +r.Value; });
        resolve({ kpi: get('KPI_Summary'), emp: get('Employees'), rec: get('Recruitment'), att: get('Attendance'), trn: get('Training'), cfg });
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

/* ════════════════════════════════════════════════════════════
   BUSINESS VALUE CALCULATIONS
════════════════════════════════════════════════════════════ */
function calcBusinessValue(cfg, kpiRows) {
  const latest = kpiRows[kpiRows.length - 1];
  const salary  = cfg.AverageSalary || 85000;
  const hc      = cfg.CurrentHeadcount || 480;
  const rev     = cfg.AnnualRevenue || 85000000;
  const rpeMult = cfg.TurnoverReplacementCostMultiplier || 1.5;
  const absMult = cfg.AbsenteeismIndirectCostMultiplier || 1.5;
  const revenuePerEmpPerDay = rev / hc / 250;

  // 1. Turnover cost
  const turnoverCost = (latest.TurnoverRate / 100) * hc * salary * rpeMult;
  const targetTurnCost = ((cfg.TargetTurnoverRate || 10) / 100) * hc * salary * rpeMult;
  const turnoverSaving = Math.max(0, turnoverCost - targetTurnCost);

  // 2. Time-to-hire productivity loss (assume 3% vacancy rate)
  const openPositions = Math.round(hc * 0.03);
  const tthLoss = latest.TimeToHire_Days * revenuePerEmpPerDay * openPositions;
  const targetTthLoss = (cfg.TargetTimeToHire || 25) * revenuePerEmpPerDay * openPositions;
  const tthSaving = Math.max(0, tthLoss - targetTthLoss);

  // 3. Engagement premium (Gallup: 21% higher profitability for highly engaged)
  const engRatio = latest.EngagementScore / (cfg.IndustryEngagementBenchmark || 66);
  const engPremium = (engRatio - 1) * 0.21 * rev;

  // 4. Absenteeism cost
  const absentDays = (latest.AbsenteeismRate / 100) * 250 * hc;
  const dailySalary = salary / 250;
  const absCost = absentDays * dailySalary * absMult;
  const targetAbsCost = ((cfg.IndustryAbsenteeismBenchmark || 3.5) / 100) * 250 * hc * dailySalary * absMult;
  const absSaving = Math.max(0, absCost - targetAbsCost);

  // 5. Revenue per employee vs benchmark
  const rpeGap = (latest.RevenuePerEmployee - (cfg.IndustryRevenuePerEmpBenchmark || 160000)) * hc;

  // 6. Training value (improvement in performance → revenue)
  const avgImprovement = 0.13; // ~13% avg score improvement
  const trainedEmp = Math.round(hc * 0.25);
  const trainingValue = avgImprovement * (rev / hc) * trainedEmp;

  return {
    turnoverCost, turnoverSaving,
    tthLoss, tthSaving,
    engPremium,
    absCost, absSaving,
    rpeGap,
    trainingValue,
    totalRevenuePremium: Math.max(0, rpeGap) + Math.max(0, engPremium),
    totalCostReduction: turnoverSaving + tthSaving + absSaving,
  };
}

/* ════════════════════════════════════════════════════════════
   KPI CARD DEFINITIONS
════════════════════════════════════════════════════════════ */
function buildKpiDefs(latest, prev, cfg, bv) {
  const trend = (cur, prv, higherGood=true) => {
    const diff = cur - prv;
    const dir = higherGood ? (diff>0?'up':'down') : (diff<0?'up':'down');
    return { dir, pct: Math.abs(diff/prv*100).toFixed(1), diff };
  };

  const vsTarget = (cur, target, higherGood=true) => {
    const ok = higherGood ? cur >= target : cur <= target;
    return { ok, label: `Target: ${target}`, color: ok ? 'green' : cur*0.9 <= target ? 'red' : 'amber' };
  };

  return {
    overview: [
      { id:'headcount',   name:'Total Headcount',    value: latest.Headcount,                  unit:'',   icon:'fa-users',         color:'accent', icolor:'accent',
        trend: trend(latest.Headcount, prev.Headcount), tgt: null,
        bvBadge: `${fmt$(cfg.AnnualRevenue)} revenue base` },
      { id:'turnover',    name:'Turnover Rate',       value: latest.TurnoverRate+'%',            unit:'%',  icon:'fa-door-open',     color:'amber',  icolor:'amber',
        trend: trend(latest.TurnoverRate, prev.TurnoverRate, false), tgt: vsTarget(latest.TurnoverRate, cfg.TargetTurnoverRate||10, false),
        bvBadge: `Cost: ${fmt$(bv.turnoverCost)}/yr` },
      { id:'engagement',  name:'Engagement Score',    value: latest.EngagementScore+'/100',      unit:'',   icon:'fa-heart-pulse',   color:'green',  icolor:'green',
        trend: trend(latest.EngagementScore, prev.EngagementScore), tgt: vsTarget(latest.EngagementScore, cfg.TargetEngagementScore||78),
        bvBadge: `Premium: ${fmt$(Math.abs(bv.engPremium))}` },
      { id:'tth',         name:'Time to Hire',        value: latest.TimeToHire_Days+'d',         unit:'d',  icon:'fa-clock',         color:'amber',  icolor:'amber',
        trend: trend(latest.TimeToHire_Days, prev.TimeToHire_Days, false), tgt: vsTarget(latest.TimeToHire_Days, cfg.TargetTimeToHire||25, false),
        bvBadge: `Loss: ${fmt$(bv.tthLoss)}/mo` },
      { id:'rpe',         name:'Revenue / Employee',  value: fmt$(latest.RevenuePerEmployee),    unit:'',   icon:'fa-chart-line',    color:'green',  icolor:'green',
        trend: trend(latest.RevenuePerEmployee, prev.RevenuePerEmployee), tgt: vsTarget(latest.RevenuePerEmployee, cfg.TargetRevenuePerEmployee||185000),
        bvBadge: `Gap vs benchmark: ${fmt$(bv.rpeGap)}` },
      { id:'absenteeism', name:'Absenteeism Rate',    value: latest.AbsenteeismRate+'%',         unit:'%',  icon:'fa-person-walking-arrow-right', color:'amber', icolor:'amber',
        trend: trend(latest.AbsenteeismRate, prev.AbsenteeismRate, false), tgt: vsTarget(latest.AbsenteeismRate, cfg.IndustryAbsenteeismBenchmark||3.5, false),
        bvBadge: `Cost: ${fmt$(bv.absCost)}/yr` },
      { id:'cph',         name:'Cost per Hire',       value: fmt$(latest.CostPerHire),           unit:'',   icon:'fa-hand-holding-dollar', color:'green', icolor:'green',
        trend: trend(latest.CostPerHire, prev.CostPerHire, false), tgt: vsTarget(latest.CostPerHire, cfg.TargetCostPerHire||4000, false),
        bvBadge: `Bench: ${fmt$(cfg.IndustryCostPerHireBenchmark||4700)}` },
      { id:'trainingROI', name:'Training ROI',        value: latest.TrainingROI+'%',             unit:'%',  icon:'fa-graduation-cap', color:'green', icolor:'green',
        trend: trend(latest.TrainingROI, prev.TrainingROI), tgt: null,
        bvBadge: `Value: ${fmt$(bv.trainingValue)}` },
    ],
  };
}

/* ════════════════════════════════════════════════════════════
   RENDER KPI CARDS
════════════════════════════════════════════════════════════ */
function renderKpiCards(containerId, defs) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = defs.map(d => {
    const trendHtml = d.trend ? `
      <span class="kpi-trend ${d.trend.dir}">
        <i class="fas fa-arrow-${d.trend.dir === 'up' ? 'up' : 'down'}-long"></i>
        ${d.trend.pct}%
      </span>` : '';
    const tgtHtml = d.tgt ? `<span class="kpi-vs-target">${d.tgt.label}</span>` : '';
    const cardColor = d.tgt ? d.tgt.color : d.color;
    return `
      <div class="kpi-card ${cardColor}">
        <div class="kpi-icon ${d.icolor}"><i class="fas ${d.icon}"></i></div>
        <div class="kpi-name">${d.name}</div>
        <div class="kpi-value">${d.value}</div>
        <div class="kpi-meta">${trendHtml}${tgtHtml}</div>
        ${d.bvBadge ? `<div class="kpi-bv-badge"><i class="fas fa-sack-dollar"></i> ${d.bvBadge}</div>` : ''}
      </div>`;
  }).join('');
}

/* ════════════════════════════════════════════════════════════
   RENDER BUSINESS VALUE PANEL
════════════════════════════════════════════════════════════ */
function renderBvPanel(containerId, title, items) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = `
    <div class="bv-panel-header"><i class="fas fa-sack-dollar"></i><h3>${title}</h3></div>
    <div class="bv-items">${items.map(it => `
      <div class="bv-item">
        <div class="bv-item-title">${it.title}</div>
        <div class="bv-item-amount">${it.amount}</div>
        <div class="bv-item-desc">${it.desc}</div>
        ${it.formula ? `<div class="bv-item-formula">${it.formula}</div>` : ''}
      </div>`).join('')}
    </div>`;
}

/* ════════════════════════════════════════════════════════════
   CHART HELPERS
════════════════════════════════════════════════════════════ */
function destroyChart(id) { if (STATE.charts[id]) { STATE.charts[id].destroy(); delete STATE.charts[id]; } }

function lineChart(id, labels, datasets, opts={}) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  STATE.charts[id] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: datasets.map(d => ({
      borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, tension: 0.35, fill: d.fill||false, ...d
    })) },
    options: {
      responsive: true, maintainAspectRatio: true,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: datasets.length > 1 }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.formattedValue}` } } },
      scales: {
        x: { grid: { color: '#1f2d45' } },
        y: { grid: { color: '#1f2d45' }, ...( opts.yLeft||{} ) },
        ...(opts.y2 ? { y2: { position: 'right', grid: { display: false }, ...opts.y2 } } : {}),
      },
      ...opts.extra,
    },
  });
}

function barChart(id, labels, datasets, opts={}) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  STATE.charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: true,
      indexAxis: opts.horizontal ? 'y' : 'x',
      plugins: { legend: { display: datasets.length > 1 } },
      scales: {
        x: { grid: { color: '#1f2d45' }, stacked: opts.stacked||false },
        y: { grid: { color: '#1f2d45' }, stacked: opts.stacked||false },
      },
      ...opts.extra,
    },
  });
}

function doughnutChart(id, labels, data, colors) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  STATE.charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#111827', borderWidth: 2, hoverOffset: 8 }] },
    options: { responsive: true, maintainAspectRatio: true, plugins: { legend: { position: 'bottom' } }, cutout: '62%' },
  });
}

function pieChart(id, labels, data, colors) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  STATE.charts[id] = new Chart(ctx, {
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#111827', borderWidth: 2 }] },
    options: { responsive: true, maintainAspectRatio: true, plugins: { legend: { position: 'bottom' } } },
  });
}

/* ════════════════════════════════════════════════════════════
   SECTION RENDERERS
════════════════════════════════════════════════════════════ */

/* ── Overview ─────────────────────────────────────────────── */
function renderOverview() {
  const { kpi, emp, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2] || kpi[0];
  const bv = calcBusinessValue(cfg, kpi);
  const defs = buildKpiDefs(latest, prev, cfg, bv);

  // BV banner
  document.getElementById('bvRevPremium').textContent = fmt$(bv.totalRevenuePremium);
  document.getElementById('bvCostReduce').textContent = fmt$(bv.totalCostReduction);
  document.getElementById('bvEngValue').textContent   = fmt$(Math.abs(bv.engPremium));

  renderKpiCards('kpiGrid', defs.overview);

  // Overview trend chart (normalised %)
  const months = kpi.map(r=>r.Month);
  const norm = (arr, min, max) => arr.map(v => ((v-min)/(max-min)*100));
  const tRates = kpi.map(r=>r.TurnoverRate);
  const tEng   = kpi.map(r=>r.EngagementScore);
  const tRpe   = kpi.map(r=>r.RevenuePerEmployee);

  lineChart('overviewChart', months, [
    { label:'Engagement Score', data: tEng, borderColor: PALETTE.green,  backgroundColor: 'rgba(34,197,94,.07)', fill: true },
    { label:'Turnover Rate %',  data: tRates, borderColor: PALETTE.red,  borderDash:[4,4] },
    { label:'Rev/Employee (K)', data: tRpe.map(v=>+(v/1000).toFixed(1)), borderColor: PALETTE.accent, yAxisID:'y2' },
  ], { y2: { ticks: { callback: v => '$'+v+'K' } } });

  // Headcount by dept
  const deptCounts = {};
  emp.forEach(e => { deptCounts[e.Department] = (deptCounts[e.Department]||0)+1; });
  const dLabels = Object.keys(deptCounts);
  barChart('headcountDeptChart', dLabels, [{
    label: 'Employees', data: dLabels.map(d=>deptCounts[d]),
    backgroundColor: DEPT_COLORS, borderRadius: 4,
  }], { horizontal: true });
}

/* ── Workforce ────────────────────────────────────────────── */
function renderWorkforce() {
  const { kpi, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2]||kpi[0];
  const bv = calcBusinessValue(cfg, kpi);

  renderKpiCards('workforceGrid', [
    { name:'Headcount',      value: latest.Headcount,          icon:'fa-users',       color:'accent', icolor:'accent', trend:{dir:'up',pct:((latest.Headcount-prev.Headcount)/prev.Headcount*100).toFixed(1)}, tgt:null, bvBadge:`Payroll: ${fmt$(latest.Headcount*(cfg.AverageSalary||85000))}` },
    { name:'Turnover Rate',  value: latest.TurnoverRate+'%',   icon:'fa-door-open',   color:'amber',  icolor:'amber',  trend:{dir:latest.TurnoverRate<prev.TurnoverRate?'up':'down', pct: Math.abs((latest.TurnoverRate-prev.TurnoverRate)/prev.TurnoverRate*100).toFixed(1)}, tgt:{label:`Target: ${cfg.TargetTurnoverRate||10}%`, color: latest.TurnoverRate<=(cfg.TargetTurnoverRate||10)?'green':'amber'}, bvBadge:`Cost: ${fmt$(bv.turnoverCost)}/yr` },
    { name:'Retention Rate', value: latest.RetentionRate+'%',  icon:'fa-shield-heart',color:'green',  icolor:'green',  trend:{dir:'up',pct:.5}, tgt:{label:'Target: 90%', color:latest.RetentionRate>=90?'green':'amber'}, bvBadge:null },
    { name:'eNPS',           value: latest.eNPS,               icon:'fa-thumbs-up',   color:'green',  icolor:'green',  trend:{dir:latest.eNPS>prev.eNPS?'up':'down',pct:1.2}, tgt:{label:'Target: 40', color:latest.eNPS>=40?'green':'amber'}, bvBadge:null },
  ]);

  renderBvPanel('workforceBvPanel', 'Business Value: Workforce & Turnover', [
    { title:'Annual Turnover Cost', amount: fmt$(bv.turnoverCost),
      desc: `Replacing ${((latest.TurnoverRate/100)*latest.Headcount).toFixed(0)} employees costs ${(cfg.TurnoverReplacementCostMultiplier||1.5)}× their annual salary in recruitment, onboarding & lost productivity.`,
      formula: `${latest.TurnoverRate}% × ${latest.Headcount} employees × $${((cfg.AverageSalary||85000)/1000).toFixed(0)}K salary × ${cfg.TurnoverReplacementCostMultiplier||1.5}× = ${fmt$(bv.turnoverCost)}` },
    { title:'Potential Saving — Hit Turnover Target', amount: fmt$(bv.turnoverSaving),
      desc: `Reducing turnover from ${latest.TurnoverRate}% to the ${cfg.TargetTurnoverRate||10}% target unlocks ${fmt$(bv.turnoverSaving)} in annual savings.`,
      formula: `(Current ${latest.TurnoverRate}% − Target ${cfg.TargetTurnoverRate||10}%) × ${latest.Headcount} × $${((cfg.AverageSalary||85000)/1000).toFixed(0)}K × ${cfg.TurnoverReplacementCostMultiplier||1.5}× = ${fmt$(bv.turnoverSaving)}` },
    { title:'Retention vs Productivity', amount: latest.RetentionRate+'%',
      desc: 'Each percentage point of retention saved avoids re-hiring, knowledge transfer loss and team disruption costs estimated at 6–9 months\' salary per role.',
      formula: null },
  ]);

  const months = kpi.map(r=>r.Month);
  const bench  = kpi.map(()=>cfg.IndustryTurnoverBenchmark||14);
  const target = kpi.map(()=>cfg.TargetTurnoverRate||10);

  lineChart('turnoverChart', months, [
    { label:'Turnover Rate %',      data: kpi.map(r=>r.TurnoverRate),  borderColor: PALETTE.red,   backgroundColor: 'rgba(239,68,68,.08)', fill: true },
    { label:'Industry Benchmark',   data: bench,  borderColor: PALETTE.amber, borderDash:[6,3] },
    { label:'Target',               data: target, borderColor: PALETTE.green, borderDash:[4,4] },
  ]);

  lineChart('headcountTrendChart', months, [
    { label:'Headcount', data: kpi.map(r=>r.Headcount), borderColor: PALETTE.accent, backgroundColor: 'rgba(6,182,212,.1)', fill: true },
  ]);
}

/* ── Recruitment ──────────────────────────────────────────── */
function renderRecruitment() {
  const { kpi, rec, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2]||kpi[0];
  const bv = calcBusinessValue(cfg, kpi);

  renderKpiCards('recruitmentGrid', [
    { name:'Time to Hire',        value: latest.TimeToHire_Days+'d',  icon:'fa-clock',              color:'amber', icolor:'amber',  trend:{dir:latest.TimeToHire_Days<prev.TimeToHire_Days?'up':'down', pct:Math.abs((latest.TimeToHire_Days-prev.TimeToHire_Days)/prev.TimeToHire_Days*100).toFixed(1)}, tgt:{label:`Target: ${cfg.TargetTimeToHire||25}d`, color:latest.TimeToHire_Days<=(cfg.TargetTimeToHire||25)?'green':'amber'}, bvBadge:`Monthly loss: ${fmt$(bv.tthLoss)}` },
    { name:'Cost per Hire',       value: fmt$(latest.CostPerHire),    icon:'fa-hand-holding-dollar', color:'green', icolor:'green',  trend:{dir:latest.CostPerHire<prev.CostPerHire?'up':'down', pct:Math.abs((latest.CostPerHire-prev.CostPerHire)/prev.CostPerHire*100).toFixed(1)}, tgt:{label:`Target: ${fmt$(cfg.TargetCostPerHire||4000)}`, color:latest.CostPerHire<=(cfg.TargetCostPerHire||4000)?'green':'amber'}, bvBadge:`Bench: ${fmt$(cfg.IndustryCostPerHireBenchmark||4700)}` },
    { name:'Offer Accept Rate',   value: latest.OfferAcceptRate+'%',  icon:'fa-file-signature',     color:'green', icolor:'green',  trend:{dir:'up',pct:1.5}, tgt:{label:'Target: 85%', color:latest.OfferAcceptRate>=85?'green':'amber'}, bvBadge:null },
    { name:'Quality of Hire',     value: latest.QualityOfHire+'/100', icon:'fa-star',               color:'green', icolor:'green',  trend:{dir:'up',pct:0.8}, tgt:{label:'Target: 80', color:latest.QualityOfHire>=80?'green':'amber'}, bvBadge:null },
  ]);

  const rev = cfg.AnnualRevenue||85000000;
  const hc  = cfg.CurrentHeadcount||480;
  const openPos = Math.round(hc*0.03);
  const rpeDay  = rev/hc/250;

  renderBvPanel('recruitmentBvPanel', 'Business Value: Recruitment Efficiency', [
    { title:'Monthly Productivity Loss — Open Roles', amount: fmt$(bv.tthLoss),
      desc: `With ~${openPos} open positions and ${latest.TimeToHire_Days}-day avg fill time, each unfilled seat loses ${fmt$(rpeDay)}/day in revenue capacity.`,
      formula: `${latest.TimeToHire_Days} days × ${fmt$(rpeDay)}/day × ${openPos} open positions = ${fmt$(bv.tthLoss)}` },
    { title:'Saving by Hitting Time-to-Hire Target', amount: fmt$(bv.tthSaving),
      desc: `Reducing from ${latest.TimeToHire_Days} to ${cfg.TargetTimeToHire||25} days saves ${fmt$(bv.tthSaving)} per month in foregone productivity.`,
      formula: `(${latest.TimeToHire_Days} − ${cfg.TargetTimeToHire||25}) days × ${fmt$(rpeDay)}/day × ${openPos} roles = ${fmt$(bv.tthSaving)}` },
    { title:'Annual Hiring Budget', amount: fmt$(latest.CostPerHire * Math.round(hc*.15)),
      desc: `Assuming ~${Math.round(hc*.15)} new hires/year at ${fmt$(latest.CostPerHire)} each. Industry avg is ${fmt$(cfg.IndustryCostPerHireBenchmark||4700)}.`,
      formula: `${Math.round(hc*.15)} hires × ${fmt$(latest.CostPerHire)} = ${fmt$(latest.CostPerHire*Math.round(hc*.15))}` },
  ]);

  const months = kpi.map(r=>r.Month);
  lineChart('tthChart', months, [
    { label:'Time to Hire (days)', data: kpi.map(r=>r.TimeToHire_Days), borderColor: PALETTE.amber, backgroundColor: 'rgba(245,158,11,.08)', fill:true },
    { label:'Target', data: months.map(()=>cfg.TargetTimeToHire||25), borderColor: PALETTE.green, borderDash:[4,4] },
    { label:'Industry Benchmark', data: months.map(()=>cfg.IndustryTimeToHireBenchmark||36), borderColor: PALETTE.red, borderDash:[6,3] },
  ]);

  lineChart('cphChart', months, [
    { label:'Cost per Hire ($)', data: kpi.map(r=>r.CostPerHire), borderColor: PALETTE.accent, backgroundColor: 'rgba(6,182,212,.08)', fill:true },
    { label:'Target', data: months.map(()=>cfg.TargetCostPerHire||4000), borderColor: PALETTE.green, borderDash:[4,4] },
  ]);

  // Hires by dept
  const depts = [...new Set(rec.map(r=>r.Department))];
  const hiresPerDept = depts.map(d => rec.filter(r=>r.Department===d).reduce((s,r)=>s+(+r.NewHires||0),0));
  barChart('hiresDeptChart', depts, [{ label:'Total Hires', data: hiresPerDept, backgroundColor: DEPT_COLORS, borderRadius: 4 }], { horizontal: true });
}

/* ── Performance ──────────────────────────────────────────── */
function renderPerformance() {
  const { kpi, emp, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2]||kpi[0];
  const bv = calcBusinessValue(cfg, kpi);

  renderKpiCards('performanceGrid', [
    { name:'Avg Performance Score', value: latest.PerformanceAvg+'/5', icon:'fa-star', color:'green', icolor:'green', trend:{dir:'up',pct:((latest.PerformanceAvg-prev.PerformanceAvg)/prev.PerformanceAvg*100).toFixed(1)}, tgt:{label:'Target: 4.0/5', color:latest.PerformanceAvg>=4?'green':'amber'}, bvBadge:'Gallup: +17% productivity' },
    { name:'Revenue/Employee',      value: fmt$(latest.RevenuePerEmployee), icon:'fa-chart-line', color:'green', icolor:'green', trend:{dir:'up',pct:((latest.RevenuePerEmployee-prev.RevenuePerEmployee)/prev.RevenuePerEmployee*100).toFixed(1)}, tgt:{label:`Target: ${fmt$(cfg.TargetRevenuePerEmployee||185000)}`, color:latest.RevenuePerEmployee>=(cfg.TargetRevenuePerEmployee||185000)?'green':'amber'}, bvBadge:`Gap: ${fmt$(bv.rpeGap)}` },
    { name:'% High Performers',     value: (emp.filter(e=>e.PerformanceScore>=4).length/emp.length*100).toFixed(0)+'%', icon:'fa-trophy', color:'accent', icolor:'accent', trend:{dir:'up',pct:.8}, tgt:{label:'Target: 30%', color:(emp.filter(e=>e.PerformanceScore>=4).length/emp.length*100)>=30?'green':'amber'}, bvBadge:null },
    { name:'Quality of Hire',       value: latest.QualityOfHire+'/100', icon:'fa-medal', color:'green', icolor:'green', trend:{dir:'up',pct:.5}, tgt:{label:'Target: 80', color:latest.QualityOfHire>=80?'green':'amber'}, bvBadge:null },
  ]);

  const hc = cfg.CurrentHeadcount||480;
  const rev = cfg.AnnualRevenue||85000000;

  renderBvPanel('performanceBvPanel', 'Business Value: Performance & Productivity', [
    { title:'Revenue per Employee vs Benchmark', amount: fmt$(bv.rpeGap),
      desc: `At ${fmt$(latest.RevenuePerEmployee)}/employee vs the ${fmt$(cfg.IndustryRevenuePerEmpBenchmark||160000)} industry benchmark, the total revenue differential across ${hc} employees.`,
      formula: `(${fmt$(latest.RevenuePerEmployee)} − ${fmt$(cfg.IndustryRevenuePerEmpBenchmark||160000)}) × ${hc} employees = ${fmt$(bv.rpeGap)}` },
    { title:'High-Performer Revenue Premium', amount: fmt$(emp.filter(e=>e.PerformanceScore>=4).length * (rev/hc) * 0.17),
      desc: `Gallup research: high performers generate 17% more than average. Your ${emp.filter(e=>e.PerformanceScore>=4).length} high performers create a productivity premium.`,
      formula: `${emp.filter(e=>e.PerformanceScore>=4).length} high performers × $${(rev/hc/1000).toFixed(0)}K avg revenue × 17% premium` },
    { title:'Performance Improvement Target Value', amount: fmt$(0.1 * rev),
      desc: 'A 10% improvement in average performance score could generate an estimated 10% uplift in overall productivity and revenue contribution.',
      formula: `10% improvement × ${fmt$(rev)} annual revenue = ${fmt$(0.1*rev)}` },
  ]);

  // Performance distribution
  const bands = [
    { label:'Exceptional (4.5–5)', min:4.5, max:5 },
    { label:'Exceeds (4–4.5)',     min:4,   max:4.5 },
    { label:'Meets (3–4)',         min:3,   max:4 },
    { label:'Below (2–3)',         min:2,   max:3 },
    { label:'PIP (<2)',            min:0,   max:2 },
  ];
  const distData = bands.map(b => emp.filter(e=>e.PerformanceScore>=b.min&&e.PerformanceScore<b.max).length);
  doughnutChart('perfDistChart', bands.map(b=>b.label), distData, [PALETTE.green, PALETTE.teal, PALETTE.accent, PALETTE.amber, PALETTE.red]);

  const months = kpi.map(r=>r.Month);
  lineChart('rpeChart', months, [
    { label:'Revenue/Employee', data: kpi.map(r=>r.RevenuePerEmployee), borderColor: PALETTE.accent, backgroundColor: 'rgba(6,182,212,.07)', fill:true },
    { label:'Industry Benchmark', data: months.map(()=>cfg.IndustryRevenuePerEmpBenchmark||160000), borderColor: PALETTE.amber, borderDash:[6,3] },
    { label:'Target', data: months.map(()=>cfg.TargetRevenuePerEmployee||185000), borderColor: PALETTE.green, borderDash:[4,4] },
  ], { yLeft: { ticks: { callback: v => fmt$(v) } } });

  lineChart('perfTrendChart', months, [
    { label:'Avg Performance Score', data: kpi.map(r=>r.PerformanceAvg), borderColor: PALETTE.green, backgroundColor: 'rgba(34,197,94,.07)', fill:true },
    { label:'Target (4.0)', data: months.map(()=>4.0), borderColor: PALETTE.accent, borderDash:[4,4] },
  ]);

  // By dept
  const depts = [...new Set(emp.map(e=>e.Department))];
  const avgScores = depts.map(d => {
    const scores = emp.filter(e=>e.Department===d).map(e=>+e.PerformanceScore);
    return +(scores.reduce((a,b)=>a+b,0)/scores.length).toFixed(2);
  });
  barChart('perfDeptChart', depts, [{ label:'Avg Score', data: avgScores, backgroundColor: DEPT_COLORS, borderRadius: 4 }], { horizontal: true });
}

/* ── Engagement ───────────────────────────────────────────── */
function renderEngagement() {
  const { kpi, att, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2]||kpi[0];
  const bv = calcBusinessValue(cfg, kpi);

  renderKpiCards('engagementGrid', [
    { name:'Engagement Score', value: latest.EngagementScore+'/100', icon:'fa-heart-pulse', color:'green', icolor:'green', trend:{dir:'up',pct:((latest.EngagementScore-prev.EngagementScore)/prev.EngagementScore*100).toFixed(1)}, tgt:{label:`Target: ${cfg.TargetEngagementScore||78}`, color:latest.EngagementScore>=(cfg.TargetEngagementScore||78)?'green':'amber'}, bvBadge:`Value: ${fmt$(Math.abs(bv.engPremium))}` },
    { name:'eNPS',             value: latest.eNPS,                   icon:'fa-thumbs-up',  color:'accent', icolor:'accent', trend:{dir:'up',pct:2}, tgt:{label:'Target: 40', color:latest.eNPS>=40?'green':'amber'}, bvBadge:null },
    { name:'Absenteeism Rate', value: latest.AbsenteeismRate+'%',    icon:'fa-person-walking-arrow-right', color:'amber', icolor:'amber', trend:{dir:latest.AbsenteeismRate<prev.AbsenteeismRate?'up':'down',pct:Math.abs((latest.AbsenteeismRate-prev.AbsenteeismRate)/prev.AbsenteeismRate*100).toFixed(1)}, tgt:{label:`Benchmark: ${cfg.IndustryAbsenteeismBenchmark||3.5}%`, color:latest.AbsenteeismRate<=(cfg.IndustryAbsenteeismBenchmark||3.5)?'green':'amber'}, bvBadge:`Cost: ${fmt$(bv.absCost)}/yr` },
    { name:'Wellness Index',   value: (100-latest.AbsenteeismRate*8).toFixed(0)+'/100', icon:'fa-shield-heart', color:'green', icolor:'green', trend:{dir:'up',pct:.5}, tgt:{label:'Target: 80', color:(100-latest.AbsenteeismRate*8)>=80?'green':'amber'}, bvBadge:null },
  ]);

  renderBvPanel('engagementBvPanel', 'Business Value: Engagement & Absenteeism', [
    { title:'Engagement Revenue Premium', amount: fmt$(Math.abs(bv.engPremium)),
      desc: `Your engagement score of ${latest.EngagementScore} is ${(latest.EngagementScore/(cfg.IndustryEngagementBenchmark||66)*100-100).toFixed(0)}% above the ${cfg.IndustryEngagementBenchmark||66}-point industry avg. Gallup research links high engagement to 21% higher profitability.`,
      formula: `(${latest.EngagementScore} / ${cfg.IndustryEngagementBenchmark||66} − 1) × 21% × ${fmt$(cfg.AnnualRevenue||85000000)} revenue = ${fmt$(Math.abs(bv.engPremium))}` },
    { title:'Annual Absenteeism Cost', amount: fmt$(bv.absCost),
      desc: `At ${latest.AbsenteeismRate}% absenteeism, employees collectively miss ${((latest.AbsenteeismRate/100)*250*(cfg.CurrentHeadcount||480)).toFixed(0)} working days/year. Indirect costs (overtime, cover) add a ${cfg.AbsenteeismIndirectCostMultiplier||1.5}× multiplier.`,
      formula: `${latest.AbsenteeismRate}% × 250 days × ${cfg.CurrentHeadcount||480} employees × $${((cfg.AverageSalary||85000)/250).toFixed(0)}/day × ${cfg.AbsenteeismIndirectCostMultiplier||1.5}× = ${fmt$(bv.absCost)}` },
    { title:'Absenteeism Saving Potential', amount: fmt$(bv.absSaving),
      desc: `Reducing to the ${cfg.IndustryAbsenteeismBenchmark||3.5}% benchmark saves ${fmt$(bv.absSaving)} per year in salary and overhead costs.`,
      formula: `(${latest.AbsenteeismRate}% − ${cfg.IndustryAbsenteeismBenchmark||3.5}%) × 250 × ${cfg.CurrentHeadcount||480} × $${((cfg.AverageSalary||85000)/250).toFixed(0)}/day × ${cfg.AbsenteeismIndirectCostMultiplier||1.5}× = ${fmt$(bv.absSaving)}` },
  ]);

  const months = kpi.map(r=>r.Month);
  lineChart('engagementChart', months, [
    { label:'Engagement Score', data: kpi.map(r=>r.EngagementScore), borderColor: PALETTE.green, backgroundColor: 'rgba(34,197,94,.08)', fill:true },
    { label:'Industry Benchmark', data: months.map(()=>cfg.IndustryEngagementBenchmark||66), borderColor: PALETTE.amber, borderDash:[6,3] },
    { label:'Target', data: months.map(()=>cfg.TargetEngagementScore||78), borderColor: PALETTE.accent, borderDash:[4,4] },
  ]);

  lineChart('absenteeismChart', months, [
    { label:'Absenteeism %', data: kpi.map(r=>r.AbsenteeismRate), borderColor: PALETTE.amber, backgroundColor: 'rgba(245,158,11,.08)', fill:true },
    { label:'Benchmark', data: months.map(()=>cfg.IndustryAbsenteeismBenchmark||3.5), borderColor: PALETTE.red, borderDash:[4,4] },
  ]);

  // By dept
  const depts = [...new Set(att.map(r=>r.Department))];
  const avgAbs = depts.map(d => {
    const rows = att.filter(r=>r.Department===d);
    return +(rows.reduce((s,r)=>s+(+r.AbsenteeismRate||0),0)/rows.length).toFixed(2);
  });
  barChart('absentDeptChart', depts, [{ label:'Absenteeism %', data: avgAbs, backgroundColor: DEPT_COLORS, borderRadius: 4 }], { horizontal: true });

  lineChart('enpsChart', months, [
    { label:'eNPS', data: kpi.map(r=>r.eNPS), borderColor: PALETTE.purple, backgroundColor: 'rgba(168,85,247,.08)', fill:true },
    { label:'Target (40)', data: months.map(()=>40), borderColor: PALETTE.green, borderDash:[4,4] },
    { label:'Zero Line', data: months.map(()=>0), borderColor: '#64748b', borderDash:[2,4] },
  ]);
}

/* ── Learning & Dev ───────────────────────────────────────── */
function renderLearning() {
  const { kpi, trn, cfg } = STATE;
  const latest = kpi[kpi.length-1], prev = kpi[kpi.length-2]||kpi[0];
  const bv = calcBusinessValue(cfg, kpi);

  const avgPre  = +(trn.reduce((s,r)=>s+(+r.PreScore||0),0)/trn.length).toFixed(1);
  const avgPost = +(trn.reduce((s,r)=>s+(+r.PostScore||0),0)/trn.length).toFixed(1);
  const avgCompl= +(trn.reduce((s,r)=>s+(+r.CompletionRate||0),0)/trn.length).toFixed(1);

  renderKpiCards('learningGrid', [
    { name:'Training ROI',          value: latest.TrainingROI+'%',        icon:'fa-chart-pie',    color:'green', icolor:'green', trend:{dir:'up',pct:((latest.TrainingROI-prev.TrainingROI)/prev.TrainingROI*100).toFixed(1)}, tgt:{label:'Target: 300%', color:latest.TrainingROI>=300?'green':'amber'}, bvBadge:`Value: ${fmt$(bv.trainingValue)}` },
    { name:'Training Hrs/Employee', value: latest.TrainingHoursPerEmp+'h',icon:'fa-clock',        color:'accent', icolor:'accent', trend:{dir:'up',pct:1.5}, tgt:{label:'Target: 40h', color:latest.TrainingHoursPerEmp>=40?'green':'amber'}, bvBadge:null },
    { name:'Completion Rate',       value: avgCompl+'%',                   icon:'fa-circle-check', color:'green', icolor:'green', trend:{dir:'up',pct:.7}, tgt:{label:'Target: 90%', color:avgCompl>=90?'green':'amber'}, bvBadge:null },
    { name:'Avg Score Improvement', value: '+'+(avgPost-avgPre).toFixed(1)+'pts', icon:'fa-arrow-trend-up', color:'green', icolor:'green', trend:{dir:'up',pct:2}, tgt:{label:'Target: +15pts', color:(avgPost-avgPre)>=15?'green':'amber'}, bvBadge:null },
  ]);

  const totalCost = trn.reduce((s,r)=>s+(+r.TotalCost||0),0);
  const roi = latest.TrainingROI;

  renderBvPanel('learningBvPanel', 'Business Value: Learning & Development ROI', [
    { title:'Training ROI', amount: latest.TrainingROI+'%',
      desc: `For every $1 invested in L&D, the organisation generates $${(roi/100+1).toFixed(2)} in performance value. Industry-leading programmes achieve 300%+.`,
      formula: `(Post-training performance value − Training cost) / Training cost × 100 = ${roi}%` },
    { title:'L&D Value Generated', amount: fmt$(bv.trainingValue),
      desc: `Training 25% of the workforce drives ~13% performance improvement. Applied to revenue per employee, this translates to ${fmt$(bv.trainingValue)} in productivity value.`,
      formula: `13% improvement × ${fmt$(cfg.AnnualRevenue||85000000)} / ${cfg.CurrentHeadcount||480} × ${Math.round((cfg.CurrentHeadcount||480)*.25)} trained employees` },
    { title:'Annual L&D Investment', amount: fmt$(totalCost),
      desc: `Total training spend across all departments and programmes in the last 12 months. Benchmark: 2–3% of payroll ($${(((cfg.AverageSalary||85000)*(cfg.CurrentHeadcount||480))*0.025/1000).toFixed(0)}K–$${(((cfg.AverageSalary||85000)*(cfg.CurrentHeadcount||480))*0.035/1000).toFixed(0)}K).`,
      formula: null },
  ]);

  const months = kpi.map(r=>r.Month);
  lineChart('trainingRoiChart', months, [
    { label:'Training ROI %', data: kpi.map(r=>r.TrainingROI), borderColor: PALETTE.green, backgroundColor: 'rgba(34,197,94,.08)', fill:true },
    { label:'Target 300%', data: months.map(()=>300), borderColor: PALETTE.accent, borderDash:[4,4] },
  ]);

  // Monthly completion rate
  const byMonth = months.map(m => {
    const rows = trn.filter(r=>r.Month===m);
    if (!rows.length) return 0;
    return +(rows.reduce((s,r)=>s+(+r.CompletionRate||0),0)/rows.length).toFixed(1);
  });
  lineChart('completionChart', months, [
    { label:'Completion Rate %', data: byMonth, borderColor: PALETTE.teal, backgroundColor: 'rgba(20,184,166,.08)', fill:true },
    { label:'Target 90%', data: months.map(()=>90), borderColor: PALETTE.green, borderDash:[4,4] },
  ]);

  // Pre/Post by dept
  const depts = [...new Set(trn.map(r=>r.Department))];
  const preAvg  = depts.map(d => { const rows=trn.filter(r=>r.Department===d); return +(rows.reduce((s,r)=>s+(+r.PreScore||0),0)/rows.length).toFixed(1); });
  const postAvg = depts.map(d => { const rows=trn.filter(r=>r.Department===d); return +(rows.reduce((s,r)=>s+(+r.PostScore||0),0)/rows.length).toFixed(1); });
  barChart('prePostChart', depts, [
    { label:'Pre-Training',  data: preAvg,  backgroundColor: 'rgba(6,182,212,.5)', borderRadius: 4 },
    { label:'Post-Training', data: postAvg, backgroundColor: PALETTE.green, borderRadius: 4 },
  ]);

  lineChart('trainingHrsChart', months, [
    { label:'Hours/Employee', data: kpi.map(r=>r.TrainingHoursPerEmp), borderColor: PALETTE.purple, backgroundColor: 'rgba(168,85,247,.08)', fill:true },
    { label:'Target 40h', data: months.map(()=>40), borderColor: PALETTE.green, borderDash:[4,4] },
  ]);
}

/* ── Diversity ────────────────────────────────────────────── */
function renderDiversity() {
  const { emp, cfg } = STATE;

  const countBy = key => {
    const m = {};
    emp.forEach(e => { const v=e[key]||'Unknown'; m[v]=(m[v]||0)+1; });
    return m;
  };

  const genderMap = countBy('Gender');
  const ageMap    = countBy('AgeGroup');
  const levelMap  = countBy('Level');
  const deptMap   = countBy('Department');

  const femPct  = ((genderMap['Female']||0)/emp.length*100).toFixed(0);
  const deiScore= Math.round(50 + (femPct-40)*0.5 + Math.random()*5);

  renderKpiCards('diversityGrid', [
    { name:'Female Representation', value: femPct+'%',  icon:'fa-venus',         color: femPct>=45?'green':'amber', icolor:'accent', trend:{dir:'up',pct:.5}, tgt:{label:'Target: 45%', color:femPct>=45?'green':'amber'}, bvBadge:'McKinsey: +15% profitability' },
    { name:'DEI Score',             value: deiScore+'/100', icon:'fa-people-group', color:'green', icolor:'green', trend:{dir:'up',pct:.8}, tgt:{label:'Target: 75', color:deiScore>=75?'green':'amber'}, bvBadge:'Deloitte: +6× innovation' },
    { name:'Age Diversity Index',   value: Object.keys(ageMap).length+'/5 groups', icon:'fa-chart-bar', color:'accent', icolor:'accent', trend:{dir:'up',pct:0}, tgt:{label:'Target: 5 groups', color:Object.keys(ageMap).length>=5?'green':'amber'}, bvBadge:null },
    { name:'Total Headcount',       value: emp.length,    icon:'fa-users',         color:'accent', icolor:'accent', trend:{dir:'up',pct:.5}, tgt:null, bvBadge:null },
  ]);

  renderBvPanel('diversityBvPanel', 'Business Value: Diversity & Inclusion', [
    { title:'Gender Diversity Business Impact', amount: fmt$((cfg.AnnualRevenue||85000000)*0.15),
      desc: `McKinsey research: companies in the top quartile for gender diversity are 15% more likely to outperform peers. Applied to ${fmt$(cfg.AnnualRevenue||85000000)} revenue.`,
      formula: `15% profitability premium × ${fmt$(cfg.AnnualRevenue||85000000)} = ${fmt$((cfg.AnnualRevenue||85000000)*0.15)}` },
    { title:'Innovation Premium — Diverse Teams', amount: fmt$((cfg.AnnualRevenue||85000000)*0.06),
      desc: 'Deloitte: inclusive teams make better decisions 87% of the time and are 6× more likely to be innovative. Estimated 6% innovation-driven revenue uplift.',
      formula: `6% innovation uplift × ${fmt$(cfg.AnnualRevenue||85000000)} = ${fmt$((cfg.AnnualRevenue||85000000)*0.06)}` },
    { title:'Talent Attraction Premium', amount: fmt$((cfg.TurnoverReplacementCostMultiplier||1.5)*(cfg.AverageSalary||85000)*Math.round((cfg.CurrentHeadcount||480)*.05)),
      desc: `Glassdoor: 76% of job seekers consider D&I important. A strong D&I brand reduces voluntary turnover by ~5%, avoiding ${fmt$((cfg.TurnoverReplacementCostMultiplier||1.5)*(cfg.AverageSalary||85000)*Math.round((cfg.CurrentHeadcount||480)*.05))} in replacement costs.`,
      formula: null },
  ]);

  const gLabels = Object.keys(genderMap);
  doughnutChart('genderChart', gLabels, gLabels.map(k=>genderMap[k]), [PALETTE.accent, PALETTE.pink, PALETTE.purple]);

  const ageOrder = ['Under 25','25-34','35-44','45-54','55+'];
  const ageCounts = ageOrder.map(a => ageMap[a]||0);
  barChart('ageChart', ageOrder, [{ label:'Employees', data: ageCounts, backgroundColor: DEPT_COLORS.slice(0,5), borderRadius: 4 }]);

  const levelOrder = ['Individual Contributor','Senior','Lead/Principal','Manager','Director','VP+'];
  const levelCounts = levelOrder.map(l=>levelMap[l]||0);
  barChart('levelChart', levelOrder, [{ label:'Employees', data: levelCounts, backgroundColor: DEPT_COLORS, borderRadius: 4 }], { horizontal: true });

  // Gender by Dept stacked
  const depts = [...new Set(emp.map(e=>e.Department))];
  const maleData   = depts.map(d => emp.filter(e=>e.Department===d&&e.Gender==='Male').length);
  const femaleData = depts.map(d => emp.filter(e=>e.Department===d&&e.Gender==='Female').length);
  const nbData     = depts.map(d => emp.filter(e=>e.Department===d&&e.Gender==='Non-Binary').length);
  barChart('genderDeptChart', depts, [
    { label:'Male',       data: maleData,   backgroundColor: 'rgba(6,182,212,.7)', borderRadius: 2 },
    { label:'Female',     data: femaleData, backgroundColor: 'rgba(236,72,153,.7)', borderRadius: 2 },
    { label:'Non-Binary', data: nbData,     backgroundColor: 'rgba(168,85,247,.7)', borderRadius: 2 },
  ], { stacked: true });

  // Gender by level
  const femByLevel  = levelOrder.map(l => emp.filter(e=>e.Level===l&&e.Gender==='Female').length);
  const maleByLevel = levelOrder.map(l => emp.filter(e=>e.Level===l&&e.Gender==='Male').length);
  barChart('genderLevelChart', levelOrder, [
    { label:'Male',   data: maleByLevel,  backgroundColor: 'rgba(6,182,212,.7)',  borderRadius: 3 },
    { label:'Female', data: femByLevel,   backgroundColor: 'rgba(236,72,153,.7)', borderRadius: 3 },
  ], { stacked: true, horizontal: true });
}

/* ════════════════════════════════════════════════════════════
   NAVIGATION
════════════════════════════════════════════════════════════ */
const SECTIONS = ['overview','workforce','recruitment','performance','engagement','learning','diversity'];
const TITLES = {
  overview:    ['Executive Overview',     'All key HR metrics connected to business value'],
  workforce:   ['Workforce Analytics',    'Headcount, turnover and retention trends'],
  recruitment: ['Recruitment Analytics',  'Hiring speed, cost, quality & offer acceptance'],
  performance: ['Performance Analytics',  'Ratings, productivity & revenue per employee'],
  engagement:  ['Engagement & Wellbeing', 'Engagement score, eNPS & absenteeism'],
  learning:    ['Learning & Development', 'Training ROI, completion & skill improvement'],
  diversity:   ['Diversity & Inclusion',  'Gender, age & level representation'],
};

const RENDERERS = {
  overview: renderOverview, workforce: renderWorkforce, recruitment: renderRecruitment,
  performance: renderPerformance, engagement: renderEngagement,
  learning: renderLearning, diversity: renderDiversity,
};

function showSection(name) {
  SECTIONS.forEach(s => {
    const el = document.getElementById(`sec-${s}`);
    if (el) el.classList.toggle('hidden', s !== name);
  });
  document.querySelectorAll('.nav-item[data-section]').forEach(a => {
    a.classList.toggle('active', a.dataset.section === name);
  });
  const [title, desc] = TITLES[name] || ['HR Dashboard',''];
  document.getElementById('sectionTitle').textContent = title;
  document.getElementById('sectionDesc').textContent  = desc;
  STATE.activeSection = name;
  RENDERERS[name]?.();
}

/* ════════════════════════════════════════════════════════════
   LIVE MODE
════════════════════════════════════════════════════════════ */
function startLive() {
  STATE.liveMode = true;
  document.getElementById('liveBadge').textContent = 'ON';
  document.getElementById('liveBadge').classList.add('on');
  document.getElementById('liveDot').classList.remove('hidden');
  STATE.liveTimer = setInterval(() => {
    const last = STATE.kpi[STATE.kpi.length-1];
    STATE.kpi.push({
      ...last,
      TurnoverRate:        +(last.TurnoverRate    + jitter(.15)).toFixed(1),
      EngagementScore:     +(last.EngagementScore + jitter(.3)).toFixed(1),
      AbsenteeismRate:     +(last.AbsenteeismRate + jitter(.08)).toFixed(1),
      TimeToHire_Days:     +(last.TimeToHire_Days + jitter(.5)).toFixed(1),
      CostPerHire:         Math.round(last.CostPerHire + jitter(50)),
      RevenuePerEmployee:  Math.round(last.RevenuePerEmployee + jitter(500)),
      PerformanceAvg:      +(last.PerformanceAvg  + jitter(.02)).toFixed(2),
      TrainingROI:         Math.round(last.TrainingROI + jitter(5)),
      eNPS:                Math.round(last.eNPS + jitter(1)),
      Headcount:           Math.round(last.Headcount + jitter(2)),
    });
    if (STATE.kpi.length > 24) STATE.kpi.shift();
    RENDERERS[STATE.activeSection]?.();
    document.getElementById('lastUpdated').textContent = 'Updated: '+new Date().toLocaleTimeString();
  }, 3000);
}

function stopLive() {
  STATE.liveMode = false;
  clearInterval(STATE.liveTimer);
  document.getElementById('liveBadge').textContent = 'OFF';
  document.getElementById('liveBadge').classList.remove('on');
  document.getElementById('liveDot').classList.add('hidden');
}

/* ════════════════════════════════════════════════════════════
   TOAST
════════════════════════════════════════════════════════════ */
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  el.classList.add('show');
  setTimeout(() => { el.classList.remove('show'); setTimeout(()=>el.classList.add('hidden'),350); }, 3000);
}

/* ════════════════════════════════════════════════════════════
   DEPARTMENT FILTER
════════════════════════════════════════════════════════════ */
function populateDeptFilter(emp) {
  const depts = [...new Set(emp.map(e=>e.Department))].sort();
  const sel = document.getElementById('deptFilter');
  sel.innerHTML = '<option value="all">All Departments</option>' +
    depts.map(d=>`<option value="${d}">${d}</option>`).join('');
}

function applyDeptFilter(dept) {
  STATE.dept = dept;
  // Filter employee-level data
  const allEmp = STATE._allEmp || STATE.emp;
  STATE._allEmp = allEmp;
  STATE.emp = dept === 'all' ? allEmp : allEmp.filter(e=>e.Department===dept);
  const allRec = STATE._allRec || STATE.rec;
  STATE._allRec = allRec;
  STATE.rec = dept === 'all' ? allRec : allRec.filter(r=>r.Department===dept);
  const allAtt = STATE._allAtt || STATE.att;
  STATE._allAtt = allAtt;
  STATE.att = dept === 'all' ? allAtt : allAtt.filter(r=>r.Department===dept);
  const allTrn = STATE._allTrn || STATE.trn;
  STATE._allTrn = allTrn;
  STATE.trn = dept === 'all' ? allTrn : allTrn.filter(r=>r.Department===dept);
  RENDERERS[STATE.activeSection]?.();
}

/* ════════════════════════════════════════════════════════════
   INIT & EVENT WIRING
════════════════════════════════════════════════════════════ */
function loadData(data) {
  Object.assign(STATE, data);
  STATE._allEmp = data.emp;
  STATE._allRec = data.rec;
  STATE._allAtt = data.att;
  STATE._allTrn = data.trn;
  populateDeptFilter(data.emp);
  document.getElementById('lastUpdated').textContent = 'Loaded: ' + new Date().toLocaleTimeString();
  showSection(STATE.activeSection);
  showToast('✅ Data loaded successfully');
}

document.addEventListener('DOMContentLoaded', () => {

  /* ── Nav clicks ── */
  document.querySelectorAll('.nav-item[data-section]').forEach(a => {
    a.addEventListener('click', e => { e.preventDefault(); showSection(a.dataset.section); });
  });

  /* ── Mobile menu ── */
  document.getElementById('menuToggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('open');
  });

  /* ── Upload triggers ── */
  const openModal = () => document.getElementById('modal').classList.remove('hidden');
  const closeModal = () => document.getElementById('modal').classList.add('hidden');
  document.getElementById('uploadBtn').addEventListener('click', openModal);
  document.getElementById('uploadNavBtn').addEventListener('click', e => { e.preventDefault(); openModal(); });
  document.getElementById('modalClose').addEventListener('click', closeModal);
  document.getElementById('modal').addEventListener('click', e => { if(e.target===document.getElementById('modal')) closeModal(); });

  /* ── Sample data btn ── */
  document.getElementById('useSampleBtn').addEventListener('click', () => {
    closeModal();
    const s = buildSampleState();
    document.getElementById('dataSourceLabel').innerHTML = '<i class="fas fa-database"></i> Sample Data';
    loadData(s);
  });

  /* ── File input ── */
  const fileInput = document.getElementById('fileInput');
  document.getElementById('dropLink').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', async () => {
    if (!fileInput.files[0]) return;
    try {
      const data = await parseExcel(fileInput.files[0]);
      document.getElementById('dataSourceLabel').innerHTML = `<i class="fas fa-file-excel"></i> ${fileInput.files[0].name}`;
      closeModal();
      loadData(data);
    } catch(e) { showToast('❌ Error parsing Excel: '+e.message); }
  });

  /* ── Drag & drop ── */
  const zone = document.getElementById('dropZone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', async e => {
    e.preventDefault(); zone.classList.remove('over');
    const file = e.dataTransfer.files[0];
    if (!file) return;
    try {
      const data = await parseExcel(file);
      document.getElementById('dataSourceLabel').innerHTML = `<i class="fas fa-file-excel"></i> ${file.name}`;
      closeModal();
      loadData(data);
    } catch(err) { showToast('❌ '+err.message); }
  });

  /* ── Live mode ── */
  document.getElementById('liveToggleBtn').addEventListener('click', e => {
    e.preventDefault();
    STATE.liveMode ? stopLive() : startLive();
  });

  /* ── Dept filter ── */
  document.getElementById('deptFilter').addEventListener('change', e => applyDeptFilter(e.target.value));

  /* ── Auto-load sample data on start ── */
  const s = buildSampleState();
  loadData(s);
});
