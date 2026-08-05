/* ============================================================
   dashboard.js — Home Dashboard (Live SharePoint version)
   ============================================================ */

'use strict';

/* ── Page Init ──────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', async () => {
  applyChartDefaults();

  try {
    // loadAllData() calls loadAllDataLive() in sharepoint.js
    // which triggers Microsoft sign-in if needed
    const data = await loadAllData();

    renderKPICards(data);
    renderMetricStrip(data);
    renderCharts(data);
    setupSearch(data);
    setupManualRefresh(data);

    // Start auto-refresh — re-renders dashboard when new data arrives
    startAutoRefresh(() => {
      renderKPICards(VulnData);
      renderMetricStrip(VulnData);
      renderCharts(VulnData);
    });

    hideLoader();
    showToast('Live data loaded from SharePoint', 'success');

  } catch (err) {
    hideLoader();
    showErrorState(err.message);
  }
});

/* ── Error State ────────────────────────────────────────────── */
function showErrorState(msg) {
  const el = document.getElementById('errorBanner');
  if (el) {
    el.style.display = 'flex';
    const txt = el.querySelector('.error-msg');
    if (txt) txt.textContent = msg;
  }
  showToast('Failed to load SharePoint data. Check config.js settings.', 'error', 8000);
}

/* ── KPI Cards ──────────────────────────────────────────────── */
function renderKPICards(data) {
  const cats = [
    { key:'application', page:'details.html?category=application' },
    { key:'device',      page:'details.html?category=device' },
    { key:'astra',       page:'details.html?category=astra' },
    { key:'database',    page:'details.html?category=database' }
  ];

  cats.forEach(({ key, page }) => {
    const card  = document.getElementById(`kpi-${key}`);
    if (!card) return;
    const stats = calcStats(data[key] || []);

    // Click → navigate
    card.onclick = () => { window.location.href = page; };

    // Main number
    const mainEl = card.querySelector('.kpi-card-main');
    if (mainEl) animateCounter(mainEl, stats.active);

    // Severity minis
    const minis  = card.querySelectorAll('.sev-mini');
    const vals   = [stats.critical, stats.high, stats.medium, stats.low];
    minis.forEach((m, i) => {
      const n = m.querySelector('.num');
      if (n && vals[i] !== undefined) animateCounter(n, vals[i], 900);
    });

    // Past due
    const pdEl = card.querySelector('.kpi-past-due');
    if (pdEl) pdEl.textContent = stats.pastDue.toLocaleString();

    // Trend
    const trendEl = card.querySelector('.kpi-trend');
    if (trendEl) {
      const critPct = stats.active > 0 ? (stats.critical / stats.active * 100) : 0;
      trendEl.className  = critPct > 5 ? 'kpi-trend up' : 'kpi-trend down';
      trendEl.innerHTML  = critPct > 5
        ? '<i class="fas fa-arrow-up"></i> High Risk'
        : '<i class="fas fa-arrow-down"></i> Improving';
    }
  });
}

/* ── Metric Strip ───────────────────────────────────────────── */
function renderMetricStrip(data) {
  let totalActive=0, totalPastDue=0, totalClosed=0, allC=0, allH=0, allM=0, allL=0;

  ['application','device','astra','database'].forEach(key => {
    const st = calcStats(data[key] || []);
    totalActive  += st.active;
    totalPastDue += st.pastDue;
    totalClosed  += st.closed;
    allC += st.critical; allH += st.high; allM += st.medium; allL += st.low;
  });

  const total = allC + allH + allM + allL || 1;

  setCount('metric-total-active',   totalActive);
  setCount('metric-past-due',       totalPastDue);
  setCount('metric-resolved',       totalClosed);
  setCount('metric-critical-count', allC);
  setCount('metric-high-count',     allH);
  setCount('metric-risk-score', Math.round((allC*10 + allH*7 + allM*4 + allL*1) / total));

  setText('pct-critical', Math.round(allC/total*100) + '%');
  setText('pct-high',     Math.round(allH/total*100) + '%');
  setText('pct-medium',   Math.round(allM/total*100) + '%');
  setText('pct-low',      Math.round(allL/total*100) + '%');
}

function setCount(id, val) {
  const el = document.getElementById(id);
  if (el) animateCounter(el, val);
}
function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

/* ── Charts ─────────────────────────────────────────────────── */
function renderCharts(data) {
  renderSeverityPie(data);
  renderCategoryBar(data);
  renderMonthlyLine(data.application || []);
  renderSLADonut(data);
  renderTopAppsBar(data.application || []);
  renderHeatmap(data.application || []);
}

function renderSeverityPie(data) {
  let c=0,h=0,m=0,l=0;
  ['application','device','astra','database'].forEach(k => {
    const st = calcStats(data[k]||[]);
    c+=st.critical; h+=st.high; m+=st.medium; l+=st.low;
  });
  buildPieChart('chartSeverityPie', {
    labels: ['Critical','High','Medium','Low'],
    data:   [c, h, m, l],
    colors: ['#ff3b30','#ff9500','#ffcc00','#34c759']
  });
}

function renderCategoryBar(data) {
  const keys   = ['application','device','astra','database'];
  const labels = ['Application','Device','ASTRA','Database'];
  const sevColors = { Critical:'#ff3b30', High:'#ff9500', Medium:'#ffcc00', Low:'#34c759' };
  const datasets = ['Critical','High','Medium','Low'].map(sev => ({
    label: sev,
    data:  keys.map(k => calcStats(data[k]||[])[sev.toLowerCase()]),
    backgroundColor: sevColors[sev] + '99',
    borderColor:     sevColors[sev],
    borderWidth: 1
  }));
  const ctx = document.getElementById('chartCategoryBar');
  if (!ctx) return;
  if (ctx._chart) ctx._chart.destroy();
  ctx._chart = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { stacked: true, grid: { color: 'rgba(0,180,255,0.05)' }, ticks: { color:'#8fb3d0' } },
        y: { stacked: true, grid: { color: 'rgba(0,180,255,0.05)' }, ticks: { color:'#8fb3d0' } }
      },
      plugins: { legend: { position:'top', labels: { color:'#8fb3d0' } } }
    }
  });
}

function renderMonthlyLine(rows) {
  buildLineChart('chartMonthlyLine', getMonthlyTrend(rows));
}

function renderSLADonut(data) {
  let within=0, atRisk=0, past=0;

  (data.astra||[]).forEach(r => {
    const s = String(r.slaStatus||'').toLowerCase();
    if (s.includes('within')) within++;
    else if (s.includes('at risk')) atRisk++;
    else if (s.includes('past'))    past++;
  });
  (data.application||[]).forEach(r => {
    const s = String(r.slaStatus||'').toLowerCase();
    if (s.includes('on track') || s.includes('within')) within++;
    else if (s.includes('past')) past++;
    else within++;
  });
  (data.device||[]).forEach(r => {
    const s = String(r.slaStatus||'').toLowerCase();
    if (s.includes('in-flight')) within++;
    else if (s.includes('miss')) past++;
    else if (s.includes('approach')) atRisk++;
    else within++;
  });

  const ctx = document.getElementById('chartSLADonut');
  if (!ctx) return;
  if (ctx._chart) ctx._chart.destroy();
  ctx._chart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Within SLA','At Risk','Past SLA'],
      datasets: [{ data:[within,atRisk,past], backgroundColor:['#34c75999','#ff950099','#ff3b3099'], borderColor:['#34c759','#ff9500','#ff3b30'], borderWidth:2, hoverOffset:8 }]
    },
    options: { responsive:true, maintainAspectRatio:false, cutout:'65%', plugins:{ legend:{ position:'bottom' } } }
  });
}

function renderTopAppsBar(rows) {
  const top = getTopApps(rows, 10);
  const ctx = document.getElementById('chartTopApps');
  if (!ctx) return;
  if (ctx._chart) ctx._chart.destroy();
  const max = top[0]?.value || 1;
  const colors = top.map(t => {
    const p = t.value/max;
    return p>0.8?'#ff3b3099': p>0.5?'#ff950099': p>0.25?'#ffcc0099':'#34c75999';
  });
  ctx._chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: top.map(t=>t.label),
      datasets: [{ data:top.map(t=>t.value), backgroundColor:colors, borderColor:colors.map(c=>c.replace('99','')), borderWidth:1 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      scales: {
        x: { grid:{ color:'rgba(0,180,255,0.05)' }, ticks:{ color:'#8fb3d0' } },
        y: { grid:{ color:'transparent' }, ticks:{ color:'#8fb3d0', font:{ size:11 } } }
      },
      plugins: { legend:{ display:false } }
    }
  });
}

function renderHeatmap(rows) {
  const ctx = document.getElementById('chartHeatmap');
  if (!ctx) return;
  const sevs = ['Critical','High','Medium','Low'];
  const matrix = {};
  sevs.forEach(s => { matrix[s] = { Critical:0, High:0, Medium:0, Low:0 }; });

  rows.forEach(r => {
    const sev   = normSeverity(r.severity||'');
    const score = parseFloat(r.score)||0;
    const impact = score>=70?'Critical': score>=50?'High': score>=30?'Medium':'Low';
    if (matrix[sev]?.[impact] !== undefined) matrix[sev][impact]++;
  });

  const bubbleData = [];
  const colors = ['#ff3b30','#ff9500','#ffcc00','#34c759'];
  sevs.forEach((row,ri) => {
    sevs.forEach((col,ci) => {
      const val = matrix[row][col];
      if (val>0) bubbleData.push({ x:ci, y:3-ri, r:Math.min(Math.sqrt(val)*3,30), label:`${row}/${col}: ${val}`, color:colors[ri] });
    });
  });

  if (ctx._chart) ctx._chart.destroy();
  ctx._chart = new Chart(ctx, {
    type:'bubble',
    data: { datasets:[{ data:bubbleData, backgroundColor:bubbleData.map(d=>d.color+'88'), borderColor:bubbleData.map(d=>d.color), borderWidth:1 }] },
    options: {
      responsive:true, maintainAspectRatio:false,
      scales: {
        x: { min:-0.5, max:3.5, grid:{ color:'rgba(0,180,255,0.08)' }, ticks:{ callback:i=>['Low','Med','High','Crit'][Math.round(i)], color:'#8fb3d0' }, title:{ display:true, text:'Impact', color:'#8fb3d0' } },
        y: { min:-0.5, max:3.5, grid:{ color:'rgba(0,180,255,0.08)' }, ticks:{ callback:i=>['Low','Med','High','Crit'][Math.round(i)], color:'#8fb3d0' }, title:{ display:true, text:'Severity', color:'#8fb3d0' } }
      },
      plugins: { legend:{ display:false }, tooltip:{ callbacks:{ label:ctx=>ctx.raw.label } } }
    }
  });
}

/* ── Manual Refresh Button ──────────────────────────────────── */
function setupManualRefresh(data) {
  const btn = document.getElementById('btnRefresh');
  if (!btn) return;
  btn.addEventListener('click', async () => {
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Refreshing...';
    try {
      const fresh = await loadAllData(true);
      renderKPICards(fresh);
      renderMetricStrip(fresh);
      renderCharts(fresh);
      showToast('Dashboard refreshed from SharePoint', 'success');
      updateLastRefreshTime();
    } catch(e) {
      showToast('Refresh failed: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
      btn.innerHTML = '<i class="fas fa-rotate-right"></i> Refresh';
    }
  });
}

/* ── Global Search ──────────────────────────────────────────── */
function setupSearch(data) {
  const input   = document.getElementById('globalSearch');
  const results = document.getElementById('searchResults');
  if (!input || !results) return;

  let timer;
  input.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      const q = input.value.trim().toLowerCase();
      if (!q || q.length < 2) { results.classList.remove('show'); return; }
      const hits = [];
      (data.application||[]).slice(0,3000).forEach(row => {
        if (hits.length >= 8) return;
        if ((row.id||'').toLowerCase().includes(q) || (row.name||'').toLowerCase().includes(q) || (row.app||'').toLowerCase().includes(q)) {
          hits.push({ title: row.id||row.app, meta:`${row.app} — ${row.severity}`, cat:'application' });
        }
      });
      results.innerHTML = hits.length
        ? hits.map(h=>`<div class="search-result-item" onclick="location.href='details.html?category=${h.cat}&search=${encodeURIComponent(h.title)}'"><div class="res-title">${h.title}</div><div class="res-meta">${h.meta}</div></div>`).join('')
        : '<div class="search-result-item"><span class="res-meta">No results found</span></div>';
      results.classList.add('show');
    }, 300);
  });

  input.addEventListener('blur', () => setTimeout(()=>results.classList.remove('show'), 200));
}
