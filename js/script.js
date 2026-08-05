/* ============================================================
   script.js — Core Utilities + Data Bridge
   Works with BOTH the live SharePoint loader (sharepoint.js)
   and keeps all chart/table/badge helper functions intact.
   ============================================================ */

'use strict';

/* ── Global Data Store ─────────────────────────────────────── */
const VulnData = {
  application : [],
  device      : [],
  astra       : [],
  database    : [],
  loaded      : false,
  loadError   : null,
  lastFetched : null
};

/* ── Compatibility shim — the dashboard/details JS call
   loadAllData() which now delegates to the SharePoint
   live loader defined in sharepoint.js
─────────────────────────────────────────────────────────── */
async function loadAllData(forceRefresh = false) {
  if (VulnData.loaded && !forceRefresh) return VulnData;
  return loadAllDataLive(forceRefresh);
}

/* ── FieldMap (kept for drill.html drill-down page) ─────────
   On the drill-down page rows are already normalized objects
   (not raw SharePoint rows), so these passthrough mappings
   just return the property directly.
─────────────────────────────────────────────────────────── */
const FieldMap = {
  application: {
    id        : r => r.id         || '',
    name      : r => r.name       || '',
    app       : r => r.app        || '',
    severity  : r => r.severity   || 'Medium',
    status    : r => r.status     || 'Open',
    slaStatus : r => r.slaStatus  || '',
    score     : r => r.score      || 0,
    dueDate   : r => r.dueDate,
    created   : r => r.created,
    desc      : r => r.desc       || '',
    rec       : r => r.rec        || '',
    category  : r => r.category   || '',
    portfolio : r => r.portfolio  || '',
    family    : r => r.family     || '',
    cve       : r => r.cve        || '',
    tool      : r => r.tool       || '',
    location  : r => r.location   || '',
    firstFound: r => r.firstFound,
    aging     : r => r.aging      || ''
  },
  astra: {
    id        : r => r.id         || '',
    name      : r => r.name       || '',
    app       : r => r.app        || '',
    severity  : r => r.severity   || 'Medium',
    status    : r => r.status     || '',
    slaStatus : r => r.slaStatus  || '',
    score     : r => r.score      || 0,
    dueDate   : r => r.dueDate,
    created   : r => r.created,
    desc      : r => r.desc       || '',
    rec       : r => r.rec        || '',
    category  : r => r.category   || '',
    portfolio : r => r.portfolio  || '',
    family    : r => r.family     || '',
    cve       : r => r.cve        || '',
    hostname  : r => r.hostname   || '',
    aging     : r => r.aging      || '',
    contract  : r => r.contract   || '',
    sla       : r => r.sla        || ''
  },
  device: {
    id        : r => r.id         || '',
    name      : r => r.name       || '',
    app       : r => r.app        || '',
    severity  : r => r.severity   || 'Medium',
    status    : r => r.status     || 'Open',
    slaStatus : r => r.slaStatus  || '',
    score     : r => r.score      || 0,
    dueDate   : r => r.dueDate,
    created   : r => r.created,
    desc      : r => r.desc       || '',
    rec       : r => r.rec        || '',
    cve       : r => r.cve        || '',
    hostname  : r => r.hostname   || '',
    ip        : r => r.ip         || '',
    env       : r => r.env        || '',
    team      : r => r.team       || '',
    firstFound: r => r.firstFound,
    lastFound : r => r.lastFound
  },
  database: {
    id        : r => r.id         || '',
    name      : r => r.name       || '',
    app       : r => r.app        || '',
    severity  : r => r.severity   || 'Medium',
    status    : r => r.status     || 'Open',
    slaStatus : r => r.slaStatus  || '',
    score     : r => r.score      || 0,
    dueDate   : r => r.dueDate,
    created   : r => r.created,
    desc      : r => r.desc       || '',
    rec       : r => r.rec        || '',
    cve       : r => r.cve        || '',
    hostname  : r => r.hostname   || '',
    env       : r => r.env        || '',
    aging     : r => r.aging      || '',
    portfolio : r => r.portfolio  || '',
    poc       : r => r.poc        || ''
  }
};

/* ── Severity Normalizer ────────────────────────────────────── */
function normSeverity(val) {
  if (!val) return 'Medium';
  const v = String(val).toLowerCase();
  if (v.includes('critical'))                        return 'Critical';
  if (v.includes('high') || v === '2 - high')        return 'High';
  if (v.includes('med')  || v === '3 - medium')      return 'Medium';
  if (v.includes('low')  || v === '4 - low')         return 'Low';
  if (v.includes('info'))                            return 'Informational';
  return 'Medium';
}

function getSeverityClass(sev) {
  const s = String(sev).toLowerCase();
  if (s === 'critical')      return 'critical';
  if (s === 'high')          return 'high';
  if (s === 'medium')        return 'medium';
  if (s === 'low')           return 'low';
  if (s === 'informational') return 'info';
  return 'medium';
}

function getStatusClass(status) {
  const s = String(status).toLowerCase();
  if (s.includes('past due') || s.includes('past sla') || s.includes('target miss')) return 'past-due';
  if (s.includes('on track') || s.includes('within') || s.includes('in-flight'))     return 'on-track';
  if (s.includes('at risk')  || s.includes('approaching'))                            return 'at-risk';
  if (s.includes('defer'))                                                             return 'deferred';
  if (s.includes('review'))                                                            return 'in-review';
  return 'on-track';
}

/* ── Date Utilities ─────────────────────────────────────────── */
function formatDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function daysUntil(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return null;
  return Math.ceil((d - new Date()) / (1000 * 60 * 60 * 24));
}

/* ── Animated Counter ───────────────────────────────────────── */
function animateCounter(el, target, duration = 1200) {
  if (!el) return;
  const start = performance.now();
  const update = now => {
    const t    = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - t, 3);
    el.textContent = Math.round(target * ease).toLocaleString();
    if (t < 1) requestAnimationFrame(update);
  };
  requestAnimationFrame(update);
}

/* ── Toast Notifications ────────────────────────────────────── */
function showToast(msg, type = 'info', duration = 3500) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const icons  = { success:'fa-check-circle', error:'fa-times-circle', warning:'fa-exclamation-triangle', info:'fa-info-circle' };
  const colors = { success:'#4cd964', error:'#ff6b6b', warning:'#ffa040', info:'#00b4ff' };
  const toast  = document.createElement('div');
  toast.className = `toast-item ${type}`;
  toast.innerHTML = `<i class="fas ${icons[type]}" style="color:${colors[type]};font-size:1rem;"></i><span>${msg}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'toastIn 0.3s ease reverse';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

/* ── Stats Calculator (works on normalized row objects) ─────── */
function calcStats(rows) {
  let critical=0, high=0, medium=0, low=0, info=0, closed=0, pastDue=0, atRisk=0;

  rows.forEach(row => {
    // Row may be a raw object (Excel path) or normalized object (SharePoint path)
    const sev = normSeverity(row.severity || row['Severity'] || row['Risk/Severity'] || row['risk_rating'] || '');
    if (sev === 'Critical')      critical++;
    else if (sev === 'High')     high++;
    else if (sev === 'Medium')   medium++;
    else if (sev === 'Low')      low++;
    else if (sev === 'Informational') info++;

    const st  = String(row.status    || row['State']     || row['state']    || '').toLowerCase();
    const sla = String(row.slaStatus || row['SLAStatus'] || row['ttr_status'] || row['SLA'] || row['Status'] || '').toLowerCase();

    if (st === 'closed' || st === 'resolved') closed++;
    if (sla.includes('past due') || sla.includes('target miss') || sla.includes('past sla')) pastDue++;
    if (sla.includes('at risk')  || sla.includes('approaching')) atRisk++;
  });

  const active = rows.length - closed;
  return { total: rows.length, active, critical, high, medium, low, info, closed, pastDue, atRisk };
}

/* ── Monthly Trend ──────────────────────────────────────────── */
function getMonthlyTrend(rows) {
  const months = {};
  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const d   = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    months[key] = 0;
  }
  rows.forEach(row => {
    const d = row.created;
    if (d instanceof Date && !isNaN(d)) {
      const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      if (key in months) months[key]++;
    }
  });
  return {
    labels: Object.keys(months).map(k => {
      const [y, m] = k.split('-');
      return new Date(+y, +m-1).toLocaleDateString('en-US', { month:'short', year:'2-digit' });
    }),
    data: Object.values(months)
  };
}

/* ── Top N Apps ─────────────────────────────────────────────── */
function getTopApps(rows, n = 10) {
  const counts = {};
  rows.forEach(row => {
    const app = row.app || 'Unknown';
    counts[app] = (counts[app] || 0) + 1;
  });
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .map(([label, value]) => ({ label, value }));
}

/* ── Chart.js Global Defaults ───────────────────────────────── */
function applyChartDefaults() {
  if (typeof Chart === 'undefined') return;
  Chart.defaults.color           = '#8fb3d0';
  Chart.defaults.borderColor     = 'rgba(0,180,255,0.08)';
  Chart.defaults.font.family     = "'Segoe UI', system-ui, sans-serif";
  Chart.defaults.font.size       = 11;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.padding       = 16;
  Chart.defaults.plugins.tooltip.backgroundColor     = 'rgba(4,13,26,0.95)';
  Chart.defaults.plugins.tooltip.borderColor         = 'rgba(0,180,255,0.3)';
  Chart.defaults.plugins.tooltip.borderWidth         = 1;
  Chart.defaults.plugins.tooltip.padding             = 12;
  Chart.defaults.plugins.tooltip.titleColor          = '#e8f4fd';
  Chart.defaults.plugins.tooltip.bodyColor           = '#8fb3d0';
  Chart.defaults.plugins.tooltip.cornerRadius        = 8;
}

/* ── Chart Builders ─────────────────────────────────────────── */
function buildPieChart(canvasId, dist) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  if (ctx._chart) ctx._chart.destroy();
  ctx._chart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: dist.labels,
      datasets: [{
        data: dist.data,
        backgroundColor: dist.colors.map(c => c + 'cc'),
        borderColor: dist.colors,
        borderWidth: 2,
        hoverOffset: 8
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: { legend: { position: 'right' } }
    }
  });
  return ctx._chart;
}

function buildLineChart(canvasId, trend) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  if (ctx._chart) ctx._chart.destroy();
  ctx._chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: trend.labels,
      datasets: [{
        label: 'New Vulnerabilities',
        data: trend.data,
        borderColor: '#00b4ff',
        backgroundColor: 'rgba(0,180,255,0.08)',
        borderWidth: 2,
        pointBackgroundColor: '#00b4ff',
        pointRadius: 4,
        tension: 0.4,
        fill: true
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { color: 'rgba(0,180,255,0.05)' }, ticks: { color: '#8fb3d0' } },
        y: { grid: { color: 'rgba(0,180,255,0.05)' }, ticks: { color: '#8fb3d0', precision: 0 } }
      },
      plugins: { legend: { display: false } }
    }
  });
  return ctx._chart;
}

/* ── URL Helpers ────────────────────────────────────────────── */
function getParam(name) {
  return new URLSearchParams(window.location.search).get(name);
}

function navigateTo(page, params = {}) {
  const url = new URL(page, window.location.href);
  Object.entries(params).forEach(([k,v]) => url.searchParams.set(k, v));
  window.location.href = url.toString();
}

/* ── Excel Export (XLSX) ────────────────────────────────────── */
function exportToExcel(data, filename = 'vulnerabilities.xlsx') {
  if (typeof XLSX === 'undefined') { showToast('XLSX library not loaded', 'error'); return; }
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Vulnerabilities');
  XLSX.writeFile(wb, filename);
  showToast('Excel file exported successfully', 'success');
}

/* ── Theme (Dark / Light) ───────────────────────────────────── */
function initTheme() {
  const saved = localStorage.getItem('portalTheme') || 'dark';
  setTheme(saved);
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('portalTheme', theme);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.querySelector('i').className = theme === 'dark' ? 'fas fa-sun' : 'fas fa-moon';
}

function toggleTheme() {
  const cur = document.documentElement.getAttribute('data-theme') || 'dark';
  setTheme(cur === 'dark' ? 'light' : 'dark');
}

/* ── Navbar Clock ───────────────────────────────────────────── */
function startClock() {
  const dateEl    = document.getElementById('navDate');
  const timeEl    = document.getElementById('navTime');
  const refreshEl = document.getElementById('lastRefresh');
  const tick = () => {
    const now = new Date();
    if (dateEl) dateEl.textContent = now.toLocaleDateString('en-US', { weekday:'short', month:'short', day:'numeric', year:'numeric' });
    if (timeEl) timeEl.textContent = now.toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit', second:'2-digit' });
  };
  tick();
  setInterval(tick, 1000);
  if (refreshEl) refreshEl.textContent = new Date().toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit' });
}

/* ── Loading Overlay ────────────────────────────────────────── */
function hideLoader() {
  const loader = document.getElementById('loadingOverlay');
  if (loader) {
    setTimeout(() => {
      loader.classList.add('hide');
      setTimeout(() => loader.remove(), 500);
    }, 600);
  }
}

/* ── DOM Ready ──────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  startClock();
  applyChartDefaults();

  const themeBtn = document.getElementById('themeToggle');
  if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

  const fab = document.getElementById('fabTop');
  if (fab) fab.addEventListener('click', () => window.scrollTo({ top:0, behavior:'smooth' }));

  const signOutBtn = document.getElementById('btnSignOut');
  if (signOutBtn) signOutBtn.addEventListener('click', async () => {
    await signOut();
    location.reload();
  });
});
