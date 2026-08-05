/* ============================================================
   details.js — Listing / Detail Table Page (Live version)
   ============================================================ */

'use strict';

let currentCategory = 'application';
let currentData     = [];
let dtInstance      = null;

const CatConfig = {
  application: {
    label: 'Application Vulnerabilities', icon:'fa-code', color:'#00b4ff',
    cols: [
      { title:'VIT Number',    key:'id' },
      { title:'App Acronym',   key:'app' },
      { title:'Vulnerability', key:'name' },
      { title:'Severity',      key:'severity',  render: severityBadge },
      { title:'Risk Score',    key:'score',     render: scoreBar },
      { title:'Status',        key:'status' },
      { title:'SLA Status',    key:'slaStatus', render: statusBadge },
      { title:'Due Date',      key:'dueDate',   render: d => formatDate(d) },
      { title:'Days Rem.',     key:'dueDate',   render: daysRemaining },
      { title:'Portfolio',     key:'portfolio' },
      { title:'Category',      key:'category' }
    ]
  },
  astra: {
    label: 'ASTRA Vulnerabilities', icon:'fa-shield-alt', color:'#ff9500',
    cols: [
      { title:'App Acronym',   key:'app' },
      { title:'Hostname',      key:'hostname' },
      { title:'Vulnerability', key:'name' },
      { title:'Severity',      key:'severity',  render: severityBadge },
      { title:'SLA',           key:'slaStatus', render: statusBadge },
      { title:'SLA Date',      key:'dueDate',   render: d => formatDate(d) },
      { title:'Aging',         key:'aging' },
      { title:'Portfolio',     key:'portfolio' },
      { title:'Contract',      key:'contract' }
    ]
  },
  device: {
    label: 'Device Vulnerabilities', icon:'fa-server', color:'#a855f7',
    cols: [
      { title:'VIT Number',    key:'id' },
      { title:'Hostname',      key:'hostname' },
      { title:'Vulnerability', key:'name' },
      { title:'Severity',      key:'severity',  render: severityBadge },
      { title:'Risk Score',    key:'score',     render: scoreBar },
      { title:'State',         key:'status' },
      { title:'TTR Status',    key:'slaStatus', render: statusBadge },
      { title:'Due Date',      key:'dueDate',   render: d => formatDate(d) },
      { title:'IP Address',    key:'ip' },
      { title:'Environment',   key:'env' }
    ]
  },
  database: {
    label: 'Database Vulnerabilities', icon:'fa-database', color:'#34c759',
    cols: [
      { title:'Number',        key:'id' },
      { title:'App Acronym',   key:'app' },
      { title:'Hostname',      key:'hostname' },
      { title:'Vulnerability', key:'name' },
      { title:'Severity',      key:'severity',  render: severityBadge },
      { title:'State',         key:'status' },
      { title:'Status',        key:'slaStatus', render: statusBadge },
      { title:'Due Date',      key:'dueDate',   render: d => formatDate(d) },
      { title:'Aging',         key:'aging' },
      { title:'Portfolio',     key:'portfolio' }
    ]
  }
};

/* ── Cell Renderers ─────────────────────────────────────────── */
function severityBadge(sev) {
  const cls = getSeverityClass(sev||'');
  return `<span class="badge-severity ${cls}"><i class="fas fa-circle" style="font-size:6px"></i> ${sev||'—'}</span>`;
}

function statusBadge(s) {
  if (!s) return '<span class="badge-status on-track">—</span>';
  return `<span class="badge-status ${getStatusClass(s)}">${s}</span>`;
}

function scoreBar(score) {
  const v = parseFloat(score)||0;
  if (!v) return '<span style="color:var(--text-muted)">—</span>';
  const color = v>=70?'#ff3b30': v>=50?'#ff9500': v>=30?'#ffcc00':'#34c759';
  return `<div class="risk-bar-wrap"><div class="risk-bar"><div class="risk-bar-fill" style="width:${Math.min(v,100)}%;background:${color}"></div></div><span class="risk-val" style="color:${color}">${Math.round(v)}</span></div>`;
}

function daysRemaining(dueDate) {
  const days = daysUntil(dueDate);
  if (days===null) return '—';
  if (days<0)  return `<span style="color:#ff3b30;font-weight:700">${days}d</span>`;
  if (days<=7) return `<span style="color:#ff9500;font-weight:700">${days}d</span>`;
  return `<span style="color:#4cd964">${days}d</span>`;
}

/* ── Init ───────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', async () => {
  currentCategory = getParam('category') || 'application';
  if (!CatConfig[currentCategory]) currentCategory = 'application';

  const cfg = CatConfig[currentCategory];

  // Update titles
  const titleEl = document.getElementById('pageTitle');
  const bcCat   = document.getElementById('breadcrumbCategory');
  if (titleEl) titleEl.innerHTML = `<i class="fas ${cfg.icon}" style="color:${cfg.color}"></i> ${cfg.label}`;
  if (bcCat)   bcCat.textContent = cfg.label;

  // Highlight active pill
  document.querySelectorAll('.category-pill').forEach(p =>
    p.classList.toggle('active', p.dataset.cat === currentCategory)
  );

  // Load live data
  const data = await loadAllData();
  currentData = data[currentCategory] || [];

  renderSummaryCards(currentData);
  populateFilters(currentData);
  buildTable(currentData);

  // Button listeners
  document.getElementById('btnApplyFilters')?.addEventListener('click', applyFilters);
  document.getElementById('btnClearFilters')?.addEventListener('click', clearFilters);
  document.getElementById('btnExportExcel')?.addEventListener('click',  doExport);
  document.getElementById('btnExportPDF')?.addEventListener('click',    () => window.print());
  document.getElementById('btnPrint')?.addEventListener('click',        () => window.print());
  document.getElementById('btnRefresh')?.addEventListener('click',      doRefresh);

  // Filter collapse
  const toggle = document.getElementById('filterToggle');
  const body   = document.getElementById('filterBody');
  if (toggle && body) {
    toggle.addEventListener('click', () => {
      const open = body.style.display !== 'none';
      body.style.display = open ? 'none' : 'block';
      toggle.querySelector('.toggle-icon').className = `fas fa-chevron-${open?'down':'up'} toggle-icon`;
    });
  }

  // Pre-set search from URL
  const preSearch = getParam('search');
  if (preSearch && dtInstance) dtInstance.search(preSearch).draw();

  hideLoader();
  showToast(`Loaded ${currentData.length.toLocaleString()} live records from SharePoint`, 'success');
});

/* ── Summary Cards ──────────────────────────────────────────── */
function renderSummaryCards(rows) {
  const st = calcStats(rows);
  [
    ['sum-critical', st.critical],
    ['sum-high',     st.high],
    ['sum-medium',   st.medium],
    ['sum-low',      st.low],
    ['sum-past-due', st.pastDue],
    ['sum-total',    st.active]
  ].forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) animateCounter(el, val);
  });
}

/* ── Filter Dropdowns ───────────────────────────────────────── */
function populateFilters(rows) {
  const unique = field => [...new Set(rows.map(r => r[field]).filter(Boolean))].sort();
  populateSel('filterSeverity', ['Critical','High','Medium','Low','Informational']);
  populateSel('filterApp',      unique('app').slice(0, 150));
  populateSel('filterStatus',   unique('status'));
  populateSel('filterSLA',      unique('slaStatus'));
}

function populateSel(id, opts) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = '<option value="">All</option>' + opts.map(o=>`<option value="${o}">${o}</option>`).join('');
}

/* ── Apply / Clear ──────────────────────────────────────────── */
function applyFilters() {
  if (!dtInstance) return;
  const sev    = document.getElementById('filterSeverity')?.value || '';
  const app    = document.getElementById('filterApp')?.value      || '';
  const status = document.getElementById('filterStatus')?.value   || '';
  const sla    = document.getElementById('filterSLA')?.value      || '';
  const kw     = document.getElementById('filterSearch')?.value   || '';

  $.fn.dataTable.ext.search = [];
  if (sev || app || status || sla) {
    $.fn.dataTable.ext.search.push((_, __, idx) => {
      const r = currentData[idx];
      if (!r) return true;
      if (sev    && r.severity  !== sev)    return false;
      if (app    && r.app       !== app)    return false;
      if (status && r.status    !== status) return false;
      if (sla    && r.slaStatus !== sla)    return false;
      return true;
    });
  }
  dtInstance.search(kw).draw();
  showToast('Filters applied', 'info');
}

function clearFilters() {
  ['filterSeverity','filterApp','filterStatus','filterSLA','filterSearch'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  $.fn.dataTable.ext.search = [];
  dtInstance?.search('').draw();
  showToast('Filters cleared', 'info');
}

/* ── DataTable ──────────────────────────────────────────────── */
function buildTable(rows) {
  const cfg  = CatConfig[currentCategory];
  const el   = document.getElementById('mainTable');
  if (!el) return;

  const columns = cfg.cols.map(col => ({
    title: col.title,
    defaultContent: '—',
    data: (row, type) => {
      const val = row[col.key];
      return type === 'display' && col.render ? col.render(val) : (val ?? '');
    }
  }));

  if (dtInstance) { dtInstance.destroy(); dtInstance = null; }

  dtInstance = $(el).DataTable({
    data: rows,
    columns,
    pageLength: 25,
    lengthMenu: [[10,25,50,100,-1],[10,25,50,100,'All']],
    order: [[0,'asc']],
    scrollX: true,
    dom: '<"dt-top-row"lf>rt<"dt-bottom-row"ip>',
    language: {
      search: '', searchPlaceholder: 'Search table...',
      info: 'Showing _START_ to _END_ of _TOTAL_ records',
      infoEmpty: 'No records', zeroRecords: 'No matching records'
    },
    createdRow: (row, rowData) => {
      row.style.cursor = 'pointer';
      row.addEventListener('click', () => drillDown(rowData));
    }
  });
}

/* ── Drill Down ─────────────────────────────────────────────── */
function drillDown(rowData) {
  try {
    sessionStorage.setItem('drilldownRow',     JSON.stringify(rowData));
    sessionStorage.setItem('drilldownCategory', currentCategory);
  } catch(e) {}
  window.location.href = `drill.html?category=${currentCategory}&id=${encodeURIComponent(rowData.id||'')}`;
}

/* ── Export ─────────────────────────────────────────────────── */
function doExport() {
  const cfg  = CatConfig[currentCategory];
  const exportData = currentData.map(r => {
    const obj = {};
    cfg.cols.forEach(col => { obj[col.title] = r[col.key] ?? ''; });
    return obj;
  });
  exportToExcel(exportData, `${currentCategory}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ── Live Refresh ───────────────────────────────────────────── */
async function doRefresh() {
  const btn = document.getElementById('btnRefresh');
  if (btn) { btn.disabled=true; btn.innerHTML='<i class="fas fa-spinner fa-spin"></i>'; }
  try {
    const data  = await loadAllData(true);
    currentData = data[currentCategory] || [];
    renderSummaryCards(currentData);
    if (dtInstance) { dtInstance.destroy(); dtInstance=null; }
    buildTable(currentData);
    showToast('Table refreshed from SharePoint', 'success');
  } catch(e) {
    showToast('Refresh failed: ' + e.message, 'error');
  } finally {
    if (btn) { btn.disabled=false; btn.innerHTML='<i class="fas fa-rotate-right"></i> Refresh'; }
  }
}
