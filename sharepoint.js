/* ============================================================
   sharepoint.js — Microsoft Graph API Data Layer
   Handles authentication, fetching, pagination, caching,
   and field-mapping from SharePoint List columns to the
   same normalized shape the dashboard already expects.
   ============================================================ */

'use strict';

/* ── Auth State ─────────────────────────────────────────────── */
const Auth = {
  account: null,        // signed-in MSAL account object
  token:   null,        // current access token string
  expiry:  0            // token expiry timestamp
};

/* ── MSAL (Microsoft Authentication Library) instance ────────
   MSAL handles the sign-in popup and token management.
   We use the CDN version loaded in index.html.
   The clientId below is the PUBLIC Graph Explorer app — it
   works for any Microsoft 365 tenant without IT registering
   anything. For production, replace with your own app ID.
──────────────────────────────────────────────────────────── */
let msalInstance = null;

const MSAL_CONFIG = {
  auth: {
    // Microsoft's own "Graph Explorer" public client ID.
    // Works for any org. No Azure app registration needed.
    clientId   : 'd3590ed6-52b3-4102-aeff-aad2292ab01c',
    authority  : 'https://login.microsoftonline.com/common',
    redirectUri: window.location.href.split('?')[0]   // current page URL
  },
  cache: {
    cacheLocation      : 'localStorage',
    storeAuthStateInCookie: false
  }
};

const GRAPH_SCOPES = ['https://graph.microsoft.com/Sites.Read.All'];

/* ── Initialise MSAL ────────────────────────────────────────── */
async function initMSAL() {
  if (msalInstance) return msalInstance;

  // Wait for MSAL library to be available (loaded via CDN script tag)
  await waitForLib('msal');
  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

  // Handle redirect response (in case of redirect-flow fallback)
  try {
    await msalInstance.handleRedirectPromise();
  } catch(e) {
    console.warn('MSAL redirect handle:', e.message);
  }

  // Check if a user is already signed in from a previous session
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    Auth.account = accounts[0];
  }

  return msalInstance;
}

/* ── Get Access Token ───────────────────────────────────────── */
async function getToken() {
  // Return cached token if still valid (with 2-min buffer)
  if (Auth.token && Date.now() < Auth.expiry - 120000) return Auth.token;

  await initMSAL();

  // Try silent token acquisition first (uses cached session)
  if (Auth.account) {
    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes : GRAPH_SCOPES,
        account: Auth.account
      });
      Auth.token  = result.accessToken;
      Auth.expiry = result.expiresOn.getTime();
      return Auth.token;
    } catch(silentErr) {
      // Silent failed — fall through to popup
      console.info('Silent token failed, trying popup:', silentErr.message);
    }
  }

  // Show sign-in popup
  try {
    const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    Auth.account = result.account;
    Auth.token   = result.accessToken;
    Auth.expiry  = result.expiresOn.getTime();
    return Auth.token;
  } catch(popupErr) {
    throw new Error('Sign-in cancelled or failed: ' + popupErr.message);
  }
}

/* ── Core Graph API Fetch ───────────────────────────────────── */
async function graphFetch(url) {
  const token = await getToken();
  const resp  = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type' : 'application/json'
    }
  });

  if (resp.status === 401) {
    // Token may have expired — clear and retry once
    Auth.token = null;
    return graphFetch(url);
  }

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(`Graph API ${resp.status}: ${err?.error?.message || resp.statusText}`);
  }

  return resp.json();
}

/* ── Fetch All Pages of a SharePoint List ───────────────────── */
async function fetchListItems(listName, selectFields = null) {
  const base = SP_CONFIG.graphBase;

  // First, resolve the list ID by name
  const listsData = await graphFetch(`${base}/lists?$select=id,name,displayName`);
  const listMeta  = listsData.value.find(
    l => l.name === listName || l.displayName === listName
  );

  if (!listMeta) {
    console.warn(`List "${listName}" not found on SharePoint site.`);
    return [];
  }

  // Build fields query
  const fieldsParam = selectFields
    ? `&$expand=fields($select=${selectFields.join(',')})`
    : '&$expand=fields';

  let url = `${base}/lists/${listMeta.id}/items?$top=${SP_CONFIG.pageSize}${fieldsParam}`;
  const allItems = [];

  // Paginate through all results (SharePoint returns @odata.nextLink when more pages exist)
  while (url) {
    const page = await graphFetch(url);
    const rows = (page.value || []).map(item => item.fields || {});
    allItems.push(...rows);
    url = page['@odata.nextLink'] || null;
  }

  return allItems;
}

/* ── Field Mappings: SharePoint columns → Dashboard shape ─────
   These map YOUR SharePoint column internal names to the
   normalized fields the dashboard charts and tables expect.
   The internal name is the column name with spaces removed.
──────────────────────────────────────────────────────────── */
const SPFieldMap = {

  application: row => ({
    id         : row.VITNumber        || row.Title || '',
    name       : row.VulnerabilityName || '',
    app        : row.ApplicationAcronym || '',
    severity   : row.Severity          || 'Medium',
    status     : row.State             || 'Open',
    slaStatus  : row.SLAStatus         || '',
    score      : parseFloat(row.RiskScore) || 0,
    dueDate    : row.ActionByDate   ? new Date(row.ActionByDate)   : null,
    created    : row.Created        ? new Date(row.Created)        : null,
    desc       : row.VulnerabilitySummary || '',
    rec        : row.Recommendation       || '',
    category   : row.Category             || '',
    portfolio  : row.ApplicationPortfolio || '',
    family     : row.ApplicationFamily    || '',
    cve        : row.CVENumber            || '',
    tool       : row.AssessmentTool       || '',
    location   : row.Location             || '',
    firstFound : row.FirstFound     ? new Date(row.FirstFound)     : null,
    aging      : row.Aging                || '',
    _raw       : row
  }),

  astra: row => ({
    id         : row.Title || row.id || '',
    name       : row.VulnDescription  || '',
    app        : row.ApplicationAcronym || '',
    severity   : row.RiskSeverity      || 'Medium',
    status     : row.SLA               || 'Within SLA',
    slaStatus  : row.SLA               || '',
    score      : 0,
    dueDate    : row.SLADate    ? new Date(row.SLADate)    : null,
    created    : row.Created    ? new Date(row.Created)    : null,
    desc       : row.VulnDescription   || '',
    rec        : '',
    category   : row.Category          || '',
    portfolio  : row.ApplicationPortfolio || '',
    family     : '',
    cve        : '',
    hostname   : row.Hostname          || '',
    aging      : row.Aging             || '',
    contract   : row.Contract          || '',
    sla        : row.SLA               || '',
    _raw       : row
  }),

  device: row => ({
    id         : row.VITNumber || row.Title || '',
    name       : row.VulnerabilityName || '',
    app        : row.DeviceCI   || '',
    severity   : normSeverity(row.RiskRating),
    status     : row.State      || 'Open',
    slaStatus  : row.TTRStatus  || '',
    score      : parseFloat(row.RiskScore) || 0,
    dueDate    : row.TTRTargetDate ? new Date(row.TTRTargetDate) : null,
    created    : row.Created       ? new Date(row.Created)       : null,
    desc       : row.VulnerabilitySummary || '',
    rec        : row.Recommendation       || '',
    cve        : row.CVEs               || '',
    hostname   : row.Hostname           || '',
    ip         : row.IPAddress          || '',
    env        : row.Environment        || '',
    team       : row.AssignmentGroup    || '',
    firstFound : null,
    lastFound  : null,
    _raw       : row
  }),

  database: row => ({
    id         : row.VITNumber || row.Title || '',
    name       : row.VulnerabilityName || '',
    app        : row.ApplicationAcronym || '',
    severity   : row.Severity           || 'Medium',
    status     : row.State              || 'Open',
    slaStatus  : row.TTRStatus          || '',
    score      : 0,
    dueDate    : row.ActionByDate  ? new Date(row.ActionByDate)  : null,
    created    : row.Created       ? new Date(row.Created)       : null,
    desc       : row.VulnerabilitySummary || '',
    rec        : '',
    cve        : '',
    hostname   : row.Hostname     || '',
    env        : '',
    aging      : row.AgingData    || '',
    portfolio  : row.AppPortfolio || '',
    poc        : row.DBPOC        || '',
    _raw       : row
  })
};

/* ── Main Live Data Loader ──────────────────────────────────── */
async function loadAllDataLive(forceRefresh = false) {

  // Return cache if fresh and not forcing refresh
  const CACHE_KEY  = 'sp_vuln_cache';
  const CACHE_TIME = 'sp_vuln_cache_time';
  const maxAge     = SP_CONFIG.autoRefreshMs || 300000;

  if (!forceRefresh) {
    try {
      const cached   = sessionStorage.getItem(CACHE_KEY);
      const cachedAt = parseInt(sessionStorage.getItem(CACHE_TIME) || '0');
      if (cached && Date.now() - cachedAt < maxAge) {
        const parsed = JSON.parse(cached);
        Object.assign(VulnData, parsed, { loaded: true });
        console.log('Loaded from session cache');
        return VulnData;
      }
    } catch(e) { /* ignore cache errors */ }
  }

  updateLoadingStatus('Connecting to SharePoint...');

  try {
    // Fetch all four lists in parallel
    const [appRaw, astraRaw, deviceRaw, dbRaw] = await Promise.all([
      fetchListItems(SP_CONFIG.lists.application),
      fetchListItems(SP_CONFIG.lists.astra),
      fetchListItems(SP_CONFIG.lists.device),
      fetchListItems(SP_CONFIG.lists.database)
    ]);

    updateLoadingStatus('Processing vulnerability data...');

    // Map raw SharePoint rows to normalized dashboard objects
    VulnData.application = appRaw.map(SPFieldMap.application);
    VulnData.astra       = astraRaw.map(SPFieldMap.astra);
    VulnData.device      = deviceRaw.map(SPFieldMap.device);
    VulnData.database    = dbRaw.map(SPFieldMap.database);
    VulnData.loaded      = true;
    VulnData.lastFetched = new Date();

    // Cache in sessionStorage
    try {
      const toCache = {
        application: VulnData.application,
        astra      : VulnData.astra,
        device     : VulnData.device,
        database   : VulnData.database
      };
      sessionStorage.setItem(CACHE_KEY,  JSON.stringify(toCache));
      sessionStorage.setItem(CACHE_TIME, Date.now().toString());
    } catch(e) { /* sessionStorage quota exceeded — skip cache */ }

    console.log(
      `Live data loaded — ` +
      `App:${VulnData.application.length} ` +
      `ASTRA:${VulnData.astra.length} ` +
      `Device:${VulnData.device.length} ` +
      `DB:${VulnData.database.length}`
    );

    return VulnData;

  } catch(err) {
    VulnData.loadError = err.message;
    console.error('SharePoint load error:', err);
    throw err;
  }
}

/* ── Auto-Refresh Setup ─────────────────────────────────────── */
function startAutoRefresh(onRefreshCallback) {
  if (!SP_CONFIG.autoRefreshMs) return;

  setInterval(async () => {
    try {
      await loadAllDataLive(true);   // force refresh bypasses cache
      if (typeof onRefreshCallback === 'function') onRefreshCallback();
      showToast('Dashboard data refreshed from SharePoint', 'success');
      updateLastRefreshTime();
    } catch(e) {
      showToast('Auto-refresh failed — will retry', 'warning');
    }
  }, SP_CONFIG.autoRefreshMs);
}

/* ── UI Helpers ─────────────────────────────────────────────── */
function updateLoadingStatus(msg) {
  const el = document.querySelector('.loader-text');
  if (el) el.textContent = msg;
}

function updateLastRefreshTime() {
  const el = document.getElementById('lastRefresh');
  if (el) el.textContent = new Date().toLocaleTimeString('en-US', {
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  });
}

function showSignInBanner() {
  const banner = document.getElementById('signInBanner');
  if (banner) banner.style.display = 'flex';
}

function hideSignInBanner() {
  const banner = document.getElementById('signInBanner');
  if (banner) banner.style.display = 'none';
}

/* ── Wait for Library Helper ────────────────────────────────── */
function waitForLib(globalName, maxWait = 10000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      if (window[globalName]) return resolve(window[globalName]);
      if (Date.now() - start > maxWait) return reject(new Error(`${globalName} did not load`));
      setTimeout(check, 100);
    };
    check();
  });
}

/* ── Sign Out ───────────────────────────────────────────────── */
async function signOut() {
  await initMSAL();
  sessionStorage.removeItem('sp_vuln_cache');
  sessionStorage.removeItem('sp_vuln_cache_time');
  if (Auth.account) {
    await msalInstance.logoutPopup({ account: Auth.account });
  }
}
