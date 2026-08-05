/* ============================================================
   sharepoint.js — Fixed version
   Clears stuck MSAL state automatically on every load
   ============================================================ */

'use strict';

let msalInstance = null;
let initPromise  = null;

const MSAL_CONFIG = {
  auth: {
    clientId   : 'd3590ed6-52b3-4102-aeff-aad2292ab01c',
    authority  : 'https://login.microsoftonline.com/accenture.com',
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: {
    cacheLocation         : 'sessionStorage',
    storeAuthStateInCookie: true
  }
};

const GRAPH_SCOPES = ['https://graph.microsoft.com/Sites.Read.All'];

const Auth = {
  account: null,
  token  : null,
  expiry : 0
};

function clearMSALStuckState() {
  const keysToRemove = [];
  for (let i = 0; i < sessionStorage.length; i++) {
    const key = sessionStorage.key(i);
    if (key && (key.includes('msal') || key.includes('interaction') || key.includes('login.windows'))) {
      keysToRemove.push(key);
    }
  }
  keysToRemove.forEach(k => sessionStorage.removeItem(k));
}

function waitForLib(globalName, maxWait = 10000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      if (window[globalName]) return resolve(window[globalName]);
      if (Date.now() - start > maxWait) return reject(new Error(globalName + ' did not load'));
      setTimeout(check, 100);
    };
    check();
  });
}

async function initMSAL() {
  if (msalInstance) return msalInstance;
  if (initPromise)  return initPromise;

  initPromise = (async () => {
    clearMSALStuckState();
    await waitForLib('msal');
    msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    try {
      const response = await msalInstance.handleRedirectPromise();
      if (response && response.account) {
        Auth.account = response.account;
        Auth.token   = response.accessToken;
        Auth.expiry  = response.expiresOn ? response.expiresOn.getTime() : Date.now() + 3600000;
      }
    } catch (e) {
      console.warn('Redirect handle:', e.message);
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0 && !Auth.account) {
      Auth.account = accounts[0];
    }
    return msalInstance;
  })();

  return initPromise;
}

async function getToken() {
  if (Auth.token && Date.now() < Auth.expiry - 120000) return Auth.token;
  const msalApp = await initMSAL();
  if (Auth.account) {
    try {
      const result = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: Auth.account });
      Auth.token  = result.accessToken;
      Auth.expiry = result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600000;
      return Auth.token;
    } catch (e) {
      console.info('Silent failed:', e.message);
    }
  }
  clearMSALStuckState();
  try {
    const result = await msalApp.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    Auth.account = result.account;
    Auth.token   = result.accessToken;
    Auth.expiry  = result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600000;
    return Auth.token;
  } catch (e) {
    throw new Error('Sign-in failed: ' + e.message);
  }
}

async function graphFetch(url) {
  const token = await getToken();
  const resp  = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }
  });
  if (resp.status === 401) { Auth.token = null; return graphFetch(url); }
  if (!resp.ok) {
    const e = await resp.json().catch(() => ({}));
    throw new Error('Graph ' + resp.status + ': ' + (e.error ? e.error.message : resp.statusText));
  }
  return resp.json();
}

async function fetchListItems(listName) {
  const base      = SP_CONFIG.graphBase;
  const listsData = await graphFetch(base + '/lists?$select=id,name,displayName');
  const listMeta  = listsData.value.find(l => l.name === listName || l.displayName === listName);
  if (!listMeta) { console.warn('List not found: ' + listName); return []; }
  let url      = base + '/lists/' + listMeta.id + '/items?$top=' + SP_CONFIG.pageSize + '&$expand=fields';
  let allItems = [];
  while (url) {
    const page = await graphFetch(url);
    allItems   = allItems.concat((page.value || []).map(item => item.fields || {}));
    url        = page['@odata.nextLink'] || null;
  }
  return allItems;
}

const SPFieldMap = {
  application: r => ({
    id: r.VITNumber||r.Title||'', name: r.VulnerabilityName||'', app: r.ApplicationAcronym||'',
    severity: r.Severity||'Medium', status: r.State||'Open', slaStatus: r.SLAStatus||'',
    score: parseFloat(r.RiskScore)||0, dueDate: r.ActionByDate?new Date(r.ActionByDate):null,
    created: r.Created?new Date(r.Created):null, desc: r.VulnerabilitySummary||'',
    rec: r.Recommendation||'', category: r.Category||'', portfolio: r.ApplicationPortfolio||'',
    family: r.ApplicationFamily||'', cve: r.CVENumber||'', tool: r.AssessmentTool||'',
    location: r.Location||'', firstFound: r.FirstFound?new Date(r.FirstFound):null,
    aging: r.Aging||'', _raw: r
  }),
  astra: r => ({
    id: r.Title||'', name: r.VulnDescription||'', app: r.ApplicationAcronym||'',
    severity: r.RiskSeverity||'Medium', status: r.SLA||'Within SLA', slaStatus: r.SLA||'',
    score: 0, dueDate: r.SLADate?new Date(r.SLADate):null, created: r.Created?new Date(r.Created):null,
    desc: r.VulnDescription||'', rec: '', category: r.Category||'', portfolio: r.ApplicationPortfolio||'',
    family: '', cve: '', hostname: r.Hostname||'', aging: r.Aging||'', contract: r.Contract||'',
    sla: r.SLA||'', _raw: r
  }),
  device: r => ({
    id: r.VITNumber||r.Title||'', name: r.VulnerabilityName||'', app: r.DeviceCI||'',
    severity: normSeverity(r.RiskRating), status: r.State||'Open', slaStatus: r.TTRStatus||'',
    score: parseFloat(r.RiskScore)||0, dueDate: r.TTRTargetDate?new Date(r.TTRTargetDate):null,
    created: r.Created?new Date(r.Created):null, desc: r.VulnerabilitySummary||'',
    rec: r.Recommendation||'', cve: r.CVEs||'', hostname: r.Hostname||'', ip: r.IPAddress||'',
    env: r.Environment||'', team: r.AssignmentGroup||'', firstFound: null, lastFound: null, _raw: r
  }),
  database: r => ({
    id: r.VITNumber||r.Title||'', name: r.VulnerabilityName||'', app: r.ApplicationAcronym||'',
    severity: r.Severity||'Medium', status: r.State||'Open', slaStatus: r.TTRStatus||'',
    score: 0, dueDate: r.ActionByDate?new Date(r.ActionByDate):null,
    created: r.Created?new Date(r.Created):null, desc: r.VulnerabilitySummary||'',
    rec: '', cve: '', hostname: r.Hostname||'', env: '', aging: r.AgingData||'',
    portfolio: r.AppPortfolio||'', poc: r.DBPOC||'', _raw: r
  })
};

async function loadAllDataLive(forceRefresh) {
  forceRefresh    = forceRefresh || false;
  const CACHE_KEY  = 'sp_vuln_cache';
  const CACHE_TIME = 'sp_vuln_cache_time';
  const maxAge     = SP_CONFIG.autoRefreshMs || 300000;

  if (!forceRefresh) {
    try {
      const cached   = sessionStorage.getItem(CACHE_KEY);
      const cachedAt = parseInt(sessionStorage.getItem(CACHE_TIME) || '0');
      if (cached && (Date.now() - cachedAt) < maxAge) {
        Object.assign(VulnData, JSON.parse(cached), { loaded: true });
        return VulnData;
      }
    } catch(e) {}
  }

  updateLoadingStatus('Signing in to Microsoft...');

  const [appRaw, astraRaw, deviceRaw, dbRaw] = await Promise.all([
    fetchListItems(SP_CONFIG.lists.application),
    fetchListItems(SP_CONFIG.lists.astra),
    fetchListItems(SP_CONFIG.lists.device),
    fetchListItems(SP_CONFIG.lists.database)
  ]);

  VulnData.application = appRaw.map(SPFieldMap.application);
  VulnData.astra       = astraRaw.map(SPFieldMap.astra);
  VulnData.device      = deviceRaw.map(SPFieldMap.device);
  VulnData.database    = dbRaw.map(SPFieldMap.database);
  VulnData.loaded      = true;
  VulnData.lastFetched = new Date();

  try {
    sessionStorage.setItem(CACHE_KEY, JSON.stringify({
      application: VulnData.application, astra: VulnData.astra,
      device: VulnData.device, database: VulnData.database
    }));
    sessionStorage.setItem(CACHE_TIME, Date.now().toString());
  } catch(e) {}

  return VulnData;
}

function startAutoRefresh(onRefreshCallback) {
  if (!SP_CONFIG.autoRefreshMs) return;
  setInterval(async () => {
    try {
      await loadAllDataLive(true);
      if (typeof onRefreshCallback === 'function') onRefreshCallback();
      showToast('Dashboard refreshed from SharePoint', 'success');
      updateLastRefreshTime();
    } catch(e) { console.warn('Auto-refresh failed:', e.message); }
  }, SP_CONFIG.autoRefreshMs);
}

function updateLoadingStatus(msg) {
  const el = document.querySelector('.loader-text');
  if (el) el.textContent = msg;
}

function updateLastRefreshTime() {
  const el = document.getElementById('lastRefresh');
  if (el) el.textContent = new Date().toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit', second:'2-digit' });
}

async function signOut() {
  sessionStorage.clear();
  if (msalInstance && Auth.account) {
    try { await msalInstance.logoutPopup({ account: Auth.account }); } catch(e) {}
  }
  location.reload();
}
