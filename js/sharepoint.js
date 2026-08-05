'use strict';

let msalInstance = null;
let isInitializing = false;
let initPromise = null;

const MSAL_CONFIG = {
  auth: {
    clientId   : 'd3590ed6-52b3-4102-aeff-aad2292ab01c',
    authority  : 'https://login.microsoftonline.com/accenture.com',
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: true
  }
};

const GRAPH_SCOPES = ['https://graph.microsoft.com/Sites.Read.All'];

const Auth = {
  account: null,
  token:   null,
  expiry:  0
};

/* ── Init MSAL once ─────────────────────────────────────────── */
async function initMSAL() {
  if (msalInstance) return msalInstance;
  if (initPromise)  return initPromise;

  initPromise = (async () => {
    await waitForLib('msal');

    // Clear any stuck MSAL interaction state before initializing
    const keysToRemove = [];
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      if (key && (
        key.includes('msal') ||
        key.includes('login') ||
        key.includes('interaction')
      )) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(k => sessionStorage.removeItem(k));

    msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

    try {
      const response = await msalInstance.handleRedirectPromise();
      if (response) {
        Auth.account = response.account;
        Auth.token   = response.accessToken;
        Auth.expiry  = response.expiresOn?.getTime() || 0;
      }
    } catch(e) {
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

/* ── Get Token ──────────────────────────────────────────────── */
async function getToken() {
  if (Auth.token && Date.now() < Auth.expiry - 120000) return Auth.token;

  const msal = await initMSAL();

  // Try silent first
  if (Auth.account) {
    try {
      const result = await msal.acquireTokenSilent({
        scopes : GRAPH_SCOPES,
        account: Auth.account
      });
      Auth.token  = result.accessToken;
      Auth.expiry = result.expiresOn?.getTime() || (Date.now() + 3600000);
      return Auth.token;
    } catch(e) {
      console.info('Silent token failed:', e.message);
    }
  }

  // Check no interaction already in progress
  const interactionStatus = sessionStorage.getItem('msal.interaction.status');
  if (interactionStatus) {
    sessionStorage.removeItem('msal.interaction.status');
  }

  // Popup sign-in
  try {
    const result = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    Auth.account = result.account;
    Auth.token   = result.accessToken;
    Auth.expiry  = result.expiresOn?.getTime() || (Date.now() + 3600000);
    return Auth.token;
  } catch(e) {
    throw new Error('Sign-in cancelled or failed: ' + e.message);
  }
}

/* ── Graph Fetch ────────────────────────────────────────────── */
async function graphFetch(url) {
  const token = await getToken();
  const resp  = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type' : 'application/json'
    }
  });

  if (resp.status === 401) {
    Auth.token = null;
    return graphFetch(url);
  }

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(`Graph ${resp.status}: ${err?.error?.message || resp.statusText}`);
  }

  return resp.json();
}

/* ── Fetch
