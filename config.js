/* ============================================================
   config.js  — YOUR SHAREPOINT CONFIGURATION
   ============================================================
   INSTRUCTIONS:
   Fill in the values below with your own SharePoint details.
   You collected these in Phase 1 Step 5 of the setup guide.

   DO NOT share this file publicly if your data is sensitive.
   ============================================================ */

const SP_CONFIG = {

  /* ── 1. Your Microsoft Tenant Name ────────────────────────
     This is the part before .sharepoint.com in your URL.
     Example: if your SharePoint URL is
       https://attglobal.sharepoint.com
     then your tenant is:
       attglobal
  ──────────────────────────────────────────────────────── */
  tenant: 'accenture',           // e.g. 'attglobal'

  /* ── 2. Your SharePoint Site Name ─────────────────────────
     The name you gave your site when you created it.
     Example: SecurityVulnerabilityPortal
  ──────────────────────────────────────────────────────── */
  siteName: 'Security Vulnerability Portal',

  /* ── 3. Your SharePoint List Names ────────────────────────
     These must match EXACTLY what you named the lists.
     Case-sensitive. No spaces.
  ──────────────────────────────────────────────────────── */
  lists: {
    application : 'ApplicationVulnerabilities',
    astra       : 'ASTRAVulnerabilities'
  },

  /* ── 4. Auto-Refresh Interval ──────────────────────────────
     How often the dashboard re-fetches data from SharePoint.
     Value is in milliseconds. 300000 = 5 minutes.
     Set to 0 to disable auto-refresh.
  ──────────────────────────────────────────────────────── */
  autoRefreshMs: 300000,

  /* ── 5. Max rows to fetch per list ────────────────────────
     SharePoint Graph API paginates at 5000 rows max.
     If you have more than 5000 rows in a list, the loader
     will automatically fetch the next page too.
     You can lower this to speed up initial load.
  ──────────────────────────────────────────────────────── */
  pageSize: 2000,

  /* ── INTERNAL — do not edit below this line ─────────────── */
  get siteUrl() {
    return `https://${this.tenant}.sharepoint.com/sites/${this.siteName}`;
  },
  get graphBase() {
    return `https://graph.microsoft.com/v1.0/sites/${this.tenant}.sharepoint.com:/sites/${this.siteName}`;
  }
};
