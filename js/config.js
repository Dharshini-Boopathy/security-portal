const SP_CONFIG = {

  tenant: 'accenture',

  siteName: 'SecurityVulnerabilityPortal',

  customDomain: 'ts.accenture.com',

  lists: {
    application : 'ApplicationVulnerabilities',
    astra       : 'ASTRAVulnerabilities',
    device      : 'DeviceVulnerabilities',
    database    : 'DatabaseVulnerabilities'
  },

  autoRefreshMs: 300000,

  pageSize: 2000,

  get siteUrl() {
    return `https://${this.customDomain}/sites/${this.siteName}`;
  },

  get graphBase() {
    return `https://graph.microsoft.com/v1.0/sites/${this.customDomain}:/sites/${this.siteName}`;
  }
};
