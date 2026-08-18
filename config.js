/* ---------------------------------------------------------------------------
 * Fill these three values in after the one-time Azure setup (see README).
 * Nothing here is a secret: a browser app has no client secret, and the
 * redirect URI registered in Azure is what stops anyone else using this ID.
 * ------------------------------------------------------------------------- */
window.CONFIG = {

  // Application (client) ID from the Azure app registration.
  clientId: "PASTE_CLIENT_ID_HERE",

  // Directory (tenant) ID from the same page. "organizations" also works, but
  // pinning the tenant keeps the sign-in page from asking which account type.
  tenantId: "PASTE_TENANT_ID_HERE",

  // The workbook's URL. In OneDrive/SharePoint, open the file, then copy the
  // address bar URL (or use Share > Copy link). Either form works.
  workbookUrl: "PASTE_WORKBOOK_URL_HERE",

  // Table inside the workbook that receives one row per day worked.
  tableName: "HoursLog",

  // Round each day's hours to this many decimals before it reaches Excel.
  decimals: 2
};
