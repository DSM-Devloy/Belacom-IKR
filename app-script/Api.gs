/**
 * Api.gs - tambahkan sebagai file baru di project Apps Script.
 * Setup : Project Settings > Script Properties > tambah  API_SECRET = (sama dengan GAS_SECRET di Vercel)
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */
const API_ALLOWED_ = [
  'checkUserLogin', 'getUsersList', 'saveUser', 'deleteUser', 'updatePassword',
  'getDashboardData', 'getUniqueTeamsList', 'getBillingDashboardData', 'getWorkOrders',
  'updatePaymentStatus', 'getTeamsList', 'getReportIKRList', 'saveWorkOrder',
  'uploadImageToDrive', 'saveReportIKRData', 'getReportDetail', 'generateWATemplate',
  'generatePDFReport', 'getTeamStatusPerformance', 'getTeamsManagementList',
  'saveTeamData', 'deleteTeamRow', 'exportWorkOrdersToExcel', 'exportReportIKRToExcel',
  'deleteWorkOrder',
  'getPriceSettings', 'savePriceSetting', 'deletePriceSetting', 'getPriceImpact'
];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('API_SECRET');
    if (!secret || req.secret !== secret) return json_({ ok: false, error: 'Unauthorized' });
    if (API_ALLOWED_.indexOf(req.fn) < 0) return json_({ ok: false, error: 'Fungsi tidak diizinkan' });
    const data = globalThis[req.fn].apply(null, req.args || []);
    return json_({ ok: true, data: data === undefined ? null : data });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
