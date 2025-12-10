// ============================================
// AR Outstanding Sheet to JSON API (Read/Write)
// Sheet ID: 15-Nz10GIXIgv_Ia2RZXADReJmVbyrjmW
// Columns (current structure):
// Client, State, Account#, Patient Name, Insurance Name, DOS, Year,
// Aging Days, Aging Bucket, S. Date, S. Aging Days, S. Aging Bucket,
// Status, Billed Amount, Balance Amount, Insurance Type, AR Comments,
// Type (Ins Call/Email/Analysis/Portal), Status Code, Action Code,
// Assigned To, Worked By, Worked Date, Follow Up Date, Service Type,
// Claim Type, Allocation Date, Allocation By
// ============================================

const API_SHEET_CONFIG = {
  id: '15-Nz10GIXIgv_Ia2RZXADReJmVbyrjmW',
  name: 'AR Outstanding',
  gid: '922894166' // tab gid from the URL
};

const API_CACHE_DURATION = 120; // seconds
const API_COLUMN_HEADERS = [
  'Client', 'State', 'Account#', 'Patient Name', 'Insurance Name',
  'DOS', 'Year', 'Aging Days', 'Aging Bucket', 'S. Date',
  'S. Aging Days', 'S. Aging Bucket', 'Status', 'Billed Amount',
  'Balance Amount', 'Insurance Type', 'AR Comments',
  'Type (Ins Call/Email/Analysis/Portal)', 'Status Code',
  'Action Code', 'Assigned To', 'Worked By', 'Worked Date',
  'Follow Up Date', 'Service Type', 'Claim Type',
  'Allocation Date', 'Allocation By'
];

// ============== GET ==============
function doGet(e) {
  const sheetId = (e && e.parameter && e.parameter.sheetId) || API_SHEET_CONFIG.id;
  const sheetGid = (e && e.parameter && e.parameter.gid) || API_SHEET_CONFIG.gid;

  const cacheKey = `ar_api_${sheetId}_${sheetGid}`;
  const cache = CacheService.getScriptCache();

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }

    const jsonData = readAndCacheSheet(sheetId, sheetGid, cacheKey, cache);
    return ContentService.createTextOutput(jsonData)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============== POST (WRITE) ==============
// Expects JSON body:
// {
//   "sheetId": "optional sheet id (defaults)",
//   "gid": "optional sheet gid (defaults)",
//   "data": { <field:value pairs> }
// }
// Upserts by Account# (unique key). If Account# missing, appends new row.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const payload = JSON.parse(e.postData.contents || '{}');
    const sheetId = payload.sheetId || API_SHEET_CONFIG.id;
    const sheetGid = payload.gid || API_SHEET_CONFIG.gid;
    const rowData = payload.data;
    if (!rowData) throw new Error('Missing data field in payload');

    const cacheKey = `ar_api_${sheetId}_${sheetGid}`;
    const sheet = getSheetByIdOrGid(sheetId, sheetGid, API_SHEET_CONFIG.name);
    if (!sheet) throw new Error('Sheet not found: gid=' + sheetGid + ' name=' + API_SHEET_CONFIG.name);

    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];

    // Map incoming data to row order based on headers
    const record = headers.map(h => rowData.hasOwnProperty(h) ? rowData[h] : '');

    // Find existing row by Account#
    const accountKey = rowData['Account#'];
    let rowIndex = -1;
    if (accountKey) {
      const accountColIndex = headers.indexOf('Account#');
      if (accountColIndex >= 0) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][accountColIndex]) === String(accountKey)) {
            rowIndex = i + 1; // 1-based
            break;
          }
        }
      }
    }

    if (rowIndex === -1) {
      sheet.appendRow(record);
    } else {
      sheet.getRange(rowIndex, 1, 1, record.length).setValues([record]);
    }

    // Refresh cache
    const cache = CacheService.getScriptCache();
    readAndCacheSheet(sheetId, sheetGid, cacheKey, cache);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Data upserted successfully',
      cached: true
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============== Helpers ==============
function readAndCacheSheet(sheetId, sheetGid, cacheKey, cache) {
  const sheet = getSheetByIdOrGid(sheetId, sheetGid, API_SHEET_CONFIG.name);
  if (!sheet) throw new Error('Sheet not found: gid=' + sheetGid + ' name=' + API_SHEET_CONFIG.name);

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    const empty = JSON.stringify({ success: true, sheetId, recordCount: 0, data: [] });
    cache.put(cacheKey, empty, API_CACHE_DURATION);
    return empty;
  }

  const headers = data[0];
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const record = {};
    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = data[i][j];
    }
    records.push(record);
  }

  const jsonData = JSON.stringify({
    success: true,
    sheetId: sheetId,
    gid: sheetGid,
    recordCount: records.length,
    timestamp: new Date().toISOString(),
    data: records
  });

  cache.put(cacheKey, jsonData, API_CACHE_DURATION);
  return jsonData;
}

function clearApiCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(`ar_api_${API_SHEET_CONFIG.id}`);
}

function warmUpApiCache() {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ar_api_${API_SHEET_CONFIG.id}_${API_SHEET_CONFIG.gid}`;
  readAndCacheSheet(API_SHEET_CONFIG.id, API_SHEET_CONFIG.gid, cacheKey, cache);
}

// Get sheet by gid (preferred) or by name fallback
function getSheetByIdOrGid(sheetId, gid, nameFallback) {
  const ss = SpreadsheetApp.openById(sheetId);
  if (gid) {
    const gidNum = Number(gid);
    const sheetByGid = ss.getSheets().find(s => s.getSheetId() === gidNum);
    if (sheetByGid) return sheetByGid;
  }
  return ss.getSheetByName(nameFallback);
}

