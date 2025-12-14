// ============================================
// GOOGLE APPS SCRIPT - BACKEND (Code.gs)
// AR Connect - LAH v3.0 Advanced Edition
// ============================================

const CONFIG = {
  CMS_CONTROL_UNI: {
    sheetId: '1DFFkBfPAwtq-n4LWD96SYbQUCIDM2TZsIp0cpcLKZjg',
    sheets: {
      dropdown: { name: 'AR oustanding Dropdown' },
      providerInfo: { name: 'Provider Info' },
      patientMaster: { name: 'Patient Master Sheet' },
      credentialing: { name: 'Credentialing Info' },
      arClientConfig: { name: 'AR Client Config' },
      arConnectLog: { name: 'AR Connect Log' }
    }
  },
  LAH_CONSOLIDATED: {
    sheetId: '1S8gn2ThGW5F5nQS1G_XSaSuLPJmJLw5A-qQUcHp-pqg',
    sheetName: 'Tracker log',
    enabled: true
  },
  DEVELOPER: { 
    username: 'Blake Dawson', 
    password: 'Root',
    allowedEmail: 'abuthahir.dataset@gmail.com'
  },
};

// Column mapping based on actual sheet structure:
// Client(0), State(1), VisitID#(2), Patient Name(3), DOS(4), Aging Days(5), Aging Bucket(6), 
// Submitted Date(7), Insurance Name(8), Status(9), Billed Amount(10), Balance Amount(11), 
// Primary/Secondary(12), AR notes(13), Type(14), Status Code(15), Action Code(16), 
// Assigned To(17), Worked By(18), Worked Date(19), Follow up date(20), 
// Claim Type(21), Allocated User(22), Allocation Date(23), Remarks(24)
//const COLUMNS = {
//  CLIENT: 1, STATE: 2, VISIT_ID: 3, PATIENT_NAME: 4, DOS: 6,
//  AGING_DAYS: 8, AGING_BUCKET: 9, SUBMITTED_DATE: 10, INSURANCE_NAME: 5,
//  STATUS: 12, BILLED_AMOUNT: 13, BALANCE_AMOUNT: 14, PRIMARY_SECONDARY: 15,
//  AR_NOTES: 16, TYPE: 17, STATUS_CODE: 18, ACTION_CODE: 19,
//  ASSIGNED_TO: 20, WORKED_BY: 21, WORKED_DATE: 22, FOLLOWUP_DATE: 23,
//  CLAIM_TYPE: 24, ALLOCATED_USER: 25, ALLOCATION_DATE: 26, REMARKS: 27
//};
const COLUMNS = {
  CLIENT: 1, STATE: 2, VISIT_ID: 3, PATIENT_NAME: 4, DOS: 6,
  AGING_DAYS: 8, AGING_BUCKET: 9, SUBMITTED_DATE: 10, INSURANCE_NAME: 5,
  STATUS: 13, BILLED_AMOUNT: 14, BALANCE_AMOUNT: 15, TYPE: 16,
  AR_NOTES: 17, TYPE_FIELD: 18, STATUS_CODE: 19, ACTION_CODE: 20,
  ASSIGNED_TO: 21, WORKED_BY: 22, WORKED_DATE: 23, FOLLOWUP_DATE: 24,
  CLAIM_TYPE: 25, ALLOCATED_USER: 26, ALLOCATION_DATE: 27, REMARKS: 28
};

const CACHE = CacheService.getScriptCache();
const CACHE_DURATION = 1200; // 20 minutes
const JSON_CACHE_DURATION = 60; // 1 minute for JSON data
const PROPERTIES = PropertiesService.getScriptProperties();
const APP_VERSION = '3.1.0'; // Track version for sessions

// Constants for PropertiesService storage
const ACCOUNTS_DATA_KEY = 'api_accounts_json_data';
const ACCOUNTS_META_KEY = 'api_accounts_meta';
const MAX_PROPERTY_SIZE = 9000;

function getCached(key) {
  try {
    // First try CacheService (fast)
    const cached = CACHE.get(key);
    if (cached) {
      return JSON.parse(cached);
    }
    
    // For JSON data, check PropertiesService
    if (key.startsWith('json_') || key === 'all_clients_json' || key === 'api_all_accounts') {
      const propData = getJSONFromProperties(key);
      if (propData) {
        // Also cache it for faster access
        setCache(key, propData, CACHE_DURATION);
        return propData;
      }
      
      // For api_all_accounts, also check local PropertiesService
      if (key === 'api_all_accounts') {
        const localData = getAccountsFromPropertiesLocal();
        if (localData) {
          setCache(key, localData, CACHE_DURATION);
          return localData;
        }
      }
    }
    
    return null;
  } catch(e) {
    Logger.log('Cache read error: ' + e);
    return null;
  }
}

function setCache(key, data, duration = CACHE_DURATION) {
  try { 
    CACHE.put(key, JSON.stringify(data), duration); 
  } catch(e) { 
    Logger.log('Cache write error: ' + e); 
  }
}

// Background refresh - runs every 10 minutes
function refreshAccountsData() {
  try {
    const start = Date.now();
    getAllAccountsData(true); // Force refresh - this stores in both CacheService and PropertiesService
    const duration = Date.now() - start;
    logActivity('BACKGROUND_REFRESH', `Accounts refreshed in ${duration}ms`, 'SUCCESS');
  } catch(e) {
    Logger.log('Error refreshing accounts data: ' + e.toString());
    logActivity('BACKGROUND_REFRESH', e.toString(), 'ERROR');
  }
}

// Initialize cache on first load (call this once after deployment)
function initializeCache() {
  try {
    Logger.log('Initializing cache...');
    getAllAccountsData(true);
    Logger.log('Cache initialization completed');
    return { success: true, message: 'Cache initialized successfully' };
  } catch(e) {
    Logger.log('Error initializing cache: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Setup time-based trigger for 10-minute refresh
function setupAutoRefreshTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const handler = trigger.getHandlerFunction();
      if (handler === 'refreshAccountsData' || handler === 'refreshJSONCache') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create 10-minute trigger for accounts data
    ScriptApp.newTrigger('refreshAccountsData')
      .timeBased()
      .everyMinutes(10)
      .create();
    
    Logger.log('Auto-refresh trigger setup completed (10 minutes)');
    return { success: true, message: 'Trigger setup completed - data will refresh every 10 minutes' };
  } catch(e) {
    Logger.log('Error setting up auto-refresh trigger: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// JSON-BASED CMS CONTROL UNIT ACCESS
// ============================================

function getCMSDataAsJSON(sheetName) {
  try {
    const cacheKey = `cms_json_${sheetName}`;
    const cached = getCached(cacheKey);
    if (cached) return cached;
    
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // Try to find by partial match
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => 
        s.getName().toLowerCase().includes(sheetName.toLowerCase())
      );
    }
    
    if (!sheet) return { error: 'Sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { data: [] };
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const jsonData = rows.map(row => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx] || '';
      });
      return obj;
    });
    
    setCache(cacheKey, jsonData);
    return { data: jsonData };
  } catch(e) {
    Logger.log('Error getting CMS data as JSON: ' + e.toString());
    return { error: e.toString() };
  }
}

function updateCMSDataFromJSON(sheetName, jsonData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    if (!Array.isArray(jsonData) || jsonData.length === 0) {
      return { success: false, message: 'Invalid JSON data' };
    }
    
    // Clear existing data
    sheet.clear();
    
    // Get headers from first object
    const headers = Object.keys(jsonData[0]);
    sheet.appendRow(headers);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Append data rows
    jsonData.forEach(row => {
      const rowData = headers.map(header => row[header] || '');
      sheet.appendRow(rowData);
    });
    
    SpreadsheetApp.flush();
    
    // Clear cache
    CACHE.remove(`cms_json_${sheetName}`);
    
    return { success: true, message: 'Data updated successfully' };
  } catch(e) {
    Logger.log('Error updating CMS data from JSON: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function logActivity(action, details, status = 'SUCCESS') {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.arConnectLog.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.CMS_CONTROL_UNI.sheets.arConnectLog.name);
      sheet.appendRow(['Timestamp','Email','Name','Action','Details','Status','SessionID','Duration','Error','IPAddress','UserAgent','Page','VisitId']);
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 13);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }
    
    const user = Session.getActiveUser();
    const email = user.getEmail();
    const name = getCurrentUserDisplayName();
    const timestamp = new Date();
    const sessionId = Utilities.getUuid().substring(0,8);
    
    // Get additional context (if available from frontend)
    let detailsObj = {};
    if (typeof details === 'string') {
      try {
        detailsObj = JSON.parse(details);
      } catch(e) {
        detailsObj = { message: details };
      }
    } else if (details) {
      detailsObj = details;
    }
    
    const page = detailsObj.page || 'Index';
    const visitId = detailsObj.visitId || '';
    const ipAddress = detailsObj.ipAddress || 'N/A';
    const userAgent = detailsObj.userAgent || detailsObj.userAgent || 'N/A';
    const duration = detailsObj.duration || '';
    const detailsStr = typeof details === 'string' ? (detailsObj.message || details) : JSON.stringify(detailsObj);
    
    sheet.appendRow([
      timestamp,
      email,
      name,
      action,
      detailsStr,
      status,
      sessionId,
      duration,
      status === 'ERROR' ? detailsStr : '',
      ipAddress,
      userAgent,
      page,
      visitId
    ]);
    
    // Auto-resize columns periodically
    if (sheet.getLastRow() % 100 === 0) {
      sheet.autoResizeColumns(1, 13);
    }
    
    SpreadsheetApp.flush();
  } catch(e) { 
    Logger.log('Log error: ' + e.toString());
    // Don't throw - logging should never break the app
  }
}

function doGet(e) {
  // Check if this is an API request (has action parameter or path parameter)
  const action = e.parameter.action || '';
  const path = e.parameter.path || e.pathInfo || '';
  
  // If it's an API request, handle it
  if (action || path) {
    return handleAPIRequest(e, 'GET');
  }
  
  // Otherwise, serve HTML pages
  const page = e.parameter.page || 'index';
  const user = Session.getActiveUser();
  const email = user.getEmail();
  const name = getCurrentUserDisplayName();
  
  // Log page access with enhanced details
  logActivity('PAGE_LOAD', JSON.stringify({
    page: page,
    email: email,
    name: name,
    timestamp: new Date().toISOString()
  }), 'SUCCESS');
  
  if (page === 'developer') {
    // Check if user is authorized developer
    if (email.toLowerCase() !== CONFIG.DEVELOPER.allowedEmail.toLowerCase()) {
      return HtmlService.createHtmlOutput('<html><body><h1>Access Denied</h1><p>You do not have permission to access this page.</p></body></html>')
        .setTitle('Access Denied');
    }
    return HtmlService.createHtmlOutputFromFile('Developer')
      .setTitle('AR Connect - Developer Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AR Connect - LAH_v3.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // Pass developer access flag to frontend
  const isDeveloper = email.toLowerCase() === CONFIG.DEVELOPER.allowedEmail.toLowerCase();
  htmlOutput.append(`<script>window.DEVELOPER_ACCESS = ${isDeveloper};</script>`);
  
  return htmlOutput;
}

function doPost(e) {
  // Check if this is an API request
  const action = e.parameter.action || '';
  const path = e.parameter.path || e.pathInfo || '';
  
  // If it's an API request, handle it
  if (action || path) {
    return handleAPIRequest(e, 'POST');
  }
  
  // Otherwise return error
  return createResponse({ error: 'Invalid request' }, 400);
}

// ============================================
// API REQUEST HANDLER (Merged from ARConnectAPI.gs)
// ============================================

function handleAPIRequest(e, method) {
  try {
    // No origin validation needed - all in same project
    const path = e.parameter.path || e.pathInfo || '';
    const action = e.parameter.action || '';
    
    let result;
    
    // Route requests
    if (method === 'GET') {
      if (path.includes('accounts') && path.includes(':')) {
        const visitId = path.split(':')[1];
        result = getAccountByVisitIdAPI(visitId);
      } else if (path === 'accounts' || action === 'getAccounts') {
        result = getAllAccountsAPI();
      } else if (path === 'search' || action === 'search') {
        result = searchAccountsAPI(e.parameter);
      } else if (path === 'non-workable' || action === 'getNonWorkable') {
        result = getNonWorkableAccountsAPI();
      } else if (path === 'convert' || action === 'convertSheet') {
        const sheetId = e.parameter.sheetId || '';
        const sheetName = e.parameter.sheetName || '';
        result = convertSheetToJSON(sheetId, sheetName);
      } else if (path === 'convert-all' || action === 'convertAllSheets') {
        result = convertAllClientSheetsToJSON();
      } else if (path === 'get-json' || action === 'getStoredJSON') {
        const clientName = e.parameter.client || '';
        result = getStoredJSONData(clientName);
      } else if (path === 'update-account-json' || action === 'updateAccountJSON') {
        const visitId = e.parameter.visitId || '';
        const accountDataStr = e.parameter.accountData || '';
        try {
          const accountData = accountDataStr ? JSON.parse(accountDataStr) : {};
          result = updateAccountJSON(visitId, accountData);
        } catch(e) {
          result = { error: 'Invalid account data: ' + e.toString() };
        }
      } else if (path === 'cache-status' || action === 'getCacheStatus') {
        result = getCacheStatus();
      } else {
        result = { error: 'Invalid endpoint' };
      }
    } else if (method === 'POST') {
      const data = e.postData ? JSON.parse(e.postData.contents) : {};
      
      if (path.includes('accounts') && path.includes('update')) {
        const visitId = path.split('/')[1];
        result = updateAccountAPI(visitId, data);
      } else if (path === 'production' || action === 'postProduction') {
        result = postToProductionAPI(data);
      } else if (path.includes('non-workable') && path.includes('approve')) {
        const visitId = path.split('/')[1];
        result = approveNonWorkableAPI(visitId, data);
      } else if (path === 'convert' || action === 'convertSheet') {
        result = convertSheetToJSON(data.sheetId, data.sheetName);
      } else if (path === 'refresh' || action === 'refreshJSON') {
        result = refreshAllJSONData();
      } else if (path === 'update-account-json' || action === 'updateAccountJSON') {
        const visitId = data.visitId || '';
        const accountData = data.accountData || {};
        result = updateAccountJSON(visitId, accountData);
      } else {
        result = { error: 'Invalid endpoint' };
      }
    }
    
    return createResponse(result, result.error ? 400 : 200);
  } catch(e) {
    Logger.log('API Error: ' + e.toString());
    return createResponse({ error: e.toString() }, 500);
  }
}

// API wrapper functions that use direct Code.gs functions
function getAllAccountsAPI() {
  const cacheKey = 'api_all_accounts';
  const cached = getCached(cacheKey);
  if (cached && Array.isArray(cached) && cached.length > 0) {
    return cached;
  }
  
  // Load from main cache
  const allAccounts = getAllAccountsData();
  if (allAccounts && Array.isArray(allAccounts)) {
    // Store in API cache
    setCache(cacheKey, allAccounts, CACHE_DURATION);
    storeAccountsInPropertiesLocal(allAccounts);
    return allAccounts;
  }
  
  return { error: 'Data not available. Please wait for background refresh.' };
}

function getAccountByVisitIdAPI(visitId) {
  const account = getAccountByVisitId(visitId);
  return account || { error: 'Account not found' };
}

function searchAccountsAPI(params) {
  try {
    const searchTerm = (params.searchTerm || params.term || '').toLowerCase().trim();
    const searchType = params.searchType || params.type || 'all';
    
    if (!searchTerm) {
      return { error: 'Search term required' };
    }
    
    const all = getAllAccountsData();
    let results = [];
    
    switch(searchType) {
      case 'visitId':
        results = all.filter(a => String(a.visitId || '').toLowerCase().includes(searchTerm));
        break;
      case 'patientName':
        results = all.filter(a => String(a.patientName || '').toLowerCase().includes(searchTerm));
        break;
      case 'insurance':
        results = all.filter(a => String(a.insurance || '').toLowerCase().includes(searchTerm));
        break;
      default:
        results = all.filter(a => 
          String(a.visitId || '').toLowerCase().includes(searchTerm) ||
          String(a.patientName || '').toLowerCase().includes(searchTerm) ||
          String(a.insurance || '').toLowerCase().includes(searchTerm)
        );
    }
    
    return results.slice(0, 500);
  } catch(e) {
    Logger.log('Error searching accounts: ' + e.toString());
    return { error: e.toString() };
  }
}

function getNonWorkableAccountsAPI() {
  try {
    const all = getAllAccountsData();
    return all.filter(a => String(a.statusCode || '').toLowerCase() === 'non-workable');
  } catch(e) {
    Logger.log('Error getting non-workable accounts: ' + e.toString());
    return { error: e.toString() };
  }
}

function updateAccountAPI(visitId, data) {
  try {
    return updateAccountWorkForm(visitId, data);
  } catch(e) {
    Logger.log('Error updating account: ' + e.toString());
    return { error: e.toString() };
  }
}

function postToProductionAPI(data) {
  try {
    return { success: true, message: 'Posted to production report' };
  } catch(e) {
    Logger.log('Error posting to production: ' + e.toString());
    return { error: e.toString() };
  }
}

function approveNonWorkableAPI(visitId, data) {
  try {
    return { success: true, message: 'Non-workable account approved' };
  } catch(e) {
    Logger.log('Error approving non-workable: ' + e.toString());
    return { error: e.toString() };
  }
}

function initializeARClientConfigSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.arClientConfig.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.CMS_CONTROL_UNI.sheets.arClientConfig.name);
      // Add headers
      sheet.appendRow(['Client', 'Sheet ID', 'Sheet Name']);
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, 3);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
      SpreadsheetApp.flush();
      logActivity('CONFIG_SHEET_CREATED', 'AR Client Config sheet created');
    }
    
    return sheet;
  } catch(e) {
    Logger.log('Error initializing AR Client Config sheet: ' + e.toString());
    throw new Error('Failed to initialize AR Client Config sheet');
  }
}

function loadARClientSheets() {
  const cacheKey = 'ar_client_config';
  const cached = getCached(cacheKey);
  if (cached) return cached;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.arClientConfig.name);
    
    // Try alternative sheet name if exact match fails
    if (!sheet) {
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => s.getName().toLowerCase().includes('ar client config') || 
                                  s.getName().toLowerCase().includes('client config'));
    }
    
    // If still not found, create it
    if (!sheet) {
      sheet = initializeARClientConfigSheet();
      // Return empty clients object if sheet was just created
      return {};
    }
    
    // Read from range A1:C (header + data)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      // Sheet exists but no data - return empty
      return {};
    }
    
    // Read range A2:C (skip header row)
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const clients = {};
    
    data.forEach((row, idx) => {
      const client = String(row[0] || '').trim();
      const sheetId = String(row[1] || '').trim();
      const sheetName = String(row[2] || '').trim();
      
      if (client && sheetId && sheetName) {
        // Validate sheet ID format
        if (sheetId.length > 10) {
          clients[client] = { client, sheetId, sheetName };
        } else {
          Logger.log(`Skipping row ${idx + 2}: Invalid sheet ID format`);
        }
      }
    });
    
    setCache(cacheKey, clients);
    logActivity('CONFIG_LOAD', `Loaded ${Object.keys(clients).length} clients`);
    return clients;
  } catch(e) {
    logActivity('CONFIG_LOAD', e.message, 'ERROR');
    // Try to initialize sheet if it doesn't exist
    try {
      initializeARClientConfigSheet();
      return {};
    } catch(initErr) {
      throw new Error('Config load error: ' + e.message);
    }
  }
}

/**
 * Preload all accounts data into cache for faster access
 * This function is called on app initialization to warm up the cache
 */
function preloadAllAccountsData() {
  try {
    const start = Date.now();
    const accounts = getAllAccountsData(false); // Use cache if available, otherwise load
    
    // Also build and cache the visitId map for faster account lookup
    const accountsMap = {};
    accounts.forEach(a => {
      const vid = String(a.visitId).trim();
      if (vid) accountsMap[vid] = a;
    });
    setCache('accounts_by_visitid', accountsMap, CACHE_DURATION);
    
    const duration = Date.now() - start;
    Logger.log(`Preloaded ${accounts.length} accounts in ${duration}ms`);
    return { success: true, count: accounts.length, duration: duration };
  } catch(e) {
    Logger.log('Error preloading accounts: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getAllAccountsData(forceRefresh = false) {
  const cacheKey = 'all_accounts';
  
  // Check cache first (unless force refresh)
  if (!forceRefresh) {
    const cached = getCached(cacheKey);
    if (cached && Array.isArray(cached) && cached.length > 0) { 
      // Don't log cache hits to reduce log noise
      return cached; 
    }
  }
  
  try {
    const start = Date.now();
    const clients = loadARClientSheets();
    let allAccounts = [];
    let processedRows = 0;
    let skippedRows = 0;
    
    // Process clients in parallel batches for better performance
    const clientNames = Object.keys(clients);
    
    clientNames.forEach(clientName => {
      const cfg = clients[clientName];
      try {
        const targetSheet = SpreadsheetApp.openById(cfg.sheetId);
        if (!targetSheet) {
          Logger.log(`Cannot open spreadsheet: ${cfg.sheetId}`);
          return;
        }
        
        let sheet = targetSheet.getSheetByName(cfg.sheetName);
        // Try alternative sheet name matching
        if (!sheet) {
          const allSheets = targetSheet.getSheets();
          sheet = allSheets.find(s => 
            s.getName().toLowerCase() === cfg.sheetName.toLowerCase() ||
            s.getName().toLowerCase().includes(cfg.sheetName.toLowerCase())
          );
        }
        
        if (!sheet) {
          Logger.log(`Sheet "${cfg.sheetName}" not found in ${clientName}`);
          return;
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          Logger.log(`No data rows in ${clientName} - ${cfg.sheetName}`);
          return;
        }
        
        // Use getValues() with specific range for better performance
        const numCols = Math.min(sheet.getLastColumn(), 30); // Limit columns for performance
        const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
        
        data.forEach((row, idx) => {
          try {
            const visitId = row[COLUMNS.VISIT_ID];
            if (!visitId || String(visitId).trim() === '') {
              skippedRows++;
              return;
            }
            
          allAccounts.push({
            client: clientName, 
            state: row[COLUMNS.STATE] || '', 
            visitId: String(visitId).trim(),
            patientName: row[COLUMNS.PATIENT_NAME] || '', 
            dos: formatDate(row[COLUMNS.DOS]),
            agingDays: row[COLUMNS.AGING_DAYS] || 0, 
            agingBucket: row[COLUMNS.AGING_BUCKET] || '',
            submittedDate: formatDate(row[COLUMNS.SUBMITTED_DATE]), 
            insurance: row[COLUMNS.INSURANCE_NAME] || '',
            status: row[COLUMNS.STATUS] || '', 
            billedAmount: parseFloat(row[COLUMNS.BILLED_AMOUNT] || 0),
            balanceAmount: parseFloat(row[COLUMNS.BALANCE_AMOUNT] || 0), 
            primarySecondary: row[12] || '', // Column 12 (index 12) - Primary/Secondary
            type: row[COLUMNS.TYPE] || '',
            arNotes: row[COLUMNS.AR_NOTES] || '', 
            statusCode: row[COLUMNS.STATUS_CODE] || '', 
            actionCode: row[COLUMNS.ACTION_CODE] || '', 
            assignedTo: row[COLUMNS.ASSIGNED_TO] || '',
            workedBy: row[COLUMNS.WORKED_BY] || '', 
            workedDate: formatDate(row[COLUMNS.WORKED_DATE]),
            followupDate: formatDate(row[COLUMNS.FOLLOWUP_DATE]), 
            claimType: row[COLUMNS.CLAIM_TYPE] || '',
            allocatedUser: row[COLUMNS.ALLOCATED_USER] || '', 
            allocationDate: formatDate(row[COLUMNS.ALLOCATION_DATE]),
            remarks: row[COLUMNS.REMARKS] || '', 
            sourceSheetId: cfg.sheetId, 
            sourceSheetName: cfg.sheetName
          });
          processedRows++;
          } catch(rowErr) {
            Logger.log(`Error processing row ${idx + 2} in ${clientName}: ${rowErr}`);
            skippedRows++;
          }
        });
      } catch(e) { 
        Logger.log(`Error processing ${clientName}: ${e.toString()}`);
        logActivity('DATA_LOAD_CLIENT_ERROR', `${clientName}: ${e.message}`, 'ERROR');
      }
    });
    
    // Store in cache (both CacheService and PropertiesService)
    setCache(cacheKey, allAccounts);
    const duration = Date.now() - start;
    logActivity('DATA_LOAD', `${allAccounts.length} accounts loaded in ${duration}ms (processed: ${processedRows}, skipped: ${skippedRows})`);
    return allAccounts;
  } catch(e) {
    logActivity('DATA_LOAD', e.message, 'ERROR');
    throw new Error('Data load error: ' + e.message);
  }
}

function searchAccounts(searchTerm, searchType = 'all', filters = {}) {
  try {
    const allData = getAllAccountsData();
    const term = String(searchTerm || '').toLowerCase().trim();
    let results = [];
    
    // If no search term but filters provided, use filters only
    if (!term && Object.keys(filters).length > 0) {
      results = allData;
    } else if (term) {
      // Apply search term based on type
    switch(searchType) {
      case 'visitId': 
        results = allData.filter(a => String(a.visitId || '').toLowerCase().includes(term)); 
        break;
      case 'patientName': 
        results = allData.filter(a => String(a.patientName || '').toLowerCase().includes(term)); 
        break;
      case 'insurance': 
        results = allData.filter(a => String(a.insurance || '').toLowerCase().includes(term)); 
        break;
      case 'claimType': 
      case 'serviceType': // Alias for Service Type
        results = allData.filter(a => String(a.claimType || '').toLowerCase().includes(term)); 
        break;
      case 'agingBucket': 
        results = allData.filter(a => String(a.agingBucket || '').toLowerCase().includes(term)); 
        break;
      case 'client': 
        results = allData.filter(a => String(a.client || '').toLowerCase().includes(term)); 
        break;
      case 'statusCode': 
        results = allData.filter(a => String(a.statusCode || '').toLowerCase().includes(term)); 
        break;
      case 'actionCode': 
        results = allData.filter(a => String(a.actionCode || '').toLowerCase().includes(term)); 
        break;
      case 'state': 
        results = allData.filter(a => String(a.state || '').toLowerCase().includes(term)); 
        break;
      case 'type': 
        results = allData.filter(a => String(a.type || '').toLowerCase().includes(term)); 
        break;
      default: 
        results = allData.filter(a => 
          String(a.visitId || '').toLowerCase().includes(term) ||
          String(a.patientName || '').toLowerCase().includes(term) ||
          String(a.insurance || '').toLowerCase().includes(term) ||
          String(a.client || '').toLowerCase().includes(term) ||
          String(a.claimType || '').toLowerCase().includes(term)
        );
    }
    } else {
      results = allData;
    }
    
    // Apply additional filters
    if (filters.client && filters.client !== 'all') {
      results = results.filter(a => String(a.client || '').toLowerCase() === String(filters.client).toLowerCase());
    }
    if (filters.statusCode && filters.statusCode !== 'all') {
      results = results.filter(a => String(a.statusCode || '').toLowerCase() === String(filters.statusCode).toLowerCase());
    }
    if (filters.state && filters.state !== 'all') {
      results = results.filter(a => String(a.state || '').toLowerCase() === String(filters.state).toLowerCase());
    }
    if (filters.type && filters.type !== 'all') {
      results = results.filter(a => String(a.type || '').toLowerCase() === String(filters.type).toLowerCase());
    }
    if (filters.minBalance !== undefined && filters.minBalance !== '') {
      const minBal = parseFloat(filters.minBalance) || 0;
      results = results.filter(a => parseFloat(a.balanceAmount || 0) >= minBal);
    }
    if (filters.maxBalance !== undefined && filters.maxBalance !== '') {
      const maxBal = parseFloat(filters.maxBalance) || Infinity;
      results = results.filter(a => parseFloat(a.balanceAmount || 0) <= maxBal);
    }
    
    logActivity('SEARCH', JSON.stringify({type: searchType, term: searchTerm, filters: filters, results: results.length}));
    return results.slice(0, 500); // Increased limit for better results
  } catch(e) {
    logActivity('SEARCH', e.message, 'ERROR');
    throw new Error('Search error: ' + e.message);
  }
}

function getAccountByVisitId(visitId) {
  try {
    // First try cache for faster access
    const cacheKey = 'accounts_by_visitid';
    let accountsMap = getCached(cacheKey);
    
    if (!accountsMap || !accountsMap[String(visitId).trim()]) {
      // Build map from all accounts if not cached
      const all = getAllAccountsData();
      accountsMap = {};
      all.forEach(a => {
        const vid = String(a.visitId).trim();
        if (vid) accountsMap[vid] = a;
      });
      // Cache the map for 5 minutes
      setCache(cacheKey, accountsMap, CACHE_DURATION);
    }
    
    const account = accountsMap[String(visitId).trim()] || null;
    if (account) logActivity('ACCOUNT_VIEW', `VisitID: ${visitId}`);
    return account;
  } catch(e) {
    Logger.log('Error in getAccountByVisitId: ' + e.toString());
    // Fallback to direct search
    const all = getAllAccountsData();
    const account = all.find(a => String(a.visitId).trim() === String(visitId).trim());
    if (account) logActivity('ACCOUNT_VIEW', `VisitID: ${visitId}`);
    return account || null;
  }
}

function getDropdownOptions() {
  const cached = getCached('dropdown_opts');
  if (cached) return cached;
  
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId).getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.dropdown.name);
    const rows = sheet.getDataRange().getValues().slice(1);
    const opts = {
      source: [...new Set(rows.map(r => r[0]).filter(v => v))],
      statusCode: [...new Set(rows.map(r => r[1]).filter(v => v))],
      actionCode: [...new Set(rows.map(r => r[2]).filter(v => v))],
      assignedTo: [...new Set(rows.map(r => r[3]).filter(v => v))],
      claimType: [...new Set(rows.map(r => r[4]).filter(v => v))],
      client: [...new Set(rows.map(r => r[5]).filter(v => v))],
      users: [...new Set(rows.map(r => r[6]).filter(v => v))]
    };
    setCache('dropdown_opts', opts);
    return opts;
  } catch(e) { throw new Error('Dropdown error: ' + e.message); }
}

function getPatientInfo(patientName) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId).getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.patientMaster.name);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues().slice(1);
    const clean = (s) => String(s || '').replace(/\s+/g, '').toLowerCase();
    const target = clean(patientName);
    const row = rows.find(r => {
      const n = clean(r[0]);
      return n === target || (target.length > 5 && (n.includes(target) || target.includes(n)));
    });
    if (!row) return null;
    return {
      patientName: row[0], dob: formatDate(row[1]), primaryInsurance: row[2], 
      primaryInsuranceId: row[3], peffectiveDate: formatDate(row[4]), secondaryInsurance: row[5],
      secondaryInsuranceId: row[6], seffectiveDate: formatDate(row[7]), remarks: row[8]
    };
  } catch(e) { return null; }
}

function getProviderInfo(client) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId).getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.providerInfo.name);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues().slice(1); // Skip header
    const row = rows.find(r => String(r[0] || '').trim().toLowerCase() === String(client || '').trim().toLowerCase());
    if (!row) return null;
    // Columns: Client(0), Facility Name(1), NPI(2), TAX(3), PTAN(4), Address(5)
    return { 
      client: row[0], 
      name: row[1] || '', 
      npi: row[2] || '', 
      tax: row[3] || '', 
      ptan: row[4] || '', 
      address: row[5] || '' 
    };
  } catch(e) { 
    Logger.log('Error getting provider info: ' + e.toString());
    return null; 
  }
}

function getCredentialingInfo(client) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId).getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.credentialing.name);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues().slice(1); // Skip header
    const row = rows.find(r => String(r[0] || '').trim().toLowerCase() === String(client || '').trim().toLowerCase());
    if (!row) return null;
    // Columns: Client(0), Payer(1), Credentialing Status(2), W9(3), W9 Updated Date(4)
    return { 
      client: row[0], 
      payer: row[1] || '', 
      credentialingStatus: row[2] || '', 
      w9Status: row[3] || '', 
      w9UpdatedDate: formatDate(row[4]) 
    };
  } catch(e) { 
    Logger.log('Error getting credentialing info: ' + e.toString());
    return null; 
  }
}

function getClaimHistory(visitId) {
  if (!CONFIG.LAH_CONSOLIDATED.enabled) return [];
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.LAH_CONSOLIDATED.sheetId).getSheetByName(CONFIG.LAH_CONSOLIDATED.sheetName);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues().slice(1);
    const matches = rows.filter(r => String(r[1]).trim() === String(visitId).trim());
    const history = matches.map(r => ({
      timestamp: formatDate(r[0]), visitId: r[1], client: r[2], statusCode: r[3],
      actionCode: r[4], assignedTo: r[5], arComments: r[6], workedBy: r[7],
      workedDate: formatDate(r[8])
    }));
    history.sort((a,b) => new Date(b.rawTimestamp) - new Date(a.rawTimestamp));
    logActivity('HISTORY_VIEW', `VisitID: ${visitId}, ${history.length} records`);
    return history;
  } catch(e) { return []; }
}

function updateAccountWorkForm(visitId, formData) {
  try {
    const account = getAccountByVisitId(visitId);
    if (!account) throw new Error("Visit ID not found");
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    const sheet = ss.getSheetByName(account.sourceSheetName);
    const data = sheet.getDataRange().getValues();
    let rowIndex = data.findIndex((r, i) => i > 0 && String(r[COLUMNS.VISIT_ID]).trim() === String(visitId).trim()) + 1;
    if (rowIndex === 0) throw new Error("Visit ID not in sheet");
    
    const user = Session.getActiveUser().getEmail();
    const userName = getCurrentUserDisplayName();
    const now = new Date();
    
    // Save old values for undo/redo
    const oldValue = {
      notes: account.arNotes || '',
      statusCode: account.statusCode || '',
      actionCode: account.actionCode || '',
      assignedTo: account.assignedTo || '',
      followupDate: account.followupDate || '',
      remarks: account.remarks || ''
    };
    
    const updates = [];
    
    // Batch all updates
    if (formData.notes !== undefined) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.AR_NOTES + 1), value: formData.notes});
    }
    if (formData.source) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.TYPE_FIELD + 1), value: formData.source});
    }
    if (formData.statusCode) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.STATUS_CODE + 1), value: formData.statusCode});
    }
    if (formData.actionCode) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.ACTION_CODE + 1), value: formData.actionCode});
    }
    if (formData.assignedTo) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.ASSIGNED_TO + 1), value: formData.assignedTo});
    }
    if (formData.followupDate) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.FOLLOWUP_DATE + 1), value: new Date(formData.followupDate)});
    }
    
    // Always update worked by and worked date
    updates.push({range: sheet.getRange(rowIndex, COLUMNS.WORKED_BY + 1), value: userName});
    updates.push({range: sheet.getRange(rowIndex, COLUMNS.WORKED_DATE + 1), value: now});
    
    // If status is Non-Workable, store remarks in Remarks column
    if (formData.statusCode === 'Non-Workable' && formData.remarks) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.REMARKS + 1), value: formData.remarks});
    }
    
    // Apply all updates at once
    updates.forEach(update => update.range.setValue(update.value));
    
    // Save action history for undo/redo
    const newValue = {
      notes: formData.notes || account.arNotes || '',
      statusCode: formData.statusCode || account.statusCode || '',
      actionCode: formData.actionCode || account.actionCode || '',
      assignedTo: formData.assignedTo || account.assignedTo || '',
      followupDate: formData.followupDate || account.followupDate || '',
      remarks: formData.remarks || account.remarks || ''
    };
    saveActionHistory(visitId, 'ACCOUNT_UPDATE', oldValue, newValue);
    
    // Post to Non Workable sheet if status is Non-Workable
    if (formData.statusCode === 'Non-Workable') {
      postToNonWorkableSheet(ss, account, formData, userName, now);
    }
    
    // Post to production report
    postToProductionReport(visitId, formData);
    
    // Single flush after all updates
    SpreadsheetApp.flush();
    
    // Update account in cache immediately (incremental update)
    const updatedAccount = getAccountByVisitId(visitId);
    if (updatedAccount) {
      updateAccountInCache(visitId, updatedAccount);
    } else {
      // If account not found, clear cache to force refresh
      CACHE.remove('all_accounts');
      PROPERTIES.deleteProperty(ACCOUNTS_META_KEY);
    }
    
    // Don't clear ar_client_config cache unnecessarily
    if (CONFIG.LAH_CONSOLIDATED.enabled) logToTrackerLog(visitId, formData);
    logActivity('ACCOUNT_UPDATE', `${visitId}: ${formData.statusCode}/${formData.actionCode}`);
    return { success: true };
  } catch(e) {
    logActivity('ACCOUNT_UPDATE', e.message, 'ERROR');
    throw new Error(e.message);
  }
}

function postToNonWorkableSheet(ss, account, formData, userName, workedDate) {
  try {
    let nonWorkableSheet = ss.getSheetByName('Non Workable');
    if (!nonWorkableSheet) {
      nonWorkableSheet = ss.insertSheet('Non Workable');
      // Add headers if new sheet
      nonWorkableSheet.appendRow([
        'Timestamp', 'Visit ID', 'Client', 'Patient Name', 'DOS', 'Insurance', 
        'Balance Amount', 'Status Code', 'Action Code', 'AR Notes', 'Remarks', 
        'Worked By', 'Worked Date', 'Allocated User', 'Allocation Date',
        'Approval Status', 'Approved/Denied By', 'Approval Comment', 'Approval Date'
      ]);
      // Format header
      const headerRange = nonWorkableSheet.getRange(1, 1, 1, 19);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }
    
    // Check if Visit ID already exists
    const data = nonWorkableSheet.getDataRange().getValues();
    const existingRow = data.findIndex((r, i) => i > 0 && String(r[1] || '').trim() === String(account.visitId).trim());
    
    const rowData = [
      new Date(), // Timestamp
      account.visitId,
      account.client || '',
      account.patientName || '',
      account.dos || '',
      account.insurance || '',
      account.balanceAmount || 0,
      formData.statusCode || '',
      formData.actionCode || '',
      formData.notes || '',
      formData.remarks || '',
      userName,
      workedDate,
      account.allocatedUser || '',
      account.allocationDate || '',
      'Pending', // Approval Status
      '', // Approved/Denied By
      '', // Approval Comment
      '' // Approval Date
    ];
    
    if (existingRow > 0) {
      // Update existing row (preserve approval columns if they exist)
      const existingData = nonWorkableSheet.getRange(existingRow + 1, 1, 1, 19).getValues()[0];
      rowData[15] = existingData[15] || 'Pending'; // Approval Status
      rowData[16] = existingData[16] || ''; // Approved/Denied By
      rowData[17] = existingData[17] || ''; // Approval Comment
      rowData[18] = existingData[18] || ''; // Approval Date
      nonWorkableSheet.getRange(existingRow + 1, 1, 1, 19).setValues([rowData]);
    } else {
      // Append new row
      nonWorkableSheet.appendRow(rowData);
    }
  } catch(e) {
    Logger.log('Error posting to Non Workable sheet: ' + e.toString());
    // Don't throw - this is a secondary operation
  }
}

function logToTrackerLog(visitId, formData) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.LAH_CONSOLIDATED.sheetId).getSheetByName(CONFIG.LAH_CONSOLIDATED.sheetName);
    const account = getAccountByVisitId(visitId);
    const now = new Date();
    sheet.appendRow([now, visitId, account.client||'', formData.statusCode||'', formData.actionCode||'', formData.assignedTo||'', formData.notes||'', Session.getActiveUser().getEmail(), now]);
    SpreadsheetApp.flush();
  } catch(e) { Logger.log('Tracker log error: '+e); }
}

function updateRemarks(visitId, remarks) {
  try {
    const account = getAccountByVisitId(visitId);
    if (!account) throw new Error("Visit ID not found");
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    const sheet = ss.getSheetByName(account.sourceSheetName);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[COLUMNS.VISIT_ID]).trim() === String(visitId).trim()) + 1;
    if (rowIndex === 0) throw new Error("Visit ID not in sheet");
    
    // Save old value for undo/redo
    const oldValue = { remarks: account.remarks || '' };
    const newValue = { remarks: remarks || '' };
    saveActionHistory(visitId, 'REMARKS_UPDATE', oldValue, newValue);
    
    // Update remarks in AR outstanding sheet
    sheet.getRange(rowIndex, COLUMNS.REMARKS + 1).setValue(remarks);
    
    // If account is Non-Workable, also update in Non Workable sheet
    if (account.statusCode === 'Non-Workable') {
      updateNonWorkableSheetRemarks(ss, visitId, remarks);
    }
    
    SpreadsheetApp.flush();
    
    // Update account in cache immediately (incremental update)
    // This updates both local cache and API JSON cache
    const updatedAccount = getAccountByVisitId(visitId);
    if (updatedAccount) {
      updateAccountInCache(visitId, updatedAccount);
    } else {
      CACHE.remove('all_accounts');
      PROPERTIES.deleteProperty(ACCOUNTS_META_KEY);
    }
    
    logActivity('REMARKS_UPDATE', `VisitID: ${visitId}`);
    return { success: true };
  } catch(e) {
    logActivity('REMARKS_UPDATE', e.message, 'ERROR');
    throw new Error(e.message);
  }
}

function updateNonWorkableSheetRemarks(ss, visitId, remarks) {
  try {
    const nonWorkableSheet = ss.getSheetByName('Non Workable');
    if (!nonWorkableSheet) return;
    
    const data = nonWorkableSheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[1] || '').trim() === String(visitId).trim());
    
    if (rowIndex > 0) {
      // Update remarks column (column 11, index 10)
      nonWorkableSheet.getRange(rowIndex + 1, 11).setValue(remarks);
    }
  } catch(e) {
    Logger.log('Error updating Non Workable sheet remarks: ' + e.toString());
  }
}

function getFilteredAccountsData(type) {
  const all = getAllAccountsData();
  const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
  const norm = (s) => String(s||'').toLowerCase().trim();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  switch(type) {
    case 'assigned': 
      return all.filter(a => norm(a.allocatedUser).includes(u));
    
    case 'worked': 
      // Worked By = User name AND Worked Date = Today
      return all.filter(a => {
        const allocatedMatch = norm(a.allocatedUser).includes(u);
        const workedByMatch = norm(a.workedBy) === u;
        if (!allocatedMatch || !workedByMatch || !a.workedDate) return false;
        
        const workedDate = new Date(a.workedDate);
        workedDate.setHours(0, 0, 0, 0);
        return workedDate.getTime() === today.getTime();
      });
    
    case 'pending': 
      // Worked By = Username AND (Worked Date = Blank OR Allocation Date - Worked Date > 0)
      return all.filter(a => {
        const allocatedMatch = norm(a.allocatedUser).includes(u);
        const workedByMatch = norm(a.workedBy) === u;
        const statusMatch = a.statusCode === 'Pending';
        
        if (!allocatedMatch || !workedByMatch || !statusMatch) return false;
        
        // If worked date is blank, it's pending
        if (!a.workedDate) return true;
        
        // If allocation date exists, check if allocation date - worked date > 0
        if (a.allocationDate) {
          const allocationDate = new Date(a.allocationDate).getTime();
          const workedDate = new Date(a.workedDate).getTime();
          return allocationDate - workedDate > 0;
        }
        
        return false;
      });
    
    case 'nonWorkable': 
      return all.filter(a => norm(a.allocatedUser).includes(u) && a.statusCode==='Non-Workable');
    
    default: 
      return [];
  }
}

function getAccountStatistics() {
  try {
    const all = getAllAccountsData();
    const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
    const norm = (s) => String(s||'').toLowerCase().trim();
    
    // Total Assigned: Allocated User contains current user name
    const assigned = all.filter(a => norm(a.allocatedUser).includes(u));
    const totalAssigned = assigned.length;
    
    // Total Worked: Worked By = User name AND Worked Date = Today (with proper date comparison)
    const totalWorked = all.filter(a => {
      const allocatedMatch = norm(a.allocatedUser).includes(u);
      const workedByMatch = norm(a.workedBy) === u;
      if (!allocatedMatch || !workedByMatch || !a.workedDate) return false;
      
      try {
        const workedDate = new Date(a.workedDate);
        workedDate.setHours(0, 0, 0, 0);
        workedDate.setMinutes(0, 0, 0);
        return workedDate.getTime() === today.getTime();
      } catch(e) {
        return false;
      }
    }).length;
    
    // Total Pending: Allocated to user AND Status = 'Pending' AND (Worked Date is blank OR Allocation Date > Worked Date)
    const totalPending = all.filter(a => {
      const allocatedMatch = norm(a.allocatedUser).includes(u);
      const statusMatch = norm(a.statusCode) === 'pending';
      
      if (!allocatedMatch || !statusMatch) return false;
      
      // If worked date is blank, it's pending
      if (!a.workedDate || String(a.workedDate).trim() === '') return true;
      
      // If allocation date exists, check if allocation date > worked date
      if (a.allocationDate) {
        try {
          const allocationDate = new Date(a.allocationDate).getTime();
          const workedDate = new Date(a.workedDate).getTime();
          return allocationDate > workedDate;
        } catch(e) {
          return false;
        }
      }
      
      return false;
    }).length;
    
    // Total Non Workable: Allocated to user AND Status Code = 'Non-Workable'
    const totalNonWorkable = all.filter(a => {
      return norm(a.allocatedUser).includes(u) && norm(a.statusCode) === 'non-workable';
    }).length;
    
    return {
      totalAssigned: totalAssigned,
      totalWorked: totalWorked,
      totalPending: totalPending,
      totalNonWorkable: totalNonWorkable,
      totalAccounts: all.length
    };
  } catch(e) {
    Logger.log('Error in getAccountStatistics: ' + e.toString());
    return {
      totalAssigned: 0,
      totalWorked: 0,
      totalPending: 0,
      totalNonWorkable: 0,
      totalAccounts: 0
    };
  }
}

function getAssignedAccounts(forceRefresh = false) {
  const all = getAllAccountsData(forceRefresh);
  const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
  return all.filter(a=>String(a.allocatedUser||'').toLowerCase().trim().includes(u));
}

function exportAccountsToCSV(accountType) {
  const data = getFilteredAccountsData(accountType);
  const headers = ['Client','State','Visit ID','Patient','DOS','Aging Days','Aging Bucket','Submitted','Insurance','Status','Billed','Balance','Type','Notes','Status Code','Action Code','Assigned','Worked By','Worked Date','Followup','Claim Type','Allocated User','Allocation Date','Remarks'];
  const csvData = [headers];
  const q = (v) => `"${String(v||'').replace(/"/g,'""')}"`;
  data.forEach(a => csvData.push([q(a.client),q(a.state),q(a.visitId),q(a.patientName),q(a.dos),q(a.agingDays),q(a.agingBucket),q(a.submittedDate),q(a.insurance),q(a.status),q('$'+a.billedAmount),q('$'+a.balanceAmount),q(a.type),q(a.arNotes),q(a.statusCode),q(a.actionCode),q(a.assignedTo),q(a.workedBy),q(a.workedDate),q(a.followupDate),q(a.claimType),q(a.allocatedUser),q(a.allocationDate),q(a.remarks)]));
  logActivity('EXPORT', `${accountType}: ${data.length} records`);
  return csvData;
}

function authenticateDeveloper(username, password) {
  const userEmail = Session.getActiveUser().getEmail();
  const emailValid = userEmail.toLowerCase() === CONFIG.DEVELOPER.allowedEmail.toLowerCase();
  const credValid = username === CONFIG.DEVELOPER.username && password === CONFIG.DEVELOPER.password;
  const valid = emailValid && credValid;
  
  logActivity('DEVELOPER_LOGIN', `User: ${username}, Email: ${userEmail}`, valid ? 'SUCCESS' : 'FAILED');
  
  if (!emailValid) {
    Logger.log(`Developer access denied for email: ${userEmail}`);
  }
  
  return valid;
}

function getActivityLogs(limit = 200) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.arConnectLog.name);
    if (!sheet) {
      Logger.log('AR Connect Log sheet not found');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data rows in log sheet');
      return [];
    }
    
    // Limit to prevent timeout
    const maxLimit = Math.min(limit, 1000);
    const startRow = Math.max(2, lastRow - maxLimit + 1);
    const numRows = lastRow - startRow + 1;
    
    if (numRows <= 0) return [];
    
    // Read data in chunks if needed
    const rows = sheet.getRange(startRow, 1, numRows, 13).getValues();
    
    return rows.reverse().map(r => ({
      timestamp: r[0] ? formatDateTime(r[0]) : '',
      email: r[1] || '',
      name: r[2] || '',
      action: r[3] || '',
      details: r[4] || '',
      status: r[5] || 'SUCCESS',
      sessionId: r[6] || '',
      duration: r[7] || '',
      error: r[8] || '',
      ipAddress: r[9] || '',
      userAgent: r[10] || '',
      page: r[11] || '',
      visitId: r[12] || '',
      rawTimestamp: r[0]
    }));
  } catch(e) { 
    Logger.log('Error getting activity logs: ' + e.toString());
    return []; 
  }
}

function getUserStatistics() {
  try {
    const logs = getActivityLogs(5000); // Get more logs for better stats
    const stats = {};
    const activeSessions = {};
    const sessionTimeout = 30 * 60 * 1000; // 30 minutes
    
    logs.forEach(log => {
      const key = log.email || log.name || 'Unknown';
      if (!stats[key]) {
        stats[key] = { 
          name: log.name || 'Unknown', 
          email: log.email || '', 
          totalActions: 0, 
          lastSeen: log.timestamp, 
          firstSeen: log.timestamp,
          lastSeenRaw: log.rawTimestamp,
          firstSeenRaw: log.rawTimestamp,
          actions: {},
          sessions: new Set(),
          errors: 0,
          pages: new Set(),
          isActive: false
        };
      }
      
      stats[key].totalActions++;
      stats[key].actions[log.action] = (stats[key].actions[log.action] || 0) + 1;
      
      if (log.sessionId) stats[key].sessions.add(log.sessionId);
      if (log.page) stats[key].pages.add(log.page);
      if (log.status === 'ERROR' || log.status === 'FAILED') stats[key].errors++;
      
      // Update last seen (most recent)
      if (log.rawTimestamp) {
        const logTime = new Date(log.rawTimestamp).getTime();
        const lastSeenTime = stats[key].lastSeenRaw ? new Date(stats[key].lastSeenRaw).getTime() : 0;
        const firstSeenTime = stats[key].firstSeenRaw ? new Date(stats[key].firstSeenRaw).getTime() : Infinity;
        
        if (logTime > lastSeenTime) {
          stats[key].lastSeen = log.timestamp;
          stats[key].lastSeenRaw = log.rawTimestamp;
        }
        if (logTime < firstSeenTime) {
          stats[key].firstSeen = log.timestamp;
          stats[key].firstSeenRaw = log.rawTimestamp;
        }
        
        // Check if session is active (within last 30 minutes)
        const now = new Date().getTime();
        if (logTime > (now - sessionTimeout)) {
          stats[key].isActive = true;
          if (log.sessionId) {
            activeSessions[log.sessionId] = {
              email: log.email,
              name: log.name,
              lastActivity: log.timestamp,
              lastActivityRaw: log.rawTimestamp,
              page: log.page || 'Index'
            };
          }
        }
      }
    });
    
    // Convert Sets to counts and add active session info
    return Object.values(stats).map(stat => ({
      name: stat.name,
      email: stat.email,
      totalActions: stat.totalActions,
      lastSeen: stat.lastSeen,
      firstSeen: stat.firstSeen,
      sessions: stat.sessions.size,
      errors: stat.errors,
      pages: Array.from(stat.pages),
      actions: stat.actions,
      isActive: stat.isActive,
      activeSessions: Object.values(activeSessions).filter(s => s.email === stat.email).length
    }));
  } catch(e) { 
    Logger.log('Error getting user statistics: ' + e.toString());
    return []; 
  }
}

function getActiveSessions() {
  try {
    const logs = getActivityLogs(500); // Reduced limit for performance
    const sessionTimeout = 30 * 60 * 1000; // 30 minutes
    const now = new Date().getTime();
    const activeSessions = {};
    
    logs.forEach(log => {
      if (log.sessionId && log.rawTimestamp) {
        try {
          const logTime = new Date(log.rawTimestamp).getTime();
          if (!isNaN(logTime) && logTime > (now - sessionTimeout)) {
            if (!activeSessions[log.sessionId] || new Date(log.rawTimestamp) > new Date(activeSessions[log.sessionId].lastActivityRaw)) {
              activeSessions[log.sessionId] = {
                sessionId: log.sessionId,
                email: log.email || '',
                name: log.name || 'Unknown',
                lastActivity: log.timestamp,
                lastActivityRaw: log.rawTimestamp,
                page: log.page || 'Index',
                action: log.action || '',
                ipAddress: log.ipAddress || '',
                userAgent: log.userAgent || '',
                version: APP_VERSION // Add version tracking
              };
            }
          }
        } catch(e) {
          Logger.log('Error processing log entry: ' + e.toString());
        }
      }
    });
    
    return Object.values(activeSessions);
  } catch(e) {
    Logger.log('Error getting active sessions: ' + e.toString());
    return [];
  }
}

function blockUser(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName('Blocked Users');
    if (!sheet) {
      sheet = ss.insertSheet('Blocked Users');
      sheet.appendRow(['Email', 'Name', 'Blocked Date', 'Blocked By']);
    }
    const user = Session.getActiveUser();
    sheet.appendRow([email, '', new Date(), user.getEmail()]);
    logActivity('USER_BLOCKED', `Blocked user: ${email}`, 'SUCCESS');
    return { success: true, message: 'User blocked successfully' };
  } catch(e) {
    Logger.log('Error blocking user: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteUserLogs(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    const sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.arConnectLog.name);
    if (!sheet) return { success: false, message: 'Log sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1] || '').toLowerCase() === email.toLowerCase()) {
        rowsToDelete.push(i + 1);
      }
    }
    
    if (rowsToDelete.length > 0) {
      rowsToDelete.forEach(row => sheet.deleteRow(row));
      logActivity('USER_LOGS_DELETED', `Deleted ${rowsToDelete.length} logs for: ${email}`, 'SUCCESS');
      return { success: true, message: `Deleted ${rowsToDelete.length} log entries` };
    }
    
    return { success: true, message: 'No logs found for this user' };
  } catch(e) {
    Logger.log('Error deleting user logs: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getAnalyticsData() {
  try {
    const logs = getActivityLogs(10000);
    const now = new Date();
    const last24h = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const last7d = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const last30d = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    
    const analytics = {
      totalUsers: new Set(logs.map(l => l.email)).size,
      totalActions: logs.length,
      errors: logs.filter(l => l.status === 'ERROR' || l.status === 'FAILED').length,
      actions24h: logs.filter(l => l.rawTimestamp && new Date(l.rawTimestamp) > last24h).length,
      actions7d: logs.filter(l => l.rawTimestamp && new Date(l.rawTimestamp) > last7d).length,
      actions30d: logs.filter(l => l.rawTimestamp && new Date(l.rawTimestamp) > last30d).length,
      topActions: {},
      topUsers: {},
      hourlyActivity: {},
      dailyActivity: {}
    };
    
    logs.forEach(log => {
      // Top actions
      analytics.topActions[log.action] = (analytics.topActions[log.action] || 0) + 1;
      
      // Top users
      if (log.email) {
        analytics.topUsers[log.email] = (analytics.topUsers[log.email] || 0) + 1;
      }
      
      // Hourly activity
      if (log.rawTimestamp) {
        const date = new Date(log.rawTimestamp);
        const hour = date.getHours();
        analytics.hourlyActivity[hour] = (analytics.hourlyActivity[hour] || 0) + 1;
        
        // Daily activity
        const dayKey = `${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()}`;
        analytics.dailyActivity[dayKey] = (analytics.dailyActivity[dayKey] || 0) + 1;
      }
    });
    
    // Sort and limit top actions/users
    analytics.topActions = Object.entries(analytics.topActions)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .reduce((obj, [key, val]) => { obj[key] = val; return obj; }, {});
    
    analytics.topUsers = Object.entries(analytics.topUsers)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .reduce((obj, [key, val]) => { obj[key] = val; return obj; }, {});
    
    return analytics;
  } catch(e) {
    Logger.log('Error getting analytics: ' + e.toString());
    return null;
  }
}

function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  try { 
    const d = new Date(date); 
    return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`; 
  } catch(e) { return ''; }
}

function formatDateTime(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    try {
      const d = new Date(date);
      return formatDateTime(d);
    } catch(e) {
      return date;
    }
  }
  try { 
    const d = new Date(date); 
    const dateStr = `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;
    const timeStr = `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
    return `${dateStr} ${timeStr}`;
  } catch(e) { return ''; }
}

function getCurrentUser() { return { email: Session.getActiveUser().getEmail(), name: getCurrentUserDisplayName() }; }

function getCurrentUserDisplayName() {
  try {
    const user = Session.getActiveUser();
    const email = user.getEmail();
    
    // Try to get name from People API (if available)
    try {
      const peopleApi = People.People.get('people/me', { personFields: 'names' });
      if (peopleApi.names && peopleApi.names.length > 0) {
        return peopleApi.names[0].displayName || peopleApi.names[0].givenName || '';
      }
    } catch(e) {
      // People API not available, continue with fallback
    }
    
    // Fallback: Try to get from sheet mapping
    try {
      const sheet = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId).getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.dropdown.name);
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][6]).toLowerCase().includes(email.split('@')[0].toLowerCase())) {
          return data[i][6];
        }
      }
    } catch(e) {
      // Sheet lookup failed, continue
    }
    
    // Final fallback: Use email prefix with proper formatting
    const emailPrefix = email.split('@')[0];
    return emailPrefix.split('.').map(part => 
      part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
    ).join(' ');
  } catch(e) { 
    // Ultimate fallback
    try {
      const email = Session.getActiveUser().getEmail();
      const emailPrefix = email.split('@')[0];
      return emailPrefix.charAt(0).toUpperCase() + emailPrefix.slice(1);
    } catch(e2) {
      return 'User';
    }
  }
}

function getSearchFilterOptions() {
  try {
    const allData = getAllAccountsData();
    const clients = [...new Set(allData.map(a => a.client).filter(c => c))].sort();
    const states = [...new Set(allData.map(a => a.state).filter(s => s))].sort();
    const statusCodes = [...new Set(allData.map(a => a.statusCode).filter(s => s))].sort();
    const types = [...new Set(allData.map(a => a.type).filter(t => t))].sort();
    const claimTypes = [...new Set(allData.map(a => a.claimType).filter(c => c))].sort();
    
    return {
      clients: clients,
      states: states,
      statusCodes: statusCodes,
      types: types,
      claimTypes: claimTypes
    };
  } catch(e) {
    Logger.log('Error getting filter options: ' + e.toString());
    return { clients: [], states: [], statusCodes: [], types: [], claimTypes: [] };
  }
}

// ============================================
// ACCESS CONTROLLER FUNCTIONS
// ============================================

function initializeAccessControllerSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.accessController.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.CMS_CONTROL_UNI.sheets.accessController.name);
      // Updated headers: Email, Name, Role, Status, API Key (deprecated), Created Date, Last Access, Version, Allowed Tabs, Allowed URLs, Created By
      sheet.appendRow([
        'Email', 'Name', 'Role', 'Status', 'API Key', 'Created Date', 
        'Last Access', 'Version', 'Allowed Tabs', 'Allowed URLs', 'Created By'
      ]);
      // Format header
      const headerRange = sheet.getRange(1, 1, 1, 11);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
      SpreadsheetApp.flush();
    } else {
      // Check if Allowed URLs column exists, if not add it
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Allowed URLs')) {
        // Add Allowed URLs column (column 10, index 9)
        sheet.getRange(1, 10).setValue('Allowed URLs');
        sheet.getRange(1, 10).setFontWeight('bold');
        sheet.getRange(1, 10).setBackground('#4285f4');
        sheet.getRange(1, 10).setFontColor('#ffffff');
        SpreadsheetApp.flush();
      }
    }
    
    return sheet;
  } catch(e) {
    Logger.log('Error initializing Access Controller: ' + e.toString());
    throw new Error('Failed to initialize Access Controller sheet');
  }
}

function checkUserAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        const status = String(data[i][3] || '').toLowerCase();
        return status === 'active';
      }
    }
    
    // If not found in Access Controller, allow access (backward compatibility)
    return true;
  } catch(e) {
    Logger.log('Error checking user access: ' + e.toString());
    return true; // Default to allow access
  }
}

function getUserRole() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        return String(data[i][2] || 'User').trim();
      }
    }
    
    return 'User'; // Default role
  } catch(e) {
    Logger.log('Error getting user role: ' + e.toString());
    return 'User';
  }
}

function getUserTabs() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        const tabsJson = String(data[i][8] || '[]');
        try {
          return JSON.parse(tabsJson);
        } catch(e) {
          // Default tabs if JSON is invalid
          return ['assigned', 'searchResults'];
        }
      }
    }
    
    // Default tabs for users not in Access Controller
    return ['assigned', 'searchResults'];
  } catch(e) {
    Logger.log('Error getting user tabs: ' + e.toString());
    return ['assigned', 'searchResults'];
  }
}

function grantAccess(email, role, tabs = [], allowedUrls = []) {
  try {
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    const userEmail = email.toLowerCase();
    const currentUser = Session.getActiveUser().getEmail();
    
    // Check if user already exists
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        // Update existing user
        sheet.getRange(i + 1, 2).setValue(data[i][1] || email.split('@')[0]); // Name
        sheet.getRange(i + 1, 3).setValue(role); // Role
        sheet.getRange(i + 1, 4).setValue('Active'); // Status
        sheet.getRange(i + 1, 7).setValue(new Date()); // Last Access
        sheet.getRange(i + 1, 8).setValue(APP_VERSION); // Version
        sheet.getRange(i + 1, 9).setValue(JSON.stringify(tabs)); // Allowed Tabs
        sheet.getRange(i + 1, 10).setValue(allowedUrls.join(',')); // Allowed URLs
        logActivity('ACCESS_GRANTED', `Updated access for: ${email}`, 'SUCCESS');
        return { success: true, message: 'Access updated successfully' };
      }
    }
    
    // Add new user
    const apiKey = Utilities.getUuid();
    sheet.appendRow([
      email,
      email.split('@')[0],
      role,
      'Active',
      apiKey,
      new Date(),
      new Date(),
      APP_VERSION,
      JSON.stringify(tabs),
      allowedUrls.join(','), // Allowed URLs
      currentUser
    ]);
    
    logActivity('ACCESS_GRANTED', `Granted access to: ${email}`, 'SUCCESS');
    return { success: true, message: 'Access granted successfully', apiKey: apiKey };
  } catch(e) {
    Logger.log('Error granting access: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function revokeAccess(email) {
  try {
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    const userEmail = email.toLowerCase();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        sheet.getRange(i + 1, 4).setValue('Blocked'); // Set status to Blocked
        logActivity('ACCESS_REVOKED', `Revoked access for: ${email}`, 'SUCCESS');
        return { success: true, message: 'Access revoked successfully' };
      }
    }
    
    return { success: false, message: 'User not found' };
  } catch(e) {
    Logger.log('Error revoking access: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateUserRole(email, role) {
  try {
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    const userEmail = email.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        sheet.getRange(i + 1, 3).setValue(role);
        logActivity('ROLE_UPDATED', `Updated role for ${email} to ${role}`, 'SUCCESS');
        return { success: true, message: 'Role updated successfully' };
      }
    }
    
    return { success: false, message: 'User not found' };
  } catch(e) {
    Logger.log('Error updating role: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateUserTabs(email, tabs) {
  try {
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    const userEmail = email.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        sheet.getRange(i + 1, 9).setValue(JSON.stringify(tabs));
        logActivity('TABS_UPDATED', `Updated tabs for ${email}`, 'SUCCESS');
        return { success: true, message: 'Tabs updated successfully' };
      }
    }
    
    return { success: false, message: 'User not found' };
  } catch(e) {
    Logger.log('Error updating tabs: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getAllAccessUsers() {
  try {
    const sheet = initializeAccessControllerSheet();
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    for (let i = 1; i < data.length; i++) {
      try {
        const tabsJson = String(data[i][8] || '[]');
        let tabs = [];
        try {
          tabs = JSON.parse(tabsJson);
        } catch(e) {
          tabs = [];
        }
        
        users.push({
          email: data[i][0] || '',
          name: data[i][1] || '',
          role: data[i][2] || 'User',
          status: data[i][3] || 'Active',
          apiKey: data[i][4] || '',
          createdDate: data[i][5] ? formatDateTime(data[i][5]) : '',
          lastAccess: data[i][6] ? formatDateTime(data[i][6]) : '',
          version: data[i][7] || APP_VERSION,
          allowedTabs: tabs,
          allowedUrls: data[i][9] ? String(data[i][9]).split(',').map(u => u.trim()) : [],
          createdBy: data[i][10] || ''
        });
      } catch(e) {
        Logger.log('Error parsing user row: ' + e.toString());
      }
    }
    
    return users;
  } catch(e) {
    Logger.log('Error getting all access users: ' + e.toString());
    return [];
  }
}

// ============================================
// UNDO/REDO SYSTEM
// ============================================

function saveActionHistory(visitId, action, oldValue, newValue) {
  try {
    const user = Session.getActiveUser().getEmail();
    const historyKey = `action_history_${visitId}_${user}`;
    const history = PROPERTIES.getProperty(historyKey);
    
    let historyArray = [];
    if (history) {
      try {
        historyArray = JSON.parse(history);
      } catch(e) {
        historyArray = [];
      }
    }
    
    historyArray.push({
      visitId: visitId,
      action: action,
      timestamp: new Date().toISOString(),
      oldValue: oldValue,
      newValue: newValue,
      userId: user
    });
    
    // Keep only last 50 actions per visit
    if (historyArray.length > 50) {
      historyArray = historyArray.slice(-50);
    }
    
    PROPERTIES.setProperty(historyKey, JSON.stringify(historyArray));
    return { success: true };
  } catch(e) {
    Logger.log('Error saving action history: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function undoLastAction(visitId) {
  try {
    const user = Session.getActiveUser().getEmail();
    const historyKey = `action_history_${visitId}_${user}`;
    const undoKey = `undo_history_${visitId}_${user}`;
    
    const history = PROPERTIES.getProperty(historyKey);
    if (!history) {
      return { success: false, message: 'No action history found' };
    }
    
    let historyArray = JSON.parse(history);
    if (historyArray.length === 0) {
      return { success: false, message: 'No actions to undo' };
    }
    
    const lastAction = historyArray.pop();
    
    // Save to undo history
    let undoArray = [];
    const undoHistory = PROPERTIES.getProperty(undoKey);
    if (undoHistory) {
      try {
        undoArray = JSON.parse(undoHistory);
      } catch(e) {
        undoArray = [];
      }
    }
    undoArray.push(lastAction);
    PROPERTIES.setProperty(undoKey, JSON.stringify(undoArray));
    
    // Update history
    PROPERTIES.setProperty(historyKey, JSON.stringify(historyArray));
    
    // Revert the change
    const account = getAccountByVisitId(visitId);
    if (!account) {
      return { success: false, message: 'Account not found' };
    }
    
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    const sheet = ss.getSheetByName(account.sourceSheetName);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[COLUMNS.VISIT_ID]).trim() === String(visitId).trim()) + 1;
    
    if (rowIndex === 0) {
      return { success: false, message: 'Account not found in sheet' };
    }
    
    // Revert old values
    const oldVal = lastAction.oldValue;
    if (oldVal.notes !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.AR_NOTES + 1).setValue(oldVal.notes || '');
    }
    if (oldVal.statusCode !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.STATUS_CODE + 1).setValue(oldVal.statusCode || '');
    }
    if (oldVal.actionCode !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.ACTION_CODE + 1).setValue(oldVal.actionCode || '');
    }
    if (oldVal.assignedTo !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.ASSIGNED_TO + 1).setValue(oldVal.assignedTo || '');
    }
    if (oldVal.followupDate !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.FOLLOWUP_DATE + 1).setValue(oldVal.followupDate ? new Date(oldVal.followupDate) : '');
    }
    if (oldVal.remarks !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.REMARKS + 1).setValue(oldVal.remarks || '');
    }
    
    // Remove from Tracker log if enabled
    if (CONFIG.LAH_CONSOLIDATED.enabled && lastAction.action === 'ACCOUNT_UPDATE') {
      try {
        const trackerSheet = SpreadsheetApp.openById(CONFIG.LAH_CONSOLIDATED.sheetId).getSheetByName(CONFIG.LAH_CONSOLIDATED.sheetName);
        if (trackerSheet) {
          const trackerData = trackerSheet.getDataRange().getValues();
          // Find and remove the most recent entry for this visitId
          for (let i = trackerData.length - 1; i >= 1; i--) {
            if (String(trackerData[i][1]).trim() === String(visitId).trim() && 
                String(trackerData[i][7]).trim().toLowerCase() === user.toLowerCase()) {
              trackerSheet.deleteRow(i + 1);
              break;
            }
          }
        }
      } catch(trackerErr) {
        Logger.log('Error removing from tracker log: ' + trackerErr.toString());
        // Don't fail undo if tracker log removal fails
      }
    }
    
    SpreadsheetApp.flush();
    
    // Update account in cache immediately (incremental update)
    // This updates both local cache and API JSON cache
    const updatedAccount = getAccountByVisitId(visitId);
    if (updatedAccount) {
      updateAccountInCache(visitId, updatedAccount);
    } else {
      CACHE.remove('all_accounts');
      const ACCOUNTS_META_KEY = 'api_accounts_meta';
      PROPERTIES.deleteProperty(ACCOUNTS_META_KEY);
    }
    
    logActivity('UNDO_ACTION', `Undid action for ${visitId}`, 'SUCCESS');
    
    return { success: true, message: 'Action undone successfully' };
  } catch(e) {
    Logger.log('Error undoing action: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function redoAction(visitId) {
  try {
    const user = Session.getActiveUser().getEmail();
    const undoKey = `undo_history_${visitId}_${user}`;
    const historyKey = `action_history_${visitId}_${user}`;
    
    const undoHistory = PROPERTIES.getProperty(undoKey);
    if (!undoHistory) {
      return { success: false, message: 'No actions to redo' };
    }
    
    let undoArray = JSON.parse(undoHistory);
    if (undoArray.length === 0) {
      return { success: false, message: 'No actions to redo' };
    }
    
    const actionToRedo = undoArray.pop();
    
    // Add back to history
    let historyArray = [];
    const history = PROPERTIES.getProperty(historyKey);
    if (history) {
      try {
        historyArray = JSON.parse(history);
      } catch(e) {
        historyArray = [];
      }
    }
    historyArray.push(actionToRedo);
    PROPERTIES.setProperty(historyKey, JSON.stringify(historyArray));
    PROPERTIES.setProperty(undoKey, JSON.stringify(undoArray));
    
    // Re-apply the change
    const account = getAccountByVisitId(visitId);
    if (!account) {
      return { success: false, message: 'Account not found' };
    }
    
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    const sheet = ss.getSheetByName(account.sourceSheetName);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[COLUMNS.VISIT_ID]).trim() === String(visitId).trim()) + 1;
    
    if (rowIndex === 0) {
      return { success: false, message: 'Account not found in sheet' };
    }
    
    // Apply new values
    const newVal = actionToRedo.newValue;
    if (newVal.notes !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.AR_NOTES + 1).setValue(newVal.notes || '');
    }
    if (newVal.statusCode !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.STATUS_CODE + 1).setValue(newVal.statusCode || '');
    }
    if (newVal.actionCode !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.ACTION_CODE + 1).setValue(newVal.actionCode || '');
    }
    if (newVal.assignedTo !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.ASSIGNED_TO + 1).setValue(newVal.assignedTo || '');
    }
    if (newVal.followupDate !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.FOLLOWUP_DATE + 1).setValue(newVal.followupDate ? new Date(newVal.followupDate) : '');
    }
    if (newVal.remarks !== undefined) {
      sheet.getRange(rowIndex, COLUMNS.REMARKS + 1).setValue(newVal.remarks || '');
    }
    
    SpreadsheetApp.flush();
    CACHE.remove('all_accounts');
    logActivity('REDO_ACTION', `Redid action for ${visitId}`, 'SUCCESS');
    
    return { success: true, message: 'Action redone successfully' };
  } catch(e) {
    Logger.log('Error redoing action: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getActionHistory(visitId) {
  try {
    const user = Session.getActiveUser().getEmail();
    const historyKey = `action_history_${visitId}_${user}`;
    const history = PROPERTIES.getProperty(historyKey);
    
    if (!history) {
      return [];
    }
    
    return JSON.parse(history);
  } catch(e) {
    Logger.log('Error getting action history: ' + e.toString());
    return [];
  }
}

// ============================================
// NEW TAB FUNCTIONS
// ============================================

function getTotalWorkedAccounts() {
  try {
    const all = getAllAccountsData();
    const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
    const norm = (s) => String(s||'').toLowerCase().trim();
    
    return all.filter(a => {
      const allocatedMatch = norm(a.allocatedUser).includes(u);
      const workedByMatch = norm(a.workedBy) === u;
      // Check if workedDate exists and is not empty
      if (!allocatedMatch || !workedByMatch || !a.workedDate || String(a.workedDate).trim() === '') {
        return false;
      }
      // Return true if workedBy matches and workedDate exists (all worked accounts, not just today)
      return true;
    });
  } catch(e) {
    Logger.log('Error getting total worked accounts: ' + e.toString());
    return [];
  }
}

function getTotalPendingAccounts() {
  try {
    const all = getAllAccountsData();
    const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
    const norm = (s) => String(s||'').toLowerCase().trim();
    
    // Get assigned, worked, and non-workable accounts
    const assigned = all.filter(a => norm(a.allocatedUser).includes(u));
    const worked = all.filter(a => {
      const allocatedMatch = norm(a.allocatedUser).includes(u);
      const workedByMatch = norm(a.workedBy) === u;
      const hasWorkedDate = a.workedDate && String(a.workedDate).trim() !== '';
      return allocatedMatch && workedByMatch && hasWorkedDate;
    });
    const nonWorkable = all.filter(a => 
      norm(a.allocatedUser).includes(u) && norm(a.statusCode) === 'non-workable'
    );
    
    // Create sets of visitIds for faster lookup
    const workedVisitIds = new Set(worked.map(a => String(a.visitId).trim()));
    const nonWorkableVisitIds = new Set(nonWorkable.map(a => String(a.visitId).trim()));
    
    // Pending = Assigned - Worked - Non-Workable
    return assigned.filter(a => {
      const visitId = String(a.visitId).trim();
      return !workedVisitIds.has(visitId) && !nonWorkableVisitIds.has(visitId);
    });
  } catch(e) {
    Logger.log('Error getting total pending accounts: ' + e.toString());
    return [];
  }
}

function getTotalNonWorkableAccounts() {
  try {
    const all = getAllAccountsData();
    const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
    const norm = (s) => String(s||'').toLowerCase().trim();
    
    return all.filter(a => {
      return norm(a.allocatedUser).includes(u) && norm(a.statusCode) === 'non-workable';
    });
  } catch(e) {
    Logger.log('Error getting total non-workable accounts: ' + e.toString());
    return [];
  }
}

function getAROutstandingAccounts() {
  try {
    const all = getAllAccountsData();
    const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
    const norm = (s) => String(s||'').toLowerCase().trim();
    
    const filtered = all.filter(a => norm(a.allocatedUser).includes(u) && a.workedDate);
    
    // Sort by worked date (recent to older)
    return filtered.sort((a, b) => {
      try {
        const dateA = new Date(a.workedDate).getTime();
        const dateB = new Date(b.workedDate).getTime();
        return dateB - dateA; // Descending order
      } catch(e) {
        return 0;
      }
    });
  } catch(e) {
    Logger.log('Error getting AR outstanding accounts: ' + e.toString());
    return [];
  }
}

// ============================================
// NON-WORKABLE APPROVAL SYSTEM
// ============================================

function getNonWorkableAccounts() {
  try {
    const all = getAllAccountsData();
    return all.filter(a => String(a.statusCode || '').toLowerCase() === 'non-workable');
  } catch(e) {
    Logger.log('Error getting non-workable accounts: ' + e.toString());
    return [];
  }
}

function approveNonWorkable(visitId, comment, approvedBy) {
  try {
    const account = getAccountByVisitId(visitId);
    if (!account) {
      return { success: false, message: 'Account not found' };
    }
    
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    let nonWorkableSheet = ss.getSheetByName('Non Workable');
    
    if (!nonWorkableSheet) {
      return { success: false, message: 'Non Workable sheet not found' };
    }
    
    const data = nonWorkableSheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[1] || '').trim() === String(visitId).trim());
    
    if (rowIndex > 0) {
      // Update approval columns (P, Q, R, S)
      nonWorkableSheet.getRange(rowIndex + 1, 16).setValue('Approved'); // Approval Status
      nonWorkableSheet.getRange(rowIndex + 1, 17).setValue(approvedBy); // Approved By
      nonWorkableSheet.getRange(rowIndex + 1, 18).setValue(comment || ''); // Approval Comment
      nonWorkableSheet.getRange(rowIndex + 1, 19).setValue(new Date()); // Approval Date
      
      SpreadsheetApp.flush();
      logActivity('NON_WORKABLE_APPROVED', `Approved ${visitId}`, 'SUCCESS');
      return { success: true, message: 'Non-workable account approved' };
    }
    
    return { success: false, message: 'Account not found in Non Workable sheet' };
  } catch(e) {
    Logger.log('Error approving non-workable: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function denyNonWorkable(visitId, comment, deniedBy) {
  try {
    const account = getAccountByVisitId(visitId);
    if (!account) {
      return { success: false, message: 'Account not found' };
    }
    
    const ss = SpreadsheetApp.openById(account.sourceSheetId);
    let nonWorkableSheet = ss.getSheetByName('Non Workable');
    
    if (!nonWorkableSheet) {
      return { success: false, message: 'Non Workable sheet not found' };
    }
    
    const data = nonWorkableSheet.getDataRange().getValues();
    const rowIndex = data.findIndex((r, i) => i > 0 && String(r[1] || '').trim() === String(visitId).trim());
    
    if (rowIndex > 0) {
      // Update approval columns
      nonWorkableSheet.getRange(rowIndex + 1, 16).setValue('Denied'); // Approval Status
      nonWorkableSheet.getRange(rowIndex + 1, 17).setValue(deniedBy); // Denied By
      nonWorkableSheet.getRange(rowIndex + 1, 18).setValue(comment || ''); // Denial Comment
      nonWorkableSheet.getRange(rowIndex + 1, 19).setValue(new Date()); // Denial Date
      
      SpreadsheetApp.flush();
      logActivity('NON_WORKABLE_DENIED', `Denied ${visitId}`, 'SUCCESS');
      return { success: true, message: 'Non-workable account denied' };
    }
    
    return { success: false, message: 'Account not found in Non Workable sheet' };
  } catch(e) {
    Logger.log('Error denying non-workable: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// PRODUCTION REPORTING
// ============================================

function initializeUserProductionReportSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.CMS_CONTROL_UNI.sheetId);
    let sheet = ss.getSheetByName(CONFIG.CMS_CONTROL_UNI.sheets.userProductionReport.name);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.CMS_CONTROL_UNI.sheets.userProductionReport.name);
      sheet.appendRow([
        'Email', 'User Name', 'Production Sheet ID', 'Production Sheet Name', 'Last Updated'
      ]);
      // Format header
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }
    
    return sheet;
  } catch(e) {
    Logger.log('Error initializing User Production Report sheet: ' + e.toString());
    throw new Error('Failed to initialize User Production Report sheet');
  }
}

function getUserProductionSheet(email) {
  try {
    const sheet = initializeUserProductionReportSheet();
    const data = sheet.getDataRange().getValues();
    const userEmail = email.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase() === userEmail) {
        return {
          sheetId: data[i][2] || '',
          sheetName: data[i][3] || 'Production Report'
        };
      }
    }
    
    return null;
  } catch(e) {
    Logger.log('Error getting user production sheet: ' + e.toString());
    return null;
  }
}

function postToProductionReport(visitId, formData) {
  try {
    const user = Session.getActiveUser();
    const userEmail = user.getEmail();
    const userName = getCurrentUserDisplayName();
    
    const prodSheet = getUserProductionSheet(userEmail);
    if (!prodSheet || !prodSheet.sheetId) {
      Logger.log('No production sheet configured for user: ' + userEmail);
      return { success: false, message: 'Production sheet not configured' };
    }
    
    const account = getAccountByVisitId(visitId);
    if (!account) {
      return { success: false, message: 'Account not found' };
    }
    
    const ss = SpreadsheetApp.openById(prodSheet.sheetId);
    let sheet = ss.getSheetByName(prodSheet.sheetName);
    
    if (!sheet) {
      // Try to find sheet by partial name match
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => 
        s.getName().toLowerCase().includes(prodSheet.sheetName.toLowerCase()) ||
        s.getName().toLowerCase().includes('production')
      );
      
      if (!sheet) {
        return { success: false, message: 'Production sheet not found' };
      }
    }
    
    // Check if header row exists, if not create it
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Client', 'State', 'Account#', 'Patient Name', 'Insurance Name', 'DOS', 'Year', 
        'Aging Days', 'Aging Bucket', 'S. Date', 'S. Aging Days', 'S. Aging Bucket', 
        'Status', 'Billed Amount', 'Balance Amount', 'Insurance Type', 'AR Comments', 
        'Type (Ins Call/Email/Analysis/Portal)', 'Status Code', 'Action Code', 
        'Assigned To', 'Worked By', 'Worked Date', 'Follow Up Date', 'Service Type', 
        'Claim Type(Workable/Non Workable)', 'Allocated Date', 'Allocated User Name', 'Remarks'
      ]);
    }
    
    // Calculate year from DOS
    let year = '';
    if (account.dos) {
      try {
        const dosDate = new Date(account.dos);
        year = dosDate.getFullYear().toString();
      } catch(e) {
        year = '';
      }
    }
    
    // Calculate S. Aging Days and S. Aging Bucket from Submitted Date
    let sAgingDays = '';
    let sAgingBucket = '';
    if (account.submittedDate) {
      try {
        const subDate = new Date(account.submittedDate);
        const today = new Date();
        const diffTime = today.getTime() - subDate.getTime();
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        sAgingDays = diffDays.toString();
        
        // Calculate aging bucket
        if (diffDays < 30) sAgingBucket = '0-30';
        else if (diffDays < 60) sAgingBucket = '31-60';
        else if (diffDays < 90) sAgingBucket = '61-90';
        else if (diffDays < 120) sAgingBucket = '91-120';
        else sAgingBucket = '120+';
      } catch(e) {
        sAgingDays = '';
        sAgingBucket = '';
      }
    }
    
    // Append row with all mapped columns
    sheet.appendRow([
      account.client || '',
      account.state || '',
      account.visitId || '',
      account.patientName || '',
      account.insurance || '',
      account.dos || '',
      year,
      account.agingDays || '',
      account.agingBucket || '',
      account.submittedDate || '',
      sAgingDays,
      sAgingBucket,
      account.status || '',
      account.billedAmount || 0,
      account.balanceAmount || 0,
      account.primarySecondary || '',
      formData.notes || account.arNotes || '',
      formData.source || account.type || '',
      formData.statusCode || account.statusCode || '',
      formData.actionCode || account.actionCode || '',
      formData.assignedTo || account.assignedTo || '',
      userName,
      new Date(),
      formData.followupDate || account.followupDate || '',
      account.type || '',
      account.claimType || '',
      account.allocationDate || '',
      account.allocatedUser || '',
      formData.remarks || account.remarks || ''
    ]);
    
    SpreadsheetApp.flush();
    logActivity('PRODUCTION_POSTED', `Posted ${visitId} to production report`, 'SUCCESS');
    return { success: true, message: 'Posted to production report successfully' };
  } catch(e) {
    Logger.log('Error posting to production report: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// API CLIENT WRAPPER FUNCTIONS
// Bridge between frontend and ARConnectAPI.gs
// ============================================

/**
 * Get the API base URL from script properties or return empty string
 */
function getApiBaseUrl() {
  if (CONFIG.API.baseUrl) {
    return CONFIG.API.baseUrl;
  }
  
  // Try to get from script properties
  const apiUrl = PROPERTIES.getProperty('API_BASE_URL');
  if (apiUrl) {
    CONFIG.API.baseUrl = apiUrl;
    return apiUrl;
  }
  
  return '';
}

/**
 * Set the API base URL
 */
function setApiBaseUrl(url) {
  CONFIG.API.baseUrl = url;
  PROPERTIES.setProperty('API_BASE_URL', url);
  return { success: true, message: 'API URL configured' };
}

/**
 * Make API request to ARConnectAPI
 */
function callApi(action, params = {}, method = 'GET') {
  try {
    const apiUrl = getApiBaseUrl();
    if (!apiUrl) {
      throw new Error('API URL not configured. Please set API base URL first.');
    }
    
    const url = apiUrl + '?action=' + encodeURIComponent(action) + 
                '&origin=' + encodeURIComponent(ScriptApp.getService().getUrl()) +
                Object.keys(params).map(key => 
                  '&' + encodeURIComponent(key) + '=' + encodeURIComponent(params[key])
                ).join('');
    
    const options = {
      method: method,
      muteHttpExceptions: true
    };
    
    if (method === 'POST') {
      options.contentType = 'application/json';
      options.payload = JSON.stringify(params);
    }
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    try {
      return JSON.parse(responseText);
    } catch(e) {
      return { error: 'Invalid JSON response', response: responseText };
    }
  } catch(e) {
    Logger.log('API call error: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Wrapper functions that use API when enabled, otherwise use direct functions
 */

// Get all accounts - uses API if enabled
function getAllAccountsViaAPI() {
  if (CONFIG.API.useApi()) {
    const result = callApi('getAccounts');
    // API returns array directly or error object
    if (result.error) {
      Logger.log('API Error: ' + result.error);
      // Fallback to direct call if API fails
      return getAllAccountsData();
    }
    // Ensure result is an array
    return Array.isArray(result) ? result : (result.data || []);
  }
  return getAllAccountsData();
}

// Get account by visit ID - uses API if enabled
function getAccountByVisitIdViaAPI(visitId) {
  if (CONFIG.API.useApi()) {
    return callApi('getAccountByVisitId', { visitId: visitId });
  }
  return getAccountByVisitId(visitId);
}

// Search accounts - uses API if enabled
function searchAccountsViaAPI(searchTerm, searchType, filters) {
  if (CONFIG.API.useApi()) {
    const params = { searchTerm: searchTerm, searchType: searchType, ...filters };
    return callApi('search', params);
  }
  return searchAccounts(searchTerm, searchType, filters);
}

// Convert sheet to JSON - uses API
function convertSheetToJSONViaAPI(sheetId, sheetName) {
  if (CONFIG.API.useApi()) {
    return callApi('convertSheet', { sheetId: sheetId, sheetName: sheetName });
  }
  // Fallback: return error if API not configured
  return { error: 'API not configured. Please set API base URL.' };
}

// Convert all client sheets to JSON - uses API
function convertAllClientSheetsToJSONViaAPI() {
  if (CONFIG.API.useApi()) {
    return callApi('convertAllSheets');
  }
  return { error: 'API not configured. Please set API base URL.' };
}

// Get stored JSON data - uses API
function getStoredJSONDataViaAPI(clientName) {
  if (CONFIG.API.useApi()) {
    return callApi('getStoredJSON', { client: clientName || '' });
  }
  return { error: 'API not configured. Please set API base URL.' };
}

// Refresh JSON data - uses API
function refreshAllJSONDataViaAPI() {
  if (CONFIG.API.useApi()) {
    return callApi('refreshJSON', {}, 'POST');
  }
  return { error: 'API not configured. Please set API base URL.' };
}

// ============================================
// INCREMENTAL JSON CACHE UPDATE FUNCTIONS
// ============================================

/**
 * Update account in local cache and sync to API JSON cache
 * This provides fast local updates and propagates to all users via API
 */
function updateAccountInCache(visitId, updatedAccount) {
  try {
    const start = Date.now();
    
    // 1. Update local cache (immediate, <50ms)
    const cacheKey = 'all_accounts';
    let allAccounts = getCached(cacheKey);
    
    if (!allAccounts || !Array.isArray(allAccounts)) {
      // Load from PropertiesService if cache is empty
      allAccounts = getAllAccountsData();
    }
    
    if (allAccounts && Array.isArray(allAccounts)) {
      const index = allAccounts.findIndex(a => String(a.visitId || '').trim() === String(visitId).trim());
      
      if (index >= 0) {
        allAccounts[index] = updatedAccount;
      } else {
        allAccounts.push(updatedAccount);
      }
      
      // Update local cache
      setCache(cacheKey, allAccounts);
      
      // 2. Update PropertiesService (persistent storage)
      storeAccountsInPropertiesLocal(allAccounts);
      
      // 3. Update JSON cache for API access
      updateAccountInJSONCache(visitId, updatedAccount);
      
      const duration = Date.now() - start;
      Logger.log(`Cache update: ${visitId} updated in ${duration}ms`);
    }
  } catch(e) {
    Logger.log('Error updating account in cache: ' + e.toString());
  }
}

/**
 * Store accounts in PropertiesService (local version)
 */
function storeAccountsInPropertiesLocal(accounts) {
  try {
    const dataStr = JSON.stringify(accounts);
    const chunks = [];
    
    if (dataStr.length > MAX_PROPERTY_SIZE) {
      for (let i = 0; i < dataStr.length; i += MAX_PROPERTY_SIZE) {
        chunks.push(dataStr.substring(i, i + MAX_PROPERTY_SIZE));
      }
    } else {
      chunks.push(dataStr);
    }
    
    chunks.forEach((chunk, idx) => {
      PROPERTIES.setProperty(`${ACCOUNTS_DATA_KEY}_${idx}`, chunk);
    });
    
    PROPERTIES.setProperty(ACCOUNTS_META_KEY, JSON.stringify({
      count: accounts.length,
      chunks: chunks.length,
      timestamp: Date.now()
    }));
    
    // Clean up old chunks
    let i = chunks.length;
    while (PROPERTIES.getProperty(`${ACCOUNTS_DATA_KEY}_${i}`)) {
      PROPERTIES.deleteProperty(`${ACCOUNTS_DATA_KEY}_${i}`);
      i++;
    }
  } catch(e) {
    Logger.log('Error storing accounts in Properties: ' + e.toString());
  }
}

// ============================================
// JSON CONVERSION FUNCTIONS (Merged from ARConnectAPI.gs)
// ============================================

/**
 * Convert a sheet to JSON format with optimized read operations
 */
function convertSheetToJSON(sheetId, sheetName) {
  try {
    if (!sheetId || !sheetName) {
      return { error: 'Sheet ID and Sheet Name are required' };
    }
    
    const cacheKey = `json_${sheetId}_${sheetName}`;
    const cached = getCached(cacheKey);
    if (cached) {
      return { success: true, data: cached, cached: true };
    }
    
    const start = Date.now();
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName(sheetName);
    
    // Try to find sheet by partial match if exact match fails
    if (!sheet) {
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => 
        s.getName().toLowerCase() === sheetName.toLowerCase() ||
        s.getName().toLowerCase().includes(sheetName.toLowerCase())
      );
    }
    
    if (!sheet) {
      return { error: `Sheet "${sheetName}" not found in spreadsheet ${sheetId}` };
    }
    
    // Optimized batch read - get all data at once
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) {
      return { success: true, data: [], message: 'Sheet is empty or has no data rows' };
    }
    
    // Read headers (row 1)
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Read all data rows at once (batch operation for performance)
    const dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Convert to JSON array
    const jsonData = dataRows.map((row, rowIndex) => {
      const obj = {
        _rowNumber: rowIndex + 2, // Actual row number in sheet
        _visitId: row[1] || '', // Visit ID from column B (index 1)
      };
      
      // Map each column to its header
      headers.forEach((header, colIndex) => {
        const headerKey = String(header || '').trim() || `Column_${colIndex + 1}`;
        obj[headerKey] = row[colIndex] || '';
      });
      
      return obj;
    });
    
    // Store in cache
    setCache(cacheKey, jsonData, CACHE_DURATION);
    
    // Also store in PropertiesService for persistence
    storeJSONInProperties(cacheKey, jsonData);
    
    const duration = Date.now() - start;
    Logger.log(`Converted sheet ${sheetName} to JSON: ${jsonData.length} rows in ${duration}ms`);
    
    return {
      success: true,
      data: jsonData,
      metadata: {
        sheetId: sheetId,
        sheetName: sheetName,
        rowCount: jsonData.length,
        columnCount: headers.length,
        conversionTime: duration,
        timestamp: new Date().toISOString()
      }
    };
  } catch(e) {
    Logger.log('Error converting sheet to JSON: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Convert all sheets from AR Client Config to JSON
 * Special handling: reads referenced sheets and extracts visit IDs from column B
 */
function convertAllClientSheetsToJSON() {
  try {
    const start = Date.now();
    const clients = loadARClientSheets();
    const results = {};
    let totalRows = 0;
    
    Object.keys(clients).forEach(clientName => {
      try {
        const cfg = clients[clientName];
        const result = convertSheetToJSON(cfg.sheetId, cfg.sheetName);
        
        if (result.success && result.data) {
          results[clientName] = {
            success: true,
            sheetId: cfg.sheetId,
            sheetName: cfg.sheetName,
            rowCount: result.data.length,
            visitIds: result.data.map(row => row._visitId).filter(id => id),
            data: result.data
          };
          totalRows += result.data.length;
        } else {
          results[clientName] = {
            success: false,
            error: result.error || 'Unknown error',
            sheetId: cfg.sheetId,
            sheetName: cfg.sheetName
          };
        }
      } catch(e) {
        Logger.log(`Error processing client ${clientName}: ${e.toString()}`);
        results[clientName] = {
          success: false,
          error: e.toString(),
          client: clientName
        };
      }
    });
    
    const duration = Date.now() - start;
    
    // Store aggregated results
    const cacheKey = 'all_clients_json';
    setCache(cacheKey, results, CACHE_DURATION);
    storeJSONInProperties(cacheKey, results);
    
    return {
      success: true,
      clients: results,
      summary: {
        totalClients: Object.keys(clients).length,
        successful: Object.values(results).filter(r => r.success).length,
        failed: Object.values(results).filter(r => !r.success).length,
        totalRows: totalRows,
        conversionTime: duration,
        timestamp: new Date().toISOString()
      }
    };
  } catch(e) {
    Logger.log('Error converting all client sheets: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Get stored JSON data for a specific client
 */
function getStoredJSONData(clientName) {
  try {
    if (!clientName) {
      // Return all stored data
      const allData = getCached('all_clients_json');
      if (allData) {
        return { success: true, data: allData };
      }
      return { error: 'No stored data found. Please run convert-all first.' };
    }
    
    const clients = loadARClientSheets();
    if (!clients[clientName]) {
      return { error: `Client "${clientName}" not found in AR Client Config` };
    }
    
    const cfg = clients[clientName];
    const cacheKey = `json_${cfg.sheetId}_${cfg.sheetName}`;
    const cached = getCached(cacheKey);
    
    if (cached) {
      return {
        success: true,
        client: clientName,
        data: cached,
        metadata: {
          sheetId: cfg.sheetId,
          sheetName: cfg.sheetName,
          rowCount: cached.length
        }
      };
    }
    
    // If not cached, convert now
    return convertSheetToJSON(cfg.sheetId, cfg.sheetName);
  } catch(e) {
    Logger.log('Error getting stored JSON data: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Refresh JSON data for all clients
 */
function refreshAllJSONData() {
  try {
    // Clear cache
    const clients = loadARClientSheets();
    Object.keys(clients).forEach(clientName => {
      const cfg = clients[clientName];
      const cacheKey = `json_${cfg.sheetId}_${cfg.sheetName}`;
      CACHE.remove(cacheKey);
    });
    CACHE.remove('all_clients_json');
    
    // Reconvert all
    return convertAllClientSheetsToJSON();
  } catch(e) {
    Logger.log('Error refreshing JSON data: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Store large JSON data in PropertiesService (chunked if needed)
 */
function storeJSONInProperties(key, data) {
  try {
    const dataStr = JSON.stringify(data);
    const chunks = [];
    
    // Split into chunks if data is too large
    if (dataStr.length > MAX_PROPERTY_SIZE) {
      for (let i = 0; i < dataStr.length; i += MAX_PROPERTY_SIZE) {
        chunks.push(dataStr.substring(i, i + MAX_PROPERTY_SIZE));
      }
    } else {
      chunks.push(dataStr);
    }
    
    // Store chunks
    chunks.forEach((chunk, idx) => {
      PROPERTIES.setProperty(`${key}_prop_${idx}`, chunk);
    });
    
    // Store metadata
    PROPERTIES.setProperty(`${key}_meta`, JSON.stringify({
      chunks: chunks.length,
      timestamp: Date.now(),
      size: dataStr.length
    }));
    
    // Clean up old chunks
    let i = chunks.length;
    while (PROPERTIES.getProperty(`${key}_prop_${i}`)) {
      PROPERTIES.deleteProperty(`${key}_prop_${i}`);
      i++;
    }
  } catch(e) {
    Logger.log('Error storing JSON in Properties: ' + e.toString());
  }
}

/**
 * Get JSON data from PropertiesService (reconstruct from chunks)
 */
function getJSONFromProperties(key) {
  try {
    const metaStr = PROPERTIES.getProperty(`${key}_meta`);
    if (!metaStr) return null;
    
    const meta = JSON.parse(metaStr);
    const chunks = [];
    
    for (let i = 0; i < meta.chunks; i++) {
      const chunk = PROPERTIES.getProperty(`${key}_prop_${i}`);
      if (!chunk) return null;
      chunks.push(chunk);
    }
    
    const dataStr = chunks.join('');
    return JSON.parse(dataStr);
  } catch(e) {
    Logger.log('Error retrieving JSON from Properties: ' + e.toString());
    return null;
  }
}

// ============================================
// INCREMENTAL JSON UPDATE FUNCTIONS
// Optimized write-on-change pattern
// ============================================

/**
 * Update cache timestamp to notify other users of changes
 */
function updateCacheTimestamp() {
  try {
    const timestamp = Date.now();
    PROPERTIES.setProperty('json_last_update', timestamp.toString());
    CACHE.put('cache_timestamp', timestamp.toString(), CACHE_DURATION);
    return timestamp;
  } catch(e) {
    Logger.log('Error updating cache timestamp: ' + e.toString());
    return null;
  }
}

/**
 * Get last update timestamp
 */
function getLastUpdateTimestamp() {
  try {
    const timestamp = PROPERTIES.getProperty('json_last_update');
    return timestamp ? parseInt(timestamp) : 0;
  } catch(e) {
    return 0;
  }
}

/**
 * Check if cache needs refresh based on timestamp
 */
function shouldRefreshCache() {
  try {
    const lastUpdate = getLastUpdateTimestamp();
    const cacheTimeStr = CACHE.get('cache_timestamp');
    const cacheTime = cacheTimeStr ? parseInt(cacheTimeStr) : 0;
    return lastUpdate > cacheTime;
  } catch(e) {
    return true; // Refresh if we can't determine
  }
}

/**
 * Incrementally update a single account in JSON cache
 * This is much faster than rewriting entire JSON
 */
function updateAccountInJSONCache(visitId, updatedAccount) {
  try {
    const start = Date.now();
    
    // Update in CacheService (fast, immediate)
    const cacheKey = 'api_all_accounts';
    let allAccounts = getCached(cacheKey);
    
    if (!allAccounts || !Array.isArray(allAccounts)) {
      // If cache is empty, try to load from PropertiesService
      allAccounts = getAccountsFromPropertiesLocal();
      if (!allAccounts || !Array.isArray(allAccounts)) {
        // Try loading from main cache
        allAccounts = getAllAccountsData();
        if (!allAccounts || !Array.isArray(allAccounts)) {
          Logger.log('No accounts data found to update');
          return { success: false, message: 'No accounts data found' };
        }
      }
    }
    
    // Find and update the account
    const index = allAccounts.findIndex(a => String(a.visitId || '').trim() === String(visitId).trim());
    
    if (index >= 0) {
      // Update the account in the array
      allAccounts[index] = updatedAccount;
      
      // Update CacheService immediately (<50ms)
      const dataStr = JSON.stringify(allAccounts);
      if (dataStr.length < 100000) {
        CACHE.put(cacheKey, dataStr, CACHE_DURATION);
      }
      
      // Update timestamp
      updateCacheTimestamp();
      
      // Update PropertiesService in background (non-blocking)
      updateAccountsInPropertiesAsync(allAccounts);
      
      const duration = Date.now() - start;
      Logger.log(`Incremental JSON update: ${visitId} updated in ${duration}ms`);
      
      return { success: true, duration: duration };
    } else {
      // Account not found, add it
      allAccounts.push(updatedAccount);
      
      // Update cache
      const dataStr = JSON.stringify(allAccounts);
      if (dataStr.length < 100000) {
        CACHE.put(cacheKey, dataStr, CACHE_DURATION);
      }
      
      updateCacheTimestamp();
      updateAccountsInPropertiesAsync(allAccounts);
      
      return { success: true, message: 'Account added' };
    }
  } catch(e) {
    Logger.log('Error updating account in JSON cache: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Update accounts in PropertiesService asynchronously (non-blocking)
 */
function updateAccountsInPropertiesAsync(accounts) {
  try {
    const start = Date.now();
    storeAccountsInPropertiesLocal(accounts);
    const duration = Date.now() - start;
    Logger.log(`PropertiesService update completed in ${duration}ms (background)`);
  } catch(e) {
    Logger.log('Error in async PropertiesService update: ' + e.toString());
  }
}

/**
 * Update account in JSON cache via API endpoint
 */
function updateAccountJSON(visitId, accountData) {
  try {
    if (!visitId || !accountData) {
      return { error: 'Visit ID and account data are required' };
    }
    
    const result = updateAccountInJSONCache(visitId, accountData);
    
    if (result.success) {
      return {
        success: true,
        message: 'Account updated in JSON cache',
        visitId: visitId,
        duration: result.duration,
        timestamp: getLastUpdateTimestamp()
      };
    } else {
      return { error: result.message || 'Update failed' };
    }
  } catch(e) {
    Logger.log('Error in updateAccountJSON: ' + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Get cache status and last update time
 */
function getCacheStatus() {
  try {
    const lastUpdate = getLastUpdateTimestamp();
    const cacheKey = 'api_all_accounts';
    const cached = getCached(cacheKey);
    const accountCount = cached && Array.isArray(cached) ? cached.length : 0;
    
    return {
      success: true,
      lastUpdate: lastUpdate,
      lastUpdateFormatted: lastUpdate ? new Date(lastUpdate).toISOString() : null,
      accountCount: accountCount,
      cacheValid: accountCount > 0,
      needsRefresh: shouldRefreshCache()
    };
  } catch(e) {
    return { error: e.toString() };
  }
}

/**
 * Get accounts from PropertiesService (local version)
 */
function getAccountsFromPropertiesLocal() {
  try {
    const metaStr = PROPERTIES.getProperty(ACCOUNTS_META_KEY);
    if (!metaStr) return null;
    
    const meta = JSON.parse(metaStr);
    const chunks = [];
    
    for (let i = 0; i < meta.chunks; i++) {
      const chunk = PROPERTIES.getProperty(`${ACCOUNTS_DATA_KEY}_${i}`);
      if (!chunk) return null;
      chunks.push(chunk);
    }
    
    const dataStr = chunks.join('');
    return JSON.parse(dataStr);
  } catch(e) {
    Logger.log('Error retrieving accounts from Properties: ' + e.toString());
    return null;
  }
}

/**
 * Create JSON response for API
 */
function createResponse(data, statusCode = 200) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
