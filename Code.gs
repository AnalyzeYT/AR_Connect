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
  }
};

// Column mapping based on actual sheet structure:
// Client(0), State(1), VisitID#(2), Patient Name(3), DOS(4), Aging Days(5), Aging Bucket(6), 
// Submitted Date(7), Insurance Name(8), Status(9), Billed Amount(10), Balance Amount(11), 
// Primary/Secondary(12), AR notes(13), Type(14), Status Code(15), Action Code(16), 
// Assigned To(17), Worked By(18), Worked Date(19), Follow up date(20), 
// Claim Type(21), Allocated User(22), Allocation Date(23), Remarks(24)
const COLUMNS = {
  CLIENT: 0, STATE: 1, VISIT_ID: 2, PATIENT_NAME: 3, DOS: 4,
  AGING_DAYS: 5, AGING_BUCKET: 6, SUBMITTED_DATE: 7, INSURANCE_NAME: 8,
  STATUS: 9, BILLED_AMOUNT: 10, BALANCE_AMOUNT: 11, PRIMARY_SECONDARY: 12,
  AR_NOTES: 13, TYPE: 14, STATUS_CODE: 15, ACTION_CODE: 16,
  ASSIGNED_TO: 17, WORKED_BY: 18, WORKED_DATE: 19, FOLLOWUP_DATE: 20,
  CLAIM_TYPE: 21, ALLOCATED_USER: 22, ALLOCATION_DATE: 23, REMARKS: 24
};

const CACHE = CacheService.getScriptCache();
const CACHE_DURATION = 300;

function getCached(key) {
  const cached = CACHE.get(key);
  return cached ? JSON.parse(cached) : null;
}

function setCache(key, data) {
  try { CACHE.put(key, JSON.stringify(data), CACHE_DURATION); } 
  catch(e) { Logger.log('Cache error: ' + e); }
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
    
    if (!sheet) throw new Error('AR Client Config sheet not found. Please check sheet name.');
    
    // Read from range A1:C (header + data)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('No data in AR Client Config (need at least header + 1 row)');
    
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
    
    if (Object.keys(clients).length === 0) {
      throw new Error('No valid configs found. Check A2:C range has Client, Sheet ID, Sheet Name columns.');
    }
    
    setCache(cacheKey, clients);
    logActivity('CONFIG_LOAD', `Loaded ${Object.keys(clients).length} clients`);
    return clients;
  } catch(e) {
    logActivity('CONFIG_LOAD', e.message, 'ERROR');
    throw new Error('Config load error: ' + e.message);
  }
}

function getAllAccountsData(forceRefresh = false) {
  const cacheKey = 'all_accounts';
  if (!forceRefresh) {
    const cached = getCached(cacheKey);
    if (cached) { 
      logActivity('DATA_LOAD', 'Cache hit - ' + cached.length + ' accounts');
      return cached; 
    }
  }
  
  try {
    const start = Date.now();
    const clients = loadARClientSheets();
    let allAccounts = [];
    let processedRows = 0;
    let skippedRows = 0;
    
    Object.keys(clients).forEach(clientName => {
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
        
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        
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
            primarySecondary: row[COLUMNS.PRIMARY_SECONDARY] || '',
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
  const allData = getAllAccountsData();
  const account = allData.find(a => String(a.visitId).trim() === String(visitId).trim());
  if (account) logActivity('ACCOUNT_VIEW', `VisitID: ${visitId}`);
  return account || null;
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
      workedDate: formatDate(r[8]), rawTimestamp: r[0]
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
    const updates = [];
    
    // Batch all updates
    if (formData.notes !== undefined) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.AR_NOTES + 1), value: formData.notes});
    }
    if (formData.source) {
      updates.push({range: sheet.getRange(rowIndex, COLUMNS.TYPE + 1), value: formData.source});
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
    
    // Post to Non Workable sheet if status is Non-Workable
    if (formData.statusCode === 'Non-Workable') {
      postToNonWorkableSheet(ss, account, formData, userName, now);
    }
    
    // Single flush after all updates
    SpreadsheetApp.flush();
    
    CACHE.remove('all_accounts');
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
        'Worked By', 'Worked Date', 'Allocated User', 'Allocation Date'
      ]);
      // Format header
      const headerRange = nonWorkableSheet.getRange(1, 1, 1, 15);
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
      account.allocationDate || ''
    ];
    
    if (existingRow > 0) {
      // Update existing row
      nonWorkableSheet.getRange(existingRow + 1, 1, 1, 15).setValues([rowData]);
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
    
    // Update remarks in AR outstanding sheet
    sheet.getRange(rowIndex, COLUMNS.REMARKS + 1).setValue(remarks);
    
    // If account is Non-Workable, also update in Non Workable sheet
    if (account.statusCode === 'Non-Workable') {
      updateNonWorkableSheetRemarks(ss, visitId, remarks);
    }
    
    SpreadsheetApp.flush();
    CACHE.remove('all_accounts');
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
  const all = getAllAccountsData();
  const u = String(getCurrentUserDisplayName()||'').toLowerCase().trim();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const f = (cond) => all.filter(a => String(a.allocatedUser||'').toLowerCase().trim().includes(u) && cond(a)).length;
  
  // Total Worked: Worked By = User name AND Worked Date = Today
  const totalWorked = all.filter(a => {
    const allocatedMatch = String(a.allocatedUser||'').toLowerCase().trim().includes(u);
    const workedByMatch = String(a.workedBy||'').toLowerCase().trim() === u;
    if (!allocatedMatch || !workedByMatch || !a.workedDate) return false;
    
    const workedDate = new Date(a.workedDate);
    workedDate.setHours(0, 0, 0, 0);
    return workedDate.getTime() === today.getTime();
  }).length;
  
  // Pending: Worked By = Username AND (Worked Date = Blank OR Allocation Date - Worked Date > 0)
  const totalPending = all.filter(a => {
    const allocatedMatch = String(a.allocatedUser||'').toLowerCase().trim().includes(u);
    const workedByMatch = String(a.workedBy||'').toLowerCase().trim() === u;
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
  }).length;
  
  return {
    totalAssigned: f(()=>true),
    totalWorked: totalWorked,
    totalPending: totalPending,
    totalNonWorkable: f(a=>a.statusCode==='Non-Workable'),
    totalAccounts: all.length
  };
}

function getAssignedAccounts() {
  const all = getAllAccountsData();
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
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const startRow = Math.max(2, lastRow - limit + 1);
    const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 13).getValues();
    
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
    const logs = getActivityLogs(1000);
    const sessionTimeout = 30 * 60 * 1000; // 30 minutes
    const now = new Date().getTime();
    const activeSessions = {};
    
    logs.forEach(log => {
      if (log.sessionId && log.rawTimestamp) {
        const logTime = new Date(log.rawTimestamp).getTime();
        if (logTime > (now - sessionTimeout)) {
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
              userAgent: log.userAgent || ''
            };
          }
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