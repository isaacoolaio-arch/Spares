// ============================================================
// OOLA SPARES - Google Apps Script Backend
// Motorcycle Spare Parts Management System
// Isaac Oola - Okum Energy Group
// ============================================================

var SPREADSHEET_ID = '1Hk3uMUW3trcGGG6mc11bkP5Yjga7X3X-eJikBTbOAyw';

var SHEETS = {
  PARTS: 'Parts',
  SUPPLIERS: 'Suppliers',
  SUPPLIER_RATES: 'SupplierRates',
  PURCHASES: 'Purchases',
  SALES: 'Sales',
  CHECKLIST: 'Checklist',
  SETTINGS: 'Settings',
  NOTIFICATIONS: 'Notifications',
  CUSTOMERS: 'Customers'
};

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const postData = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  const data = Object.assign({}, params, postData);
  const action = data.action;

  let result;
  try {
    switch (action) {
      // AUTH
      case 'verifyPin': result = verifyPin(data); break;
      case 'getUsers': result = getUsers(data); break;
      case 'generateEditOtp': result = generateEditOtp(data); break;
      case 'requestEditOtp':  result = requestEditOtp(data); break;
      case 'sendOtpNotification':  result = sendOtpNotification(data); break;
      case 'getMyNotifications':   result = getMyNotifications(data); break;
      case 'markNotificationRead': result = markNotificationRead(data); break;
      case 'verifyEditOtp': result = verifyEditOtp(data); break;
      case 'getSettings': result = getSettings(); break;
      case 'saveSettings': result = saveSettings(data); break;

      // PARTS
      case 'getParts': result = getParts(); break;
      case 'savePart': result = savePart(data); break;
      case 'deletePart': result = deletePart(data); break;

      // SUPPLIERS
      case 'getSuppliers': result = getSuppliers(); break;
      case 'saveSupplier': result = saveSupplier(data); break;
      case 'deleteSupplier': result = deleteSupplier(data); break;

      // SUPPLIER RATES
      case 'getSupplierRates': result = getSupplierRates(data); break;
      case 'saveSupplierRate': result = saveSupplierRate(data); break;
      case 'deleteSupplierRate': result = deleteSupplierRate(data); break;

      // PURCHASES
      case 'getPurchases': result = getPurchases(data); break;
      case 'getAll':          result = getAll(); break;
      case 'getCoreData':     result = getCoreData(); break;
      case 'getTransactions': result = getTransactions(); break;
      case 'fixOpeningStockDoubling': result = fixOpeningStockDoubling(); break;
      case 'saveBarcode':   result = saveBarcode(data); break;
      case 'savePartImage':   result = savePartImage(data); break;
      case 'removePartImage': result = removePartImage(data); break;
      case 'fixImageUrls':    result = fixImageUrls(); break;
      case 'copyPartImageUrl':result = copyPartImageUrl(data); break;
      case 'getCustomers':        result = getCustomers(); break;
      case 'saveCustomer':        result = saveCustomer(data); break;
      case 'deleteCustomer':      result = deleteCustomer(data); break;
      case 'getCustomerHistory':  result = getCustomerHistory(data); break;
      case 'savePurchase': result = savePurchase(data); break;
      case 'updatePurchase': result = updatePurchase(data); break;
      case 'deletePurchase': result = deletePurchase(data); break;

      // SALES
      case 'getSales': result = getSales(data); break;
      case 'saveSale': result = saveSale(data); break;
      case 'updateSale': result = updateSale(data); break;
      case 'deleteSale': result = deleteSale(data); break;

      // CHECKLIST
      case 'getChecklist': result = getChecklist(); break;
      case 'saveChecklistItem':    result = saveChecklistItem(data); break;
      case 'bulkSaveChecklist':    result = bulkSaveChecklist(data); break;
      case 'updateChecklistStatus': result = updateChecklistStatus(data); break;
      case 'deleteChecklistItem': result = deleteChecklistItem(data); break;
      case 'clearAllChecklist':    result = clearAllChecklist(); break;

      // DASHBOARD
      case 'getDashboard': result = getDashboard(); break;

      // SETUP
      case 'setupSheets': result = setupSheets(); break;
      case 'clearParts': result = clearParts(); break;
      case 'migratePartsSheet': result = migratePartsSheet(); break;
      case 'fixMovementColumn': result = fixMovementColumn(); break;
      case 'cleanPartsData': result = cleanPartsData(); break;
      case 'fixStockQty': result = fixStockQty(data); break;
      case 'fixMixedColumns': result = fixMixedColumns(); break;
      case 'reorderPartsColumns': result = reorderPartsColumns(); break;
      case 'sortParts': sortPartsSheet(); result = {success:true, message:'Parts sorted'}; break;
      case 'removeOfflineDuplicates': result = removeOfflineDuplicates(); break;
      case 'fixAllIssues': result = fixAllIssues(); break;
      case 'resetHighQty': result = resetHighQty(); break;
      case 'fixStockData': result = fixStockData(); break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SETUP
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const configs = [
    { name: SHEETS.PARTS, headers: ['PartID','ItemName','Size','Brand','Category','CostPrice','SellingPrice','StockQty','ReorderLevel','Unit','Movement','SupplierName','CreatedAt','Barcode','ImageUrl'] },
    { name: SHEETS.SUPPLIERS, headers: ['SupplierID','Name','Contact','Location','Email','Notes','CreatedAt'] },
    { name: SHEETS.SUPPLIER_RATES, headers: ['RateID','SupplierID','SupplierName','PartID','PartName','UnitPrice','Notes','UpdatedAt'] },
    { name: SHEETS.PURCHASES, headers: ['PurchaseID','Date','PartID','PartName','SupplierID','SupplierName','Qty','CostPrice','TotalCost','Notes'] },
    { name: SHEETS.SALES, headers: ['SaleID','Date','PartID','PartName','Qty','SellingPrice','TotalAmount','CustomerID','CustomerName','Notes','RecordedBy'] },
    { name: SHEETS.CHECKLIST, headers: ['ItemID','PartID','PartName','TargetQty','Status','Priority','Notes','UpdatedAt'] },
    { name: SHEETS.SETTINGS, headers: ['Key','Value'] },
    { name: SHEETS.NOTIFICATIONS, headers: ['NotifID','ToUser','FromUser','Type','Message','OTP','CreatedAt','Read'] },
    { name: SHEETS.CUSTOMERS, headers: ['CustomerID','Name','Phone','Email','Address','Notes','CreatedAt','TotalPurchases','TotalSpent'] }
  ];

  configs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
      sheet.appendRow(cfg.headers);
      sheet.getRange(1, 1, 1, cfg.headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(cfg.headers);
      sheet.getRange(1, 1, 1, cfg.headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    } else {
      // Sheet exists with data — add any missing columns
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      cfg.headers.forEach(h => {
        if (!existingHeaders.includes(h)) {
          const newCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newCol).setValue(h).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
          Logger.log('Added missing column: ' + h + ' to ' + cfg.name);
        }
      });
    }
  });

  // Default settings
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  const existing = settingsSheet.getLastRow();
  if (existing <= 1) {
    settingsSheet.appendRow(['shopName', 'Oola Spares']);
    settingsSheet.appendRow(['pin', '1234']); // stored as text
    // Force pin cell to text format
    const pinRow = settingsSheet.getLastRow();
    settingsSheet.getRange(pinRow, 2).setNumberFormat('@STRING@');
    settingsSheet.appendRow(['currency', 'UGX']);
    settingsSheet.appendRow(['ownerName', 'Isaac Oola']);
    settingsSheet.appendRow(['phone', '']);
  }

  return { success: true, message: 'Sheets setup complete' };
}

// ============================================================
// AUTH & SETTINGS
// ============================================================
function verifyPin(data) {
  const settings = getSettingsMap();
  const incoming = String(data.pin || '').trim();

  // Support up to 5 staff users
  const users = [
    { role: 'admin', name: settings.adminName || 'Admin', pin: String(settings.pin || '1234').trim() },
  ];
  for (var i = 1; i <= 5; i++) {
    var nameKey = 'user' + i + 'Name';
    var pinKey  = 'user' + i + 'Pin';
    var uName   = settings[nameKey] || ('User ' + i);
    var uPin    = String(settings[pinKey] || '').trim();
    users.push({ role: 'user' + i, name: uName, pin: uPin });
  }

  // If a specific role is requested (login screen), only match that role
  if (data.role) {
    const u = users.find(u => u.role === data.role);
    if (!u || !u.pin) return { success: false, error: 'No PIN set for this user.' };
    if (incoming === u.pin) return { success: true, role: u.role, name: u.name };
    return { success: false, error: 'Wrong PIN. Try again.' };
  }

  // Generic check (used by OTP admin verify) — match admin PIN only
  const admin = users[0];
  if (incoming === admin.pin) return { success: true, role: 'admin', name: admin.name };
  return { success: false, error: 'Wrong PIN' };
}

function getUsers(data) {
  const settings = getSettingsMap();
  return { success: true, data: {
    adminName: settings.adminName || 'Admin',
    user1Name: settings.user1Name || '',
    user2Name: settings.user2Name || '',
    user1HasPin: !!(settings.user1Pin),
    user2HasPin: !!(settings.user2Pin),
  }};
}

// Called by ADMIN to generate OTP (optionally emails it to themselves too)
function generateEditOtp(data) {
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const expiry = new Date().getTime() + 5 * 60 * 1000;
  const recordId   = data && data.recordId   ? String(data.recordId)   : '';
  const recordType = data && data.recordType ? String(data.recordType) : '';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = sheet.getDataRange().getValues();

  let otpRow = -1, expiryRow = -1, recordIdRow = -1, recordTypeRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'editOtp')           otpRow        = i + 1;
    if (rows[i][0] === 'editOtpExpiry')     expiryRow     = i + 1;
    if (rows[i][0] === 'editOtpRecordId')   recordIdRow   = i + 1;
    if (rows[i][0] === 'editOtpRecordType') recordTypeRow = i + 1;
  }

  if (otpRow > 0)        sheet.getRange(otpRow, 2).setValue(otp);
  else                   sheet.appendRow(['editOtp', otp]);
  if (expiryRow > 0)     sheet.getRange(expiryRow, 2).setValue(expiry);
  else                   sheet.appendRow(['editOtpExpiry', expiry]);
  if (recordIdRow > 0)   sheet.getRange(recordIdRow, 2).setValue(recordId);
  else                   sheet.appendRow(['editOtpRecordId', recordId]);
  if (recordTypeRow > 0) sheet.getRange(recordTypeRow, 2).setValue(recordType);
  else                   sheet.appendRow(['editOtpRecordType', recordType]);

  return { success: true, otp: otp };
}

// Called by STAFF to notify admin they need an OTP
// Sends email to admin with requester name — does NOT generate the OTP
function requestEditOtp(data) {
  const settings = getSettingsMap();
  const adminEmail  = settings.adminEmail || '';
  const requesterName = data.requesterName || 'A staff member';
  const recordDetail  = data.recordDetail  || 'is requesting edit access';
  const recordId      = data.recordId      || '';
  const recordType    = data.recordType    || '';
  const shopName = settings.shopName || 'Oola Spares';
  const message  = requesterName + ' ' + recordDetail;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  if (sheet) {
    const id = 'NOTIF-' + Date.now();
    // Store recordId||recordType in OTP column so admin can retrieve it
    const recordPayload = recordId ? `${recordId}||${recordType}` : '';
    sheet.appendRow([id, 'admin', requesterName, 'otp_request',
      message, recordPayload, new Date().toISOString(), false]);
  }

  if (!adminEmail) return { success: true, emailSent: false };
  try {
    MailApp.sendEmail(adminEmail,
      `🔔 ${shopName} — Edit Access Requested`,
      `${message}\n\nOpen the app to generate a one-time code for them.\n\n${new Date().toLocaleString()}`
    );
    return { success: true, emailSent: true };
  } catch(e) {
    return { success: true, emailSent: false };
  }
}

// Called by ADMIN to send a generated OTP directly to a specific user's inbox
function sendOtpNotification(data) {
  const toUser     = data.toUser;
  const otp        = data.otp;
  const fromUser   = data.fromUser || 'Admin';
  const recordId   = data.recordId   || '';
  const recordType = data.recordType || '';
  const recordLabel = data.recordLabel || '';
  if (!toUser || !otp) return { success: false, error: 'Missing toUser or otp' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  if (!sheet) return { success: false, error: 'Notifications sheet missing' };

  // Clear old unread OTP notifications for this user
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const toCol   = headers.indexOf('ToUser');
  const typeCol = headers.indexOf('Type');
  const readCol = headers.indexOf('Read');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][toCol] === toUser && rows[i][typeCol] === 'otp') {
      const isUnread = rows[i][readCol] === false || rows[i][readCol] === 'false' || rows[i][readCol] === '';
      if (isUnread) sheet.getRange(i + 1, readCol + 1).setValue(true);
    }
  }

  // Message includes what record the OTP is for
  const message = recordLabel
    ? `Edit code for: ${recordLabel}`
    : 'Admin sent you a one-time edit code';

  // Store recordId||recordType in OTP field so frontend can pass it back on verify
  const otpPayload = recordId ? `${otp}||${recordId}||${recordType}` : otp;

  const id = 'NOTIF-' + Date.now();
  sheet.appendRow([id, toUser, fromUser, 'otp', message, otpPayload, new Date().toISOString(), false]);

  return { success: true };
}

// Called by each user on app load / inbox open to get their unread notifications
function getMyNotifications(data) {
  const user = data.user; // 'admin', 'user1', 'user2'
  if (!user) return { success: false, error: 'Missing user' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  if (!sheet) return { success: true, data: [] };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, data: [] };
  const headers = rows[0];

  const notifs = rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  }).filter(n => {
    const matchUser = n.ToUser === user;
    const isUnread = n.Read === false || n.Read === 'false' || n.Read === '' || n.Read === 0;
    return matchUser && isUnread;
  }).map(n => ({
    ...n,
    Read: 'false', // normalise to string for frontend
    OTP: String(n.OTP || ''),
    NotifID: String(n.NotifID || ''),
    CreatedAt: String(n.CreatedAt || '')
  }));

  return { success: true, data: notifs };
}

function markNotificationRead(data) {
  const notifId = String(data.notifId || '');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  if (!sheet) return { success: false };

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('NotifID');
  const readCol = headers.indexOf('Read');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === notifId) {
      sheet.getRange(i + 1, readCol + 1).setValue(true);
      return { success: true };
    }
  }
  return { success: false, error: 'Notification not found' };
}

function verifyEditOtp(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = sheet.getDataRange().getValues();

  let storedOtp = '', storedExpiry = 0, storedRecordId = '', storedRecordType = '';
  let otpRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'editOtp')           { storedOtp = String(rows[i][1] || '').trim(); otpRow = i + 1; }
    if (rows[i][0] === 'editOtpExpiry')     storedExpiry     = Number(rows[i][1]) || 0;
    if (rows[i][0] === 'editOtpRecordId')   storedRecordId   = String(rows[i][1] || '').trim();
    if (rows[i][0] === 'editOtpRecordType') storedRecordType = String(rows[i][1] || '').trim();
  }

  const now = new Date().getTime();
  const incoming = String(data.otp || '').trim();
  const incomingRecordId   = String(data.recordId   || '').trim();
  const incomingRecordType = String(data.recordType || '').trim();

  if (!storedOtp) return { success: false, error: 'No OTP found. Generate a new one.' };

  if (now > storedExpiry) {
    if (otpRow > 0) sheet.getRange(otpRow, 2).setValue('');
    return { success: false, error: 'OTP expired. Request a new one.' };
  }

  if (incoming !== storedOtp) return { success: false, error: 'Incorrect OTP. Try again.' };

  // OTP correct — now check it's for the right record (if record was specified)
  if (storedRecordId && incomingRecordId && storedRecordId !== incomingRecordId) {
    return { success: false, error: 'This OTP is for a different record.' };
  }

  // ✓ All good — wipe OTP so it can never be reused
  if (otpRow > 0) sheet.getRange(otpRow, 2).setValue('');

  return {
    success: true,
    recordId:   storedRecordId,
    recordType: storedRecordType
  };
}

function getSettings() {
  const map = getSettingsMap();
  return { success: true, data: map };
}

function getSettingsMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  rows.forEach((row, i) => {
    if (i > 0 && row[0]) map[row[0]] = row[1];
  });
  return map;
}

function saveSettings(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = sheet.getDataRange().getValues();
  const keys = ['shopName', 'pin', 'currency', 'ownerName', 'phone', 'adminEmail',
                 'adminName', 'user1Name', 'user1Pin', 'user2Name', 'user2Pin'];
  keys.forEach(key => {
    if (data[key] !== undefined) {
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(data[key]);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, data[key]]);
    }
  });
  return { success: true };
}

// ============================================================
// GET ALL — reads every sheet in a single SpreadsheetApp.openById call
// This replaces 7 separate API calls with one, eliminating GAS quota issues
// and making data load fast and reliable every time
// ============================================================
function getAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function sheetToObjects(sheetName) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return [];
      const headers = rows[0];
      return rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
    } catch(e) {
      Logger.log('getAll sheet error [' + sheetName + ']: ' + e.message);
      return [];
    }
  }

  // Parts — must build Name field same way getParts does
  const partsSheet = ss.getSheetByName(SHEETS.PARTS);
  const partsRows = partsSheet ? partsSheet.getDataRange().getValues() : [];
  const parts = partsRows.length > 1 ? partsRows.slice(1).map(row => {
    const obj = {};
    partsRows[0].forEach((h, i) => { obj[h] = row[i]; });
    if (!obj.PartID) return null;
    const isNewFmt = partsRows[0][1] === 'ItemName';
    if (isNewFmt) {
      const isSizeFirst = partsRows[0][2] === 'Size';
      if (!isSizeFirst) { const tmp = obj.Brand; obj.Brand = obj.Size; obj.Size = tmp; }
      const nameParts = [obj.ItemName];
      if (obj.Size) nameParts.push(obj.Size);
      if (obj.Brand) nameParts.push('- ' + obj.Brand);
      obj.Name = nameParts.join(' ').replace(/\s+/g,' ').trim();
      obj.PartNo = obj.Brand || '';
      obj.Description = obj.Size || '';
      if (!obj.Movement) obj.Movement = 'Medium';
    } else {
      obj.ItemName = obj.Name || '';
      obj.Brand = obj.PartNo || '';
      obj.Size = obj.Description || '';
      if (!obj.Movement) obj.Movement = 'Medium';
    }
    return obj;
  }).filter(Boolean) : [];

  // Suppliers
  const suppliers = sheetToObjects(SHEETS.SUPPLIERS).filter(s => s.SupplierID);

  // Supplier Rates
  const supplierRates = sheetToObjects(SHEETS.SUPPLIER_RATES).filter(r => r.RateID || r.PartID);

  // Sales
  const sales = sheetToObjects(SHEETS.SALES).filter(s => s.SaleID);

  // Purchases
  const purchases = sheetToObjects(SHEETS.PURCHASES).filter(p => p.PurchaseID);

  // Checklist
  const checklist = sheetToObjects(SHEETS.CHECKLIST).filter(c => c.ItemID);

  // Settings
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  const settings = {};
  if (settingsSheet) {
    const sRows = settingsSheet.getDataRange().getValues();
    sRows.forEach((row, i) => {
      if (i > 0 && row[0]) settings[row[0]] = row[1];
    });
  }

  return {
    success: true,
    parts: parts.map(p => ({
      PartID: String(p.PartID||''), Name: p.Name||'', ItemName: p.ItemName||'',
      Size: p.Size||'', Brand: p.Brand||'', PartNo: p.PartNo||'',
      Category: p.Category||'', CostPrice: p.CostPrice||0, SellingPrice: p.SellingPrice||0,
      StockQty: p.StockQty||0, ReorderLevel: p.ReorderLevel||5,
      Unit: p.Unit||'Piece', Movement: p.Movement||'Medium',
      SupplierName: p.SupplierName||'', CreatedAt: String(p.CreatedAt||''),
      Barcode: String(p.Barcode||'')
    })),
    suppliers: suppliers.map(s => ({
      SupplierID: String(s.SupplierID||''), Name: s.Name||'',
      Phone: s.Phone||'', Location: s.Location||'', Email: s.Email||'', Notes: s.Notes||''
    })),
    supplierRates: supplierRates.map(r => ({
      RateID: String(r.RateID||''), SupplierID: String(r.SupplierID||''),
      SupplierName: r.SupplierName||'', PartID: String(r.PartID||''),
      PartName: r.PartName||'', UnitPrice: r.UnitPrice||0, Notes: r.Notes||''
    })),
    sales: sales.map(s => ({
      SaleID: String(s.SaleID||''), Date: String(s.Date||''),
      PartID: String(s.PartID||''), PartName: s.PartName||'',
      Qty: s.Qty||0, SellingPrice: s.SellingPrice||0, TotalAmount: s.TotalAmount||0,
      Notes: s.Notes||''
    })),
    purchases: purchases.map(p => ({
      PurchaseID: String(p.PurchaseID||''), Date: String(p.Date||''),
      PartID: String(p.PartID||''), PartName: p.PartName||'',
      SupplierID: String(p.SupplierID||''), SupplierName: p.SupplierName||'',
      Qty: p.Qty||0, CostPrice: p.CostPrice||0, TotalCost: p.TotalCost||0,
      Notes: p.Notes||''
    })),
    checklist: checklist.map(c => ({
      ItemID: String(c.ItemID||''), PartID: String(c.PartID||''),
      PartName: c.PartName||'', TargetQty: c.TargetQty||0,
      Status: c.Status||'', Priority: c.Priority||'', Notes: c.Notes||''
    })),
    customers: sheetToObjects(SHEETS.CUSTOMERS).filter(c => c.CustomerID).map(c => ({
      CustomerID: String(c.CustomerID||''), Name: c.Name||'',
      Phone: String(c.Phone||''), Email: c.Email||'',
      Address: c.Address||'', Notes: c.Notes||'',
      CreatedAt: String(c.CreatedAt||''),
      TotalPurchases: c.TotalPurchases||0, TotalSpent: c.TotalSpent||0
    })),
    settings
  };
}

// Fix stock quantities doubled by opening stock migration
// Subtracts the opening stock purchase qty from each affected part
function fixOpeningStockDoubling() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const purchSheet = ss.getSheetByName(SHEETS.PURCHASES);
  const partsSheet = ss.getSheetByName(SHEETS.PARTS);

  const purchRows = purchSheet.getDataRange().getValues();
  const purchHeaders = purchRows[0];
  const partIdCol   = purchHeaders.indexOf('PartID');
  const qtyCol      = purchHeaders.indexOf('Qty');
  const notesCol    = purchHeaders.indexOf('Notes');

  // Find all opening stock purchases
  const openingStockByPart = {};
  purchRows.slice(1).forEach(row => {
    const notes = String(row[notesCol]||'').toLowerCase();
    if (notes === 'opening stock') {
      const partId = String(row[partIdCol]||'');
      const qty    = parseInt(row[qtyCol]) || 0;
      openingStockByPart[partId] = (openingStockByPart[partId]||0) + qty;
    }
  });

  const partsRows = partsSheet.getDataRange().getValues();
  const partsHeaders = partsRows[0];
  const pIdCol  = partsHeaders.indexOf('PartID');
  const pQtyCol = partsHeaders.indexOf('StockQty');

  let fixed = 0;
  partsRows.slice(1).forEach((row, i) => {
    const partId = String(row[pIdCol]||'');
    if (openingStockByPart[partId]) {
      const currentQty = parseInt(row[pQtyCol]) || 0;
      const openingQty = openingStockByPart[partId];
      const correctedQty = currentQty - openingQty;
      if (correctedQty >= 0) {
        partsSheet.getRange(i + 2, pQtyCol + 1).setValue(correctedQty);
        fixed++;
      }
    }
  });

  return { success: true, message: `Fixed ${fixed} parts — subtracted opening stock quantities` };
}
// Smaller payload: no transaction history
function getCoreData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  function sheetToObjects(sheetName) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return [];
      const headers = rows[0];
      return rows.slice(1).map(row => { const obj = {}; headers.forEach((h,i) => { obj[h] = row[i]; }); return obj; });
    } catch(e) { return []; }
  }

  // Parts (same logic as getAll)
  const partsSheet = ss.getSheetByName(SHEETS.PARTS);
  const partsRows  = partsSheet ? partsSheet.getDataRange().getValues() : [];
  const parts = partsRows.length > 1 ? partsRows.slice(1).map(row => {
    const obj = {}; partsRows[0].forEach((h,i) => { obj[h] = row[i]; });
    if (!obj.PartID) return null;
    const isNewFmt = partsRows[0][1] === 'ItemName';
    if (isNewFmt) {
      const isSizeFirst = partsRows[0][2] === 'Size';
      if (!isSizeFirst) { const tmp = obj.Brand; obj.Brand = obj.Size; obj.Size = tmp; }
      const np = [obj.ItemName]; if (obj.Size) np.push(obj.Size); if (obj.Brand) np.push('- '+obj.Brand);
      obj.Name = np.join(' ').replace(/\s+/g,' ').trim(); obj.PartNo = obj.Brand||''; obj.Description = obj.Size||'';
      if (!obj.Movement) obj.Movement = 'Medium';
    } else { obj.ItemName = obj.Name||''; obj.Brand = obj.PartNo||''; obj.Size = obj.Description||''; }
    return obj;
  }).filter(Boolean) : [];

  const suppliers    = sheetToObjects(SHEETS.SUPPLIERS).filter(s => s.SupplierID);
  const supplierRates = sheetToObjects(SHEETS.SUPPLIER_RATES).filter(r => r.RateID||r.PartID);
  const checklist    = sheetToObjects(SHEETS.CHECKLIST).filter(c => c.ItemID);
  const customers    = sheetToObjects(SHEETS.CUSTOMERS).filter(c => c.CustomerID);

  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  const settings = {};
  if (settingsSheet) { const sr = settingsSheet.getDataRange().getValues(); sr.forEach((r,i) => { if (i>0 && r[0]) settings[r[0]] = r[1]; }); }

  return {
    success: true,
    parts: parts.map(p => ({
      PartID: String(p.PartID||''), Name: p.Name||'', ItemName: p.ItemName||'',
      Size: p.Size||'', Brand: p.Brand||'', PartNo: p.PartNo||'', Category: p.Category||'',
      CostPrice: p.CostPrice||0, SellingPrice: p.SellingPrice||0,
      StockQty: p.StockQty||0, ReorderLevel: p.ReorderLevel||5,
      Unit: p.Unit||'Piece', Movement: p.Movement||'Medium',
      SupplierName: p.SupplierName||'', Barcode: String(p.Barcode||''),
      ImageUrl: String(p.ImageUrl||'')
    })),
    suppliers: suppliers.map(s => ({ SupplierID: String(s.SupplierID||''), Name: s.Name||'', Phone: s.Phone||'', Location: s.Location||'', Email: s.Email||'', Notes: s.Notes||'' })),
    supplierRates: supplierRates.map(r => ({ RateID: String(r.RateID||''), SupplierID: String(r.SupplierID||''), SupplierName: r.SupplierName||'', PartID: String(r.PartID||''), PartName: r.PartName||'', UnitPrice: r.UnitPrice||0 })),
    checklist: checklist.map(c => ({ ItemID: String(c.ItemID||''), PartID: String(c.PartID||''), PartName: c.PartName||'', TargetQty: c.TargetQty||0, Status: c.Status||'', Priority: c.Priority||'', Notes: c.Notes||'' })),
    customers: customers.map(c => ({ CustomerID: String(c.CustomerID||''), Name: c.Name||'', Phone: String(c.Phone||''), Notes: c.Notes||'', TotalPurchases: c.TotalPurchases||0, TotalSpent: c.TotalSpent||0 })),
    settings
  };
}

// getTransactions — Sales + Purchases only
// Separate call so the payload stays small
function getTransactions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  function sheetToObjects(sheetName) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return [];
      const headers = rows[0];
      return rows.slice(1).map(row => { const obj = {}; headers.forEach((h,i) => { obj[h] = row[i]; }); return obj; });
    } catch(e) { return []; }
  }
  const sales     = sheetToObjects(SHEETS.SALES).filter(s => s.SaleID);
  const purchases = sheetToObjects(SHEETS.PURCHASES).filter(p => p.PurchaseID);
  return {
    success: true,
    sales: sales.map(s => ({
      SaleID: String(s.SaleID||''), Date: String(s.Date||''),
      PartID: String(s.PartID||''), PartName: s.PartName||'',
      Qty: s.Qty||0, SellingPrice: s.SellingPrice||0, TotalAmount: s.TotalAmount||0,
      CustomerID: String(s.CustomerID||''), CustomerName: s.CustomerName||'',
      Notes: s.Notes||'', RecordedBy: s.RecordedBy||''
    })),
    purchases: purchases.map(p => ({
      PurchaseID: String(p.PurchaseID||''), Date: String(p.Date||''),
      PartID: String(p.PartID||''), PartName: p.PartName||'',
      SupplierID: String(p.SupplierID||''), SupplierName: p.SupplierName||'',
      Qty: p.Qty||0, CostPrice: p.CostPrice||0, TotalCost: p.TotalCost||0, Notes: p.Notes||''
    }))
  };
}

// ============================================================
// CUSTOMERS
// ============================================================
function saveBarcode(data) {
  // Saves a barcode/QR string to a part
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('PartID');
  let barcodeCol = headers.indexOf('Barcode');

  // Add Barcode column if missing
  if (barcodeCol < 0) {
    barcodeCol = headers.length;
    sheet.getRange(1, barcodeCol + 1).setValue('Barcode');
  }

  // Check barcode not already used by another part
  if (data.Barcode) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][barcodeCol]||'').trim() === String(data.Barcode).trim() &&
          String(rows[i][idCol]) !== String(data.PartID)) {
        const existingName = rows[i][headers.indexOf('ItemName')] || rows[i][headers.indexOf('Name')] || '';
        return { success: false, error: 'Barcode already assigned to: ' + existingName };
      }
    }
  }

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.PartID)) {
      sheet.getRange(i + 1, barcodeCol + 1).setValue(data.Barcode || '');
      return { success: true };
    }
  }
  return { success: false, error: 'Part not found' };
}

function copyPartImageUrl(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PARTS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idCol  = headers.indexOf('PartID');
    let imgCol   = headers.indexOf('ImageUrl');
    if (imgCol < 0) {
      imgCol = headers.length;
      sheet.getRange(1, imgCol + 1).setValue('ImageUrl');
    }
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.PartID)) {
        sheet.getRange(i + 1, imgCol + 1).setValue(data.ImageUrl);
        return { success: true };
      }
    }
    return { success: false, error: 'Part not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function fixImageUrls() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const imgCol = headers.indexOf('ImageUrl');
  if (imgCol < 0) return { success: true, message: 'No ImageUrl column found' };

  let fixed = 0;
  for (let i = 1; i < rows.length; i++) {
    const url = String(rows[i][imgCol] || '');
    // Match old format: drive.google.com/uc?export=view&id=FILEID
    const match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match && url.includes('drive.google.com/uc')) {
      const fileId = match[1];
      const newUrl = 'https://lh3.googleusercontent.com/d/' + fileId + '=w500';
      sheet.getRange(i + 1, imgCol + 1).setValue(newUrl);
      fixed++;
    }
  }
  return { success: true, message: `Fixed ${fixed} image URLs` };
}

function savePartImage(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PARTS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idCol  = headers.indexOf('PartID');
    let imgCol   = headers.indexOf('ImageUrl');

    // Add ImageUrl column if missing
    if (imgCol < 0) {
      imgCol = headers.length;
      sheet.getRange(1, imgCol + 1).setValue('ImageUrl');
    }

    // Find the part row
    let partRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.PartID)) {
        partRow = i + 1; // 1-indexed for Sheets
        break;
      }
    }
    if (partRow < 0) return { success: false, error: 'Part not found' };

    // Save base64 image to Google Drive
    const base64Data = data.imageBase64.replace(/^data:image\/\w+;base64,/, '');
    const mimeType   = data.imageBase64.match(/data:(image\/\w+);/)?.[1] || 'image/jpeg';
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      'part_' + data.PartID + '.jpg'
    );

    // Save to a folder in Drive
    const folderName = 'OolaSpares_PartImages';
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    // Delete old image if exists
    const partName = String(data.PartID);
    const existingFiles = folder.getFilesByName('part_' + partName + '.jpg');
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    // Upload new image
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    const imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId + '=w500';

    // Save URL to Parts sheet
    sheet.getRange(partRow, imgCol + 1).setValue(imageUrl);

    return { success: true, imageUrl: imageUrl };
  } catch(e) {
    Logger.log('savePartImage error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function removePartImage(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PARTS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idCol  = headers.indexOf('PartID');
    const imgCol = headers.indexOf('ImageUrl');
    if (imgCol < 0) return { success: true };
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.PartID)) {
        sheet.getRange(i + 1, imgCol + 1).setValue('');
        return { success: true };
      }
    }
    return { success: false, error: 'Part not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getCustomers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CUSTOMERS);
  if (!sheet) return { success: true, data: [] };
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  const customers = rows.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]); return obj;
  }).filter(c => c.CustomerID);
  return { success: true, data: customers };
}

function saveCustomer(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CUSTOMERS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('CustomerID');

  if (data.CustomerID) {
    // Update existing
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.CustomerID)) {
        headers.forEach((h, j) => {
          if (h !== 'CustomerID' && h !== 'CreatedAt' && h !== 'TotalPurchases' && h !== 'TotalSpent' && data[h] !== undefined) {
            sheet.getRange(i+1, j+1).setValue(data[h]);
          }
        });
        return { success: true, id: data.CustomerID };
      }
    }
  }

  // New customer
  const id = 'CUS-' + Date.now();
  sheet.appendRow([
    id, data.Name||'', data.Phone||'', data.Email||'',
    data.Address||'', data.Notes||'',
    new Date().toISOString(), 0, 0
  ]);
  return { success: true, id: id };
}

function deleteCustomer(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CUSTOMERS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('CustomerID');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.CustomerID)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: 'Customer not found' };
}

function getCustomerHistory(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const salesSheet = ss.getSheetByName(SHEETS.SALES);
  const rows = salesSheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  const sales = rows.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]); return obj;
  }).filter(s => String(s.CustomerID) === String(data.CustomerID));
  return { success: true, data: sales };
}

// ============================================================
// PARTS
// ============================================================
function getParts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  const parts = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    // Normalize field names for frontend
    const isNewFmt = headers[1] === 'ItemName';
    if (isNewFmt) {
      // New format: B=ItemName, C=Size, D=Brand, E=Category, K=Movement, L=SupplierName
      // Detect if old new-format (Brand before Size) or latest (Size before Brand)
      const isSizeFirst = headers[2] === 'Size';
      if (!isSizeFirst) {
        // Old new-format: C=Brand, D=Size — swap for display
        const tmp = obj.Brand; obj.Brand = obj.Size; obj.Size = tmp;
      }
      const parts = [obj.ItemName];
      if (obj.Size) parts.push(obj.Size);
      if (obj.Brand) parts.push('- ' + obj.Brand);
      obj.Name = parts.join(' ').replace(/\s+/g,' ').trim();
      obj.PartNo = obj.Brand || '';
      obj.Description = obj.Size || '';
      if (!obj.Movement) obj.Movement = 'Medium';
    } else {
      obj.ItemName = obj.Name || '';
      obj.Brand = obj.PartNo || '';
      obj.Size = obj.Description || '';
      if (!obj.Movement) obj.Movement = 'Medium';
    }
    return obj;
  }).filter(p => p.PartID);
  return { success: true, data: parts };
}

function savePart(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  const now = formatDate(new Date());
  const hdrs = rows[0];

  const isNewFmt = hdrs[1] === 'ItemName';
  // Detect column order: new standard has Size(C) before Brand(D)
  const sizeIsCol2 = hdrs[2] === 'Size'; // col C = index 2

  // Extract clean fields — support all field name variants
  const itemName   = (data.ItemName || '').trim() || (data.Name ? data.Name.replace(/\s*-\s*[^-]+$/, '') : '').trim();
  const brand      = (data.Brand || data.PartNo || '').trim();
  const size       = (data.Size || data.Description || '').trim();
  const category   = (data.Category || '').trim();
  const cost       = parseFloat(data.CostPrice) || 0;
  const sell       = parseFloat(data.SellingPrice) || 0;
  const qty        = parseInt(data.StockQty) || 0;
  const reorder    = parseInt(data.ReorderLevel) || 5;
  const unit       = (data.Unit || 'Piece').trim();
  const move       = (data.Movement || 'Medium').trim();
  const supplier   = (data.SupplierName || '').trim();

  if (data.PartID) {
    // Update existing row
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.PartID)) {
        // Standard order: PartID, ItemName, Size, Brand, Category, CostPrice, SellingPrice, StockQty, ReorderLevel, Unit, Movement, SupplierName, CreatedAt
        const hasSupCol = hdrs.includes('SupplierName');
        const existSup  = hasSupCol ? (rows[i][hdrs.indexOf('SupplierName')] || '') : '';
        const existDate = rows[i][hdrs.indexOf('CreatedAt')] || now;
        sheet.getRange(i + 1, 1, 1, 13).setValues([[
          data.PartID,                    // A: PartID
          itemName,                       // B: ItemName
          size,                           // C: Size
          brand,                         // D: Brand
          category || rows[i][4] || '',  // E: Category
          cost,                          // F: CostPrice
          sell,                          // G: SellingPrice
          qty,                           // H: StockQty
          reorder || 5,                  // I: ReorderLevel
          unit,                          // J: Unit
          move,                          // K: Movement
          supplier || existSup || '',    // L: SupplierName
          existDate                      // M: CreatedAt
        ]]);
        sortPartsSheet();
        // Auto-update SupplierRates if supplier provided
        if (supplier && cost > 0) {
          updateSupplierRate(data.PartID, itemName + (size?' '+size:'') + (brand?' - '+brand:''), supplier, cost);
        }
        return { success: true, message: 'Part updated' };
      }
    }
    return { success: false, error: 'Part not found: ' + data.PartID };
  }

  // New part — append row
  const id = 'PRT-' + Date.now();
  const now2 = formatDate(new Date());
  if (isNewFmt) {
    // Correct order: PartID, ItemName, Size, Brand, Category, Cost, Sell, Qty, Reorder, Unit, Movement, SupplierName, CreatedAt
    sheet.appendRow([id, itemName, size, brand, category, cost, sell, qty, reorder, unit, move, supplier, now2]);
  } else {
    const fullName2 = [itemName, size, brand ? '- ' + brand : ''].filter(Boolean).join(' ').trim();
    sheet.appendRow([id, fullName2, category, brand, cost, sell, qty, reorder, unit, size, move, now2]);
  }
  sortPartsSheet();
  // Auto-update SupplierRates if supplier + cost price provided
  if (supplier && cost > 0) {
    updateSupplierRate(id || data.PartID, itemName + (size?' '+size:'') + (brand?' - '+brand:''), supplier, cost);
  }
  return { success: true, id: id, message: 'Part saved' };
}

function deletePart(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.PartID) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Part not found' };
}

// ============================================================
// SUPPLIERS
// ============================================================
function getSuppliers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  const suppliers = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(s => s.SupplierID);
  return { success: true, data: suppliers };
}

function saveSupplier(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
  const rows = sheet.getDataRange().getValues();
  const now = formatDate(new Date());

  if (data.SupplierID) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.SupplierID) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          data.SupplierID, data.Name, data.Contact, data.Location,
          data.Email || '', data.Notes || '', rows[i][6]
        ]]);
        return { success: true, message: 'Supplier updated' };
      }
    }
  }

  const id = 'SUP-' + Date.now();
  sheet.appendRow([id, data.Name, data.Contact, data.Location, data.Email || '', data.Notes || '', now]);
  return { success: true, id: id, message: 'Supplier saved' };
}

function deleteSupplier(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.SupplierID) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Supplier not found' };
}

// ============================================================
// SUPPLIER RATES
// ============================================================
function getSupplierRates(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIER_RATES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  let rates = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(r => r.RateID);
  if (data.PartID) rates = rates.filter(r => r.PartID === data.PartID);
  if (data.SupplierID) rates = rates.filter(r => r.SupplierID === data.SupplierID);
  return { success: true, data: rates };
}

function saveSupplierRate(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIER_RATES);
  const rows = sheet.getDataRange().getValues();
  const now = formatDate(new Date());

  if (data.RateID) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.RateID) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          data.RateID, data.SupplierID, data.SupplierName,
          data.PartID, data.PartName,
          parseFloat(data.UnitPrice) || 0,
          data.Notes || '', now
        ]]);
        return { success: true, message: 'Rate updated' };
      }
    }
  }

  const id = 'RAT-' + Date.now();
  sheet.appendRow([
    id, data.SupplierID, data.SupplierName,
    data.PartID, data.PartName,
    parseFloat(data.UnitPrice) || 0,
    data.Notes || '', now
  ]);
  return { success: true, id: id, message: 'Rate saved' };
}

function deleteSupplierRate(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SUPPLIER_RATES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.RateID) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Rate not found' };
}

// ============================================================
// PURCHASES
// ============================================================
function getPurchases(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PURCHASES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  let purchases = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(p => p.PurchaseID);
  if (data.from && data.to) {
    purchases = purchases.filter(p => p.Date >= data.from && p.Date <= data.to);
  }
  return { success: true, data: purchases };
}

function savePurchase(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PURCHASES);
  const id = 'PUR-' + Date.now();
  const qty = parseInt(data.Qty) || 0;
  const cost = parseFloat(data.CostPrice) || 0;

  sheet.appendRow([
    id, data.Date, data.PartID, data.PartName,
    data.SupplierID, data.SupplierName,
    qty, cost, qty * cost, data.Notes || ''
  ]);

  // Skip stock update for opening stock migration (stock already exists)
  if (data.noStockUpdate !== 'true') {
    updatePartStock(data.PartID, qty, 'add');
  }

  // Update cost price if provided
  if (data.updateCostPrice === 'true') {
    updatePartCostPrice(data.PartID, cost);
  }

  return { success: true, id: id, message: 'Purchase recorded' };
}

function updatePartStock(partId, qty, operation) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  const hdrs = rows[0];
  const qtyCol = hdrs.indexOf('StockQty');
  const qtyColNum = qtyCol >= 0 ? qtyCol + 1 : 8;
  const qtyIdx = qtyCol >= 0 ? qtyCol : 7;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === partId) {
      const current = parseInt(rows[i][qtyIdx]) || 0;
      const newQty = operation === 'add' ? current + qty : Math.max(0, current - qty);
      sheet.getRange(i + 1, qtyColNum).setValue(newQty);
      return;
    }
  }
}

function updatePartCostPrice(partId, price) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === partId) {
      const hdrs2 = rows[0];
      const costCol = hdrs2.indexOf('CostPrice');
      const costColNum = costCol >= 0 ? costCol + 1 : 6;
      sheet.getRange(i + 1, costColNum).setValue(price);
      return;
    }
  }
}

// ============================================================
// SALES
// ============================================================
function getSales(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SALES);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  let sales = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(s => s.SaleID);
  if (data.from && data.to) {
    sales = sales.filter(s => s.Date >= data.from && s.Date <= data.to);
  }
  return { success: true, data: sales };
}

function saveSale(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SALES);
  const id = 'SAL-' + Date.now();
  const qty = parseInt(data.Qty) || 0;
  const price = parseFloat(data.SellingPrice) || 0;
  const total = qty * price;

  sheet.appendRow([
    id, data.Date, data.PartID, data.PartName,
    qty, price, total,
    data.CustomerID || '', data.CustomerName || '',
    data.Notes || '', data.RecordedBy || ''
  ]);

  // Deduct stock
  updatePartStock(data.PartID, qty, 'deduct');

  // Update customer stats if linked
  if (data.CustomerID) {
    updateCustomerStats(ss, data.CustomerID, 1, total);
  }

  return { success: true, id: id, message: 'Sale recorded' };
}

function updateCustomerStats(ss, customerId, purchaseDelta, spentDelta) {
  try {
    const sheet = ss.getSheetByName(SHEETS.CUSTOMERS);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idCol = headers.indexOf('CustomerID');
    const tpCol = headers.indexOf('TotalPurchases');
    const tsCol = headers.indexOf('TotalSpent');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(customerId)) {
        if (tpCol >= 0) sheet.getRange(i+1, tpCol+1).setValue((parseInt(rows[i][tpCol])||0) + purchaseDelta);
        if (tsCol >= 0) sheet.getRange(i+1, tsCol+1).setValue((parseFloat(rows[i][tsCol])||0) + spentDelta);
        break;
      }
    }
  } catch(e) { Logger.log('updateCustomerStats error: ' + e.message); }
}

function updateSale(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SALES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('SaleID');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === data.SaleID) {
      const oldQty = parseInt(rows[i][headers.indexOf('Qty')]) || 0;
      const oldPartId = rows[i][headers.indexOf('PartID')];
      const newQty = parseInt(data.Qty) || 0;
      const newPrice = parseFloat(data.SellingPrice) || 0;

      // Reverse old stock deduction, apply new one
      updatePartStock(oldPartId, oldQty, 'add');
      updatePartStock(data.PartID, newQty, 'deduct');

      sheet.getRange(i + 1, 1, 1, headers.length).setValues([[
        data.SaleID, data.Date, data.PartID, data.PartName,
        newQty, newPrice, newQty * newPrice, data.Notes || ''
      ]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Sale not found' };
}

function deleteSale(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SALES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('SaleID');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === data.SaleID) {
      const qty = parseInt(rows[i][headers.indexOf('Qty')]) || 0;
      const partId = rows[i][headers.indexOf('PartID')];
      // Restore stock
      updatePartStock(partId, qty, 'add');
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Sale not found' };
}

function updatePurchase(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PURCHASES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('PurchaseID');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === data.PurchaseID) {
      const oldQty = parseInt(rows[i][headers.indexOf('Qty')]) || 0;
      const oldPartId = rows[i][headers.indexOf('PartID')];
      const newQty = parseInt(data.Qty) || 0;
      const newCost = parseFloat(data.CostPrice) || 0;

      // Reverse old stock addition, apply new one
      updatePartStock(oldPartId, oldQty, 'deduct');
      updatePartStock(data.PartID, newQty, 'add');

      if (data.updateCostPrice === 'true') {
        updatePartCostPrice(data.PartID, newCost);
      }

      sheet.getRange(i + 1, 1, 1, headers.length).setValues([[
        data.PurchaseID, data.Date, data.PartID, data.PartName,
        data.SupplierID || '', data.SupplierName || '',
        newQty, newCost, newQty * newCost, data.Notes || ''
      ]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Purchase not found' };
}

function deletePurchase(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PURCHASES);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('PurchaseID');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === data.PurchaseID) {
      const qty = parseInt(rows[i][headers.indexOf('Qty')]) || 0;
      const partId = rows[i][headers.indexOf('PartID')];
      // Reverse stock addition
      updatePartStock(partId, qty, 'deduct');
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Purchase not found' };
}

// ============================================================
// CHECKLIST
// ============================================================
function getChecklist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };
  const headers = rows[0];
  const items = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(c => c.ItemID);
  return { success: true, data: items };
}

function saveChecklistItem(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const now = formatDate(new Date());

  if (data.ItemID) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.ItemID) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          data.ItemID, data.PartID, data.PartName,
          parseInt(data.TargetQty) || 1,
          data.Status || 'Pending',
          data.Priority || 'Normal',
          data.Notes || '', now
        ]]);
        return { success: true };
      }
    }
  }

  const id = 'CHK-' + Date.now();
  sheet.appendRow([
    id, data.PartID, data.PartName,
    parseInt(data.TargetQty) || 1,
    'Pending', data.Priority || 'Normal',
    data.Notes || '', now
  ]);
  return { success: true, id: id };
}

function bulkSaveChecklist(data) {
  // data.items = array of {PartID, PartName, TargetQty, Priority, Status, Notes}
  // data.skipExisting = true — skip PartIDs already in checklist
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const now = formatDate(new Date());

  // Get existing PartIDs to skip duplicates
  const existing = new Set();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1]) existing.add(String(rows[i][1]));
  }

  const items = (data.items || []).filter(item => !existing.has(String(item.PartID)));
  if (items.length === 0) return { success: true, added: 0 };

  const newRows = items.map(item => [
    'CHK-' + Date.now() + '-' + Math.random().toString(36).substr(2,5),
    item.PartID, item.PartName,
    parseInt(item.TargetQty) || 1,
    item.Status || 'Pending',
    item.Priority || 'Normal',
    item.Notes || '', now
  ]);

  // Write all rows at once
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, newRows.length, 8).setValues(newRows);

  return { success: true, added: newRows.length };
}

function updateChecklistStatus(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  const now = formatDate(new Date());
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.ItemID) {
      sheet.getRange(i + 1, 5).setValue(data.Status);
      sheet.getRange(i + 1, 8).setValue(now);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
}

function deleteChecklistItem(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.ItemID) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
}

function clearAllChecklist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return { success: true };
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboard() {
  const parts = getParts().data;
  const sales = getSales({}).data;
  const purchases = getPurchases({}).data;
  const checklist = getChecklist().data;

  const today = formatDate(new Date());
  const todaySales = sales.filter(s => s.Date === today);
  const todayPurchases = purchases.filter(p => p.Date === today);

  const lowStock = parts.filter(p => parseInt(p.StockQty) <= parseInt(p.ReorderLevel));
  const totalParts = parts.length;
  const totalStockValue = parts.reduce((sum, p) => sum + (parseFloat(p.CostPrice) * parseInt(p.StockQty) || 0), 0);

  const todayRevenue = todaySales.reduce((sum, s) => sum + (parseFloat(s.TotalAmount) || 0), 0);
  const todaySpend = todayPurchases.reduce((sum, p) => sum + (parseFloat(p.TotalCost) || 0), 0);

  const pendingChecklist = checklist.filter(c => c.Status === 'Pending').length;

  return {
    success: true,
    data: {
      totalParts,
      totalStockValue,
      lowStockCount: lowStock.length,
      lowStockItems: lowStock.slice(0, 5),
      todayRevenue,
      todaySpend,
      todaySalesCount: todaySales.length,
      pendingChecklist,
      recentSales: sales.slice(-5).reverse()
    }
  };
}


// ============================================================
// MIGRATE PARTS SHEET to new column format
// Run once: ?action=migratePartsSheet
// ============================================================
function migratePartsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, message: 'Nothing to migrate' };

  const hdrs = rows[0];
  const isAlreadyNew = hdrs[1] === 'ItemName';
  if (isAlreadyNew) return { success: true, message: 'Already in new format' };

  // Old format: PartID, Name, Category, PartNo, CostPrice, SellingPrice, StockQty, ReorderLevel, Unit, Description, Movement, CreatedAt
  // New format: PartID, ItemName, Brand, Size, Category, CostPrice, SellingPrice, StockQty, ReorderLevel, Unit, Movement, CreatedAt

  // Update header row
  sheet.getRange(1, 1, 1, 12).setValues([[
    'PartID','ItemName','Brand','Size','Category',
    'CostPrice','SellingPrice','StockQty','ReorderLevel',
    'Unit','Movement','CreatedAt'
  ]]);
  sheet.getRange(1,1,1,12).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  let migrated = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue; // skip empty rows

    const fullName = String(row[1] || '');
    const oldPartNo = String(row[3] || ''); // brand was in PartNo
    const oldDesc = String(row[9] || '');   // size was in Description
    const oldCat = String(row[2] || '');
    const oldMove = String(row[10] || 'Medium');
    const oldCreated = row[11] || '';

    // Parse ItemName and Brand from full Name
    let itemName = fullName, brand = oldPartNo, size = oldDesc;
    if (!brand && fullName.includes(' - ')) {
      const pts = fullName.split(' - ');
      brand = pts[pts.length-1].trim();
      itemName = pts.slice(0,-1).join(' - ').trim();
    }
    // Remove size from itemName if it's repeated
    if (size && itemName.includes(size)) {
      itemName = itemName.replace(size, '').replace(/\s+/,' ').trim();
    }

    sheet.getRange(i+1, 1, 1, 12).setValues([[
      row[0], itemName, brand, size, oldCat,
      row[4], row[5], row[6], row[7], row[8],
      oldMove || 'Medium', oldCreated
    ]]);
    migrated++;
  }

  return { success: true, message: 'Migrated ' + migrated + ' rows to new format' };
}




// ============================================================
// FIX STOCK DATA — corrects wrong StockQty/SellingPrice from bad import
// Run once: ?action=fixStockData
// ============================================================
function fixStockData() {
  const FIX_DATA = [
    {name:'Spark plug A7', cost:1200, sell:2000, qty:10},{name:'Spark plug C7', cost:1200, sell:3000, qty:10},
    {name:'Spark plug D8 Short', cost:1200, sell:3000, qty:10},{name:'Spark plug D8 Long', cost:1200, sell:3000, qty:10},
    {name:'Spark plug BP67 Short', cost:1300, sell:3000, qty:10},{name:'Spark plug BP67 Long', cost:1300, sell:3000, qty:10},
    {name:'Clutch wire', cost:500, sell:1000, qty:20},{name:'Raise wire', cost:500, sell:1000, qty:20},
    {name:'Raise cable', cost:2500, sell:4000, qty:10},{name:'Brake cable', cost:2500, sell:4000, qty:10},
    {name:'Clutch cable', cost:2500, sell:4000, qty:10},{name:'Choke cable', cost:2500, sell:4000, qty:10},
    {name:'Tube 2.75-17', cost:4900, sell:7000, qty:10},{name:'Tube 3.00-17', cost:4900, sell:7000, qty:10},
    {name:'Tube 3.00-18', cost:4900, sell:8000, qty:10},{name:'Tube 2.75-21', cost:8500, sell:15000, qty:5},
    {name:'Tube 4.10-28', cost:8500, sell:15000, qty:5},{name:'Tyre 2.75-17', cost:29000, sell:45000, qty:5},
    {name:'Tyre 3.00-18', cost:55000, sell:65000, qty:5},{name:'Tyre 300-17', cost:55000, sell:65000, qty:5},
    {name:'Mirror', cost:2250, sell:6000, qty:10},{name:'Dimmer switch', cost:7500, sell:12000, qty:5},
    {name:'Rubber boot', cost:2800, sell:5000, qty:5},{name:'Gear lever', cost:3500, sell:5000, qty:5},
    {name:'Gear shaft', cost:8500, sell:15000, qty:5},{name:'Kick starter shaft', cost:15000, sell:25000, qty:3},
    {name:'Kick starter', cost:7500, sell:10000, qty:5},{name:'Front wheel axle', cost:2500, sell:5000, qty:10},
    {name:'Rear wheel axle', cost:3000, sell:5000, qty:10},{name:'Paper gasket Magnet side', cost:500, sell:2000, qty:30},
    {name:'Paper gasket Clutch plate side', cost:500, sell:2000, qty:30},{name:'Paper gasket slip part', cost:500, sell:2000, qty:30},
    {name:'Steel gasket', cost:500, sell:3000, qty:30},{name:'Clutch plate', cost:3300, sell:5000, qty:20},
    {name:'Engine Valve', cost:2800, sell:5000, qty:10},{name:'Bulb holder', cost:2700, sell:5000, qty:10},
    {name:'Head lamp bulb', cost:700, sell:1000, qty:20},{name:'Head lamp case', cost:2500, sell:5000, qty:10},
    {name:'Indicator bulb', cost:250, sell:500, qty:30},{name:'Indicator', cost:2750, sell:3000, qty:20},
    {name:'Tail light', cost:5000, sell:10000, qty:5},{name:'Head lamp complete', cost:10000, sell:15000, qty:5},
    {name:'Tail light switch', cost:2500, sell:5000, qty:5},{name:'Front sprocket', cost:1300, sell:3000, qty:30},
    {name:'Rear sprocket', cost:7000, sell:10000, qty:30},{name:'Chain', cost:10000, sell:13000, qty:10},
    {name:'Chain adjuster', cost:1250, sell:2000, qty:20},{name:'Chain guard', cost:19000, sell:25000, qty:2},
    {name:'Sprocket set', cost:28000, sell:35000, qty:5},{name:'Dush board Old mode', cost:10000, sell:15000, qty:10},
    {name:'Dush board New mode', cost:19000, sell:35000, qty:5},{name:'Switch Old mode', cost:2700, sell:5000, qty:10},
    {name:'Switch New mode', cost:10000, sell:20000, qty:5},{name:'Brake pad/shoe Standard', cost:3300, sell:5000, qty:20},
    {name:'Brake pad/shoe X-Large', cost:5500, sell:7000, qty:20},{name:'Piston', cost:9000, sell:15000, qty:10},
    {name:'Valve seal pair', cost:500, sell:2000, qty:30},{name:'Brake pedal', cost:7000, sell:10000, qty:10},
    {name:'Brake line /spring', cost:9500, sell:3000, qty:10},{name:'Single stand spring', cost:500, sell:2000, qty:20},
    {name:'Double stand spring', cost:500, sell:2000, qty:20},{name:'Single stand', cost:6500, sell:15000, qty:5},
    {name:'Double stand', cost:12000, sell:25000, qty:2},{name:'Horn', cost:2800, sell:5000, qty:20},
    {name:'Neck bearing', cost:3500, sell:10000, qty:5},{name:'Centre bolt', cost:2500, sell:5000, qty:10},
    {name:'Plug cap', cost:1500, sell:3000, qty:10},{name:'Gasket maker Small size', cost:1300, sell:3000, qty:20},
    {name:'Gasket maker Big size', cost:3800, sell:5000, qty:20},{name:'Wire system Bajaj', cost:10000, sell:15000, qty:5},
    {name:'Wire system CG 125', cost:10000, sell:15000, qty:5},{name:'Clutch lever', cost:1500, sell:3000, qty:20},
    {name:'Brake lever', cost:1500, sell:3000, qty:10},{name:'Raise handle Right', cost:2750, sell:5000, qty:10},
    {name:'Raise handle Left', cost:2750, sell:5000, qty:20},{name:'Handle bar', cost:8500, sell:15000, qty:5},
    {name:'Rear shock absorber', cost:50000, sell:60000, qty:2},{name:'Front shock absorber', cost:100000, sell:100000, qty:1},
    {name:'Seal shock absorber', cost:2800, sell:5000, qty:10},{name:'Pipe shock obsorber', cost:34000, sell:55000, qty:2},
    {name:'Oil seal', cost:2000, sell:5000, qty:20},{name:'Leg guard', cost:13200, sell:25000, qty:2},
    {name:'Front mud guard', cost:14000, sell:25000, qty:2},{name:'Fly wheel', cost:25000, sell:35000, qty:3},
    {name:'Clutch disk', cost:8500, sell:15000, qty:5},{name:'Cylinder head', cost:130000, sell:150000, qty:5},
    {name:'Cylinder block', cost:140000, sell:65000, qty:5},{name:'Crank shaft', cost:45000, sell:80000, qty:2},
    {name:'Timing chain', cost:2800, sell:5000, qty:20},{name:'Oil pump cap Metalic', cost:23000, sell:35000, qty:2},
    {name:'Tensioner block', cost:3500, sell:5000, qty:10},{name:'Rocker armor', cost:10000, sell:25000, qty:3},
    {name:'Tappet arm', cost:10000, sell:25000, qty:3},{name:'Tappet bolt', cost:600, sell:2000, qty:10},
    {name:'Oil pump', cost:9000, sell:15000, qty:3},{name:'Timing sprocket', cost:5000, sell:15000, qty:5},
    {name:'Sprocket shaft complete', cost:10000, sell:20000, qty:3},{name:'Front sprocket cover', cost:3500, sell:5000, qty:5},
    {name:'Air clearer', cost:3000, sell:5000, qty:10},{name:'Brake rod/ line', cost:1300, sell:3000, qty:20},
    {name:'Side cover', cost:10000, sell:20000, qty:5},{name:'Side cover lock', cost:2700, sell:5000, qty:5},
    {name:'Bearing front hub 6200', cost:1400, sell:3000, qty:10},{name:'Bearing front hub 6300', cost:1400, sell:3000, qty:10},
    {name:'Clutch bearing 628', cost:1300, sell:3000, qty:10},{name:'Bearing rear Hub 6202', cost:1400, sell:3000, qty:10},
    {name:'Bearing rear Hub 6302', cost:1500, sell:3000, qty:10},{name:'Crank shaft bearing 6304', cost:3500, sell:5000, qty:10},
    {name:'Cam shaft bearing 6003', cost:1300, sell:3000, qty:10},{name:'Cam shaft bearing 6002', cost:1300, sell:3000, qty:10},
    {name:'Cam shaft bearing 6201', cost:1300, sell:3000, qty:10},
    {name:'Engine Oil Accel', cost:190000, sell:190000, qty:1},{name:'Engine Oil Total', cost:220000, sell:220000, qty:1},
    {name:'Engine Oil Helix', cost:225000, sell:225000, qty:1},{name:'Engine Oil Advance', cost:190000, sell:190000, qty:1},
    {name:'Rear hub', cost:0, sell:0, qty:3},{name:'Front hub', cost:0, sell:0, qty:2},
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  const hdrs = rows[0];

  const iItem = hdrs.indexOf('ItemName');
  const iBrand = hdrs.indexOf('Brand');
  const iSize = hdrs.indexOf('Size');
  const iCost = hdrs.indexOf('CostPrice');
  const iSell = hdrs.indexOf('SellingPrice');
  const iQty = hdrs.indexOf('StockQty');
  const iMove = hdrs.indexOf('Movement');

  // Build lookup from fix data
  const fixMap = {};
  FIX_DATA.forEach(d => { fixMap[d.name.toLowerCase().trim()] = d; });

  let fixed = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;

    // Build name to match: ItemName + Size (without brand for base items)
    const item = String(row[iItem]||'').trim();
    const size = String(row[iSize]||'').trim();
    const brand = String(row[iBrand]||'').trim();

    // Try different name combinations
    const candidates = [
      (item + (size ? ' ' + size : '')).toLowerCase().trim(),
      item.toLowerCase().trim(),
    ];

    let fix = null;
    for (const c of candidates) {
      if (fixMap[c]) { fix = fixMap[c]; break; }
    }

    if (fix) {
      // Only fix if StockQty looks wrong (>100 usually means it got price data)
      const currentQty = parseInt(row[iQty]) || 0;
      const needsQtyFix = currentQty > 100 || currentQty < 0;
      const needsPriceFix = fix.sell > 0 && (parseFloat(row[iSell])||0) === 0;

      if (needsQtyFix) sheet.getRange(i+1, iQty+1).setValue(fix.qty);
      if (needsPriceFix) {
        sheet.getRange(i+1, iCost+1).setValue(fix.cost);
        sheet.getRange(i+1, iSell+1).setValue(fix.sell);
      }
      if (needsQtyFix || needsPriceFix) fixed++;
    }
  }

  return { success: true, message: 'Fixed ' + fixed + ' rows with wrong stock/price data' };
}




// ============================================================
// FIX ALL CURRENT SHEET ISSUES
// Run: ?action=fixAllIssues
// ============================================================
function fixAllIssues() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, message: 'Nothing to fix' };

  const hdrs = rows[0];
  const iItem  = hdrs.indexOf('ItemName');
  const iSize  = hdrs.indexOf('Size');
  const iBrand = hdrs.indexOf('Brand');
  const iCat   = hdrs.indexOf('Category');
  const iQty   = hdrs.indexOf('StockQty');
  const iMove  = hdrs.indexOf('Movement');
  const iDate  = hdrs.indexOf('CreatedAt');

  const KNOWN_BRANDS = ['CC','Titan','RZ','Kevla','MRF','Crocodile','DZ','D&K','Yog','Kevil',
    'TZL','NYC','DJM','NGK','Bosch','Generic','KQK','K&K','J&L','Oscally','Accel',
    'Total','Helix','Advance','X-Large','Bajaj'];

  // Sizes that look like brands but are actually size/variant descriptors
  const SIZE_DESCRIPTORS = ['Metalic','Old Mode','New Mode','Standard','X-Large',
    'Short','Long','Front','Rear','Left','Right','Pair','Set','Small Size','Big Size',
    'Magnet Side','Clutch Plate Side','Slip Part','Bearing 628','Bearing 6200',
    'Bearing 6300','Bearing 6202','Bearing 6302','Bearing 6304','Bearing 6003',
    'Bearing 6002','Bearing 6201','1 Litre','4 Litres','Quarts',
    'OLA Oil'  // OLA Oil is a brand name that ended up in size col
  ];

  // Items where size col contains wrong data (brand names)
  // Engine Oil: Accel/Total/Helix/Advance/OLA Oil are brands, not sizes
  const BRAND_IN_SIZE_ITEMS = ['Engine Oil'];

  const validMoves = ['Fast','Medium','Slow','Dead'];
  let fixed = 0, removed = 0;

  // Pass 1: Fix individual row issues
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || String(r[0]).trim() === '') continue;

    let itemName = String(r[iItem]||'').trim();
    let size     = String(r[iSize]||'').trim();
    let brand    = String(r[iBrand]||'').trim();
    let changed  = false;

    // Fix: if size === brand (duplicate), clear size
    if (size && brand && size.toLowerCase() === brand.toLowerCase()) {
      sheet.getRange(i+1, iSize+1).setValue('');
      size = '';
      changed = true;
    }

    // Fix: if size is a known brand name, it belongs in Brand col
    if (size && KNOWN_BRANDS.some(b => b.toLowerCase() === size.toLowerCase())) {
      if (!brand) {
        sheet.getRange(i+1, iBrand+1).setValue(size);
        brand = size;
      }
      sheet.getRange(i+1, iSize+1).setValue('');
      size = '';
      changed = true;
    }

    // Fix: Engine Oil specific - size like "OLA Oil", "Quarts" should be cleared
    // These are unit descriptions that ended up in Size col
    if (itemName.toLowerCase().includes('engine oil') && size && size !== brand) {
      // For Engine Oil, size is meaningless - clear it
      sheet.getRange(i+1, iSize+1).setValue('');
      size = '';
      changed = true;
    }

    // Fix: if brand is empty but size looks like a brand name (known brand)
    if (!brand && size && KNOWN_BRANDS.some(b => b.toLowerCase() === size.toLowerCase())) {
      sheet.getRange(i+1, iBrand+1).setValue(size);
      sheet.getRange(i+1, iSize+1).setValue('');
      changed = true;
    }

    // Fix: Movement contains date → reset to Medium
    const movVal = String(r[iMove]||'');
    if (!validMoves.includes(movVal.trim()) || movVal.includes('GMT') || movVal.includes('2026')) {
      sheet.getRange(i+1, iMove+1).setValue('Medium');
      changed = true;
    }

    // Fix: CreatedAt full timestamp → clean DD-Mon-YYYY format
    if (iDate >= 0) {
      const dateVal = String(r[iDate]||'');
      if (dateVal && (dateVal.includes('GMT') || dateVal.includes('T00:00') || dateVal.includes('00:00:00'))) {
        try {
          var d = new Date(dateVal);
          if (!isNaN(d.getTime())) {
            // Convert to EAT (UTC+3) and format as DD-Mon-YYYY
            var eatOffset = 3 * 60 * 60 * 1000;
            var eatDate = new Date(d.getTime() + eatOffset);
            var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            var clean = String(eatDate.getUTCDate()).padStart(2,'0') + '-' + months[eatDate.getUTCMonth()] + '-' + eatDate.getUTCFullYear();
            sheet.getRange(i+1, iDate+1).setValue(clean);
            changed = true;
          }
        } catch(e) {}
      }
    }

    // Fix: qty > 200 → reset to 0
    const qty = parseInt(r[iQty]);
    if (!isNaN(qty) && qty > 200) {
      sheet.getRange(i+1, iQty+1).setValue(0);
      changed = true;
    }

    if (changed) fixed++;
  }

  // Pass 2: Remove duplicates (same ItemName+Size+Brand)
  const freshRows = sheet.getDataRange().getValues();
  const seen = {};
  const toDelete = [];

  for (let i = 1; i < freshRows.length; i++) {
    const r = freshRows[i];
    if (!r[0] || String(r[0]).trim() === '') { toDelete.push(i+1); continue; }
    const key = String(r[iItem]||'').trim().toLowerCase() + '|' +
                String(r[iSize]||'').trim().toLowerCase() + '|' +
                String(r[iBrand]||'').trim().toLowerCase();
    if (seen[key] !== undefined) {
      // Keep one with more stock, delete other
      const prevQty = parseInt(freshRows[seen[key]][iQty])||0;
      const currQty = parseInt(r[iQty])||0;
      if (currQty > prevQty) {
        toDelete.push(seen[key]+1);
        seen[key] = i;
      } else {
        toDelete.push(i+1);
      }
      removed++;
    } else {
      seen[key] = i;
    }
  }

  // Delete from bottom to top
  toDelete.sort((a,b) => b-a);
  toDelete.forEach(r => sheet.deleteRow(r));

  return {
    success: true,
    message: 'Fixed ' + fixed + ' rows, removed ' + removed + ' duplicates, deleted ' + (toDelete.length - removed) + ' empty rows'
  };
}

// ============================================================
// REMOVE OFFLINE DUPLICATES
// Run: ?action=removeOfflineDuplicates
// ============================================================
function removeOfflineDuplicates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, message: 'Nothing to clean' };

  const hdrs = rows[0];
  const iItem  = hdrs.indexOf('ItemName');
  const iSize  = hdrs.indexOf('Size');
  const iBrand = hdrs.indexOf('Brand');
  const iDate  = hdrs.indexOf('CreatedAt');

  // Group rows by key = itemName+size+brand
  const seen = {};
  const toDelete = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const key = String(r[iItem]||'').trim().toLowerCase() + '|' +
                String(r[iSize]||'').trim().toLowerCase() + '|' +
                String(r[iBrand]||'').trim().toLowerCase();

    if (seen[key] !== undefined) {
      // Duplicate — keep the one with earlier/better date, delete the other
      // Prefer rows with qty > 0 or earlier creation date
      const existRow = rows[seen[key]];
      const existQty = parseInt(existRow[hdrs.indexOf('StockQty')])||0;
      const currQty  = parseInt(r[hdrs.indexOf('StockQty')])||0;

      if (currQty > existQty) {
        // Current row has better qty — delete the old one
        toDelete.push(seen[key] + 1); // 1-based
        seen[key] = i;
      } else {
        toDelete.push(i + 1); // 1-based
      }
    } else {
      seen[key] = i;
    }
  }

  // Delete from bottom to top
  toDelete.sort((a, b) => b - a);
  toDelete.forEach(r => sheet.deleteRow(r));

  return { success: true, message: 'Removed ' + toDelete.length + ' duplicate rows' };
}

// ============================================================
// REORDER PARTS COLUMNS to new standard order
// Run: ?action=reorderPartsColumns
// ============================================================
function reorderPartsColumns() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return { success: true, message: 'Nothing to reorder' };

  const hdrs = allData[0];
  // Current indices
  const iID   = hdrs.indexOf('PartID');
  const iItem = hdrs.indexOf('ItemName');
  const iBrand= hdrs.indexOf('Brand');
  const iSize = hdrs.indexOf('Size');
  const iCat  = hdrs.indexOf('Category');
  const iCost = hdrs.indexOf('CostPrice');
  const iSell = hdrs.indexOf('SellingPrice');
  const iQty  = hdrs.indexOf('StockQty');
  const iReorder = hdrs.indexOf('ReorderLevel');
  const iUnit = hdrs.indexOf('Unit');
  const iMove = hdrs.indexOf('Movement');
  const iSup  = hdrs.indexOf('SupplierName');
  const iDate = hdrs.indexOf('CreatedAt');

  // New order: PartID, ItemName, Size, Brand, Category, CostPrice, SellingPrice, StockQty, ReorderLevel, Unit, Movement, SupplierName, CreatedAt
  const NEW_HEADERS = ['PartID','ItemName','Size','Brand','Category','CostPrice','SellingPrice','StockQty','ReorderLevel','Unit','Movement','SupplierName','CreatedAt'];

  // Check if already in correct order
  const already = NEW_HEADERS.every((h, i) => hdrs[i] === h || (h === 'SupplierName' && !hdrs.includes('SupplierName')));

  const newRows = [NEW_HEADERS];
  for (let i = 1; i < allData.length; i++) {
    const r = allData[i];
    if (!r[iID] || String(r[iID]).trim() === '') continue;
    newRows.push([
      r[iID]   || '',
      r[iItem] || '',
      r[iSize] || '',   // Size now before Brand
      r[iBrand]|| '',
      r[iCat]  || '',
      r[iCost] || 0,
      r[iSell] || 0,
      r[iQty]  || 0,
      r[iReorder] || 5,
      r[iUnit] || 'Piece',
      r[iMove] || 'Medium',
      iSup >= 0 ? (r[iSup] || '') : '',
      iDate >= 0 ? (r[iDate] || '') : ''
    ]);
  }

  // Clear and rewrite
  sheet.clearContents();
  sheet.getRange(1, 1, newRows.length, NEW_HEADERS.length).setValues(newRows);
  // Style header
  sheet.getRange(1, 1, 1, NEW_HEADERS.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');

  return { success: true, message: 'Reordered ' + (newRows.length-1) + ' rows. New order: PartID, ItemName, Size, Brand, Category, CostPrice, SellingPrice, StockQty, ReorderLevel, Unit, Movement, SupplierName, CreatedAt' };
}

// ============================================================
// CLEAN DUPLICATE & BAD DATA
// Run once: ?action=cleanPartsData
// ============================================================
function cleanPartsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, message: 'Nothing to clean' };

  const hdrs = rows[0];
  const iID   = hdrs.indexOf('PartID');
  const iItem = hdrs.indexOf('ItemName');
  const iBrand= hdrs.indexOf('Brand');
  const iSize = hdrs.indexOf('Size');
  const iCat  = hdrs.indexOf('Category');
  const iCost = hdrs.indexOf('CostPrice');
  const iSell = hdrs.indexOf('SellingPrice');
  const iQty  = hdrs.indexOf('StockQty');
  const iMove = hdrs.indexOf('Movement');
  const iDate = hdrs.indexOf('CreatedAt');

  const validMoves = ['Fast','Medium','Slow','Dead'];
  const rowsToDelete = [];
  const seen = {}; // key = itemName+brand+size → row index (keep last/best)

  // SIZE NORMALIZATION MAP
  const sizeNorm = {
    '2.75': '2.75-17',
    '3':    '3.00-18',
    '300':  '300-17',
    '3.00': '3.00-17',
  };

  // ITEM NAME NORMALIZATION (lowercase key → correct name)
  const itemNorm = {
    'tyre .00': 'Tyre',
    'tyre .00 - cc': 'Tyre',
    'tyre .00 - kevil': 'Tyre',
    'tyre .00 - kevla': 'Tyre',
  };

  // OLD NO-BRAND TYRE ENTRIES TO DELETE (ItemName patterns, no real brand)
  // These were added before brand system existed and are now replaced by branded entries
  const oldNoBrandPatterns = [
    { item: 'tyre 2.75-17', brand: '2.75-17', size: '' },
    { item: 'tyre 3.00-18', brand: '3.00-18', size: '' },
    { item: 'tyre 300-17',  brand: '300-17',  size: '' },
    { item: 'tyre 3.00-17', brand: '3.00-17', size: '' },
    // Tubes with size as brand (old format) - only delete if branded versions exist
  ];

  let fixedSizes = 0, fixedMoves = 0, deletedDups = 0;

  // First pass: normalize sizes, item names, movement
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[iID]) continue;

    let changed = false;

    // Fix size
    const rawSize = String(row[iSize] || '').trim();
    if (sizeNorm[rawSize]) {
      sheet.getRange(i+1, iSize+1).setValue(sizeNorm[rawSize]);
      row[iSize] = sizeNorm[rawSize];
      fixedSizes++;
      changed = true;
    }

    // Fix item name
    const rawItem = String(row[iItem] || '').trim().toLowerCase();
    for (const [bad, good] of Object.entries(itemNorm)) {
      if (rawItem === bad) {
        sheet.getRange(i+1, iItem+1).setValue(good);
        row[iItem] = good;
        changed = true;
      }
    }

    // Fix Movement if it contains a date or invalid value
    const movVal = String(row[iMove] || '');
    const isDateInMove = movVal.includes('GMT') || movVal.includes('2026') || movVal.includes('May') || movVal.includes('Apr') || movVal.includes('2025');
    if (isDateInMove || !validMoves.includes(movVal.trim())) {
      sheet.getRange(i+1, iMove+1).setValue('Medium');
      row[iMove] = 'Medium';
      fixedMoves++;
    }
  }

  // Delete old no-brand tyre entries that have been replaced by branded ones
  const freshCheck = sheet.getDataRange().getValues();
  const freshHdrs = freshCheck[0];
  const fiItem = freshHdrs.indexOf('ItemName');
  const fiBrand = freshHdrs.indexOf('Brand');
  const fiSize = freshHdrs.indexOf('Size');

  // Check which branded tyres exist
  const brandedTyreSizes = new Set();
  for (let i = 1; i < freshCheck.length; i++) {
    const r = freshCheck[i];
    if (!r[0]) continue;
    const iN = String(r[fiItem]||'').toLowerCase().trim();
    const br = String(r[fiBrand]||'').toLowerCase().trim();
    const sz = String(r[fiSize]||'').toLowerCase().trim();
    if (iN === 'tyre' && br && !['2.75-17','3.00-17','3.00-18','300-17','2.75','3','300'].includes(br)) {
      brandedTyreSizes.add(sz); // e.g. "2.75-17", "3.00-17"
    }
  }

  const oldNoBrandToDelete = [];
  for (let i = 1; i < freshCheck.length; i++) {
    const r = freshCheck[i];
    if (!r[0]) continue;
    const iN = String(r[fiItem]||'').toLowerCase().trim();
    const br = String(r[fiBrand]||'').toLowerCase().trim();
    const sz = String(r[fiSize]||'').toLowerCase().trim();

    // Old tyre entries where size was put in Brand column and no real brand
    if (iN.startsWith('tyre ') && ['2.75-17','3.00-17','3.00-18','300-17','2.75','3','300','3.00'].includes(br) && sz === '') {
      // Extract size from ItemName or Brand
      const extractedSize = br.includes('-') ? br : (sizeNorm[br] || br);
      if (brandedTyreSizes.has(extractedSize) || brandedTyreSizes.size > 0) {
        oldNoBrandToDelete.push(i + 1); // 1-based, delete later
      }
    }
  }

  // Also fix "Tyre .00 - X" entries — rename ItemName to "Tyre"
  for (let i = 1; i < freshCheck.length; i++) {
    const r = freshCheck[i];
    if (!r[0]) continue;
    const iN = String(r[fiItem]||'').trim();
    if (iN.toLowerCase().startsWith('tyre .00')) {
      sheet.getRange(i+1, fiItem+1).setValue('Tyre');
      fixedSizes++;
    }
  }

  // Delete old no-brand entries from bottom to top
  oldNoBrandToDelete.sort((a,b) => b - a);
  for (const rowNum of oldNoBrandToDelete) {
    sheet.deleteRow(rowNum);
    deletedDups++;
  }

  // Second pass: find duplicates
  // Key = normalize(itemName + '|' + brand + '|' + size)
  const normalize = s => String(s||'').toLowerCase().trim().replace(/\s+/g,' ');

  const seenKeys = {}; // key → array of row indices (1-based)

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[iID]) continue;

    const key = normalize(row[iItem]) + '|' + normalize(row[iBrand]) + '|' + normalize(row[iSize]);
    if (!seenKeys[key]) seenKeys[key] = [];
    seenKeys[key].push(i); // 0-based data index
  }

  // For each duplicate group, keep the one with:
  // 1. Most recent date, OR
  // 2. Higher stock qty, OR
  // 3. Non-zero prices
  const toDeleteIndices = [];

  for (const [key, indices] of Object.entries(seenKeys)) {
    if (indices.length <= 1) continue;

    // Score each row: higher = better to keep
    const scored = indices.map(i => {
      const row = rows[i];
      const hasPrice = (parseFloat(row[iCost])||0) > 0 && (parseFloat(row[iSell])||0) > 0;
      const qty = parseInt(row[iQty]) || 0;
      const dateStr = String(row[iDate] || '');
      const dateMs = dateStr.includes('2026') || dateStr.includes('2025')
        ? new Date(dateStr).getTime() || 0 : 0;
      const score = (hasPrice ? 1000 : 0) + qty + dateMs/1e12;
      return { i, score };
    });

    scored.sort((a,b) => b.score - a.score);
    // Keep first (best), delete rest
    for (let k = 1; k < scored.length; k++) {
      toDeleteIndices.push(scored[k].i + 1); // convert to 1-based row
      deletedDups++;
    }
  }

  // Delete rows from bottom to top to preserve indices
  toDeleteIndices.sort((a,b) => b - a);
  for (const rowNum of toDeleteIndices) {
    sheet.deleteRow(rowNum);
  }

  // Remove empty rows (rows with no PartID)
  let emptyDeleted = 0;
  const freshRows2 = sheet.getDataRange().getValues();
  const emptyToDelete2 = [];
  for (let i = 1; i < freshRows2.length; i++) {
    if (!freshRows2[i][0] || String(freshRows2[i][0]).trim() === '') {
      emptyToDelete2.push(i + 1);
    }
  }
  emptyToDelete2.sort((a, b) => b - a);
  for (const rowNum of emptyToDelete2) {
    sheet.deleteRow(rowNum);
    emptyDeleted++;
  }

  // Sort remaining rows by ItemName then Brand then Size
  const sortData = sheet.getDataRange().getValues();
  if (sortData.length > 2) {
    const header = sortData[0];
    const dataOnly = sortData.slice(1).filter(r => r[0] && String(r[0]).trim());
    const iiN = header.indexOf('ItemName');
    const iiBr = header.indexOf('Brand');
    const iiSz = header.indexOf('Size');
    dataOnly.sort((a, b) => {
      const nameA = String(a[iiN]||'').toLowerCase();
      const nameB = String(b[iiN]||'').toLowerCase();
      if (nameA !== nameB) return nameA.localeCompare(nameB);
      const brandA = String(a[iiBr]||'').toLowerCase();
      const brandB = String(b[iiBr]||'').toLowerCase();
      if (brandA !== brandB) return brandA.localeCompare(brandB);
      return String(a[iiSz]||'').localeCompare(String(b[iiSz]||''));
    });
    // Write sorted data back
    sheet.clearContents();
    sheet.appendRow(header);
    sheet.getRange(1,1,1,header.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    if (dataOnly.length > 0) {
      sheet.getRange(2, 1, dataOnly.length, header.length).setValues(dataOnly);
    }
  }

  return {
    success: true,
    message: 'Done: ' + deletedDups + ' old/duplicate entries removed, ' +
             fixedSizes + ' names/sizes fixed, ' +
             fixedMoves + ' movement values fixed, ' +
             emptyDeleted + ' empty rows cleared. Sheet sorted by ItemName.'
  };
}



// ============================================================
// FIX MIXED COLUMNS — fixes Brand/Size/ItemName mix-ups
// Run: ?action=fixMixedColumns
// ============================================================
function fixMixedColumns() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return { success: true, message: 'Nothing to fix' };

  const hdrs = allData[0];
  const iID    = hdrs.indexOf('PartID');
  const iItem  = hdrs.indexOf('ItemName');
  const iBrand = hdrs.indexOf('Brand');
  const iSize  = hdrs.indexOf('Size');
  const iCat   = hdrs.indexOf('Category');

  // Known brands list
  const KNOWN_BRANDS = ['CC','Titan','RZ','Kevla','MRF','Crocodile','DZ','D&K','Yog','Kevil',
    'TZL','NYC','DJM','NGK','Bosch','Generic','Yog','KQK','K&K','J&L','Oscally','Accel',
    'Total','Helix','Advance','X-Large'];

  // Known item names (base names without brand/size)
  const KNOWN_ITEMS = ['Spark Plug','Dimmer Switch','Head Lamp Bulb','Indicator Bulb','Plug Cap',
    'Switch','Bulb Holder','Indicator','Tail Light','Tail Light Switch','Horn','Wire System',
    'Head Lamp Case','Head Lamp Complete','Dashboard','Clutch Plate','Front Sprocket',
    'Rear Sprocket','Chain','Clutch Wire','Sprocket Set','Chain Adjuster','Clutch Lever',
    'Clutch Disk','Gear Lever','Kick Starter','Gear Shaft','Kick Starter Shaft','Brake Cable',
    'Brake Pad/Shoe','Brake Lever','Brake Rod/Line','Brake Line Spring','Brake Pedal',
    'Tube','Tyre','Front Wheel Axle','Rear Wheel Axle','Front Hub','Rear Hub','Choke Cable',
    'Raise Wire','Raise Cable','Raise Handle','Carburettor','Air Cleaner','Engine Oil',
    'Front Shock Absorber','Rear Shock Absorber','Seal Shock Absorber','Pipe Shock Absorber',
    'Paper Gasket','Steel Gasket','Gasket Maker','Neck Bearing','Oil Seal','Bearing Front Hub',
    'Bearing Rear Hub','Clutch Bearing','Crank Shaft Bearing','Cam Shaft Bearing',
    'Engine Valve','Piston','Valve','Fly Wheel','Cylinder Head','Cylinder Block','Crank Shaft',
    'Timing Chain','Oil Pump Cap','Tensioner Block','Rocker Armor','Tappet Arm','Tappet Bolt',
    'Oil Pump','Timing Sprocket','Front Mud Guard','Chain Guard','Side Cover','Handle Bar',
    'Leg Guard','Mirror','Centre Bolt','Single Stand','Double Stand','Clutch cable',
    'Brake pad/shoe','Gasket maker','Gear shaft','Head lamp bulb','Head lamp case',
    'Head lamp complete','Crank shaft','Cylinder block','Cylinder head'];

  const isBrand = (val) => {
    if (!val) return false;
    const v = val.trim();
    return KNOWN_BRANDS.some(b => b.toLowerCase() === v.toLowerCase());
  };

  const isSize = (val) => {
    if (!val) return false;
    const v = String(val).trim();
    // Sizes contain numbers, dashes, dots or known size words
    return /\d/.test(v) || /^(Standard|X-Large|Short|Long|Front|Rear|Left|Right|Pair|Set|Small|Big|Small size|Big size|Metalic|Old Mode|New Mode)$/i.test(v);
  };

  const isItemName = (val) => {
    if (!val) return false;
    const v = String(val).trim();
    return KNOWN_ITEMS.some(item => item.toLowerCase() === v.toLowerCase()) || v.length > 6;
  };

  let fixed = 0;

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (!row[iID] || String(row[iID]).trim() === '') continue;

    var itemName = String(row[iItem] || '').trim();
    var brand    = String(row[iBrand] || '').trim();
    var size     = String(row[iSize]  || '').trim();
    var changed  = false;

    // PATTERN 1: Brand is in ItemName col, real ItemName is in Size col
    // e.g. ItemName="D&K", Brand="D&K", Size="Gear lever"
    if (isBrand(itemName) && isItemName(size) && !isItemName(itemName)) {
      var realItem = size;
      var realBrand = itemName;
      sheet.getRange(i+1, iItem+1).setValue(realItem);
      sheet.getRange(i+1, iBrand+1).setValue(realBrand);
      sheet.getRange(i+1, iSize+1).setValue('');
      fixed++; changed = true;
    }

    // PATTERN 2: ItemName is actually a brand name, real item is in Size
    // e.g. ItemName="Crocodile", Size="Head lamp bulb"
    else if (isBrand(itemName) && isItemName(size)) {
      sheet.getRange(i+1, iItem+1).setValue(size);
      sheet.getRange(i+1, iBrand+1).setValue(itemName);
      sheet.getRange(i+1, iSize+1).setValue('');
      fixed++; changed = true;
    }

    // PATTERN 3: ItemName contains "ItemName - Brand - Brand" duplication
    // e.g. ItemName="Chain - Yog", Brand="Yog" — remove brand from ItemName
    if (!changed && brand && itemName.endsWith(' - ' + brand)) {
      sheet.getRange(i+1, iItem+1).setValue(itemName.slice(0, -(3 + brand.length)).trim());
      fixed++; changed = true;
    }

    // PATTERN 4: Size is empty — set to empty string (not "Standard")
    // Only set Standard for brake pads/shoes explicitly
    if (!changed) {
      var freshItem = String(sheet.getRange(i+1, iItem+1).getValue() || '').trim().toLowerCase();
      if (!size && (freshItem.includes('brake pad') || freshItem.includes('brake pad/shoe'))) {
        if (!size) {
          sheet.getRange(i+1, iSize+1).setValue('Standard');
          fixed++;
        }
      }
    }
  }

  return {
    success: true,
    message: 'Fixed ' + fixed + ' rows with mixed columns'
  };
}

// ============================================================
// FIX STOCK QTY + ITEM NAMES — aggressive version
// Run: ?action=fixStockQty
// ============================================================
function fixStockQty(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return { success: true, message: 'Nothing to fix' };

  const hdrs = allData[0];
  const iID    = hdrs.indexOf('PartID');
  const iItem  = hdrs.indexOf('ItemName');
  const iBrand = hdrs.indexOf('Brand');
  const iSize  = hdrs.indexOf('Size');
  const iQty   = hdrs.indexOf('StockQty');

  const CATS = ['Engine','Transmission','Brakes','Electrical','Body/Panel','Suspension',
                'Fuel System','Tyres & Tubes','Filters','Bearings & Seals','Gaskets','Other'];

  // Reset ALL stock quantities — anything that looks wrong
  // "Wrong" = not a whole number between 0 and 200
  const maxReasonable = 200;
  let fixedQty = 0, fixedNames = 0;

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (!row[iID] || String(row[iID]).trim() === '') continue;

    // Fix qty — reset anything above maxReasonable to 0
    var rawQty = row[iQty];
    var qty = parseInt(rawQty);
    if (!isNaN(qty) && qty > maxReasonable) {
      sheet.getRange(i + 1, iQty + 1).setValue(0);
      fixedQty++;
    }

    // Fix item names
    if (iItem >= 0) {
      var itemName = String(row[iItem] || '').trim();
      var brand    = String(row[iBrand] || '').trim();
      var size     = String(row[iSize]  || '').trim();
      var changed  = false;
      var original = itemName;

      // Remove size text that appears inside item name
      if (size && size.length > 1) {
        var sIdx = itemName.toLowerCase().indexOf(size.toLowerCase());
        if (sIdx >= 0) {
          itemName = (itemName.slice(0, sIdx) + itemName.slice(sIdx + size.length));
          changed = true;
        }
      }

      // Remove brand text that appears inside item name: "Item - Brand"
      if (brand && itemName.indexOf(' - ' + brand) >= 0) {
        itemName = itemName.replace(' - ' + brand, '');
        changed = true;
      }
      // Also handle "- Brand" at start
      if (brand && itemName.startsWith('- ')) {
        itemName = itemName.slice(2);
        changed = true;
      }

      // Remove category names merged in
      for (var c = 0; c < CATS.length; c++) {
        var cat = CATS[c];
        if (itemName.indexOf(' - ' + cat) >= 0) {
          itemName = itemName.replace(' - ' + cat, '');
          changed = true;
        }
        if (itemName.indexOf(' ' + cat) >= 0 && itemName.endsWith(cat)) {
          itemName = itemName.slice(0, itemName.lastIndexOf(' ' + cat));
          changed = true;
        }
      }

      // Clean up spaces
      itemName = itemName.replace(/\s+/g, ' ').trim();

      if (changed && itemName && itemName !== original) {
        sheet.getRange(i + 1, iItem + 1).setValue(itemName);
        fixedNames++;
      }
    }
  }

  return {
    success: true,
    message: 'Fixed: ' + fixedQty + ' qty values reset to 0, ' + fixedNames + ' item names cleaned'
  };
}
// ============================================================
// FIX MOVEMENT COLUMN — repairs rows where date landed in Movement
// Run once: ?action=fixMovementColumn
// ============================================================
function fixMovementColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, message: 'Nothing to fix' };

  const hdrs = rows[0];
  // Find Movement and CreatedAt column indices (0-based)
  const movCol = hdrs.indexOf('Movement');
  const dateCol = hdrs.indexOf('CreatedAt');

  if (movCol < 0) return { success: false, error: 'Movement column not found' };

  const validMoves = ['Fast', 'Medium', 'Slow', 'Dead'];
  let fixed = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;

    const movVal = String(row[movCol] || '');
    const dateVal = String(row[dateCol] || '');

    // If Movement column contains a date string, fix it
    const isDate = movVal.includes('GMT') || movVal.includes('2026') || movVal.includes('2025') || movVal.includes('May') || movVal.includes('Apr');
    const isValidMove = validMoves.includes(movVal.trim());

    if (isDate || !isValidMove) {
      // Set Movement to Medium (default) and fix CreatedAt if needed
      sheet.getRange(i + 1, movCol + 1).setValue('Medium');

      // If CreatedAt is empty but Movement had a date, move date to CreatedAt
      if (isDate && (!dateVal || dateVal === '')) {
        sheet.getRange(i + 1, dateCol + 1).setValue(movVal);
      } else if (isDate && dateCol >= 0) {
        // Keep existing date, just clear Movement to Medium
      }
      fixed++;
    }
  }

  return { success: true, message: 'Fixed ' + fixed + ' rows with bad Movement values' };
}


// ============================================================
// NUCLEAR RESET — force reset ALL qty > 100 to 0
// Run: ?action=resetHighQty
// ============================================================
function resetHighQty() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  const data = sheet.getDataRange().getValues();
  const hdrs = data[0];
  
  // Find StockQty column - try both col H (index 7) and by header name
  let qtyCol = hdrs.indexOf('StockQty');
  if (qtyCol < 0) qtyCol = 7; // default to col H
  
  const fixed = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const qty = parseFloat(data[i][qtyCol]);
    if (!isNaN(qty) && qty > 100) {
      sheet.getRange(i + 1, qtyCol + 1).setValue(0);
      fixed.push(String(data[i][1]) + ': ' + qty + ' → 0');
    }
  }
  
  return { 
    success: true, 
    message: 'Reset ' + fixed.length + ' items',
    fixed: fixed
  };
}

// ============================================================
// CLEAR PARTS
// ============================================================
function clearParts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.PARTS);
  if (!sheet) return { success: false, error: 'Parts sheet not found' };
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return { success: true, message: 'Parts sheet cleared' };
}



// ============================================================
// AUTO-UPDATE SUPPLIER RATES
// ============================================================
function updateSupplierRate(partId, partName, supplierName, unitPrice) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rateSheet = ss.getSheetByName(SHEETS.SUPPLIER_RATES);
    if (!rateSheet) return;

    // Find the supplier
    const supSheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    const supRows = supSheet ? supSheet.getDataRange().getValues() : [];
    let supId = '';
    for (let i = 1; i < supRows.length; i++) {
      if (String(supRows[i][1]||'').trim().toLowerCase() === supplierName.toLowerCase()) {
        supId = String(supRows[i][0]);
        break;
      }
    }

    // Check if rate already exists for this part+supplier
    const rateRows = rateSheet.getDataRange().getValues();
    const rateHdrs = rateRows[0];
    const iRateID   = rateHdrs.indexOf('RateID');
    const iSupID    = rateHdrs.indexOf('SupplierID');
    const iSupName  = rateHdrs.indexOf('SupplierName');
    const iPartID   = rateHdrs.indexOf('PartID');
    const iPartName = rateHdrs.indexOf('PartName');
    const iPrice    = rateHdrs.indexOf('UnitPrice');
    const iUpdated  = rateHdrs.indexOf('UpdatedAt');

    const now = formatDate(new Date());

    for (let i = 1; i < rateRows.length; i++) {
      const r = rateRows[i];
      const sameSupplier = supId ? String(r[iSupID]) === supId : String(r[iSupName]||'').toLowerCase() === supplierName.toLowerCase();
      const samePart = String(r[iPartID]) === String(partId);
      if (sameSupplier && samePart) {
        // Update existing rate
        if (iPrice >= 0) rateSheet.getRange(i+1, iPrice+1).setValue(unitPrice);
        if (iUpdated >= 0) rateSheet.getRange(i+1, iUpdated+1).setValue(now);
        return;
      }
    }

    // No existing rate — add new one
    const newRateId = 'RAT-' + Date.now();
    rateSheet.appendRow([
      newRateId,
      supId || '',
      supplierName,
      partId,
      partName,
      unitPrice,
      '',   // Notes
      now
    ]);
  } catch(e) {
    Logger.log('updateSupplierRate error: ' + e.message);
  }
}

// ============================================================
// SORT PARTS SHEET by ItemName → Size → Brand
// ============================================================
function sortPartsSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PARTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) return;
    // Sort by col B (ItemName), then C (Size), then D (Brand) — all ascending
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    range.sort([{column: 2, ascending: true}, {column: 3, ascending: true}, {column: 4, ascending: true}]);
  } catch(e) {
    Logger.log('Sort failed: ' + e.message);
  }
}

// ============================================================
// UTILS
// ============================================================
function formatDate(date) {
  const d = date.getDate().toString().padStart(2, '0');
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const m = months[date.getMonth()];
  const y = date.getFullYear();
  return `${d}-${m}-${y}`;
}
