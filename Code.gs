// ============================================================
// குடும்ப பிரார்த்தனை நிகழ்வு — Family Prayer Event Tracker
// Google Apps Script Backend (Code.gs)
// ============================================================
// SETUP:
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Delete any existing code in Code.gs
// 3. Paste this entire file
// 4. Click Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL → paste in app.jsx as APPS_SCRIPT_URL
// ============================================================

// Sheet names — must match exactly
var SHEETS = {
  MASTER: "Master",
  FAMILIES: "Families",
  INCOME: "Income",
  EXPENSES: "Expenses",
  AUCTION: "Auction",
  LEDGER: "Ledger",
  ACCESS: "AccessControl",
  AUDIT: "AuditLog",
  LOANS: "Loans",
  FILES: "Files"
};

// Drive folder (bytes live here; Sheet holds only the index row)
var FILES_FOLDER_PROP = "SANGAM_FILES_FOLDER_ID";
var FILES_FOLDER_NAME = "Nagammal Sangam Files";
var MAX_FILE_BYTES = 10 * 1024 * 1024; // 10 MB cap per upload
var ALLOWED_FILE_TYPES = [
  "application/pdf",
  "image/jpeg", "image/png", "image/gif", "image/webp", "image/heic",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
  "application/vnd.ms-excel", // .xls
  "text/csv"
];

// ============================================================
// ENTRY POINTS
// ============================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var params;
    if (e.postData) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter || {};
    }

    var action = params.action || "";
    var email = (params.email || "").toLowerCase().trim();

    // Validate user
    var userRole = getUserRole(email);

    // Public actions (no role needed)
    if (action === "checkAccess") {
      return jsonResponse({ success: true, role: userRole, email: email });
    }

    // All other actions need at least Viewer role
    if (!userRole) {
      // For loadAll, still return role info so frontend can show proper message
      if (action === "loadAll") {
        return jsonResponse({ success: true, data: { role: null, email: email, debug: "Email not found in AccessControl sheet" } });
      }
      return jsonResponse({ success: false, error: "Access denied. Email not in AccessControl." });
    }

    // Route actions
    switch (action) {
      // --- READ (Viewer+) ---
      case "loadAll":
        return handleLoadAll(email, userRole);

      // --- DEBUG (temporary) ---
      case "debugAccess":
        var debugSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
        var debugData = debugSheet ? debugSheet.getDataRange().getValues() : [];
        var debugInfo = [];
        for (var d = 0; d < debugData.length; d++) {
          debugInfo.push(debugData[d].map(function(c) { return String(c).trim(); }));
        }
        return jsonResponse({ success: true, data: { rows: debugInfo, searchEmail: email, foundRole: userRole } });

      // --- WRITE (Editor+) ---
      case "addIncome":
        return requireRole(userRole, "Editor", function() { return handleAddIncome(params.data, email); });
      case "addExpense":
        return requireRole(userRole, "Editor", function() { return handleAddExpense(params.data, email); });
      case "addFamily":
        return requireRole(userRole, "Editor", function() { return handleAddFamily(params.data, email); });
      case "bulkAddFamilies":
        return requireRole(userRole, "Admin", function() { return handleBulkAddFamilies(params.data, email); });
      case "addAuctionItem":
        return requireRole(userRole, "Admin", function() { return handleAddAuctionItem(params.data, email); });
      case "enterBid":
        return requireRole(userRole, "Editor", function() { return handleEnterBid(params.data, email); });
      case "uploadAuctionPhoto":
        return requireRole(userRole, "Editor", function() { return handleUploadAuctionPhoto(params.data, email); });

      // --- ADMIN: Edit/Delete Income & Expenses ---
      case "editIncome":
        return requireRole(userRole, "Admin", function() { return handleEditIncome(params.data, email); });
      case "deleteIncome":
        return requireRole(userRole, "Admin", function() { return handleDeleteIncome(params.data, email); });
      case "editExpense":
        return requireRole(userRole, "Admin", function() { return handleEditExpense(params.data, email); });
      case "deleteExpense":
        return requireRole(userRole, "Admin", function() { return handleDeleteExpense(params.data, email); });

      case "saveEventConfig":
        return requireRole(userRole, "Admin", function() { return handleSaveEventConfig(params.data, email); });

      case "saveReconciliation":
        return requireRole(userRole, "Admin", function() { return handleSaveReconciliation(params.data, email); });

      case "createFinancialYear":
        return requireRole(userRole, "Admin", function() { return handleCreateFinancialYear(params.data, email); });
      case "confirmIncome":
        return requireRole(userRole, "Editor", function() { return handleConfirmIncome(params.data, email); });
      case "collectReceivable":
        return requireRole(userRole, "Editor", function() { return handleCollectReceivable(params.data, email); });
      case "recalcInterest":
        return requireRole(userRole, "Admin", function() { return handleRecalcInterest(params.data, email); });

      case "seedAuctionItems":
        return requireRole(userRole, "Admin", function() { return handleSeedAuctionItems(params.data, email); });

      // --- FILES (shared documents) ---
      case "uploadFile":
        return requireRole(userRole, "Admin", function() { return handleUploadFile(params.data, email); });
      case "getFile":
        return handleGetFile(params.data, email); // any logged-in role (Viewer+)
      case "deleteFile":
        return requireRole(userRole, "Admin", function() { return handleDeleteFile(params.data, email); });

      // --- LOANS (Pangali Advances) ---
      case "addLoan":
        return requireRole(userRole, "Admin", function() { return handleAddLoan(params.data, email); });
      case "editLoan":
        return requireRole(userRole, "Admin", function() { return handleEditLoan(params.data, email); });
      case "deleteLoan":
        return requireRole(userRole, "Admin", function() { return handleDeleteLoan(params.data, email); });

      default:
        return jsonResponse({ success: false, error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ============================================================
// AUTH & ROLES
// ============================================================

function getUserRole(email) {
  if (!email) return null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  
  // Find Email and Role columns by header name (row 0)
  var headers = data[0];
  var emailCol = -1;
  var roleCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var hdr = String(headers[h]).trim().toLowerCase();
    if (hdr === "email") emailCol = h;
    if (hdr === "role") roleCol = h;
  }
  if (emailCol === -1 || roleCol === -1) return null;
  
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol]).toLowerCase().trim();
    if (rowEmail === email) {
      return String(data[i][roleCol]).trim();
    }
  }
  return null;
}

function requireRole(currentRole, minimumRole, handler) {
  var hierarchy = { "Viewer": 1, "Editor": 2, "Admin": 3 };
  if ((hierarchy[currentRole] || 0) >= (hierarchy[minimumRole] || 99)) {
    return handler();
  }
  return jsonResponse({ success: false, error: "Insufficient permissions. Need " + minimumRole + " role." });
}

// ============================================================
// LOAD ALL DATA
// ============================================================

function handleLoadAll(email, userRole) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var result = {
    master: readMaster(ss),
    families: readSheet(ss, SHEETS.FAMILIES),
    income: readSheet(ss, SHEETS.INCOME),
    expenses: readSheet(ss, SHEETS.EXPENSES),
    auction: readSheet(ss, SHEETS.AUCTION),
    access: readSheet(ss, SHEETS.ACCESS),
    ledger: readSheet(ss, SHEETS.LEDGER),
    loans: readSheet(ss, SHEETS.LOANS),
    files: readSheet(ss, SHEETS.FILES),
    role: userRole,
  };

  // Audit log only for Admin
  if (userRole === "Admin") {
    result.auditLog = readSheet(ss, SHEETS.AUDIT);
  }

  writeAudit("LOGIN", "Auth", email + " loaded data", email);

  return jsonResponse({ success: true, data: result });
}

function readMaster(ss) {
  var sheet = ss.getSheetByName(SHEETS.MASTER);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var headers = data[0];
  var result = {};
  for (var col = 0; col < headers.length; col++) {
    var key = String(headers[col]).trim();
    if (!key) continue;
    var values = [];
    for (var row = 1; row < data.length; row++) {
      var val = String(data[row][col] || "").trim();
      if (val) values.push(val);
    }
    result[key] = values;
  }
  return result;
}

function readSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j] !== undefined ? String(data[i][j]) : "";
    }
    rows.push(obj);
  }
  return rows;
}

// ============================================================
// INCOME
// ============================================================

function handleAddIncome(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  var txnId = generateId("DON");
  var ts = now();

  var year = data.year || String(new Date().getFullYear());

  var row = [
    txnId,
    year,
    data.day,
    data.date || "",
    data.donorType,
    data.familyId || "",
    data.donorName,
    data.category,
    data.amount || 0,
    data.description || "",
    email,
    ts
  ];
  sheet.appendRow(row);

  // If Status column exists, mark as Confirmed (the default for manual entries)
  var statusIdx = getHeaderIndex(sheet, "Status");
  if (statusIdx >= 0) {
    sheet.getRange(sheet.getLastRow(), statusIdx + 1).setValue("Confirmed");
  }

  writeAudit("CREATE", "Income", txnId + " ₹" + data.amount + " " + data.category + " by " + data.donorName, email);

  return jsonResponse({ success: true, data: { txnId: txnId } });
}

// ============================================================
// EXPENSES
// ============================================================

function handleAddExpense(data, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EXPENSES);
  var txnId = generateId("EXP");
  var ts = now();

  var year = data.year || String(new Date().getFullYear());

  var row = [
    txnId,
    year,
    data.day,
    data.date || "",
    data.category,
    data.description,
    data.amount,
    data.vendor || "",
    email,
    ts
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Expenses", txnId + " ₹" + data.amount + " " + data.category + " - " + data.description, email);

  return jsonResponse({ success: true, data: { txnId: txnId } });
}

// ============================================================
// EDIT / DELETE INCOME (Admin only)
// ============================================================

function handleEditIncome(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  var allData = sheet.getDataRange().getValues();
  var txnId = data.txnId;

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === txnId) {
      var rowNum = i + 1;
      var year = data.year || String(allData[i][1]);
      var row = [
        txnId,
        year,
        data.day,
        data.date || "",
        data.donorType,
        data.familyId || "",
        data.donorName,
        data.category,
        data.amount || 0,
        data.description || "",
        String(allData[i][10]),  // keep original EnteredBy
        now()                    // update timestamp
      ];
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      writeAudit("EDIT", "Income", txnId + " edited by admin", email);
      return jsonResponse({ success: true, data: { txnId: txnId } });
    }
  }
  return jsonResponse({ success: false, error: "Transaction not found: " + txnId });
}

function handleDeleteIncome(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  var allData = sheet.getDataRange().getValues();
  var txnId = data.txnId;

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === txnId) {
      sheet.deleteRow(i + 1);
      writeAudit("DELETE", "Income", txnId + " deleted by admin", email);
      return jsonResponse({ success: true, data: { txnId: txnId } });
    }
  }
  return jsonResponse({ success: false, error: "Transaction not found: " + txnId });
}

// ============================================================
// EDIT / DELETE EXPENSES (Admin only)
// ============================================================

function handleEditExpense(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.EXPENSES);
  var allData = sheet.getDataRange().getValues();
  var txnId = data.txnId;

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === txnId) {
      var rowNum = i + 1;
      var year = data.year || String(allData[i][1]);
      var row = [
        txnId,
        year,
        data.day,
        data.date || "",
        data.category,
        data.description,
        data.amount,
        data.vendor || "",
        String(allData[i][8]),  // keep original EnteredBy
        now()                   // update timestamp
      ];
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      writeAudit("EDIT", "Expenses", txnId + " edited by admin", email);
      return jsonResponse({ success: true, data: { txnId: txnId } });
    }
  }
  return jsonResponse({ success: false, error: "Transaction not found: " + txnId });
}

function handleDeleteExpense(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.EXPENSES);
  var allData = sheet.getDataRange().getValues();
  var txnId = data.txnId;

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === txnId) {
      sheet.deleteRow(i + 1);
      writeAudit("DELETE", "Expenses", txnId + " deleted by admin", email);
      return jsonResponse({ success: true, data: { txnId: txnId } });
    }
  }
  return jsonResponse({ success: false, error: "Transaction not found: " + txnId });
}

// ============================================================
// FAMILIES
// ============================================================

function handleAddFamily(data, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FAMILIES);
  var famId = generateId("FAM");

  var row = [
    famId,
    data.familyName,
    data.phone,
    data.whatsApp || "",
    data.address
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Families", famId + " " + data.familyName, email);

  return jsonResponse({ success: true, data: { familyId: famId } });
}

function handleBulkAddFamilies(data, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FAMILIES);
  var families = data.families || [];
  var count = 0;
  for (var i = 0; i < families.length; i++) {
    var f = families[i];
    var row = [
      f.familyId || generateId("FAM"),
      f.familyName || "",
      f.phone || "",
      f.whatsApp || "",
      f.address || ""
    ];
    sheet.appendRow(row);
    count++;
  }
  writeAudit("BULK_CREATE", "Families", count + " families added", email);
  return jsonResponse({ success: true, data: { count: count } });
}

// ============================================================
// AUCTION — Add Item (Admin only)
// ============================================================

function handleAddAuctionItem(data, email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUCTION);
  var nextItemNo = sheet.getLastRow(); // header row = 1, so lastRow = current item count + 1
  var year = data.year || String(new Date().getFullYear());

  var row = [
    nextItemNo,
    year,
    data.itemName,
    data.donatedBy || "",
    "", "", "", "", "", "", "", ""
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Auction", "Manual add #" + nextItemNo + " \"" + data.itemName + "\"", email);

  return jsonResponse({ success: true, data: { itemNo: nextItemNo } });
}

// ============================================================
// AUCTION — Enter/Confirm Bid
// ============================================================

function handleEnterBid(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUCTION);
  var rowIndex = data.rowIndex + 2; // 0-based data row → 1-based sheet row (header + 1)

  // Auction columns (0-indexed):
  // A=ItemNo(0), B=Year(1), C=ItemDescription(2), D=DonatedBy(3),
  // E=WinnerFamilyID(4), F=WinnerName(5), G=BidAmount(6),
  // H=EnteredBy1(7), I=ConfirmedBy2(8), J=Timestamp1(9), K=Timestamp2(10), L=Status(11)

  var range = sheet.getRange(rowIndex, 1, 1, 12);
  var values = range.getValues()[0];

  var enteredBy1 = String(values[7]).trim();   // Column H: EnteredBy1
  var confirmedBy2 = String(values[8]).trim(); // Column I: ConfirmedBy2
  var ts = now();

  if (!enteredBy1) {
    // First entry — enter bid
    values[4] = data.familyId || "";     // WinnerFamilyID
    values[5] = data.winnerName;         // WinnerName
    values[6] = data.amount;             // BidAmount
    values[7] = email;               // EnteredBy1
    values[9] = ts;                      // Timestamp1
    values[11] = "Pending";              // Status

    range.setValues([values]);
    writeAudit("BID", "Auction", "#" + values[0] + " \"" + values[2] + "\" ₹" + data.amount + " by " + data.winnerName, email);

    return jsonResponse({ success: true, data: { status: "Pending" } });

  } else if (enteredBy1.toLowerCase() !== email.toLowerCase() && !confirmedBy2) {
    // Second person — confirm
    values[8] = email;               // ConfirmedBy2
    values[10] = ts;                     // Timestamp2
    values[11] = "Confirmed";            // Status

    range.setValues([values]);
    writeAudit("CONFIRM", "Auction", "#" + values[0] + " \"" + values[2] + "\" ₹" + values[6] + " to " + values[5], email);

    return jsonResponse({ success: true, data: { status: "Confirmed" } });

  } else if (enteredBy1.toLowerCase() === email.toLowerCase()) {
    return jsonResponse({ success: false, error: "Cannot confirm your own entry." });

  } else {
    return jsonResponse({ success: false, error: "Bid already confirmed." });
  }
}

// ============================================================
// AUDIT LOG
// ============================================================

// ============================================================
// AUCTION — Upload Photo
// ============================================================

function handleUploadAuctionPhoto(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUCTION);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find or create Photo column
  var photoCol = headers.indexOf("Photo");
  if (photoCol === -1) {
    // Add Photo header
    photoCol = headers.length;
    sheet.getRange(1, photoCol + 1).setValue("Photo");
  }
  
  // Find the row for this item
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.itemNo)) {
      sheet.getRange(i + 1, photoCol + 1).setValue(data.photo);
      writeAudit("UPDATE", "Auction", "Photo uploaded for item #" + data.itemNo, email);
      return jsonResponse({ success: true, data: { itemNo: data.itemNo } });
    }
  }
  
  return jsonResponse({ success: false, error: "Item #" + data.itemNo + " not found" });
}

// ============================================================
// FILES — Admin upload, any-logged-in download (Drive-backed)
// ============================================================

// Returns the shared Drive folder, creating it once and caching the id.
function getFilesFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(FILES_FOLDER_PROP);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* folder gone; recreate */ }
  }
  var folder = DriveApp.createFolder(FILES_FOLDER_NAME);
  props.setProperty(FILES_FOLDER_PROP, folder.getId());
  return folder;
}

// Ensures the Files index tab exists with the right header row.
function getFilesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.FILES);
  var headers = ["FileID", "FileName", "FileType", "Size", "DriveFileID",
                 "Category", "Description", "UploadedBy", "UploadedDate", "Year", "Status"];
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.FILES);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handleUploadFile(data, email) {
  try {
    if (!data || !data.base64 || !data.fileName) {
      return jsonResponse({ success: false, error: "Missing file data." });
    }
    var mime = String(data.fileType || "").trim();
    if (ALLOWED_FILE_TYPES.indexOf(mime) === -1) {
      return jsonResponse({ success: false, error: "File type not allowed: " + (mime || "unknown") + ". Allowed: PDF, images, Excel/CSV." });
    }

    // base64 may arrive as a data URL ("data:...;base64,XXXX") or raw base64
    var b64 = String(data.base64);
    var comma = b64.indexOf(",");
    if (b64.substring(0, 5) === "data:" && comma !== -1) b64 = b64.substring(comma + 1);

    var bytes = Utilities.base64Decode(b64);
    if (bytes.length > MAX_FILE_BYTES) {
      return jsonResponse({ success: false, error: "File too large (" + Math.round(bytes.length / 1048576 * 10) / 10 + " MB). Max 10 MB." });
    }

    var blob = Utilities.newBlob(bytes, mime, data.fileName);
    var folder = getFilesFolder_();
    var driveFile = folder.createFile(blob);
    // Anyone with the link can view (download is link-gated by the app login UI)
    try { driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var fileId = generateId("FILE");
    var sheet = getFilesSheet_();
    var uploadDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var row = [
      fileId,
      String(data.fileName),
      mime,
      bytes.length,
      driveFile.getId(),
      String(data.category || ""),
      String(data.description || ""),
      email,
      uploadDate,
      String(data.year || ""),
      "Active"
    ];
    var newRowNo = sheet.getLastRow() + 1;
    sheet.getRange(newRowNo, 1, 1, row.length).setValues([row]);
    // Keep date + year as plain text
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var dIdx = hdr.indexOf("UploadedDate");
    if (dIdx !== -1) sheet.getRange(newRowNo, dIdx + 1).setNumberFormat("@").setValue(uploadDate);

    writeAudit("UPLOAD", "Files", "Uploaded " + data.fileName + " (" + fileId + ")", email);
    return jsonResponse({ success: true, data: { fileId: fileId, driveFileId: driveFile.getId() } });
  } catch (err) {
    return jsonResponse({ success: false, error: "Upload failed: " + err.message });
  }
}

// Returns the file bytes as base64 so the browser can download even if the
// Drive file's sharing is restricted. Any logged-in role may call this.
function handleGetFile(data, email) {
  try {
    if (!data || !data.fileId) return jsonResponse({ success: false, error: "Missing fileId." });
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FILES);
    if (!sheet) return jsonResponse({ success: false, error: "No files." });
    var all = sheet.getDataRange().getValues();
    var hdr = all[0];
    var idIdx = hdr.indexOf("FileID");
    var driveIdx = hdr.indexOf("DriveFileID");
    var nameIdx = hdr.indexOf("FileName");
    var typeIdx = hdr.indexOf("FileType");
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][idIdx]) === String(data.fileId)) {
        var driveFile = DriveApp.getFileById(String(all[i][driveIdx]));
        var blob = driveFile.getBlob();
        writeAudit("DOWNLOAD", "Files", "Downloaded " + all[i][nameIdx] + " (" + data.fileId + ")", email);
        return jsonResponse({
          success: true,
          data: {
            fileName: String(all[i][nameIdx]),
            fileType: String(all[i][typeIdx]),
            base64: Utilities.base64Encode(blob.getBytes())
          }
        });
      }
    }
    return jsonResponse({ success: false, error: "File not found: " + data.fileId });
  } catch (err) {
    return jsonResponse({ success: false, error: "Download failed: " + err.message });
  }
}

function handleDeleteFile(data, email) {
  try {
    if (!data || !data.fileId) return jsonResponse({ success: false, error: "Missing fileId." });
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FILES);
    if (!sheet) return jsonResponse({ success: false, error: "No files." });
    var all = sheet.getDataRange().getValues();
    var hdr = all[0];
    var idIdx = hdr.indexOf("FileID");
    var driveIdx = hdr.indexOf("DriveFileID");
    var nameIdx = hdr.indexOf("FileName");
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][idIdx]) === String(data.fileId)) {
        try { DriveApp.getFileById(String(all[i][driveIdx])).setTrashed(true); } catch (e) {}
        writeAudit("DELETE", "Files", "Deleted " + all[i][nameIdx] + " (" + data.fileId + ")", email);
        sheet.deleteRow(i + 1);
        return jsonResponse({ success: true, data: { fileId: data.fileId } });
      }
    }
    return jsonResponse({ success: false, error: "File not found: " + data.fileId });
  } catch (err) {
    return jsonResponse({ success: false, error: "Delete failed: " + err.message });
  }
}

function writeAudit(action, module, details, email) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUDIT);
    if (sheet) {
      sheet.appendRow([now(), email, action, module, details]);
    }
  } catch (e) {
    // Silently fail — audit should never break main operations
    Logger.log("Audit write failed: " + e.message);
  }
}

// ============================================================
// HELPERS
// ============================================================

function generateId(prefix) {
  return prefix + "-" + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy, h:mm a");
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SAVE EVENT CONFIG (Ledger year + Day1 date)
// ============================================================

function handleSaveEventConfig(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LEDGER);
  if (!sheet) throw new Error("Ledger sheet not found");

  var year = String(data.year);
  var day1 = String(data.day1);

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var yearCol = headers.indexOf("Year");
  var day1Col = headers.indexOf("Day1");

  if (yearCol === -1 || day1Col === -1) throw new Error("Ledger sheet missing Year or Day1 columns");

  // Find existing row for this year
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][yearCol]) === year) {
      // Update existing row
      sheet.getRange(i + 1, day1Col + 1).setValue(day1);
      writeAudit("UPDATE", "Ledger", "Updated Day1 for " + year + " to " + day1, email);
      return jsonResponse({ success: true, data: { year: year, day1: day1 } });
    }
  }

  // Create new row
  var newRow = [];
  for (var c = 0; c < headers.length; c++) {
    if (c === yearCol) newRow.push(Number(year));
    else if (c === day1Col) newRow.push(day1);
    else newRow.push("");
  }
  sheet.appendRow(newRow);
  writeAudit("CREATE", "Ledger", "Created entry for " + year + " Day1=" + day1, email);
  return jsonResponse({ success: true, data: { year: year, day1: day1 } });
}

// ============================================================
// SAVE RECONCILIATION (Balance Sheet / ஐந்தொகை — Bank Deposit + Petty Cash)
// ============================================================
// Writes admin-entered Bank Deposit and Petty Cash figures to
// the Ledger sheet for the given year. Auto-creates the
// BankDeposit and PettyCash columns on first use so no manual
// schema migration is needed.
// ============================================================

function handleSaveReconciliation(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LEDGER);
  if (!sheet) throw new Error("Ledger sheet not found");

  var year = String(data.year);
  var bankDeposit = Number(data.bankDeposit) || 0;
  var pettyCash = Number(data.pettyCash) || 0;
  var notes = data.notes !== undefined ? String(data.notes) : null; // null = don't touch

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var yearCol = headers.indexOf("Year");
  if (yearCol === -1) throw new Error("Ledger sheet missing Year column");

  // Auto-create BankDeposit and PettyCash columns if missing
  var bdCol = headers.indexOf("BankDeposit");
  if (bdCol === -1) {
    bdCol = headers.length;
    sheet.getRange(1, bdCol + 1).setValue("BankDeposit");
    headers.push("BankDeposit");
  }
  var pcCol = headers.indexOf("PettyCash");
  if (pcCol === -1) {
    pcCol = headers.length;
    sheet.getRange(1, pcCol + 1).setValue("PettyCash");
    headers.push("PettyCash");
  }
  // Auto-create Notes column if missing
  var notesCol = headers.indexOf("Notes");
  if (notesCol === -1) {
    notesCol = headers.length;
    sheet.getRange(1, notesCol + 1).setValue("Notes");
    headers.push("Notes");
  }

  // Find existing row for this year
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][yearCol]) === year) {
      sheet.getRange(i + 1, bdCol + 1).setValue(bankDeposit);
      sheet.getRange(i + 1, pcCol + 1).setValue(pettyCash);
      if (notes !== null) sheet.getRange(i + 1, notesCol + 1).setValue(notes);
      writeAudit("UPDATE", "Ledger",
        "Balance Sheet reconciliation " + year + ": BankDeposit=" + bankDeposit + ", PettyCash=" + pettyCash + (notes !== null ? ", Notes updated" : ""),
        email);
      return jsonResponse({ success: true, data: { year: year, bankDeposit: bankDeposit, pettyCash: pettyCash, notes: notes || "" } });
    }
  }

  // No existing row — create new one
  var newRow = [];
  for (var c = 0; c < headers.length; c++) {
    if (c === yearCol) newRow.push(Number(year));
    else if (c === bdCol) newRow.push(bankDeposit);
    else if (c === pcCol) newRow.push(pettyCash);
    else if (c === notesCol) newRow.push(notes || "");
    else newRow.push("");
  }
  sheet.appendRow(newRow);
  writeAudit("CREATE", "Ledger",
    "Balance Sheet reconciliation " + year + " (new row): BankDeposit=" + bankDeposit + ", PettyCash=" + pettyCash,
    email);
  return jsonResponse({ success: true, data: { year: year, bankDeposit: bankDeposit, pettyCash: pettyCash, notes: notes || "" } });
}

// ============================================================
// CREATE FINANCIAL YEAR (FY rollover)
// ============================================================
// Creates a new Ledger row for max(year)+1 and pre-populates
// Income rows for each family:
//   - Pulli Vari (Amount = pulliDefault, Status = Expected)
// Plus each previous-year Pen Pulli Varavu donor (by donor name):
//   - Pen Pulli Varavu (Amount = penPulliDefault, Status = Expected)
// And for each family with an outstanding loan from previous year:
//   - Cash Receivable (Amount = LastYearReceivable + LoanAmount)
//   - Interest Receivable (Amount = InterestAmount)
//
// All rolled-over rows are marked Status="Expected" so they do NOT
// inflate income totals until an admin confirms collection.
// ============================================================

function handleCreateFinancialYear(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledgerSh = ss.getSheetByName(SHEETS.LEDGER);
  var incomeSh = ss.getSheetByName(SHEETS.INCOME);
  var familiesSh = ss.getSheetByName(SHEETS.FAMILIES);
  var loansSh = ss.getSheetByName(SHEETS.LOANS);
  if (!ledgerSh || !incomeSh || !familiesSh || !loansSh) throw new Error("Required sheet missing");

  // Determine new year
  var prevYear = Number(data.prevYear) || (new Date().getFullYear());
  var ledgerData = ledgerSh.getDataRange().getValues();
  var ledgerHeaders = ledgerData[0];
  var yearIdx = ledgerHeaders.indexOf("Year");
  if (yearIdx === -1) throw new Error("Ledger sheet missing Year column");

  var maxYear = prevYear;
  for (var i = 1; i < ledgerData.length; i++) {
    var y = Number(ledgerData[i][yearIdx]) || 0;
    if (y > maxYear) maxYear = y;
  }
  var newYear = maxYear + 1;

  // Validate that newYear isn't already in Ledger
  for (var j = 1; j < ledgerData.length; j++) {
    if (Number(ledgerData[j][yearIdx]) === newYear) {
      throw new Error("FY " + newYear + " already exists in Ledger");
    }
  }

  // Ensure Status column exists on Income sheet (flush so subsequent reads see it)
  ensureColumn(incomeSh, "Status");
  SpreadsheetApp.flush();

  // Default amounts (passed from frontend; falls back to defaults)
  var pulliDefault = Number(data.pulliDefault) || 51;
  var penPulliDefault = Number(data.penPulliDefault) || 2;

  // Read families
  var families = readSheetRows(familiesSh);
  // Read previous year's loans, indexed by FamilyID
  var loans = readSheetRows(loansSh);
  var prevLoansByFamily = {};
  loans.forEach(function(l) {
    if (Number(l.Year) === prevYear) {
      var fid = String(l.FamilyID || "").trim();
      if (fid) prevLoansByFamily[fid] = l;
    }
  });

  var prevIncome = readSheetRows(incomeSh);

  // Build previous-year receivable state per family from the Income rows.
  // For each family we track:
  //   cashExpected  — the obligation anchor (Expected/blank Cash Receivable)
  //   cashCollected — sum of Collected Cash Receivable rows
  //   intCollected  — interest already collected (the single interest row's
  //                   Amount once it is Collected)
  // Carry-forward rule (per user):
  //   unpaid principal  = cashExpected − cashCollected
  //   unpaid interest   = full-year interest on the UNPAID principal
  //                       − interest already collected
  //   next LastYearReceivable = unpaid principal + unpaid interest
  //   (uncollected interest folds into principal and compounds)
  var prevRecvByFamily = {};
  function recvFor(fid) {
    if (!prevRecvByFamily[fid]) prevRecvByFamily[fid] = {
      cashExpected: 0, cashCollected: 0, intCollected: 0, hasAny: false,
    };
    return prevRecvByFamily[fid];
  }
  prevIncome.forEach(function(r) {
    if (Number(r.Year) !== prevYear) return;
    var fid = String(r.FamilyID || "").trim();
    if (!fid) return;
    var cat = String(r.Category || "").trim().toLowerCase();
    var st = String(r.Status || "").trim();
    var amt = Number(r.Amount) || 0;
    if (cat === "cash receivable") {
      var rec = recvFor(fid); rec.hasAny = true;
      if (st === "Collected") rec.cashCollected += amt;
      else rec.cashExpected += amt; // Expected or legacy-blank = obligation
    } else if (cat === "interest receivable") {
      var rec2 = recvFor(fid); rec2.hasAny = true;
      if (st === "Collected" || st === "Confirmed") rec2.intCollected += amt;
      // Expected interest rows carry no collected value
    }
  });

  // Read previous year's Pen Pulli Varavu donors (by donor name — these are
  // external women donors, not the family roster). Deduplicate by donor name
  // so a donor who appears twice last year is carried forward once.
  var prevPenPulliDonors = [];
  var seenPenPulli = {};
  prevIncome.forEach(function(r) {
    if (Number(r.Year) === prevYear &&
        String(r.Category || "").trim().toLowerCase() === "pen pulli varavu") {
      var dname = String(r.DonorName || "").trim();
      var key = dname.toLowerCase();
      if (dname && !seenPenPulli[key]) {
        seenPenPulli[key] = true;
        prevPenPulliDonors.push({
          DonorName: dname,
          DonorType: String(r.DonorType || "External").trim() || "External",
          FamilyID: String(r.FamilyID || "").trim(),
        });
      }
    }
  });

  // Build new rows for Income sheet
  var incomeHeaders = incomeSh.getRange(1, 1, 1, incomeSh.getLastColumn()).getValues()[0];
  var ts = now();
  var newRows = [];
  var counts = { pulli: 0, penPulli: 0, cashRec: 0, intRec: 0, families: families.length };

  families.forEach(function(f) {
    var fid = String(f.FamilyID || "").trim();
    var fname = String(f.FamilyName || "").trim();
    if (!fid) return;

    // Pulli Vari (men's membership) — Expected
    newRows.push(buildIncomeRow(incomeHeaders, {
      TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
      DonorType: "Family", FamilyID: fid, DonorName: fname,
      Category: "Pulli Vari", Amount: pulliDefault, Description: "",
      EnteredBy: email, Timestamp: ts, Status: "Expected",
    }));
    counts.pulli++;

    // Pen Pulli Varavu is carried forward separately (by donor name, after
    // this family loop) — it's tracked by external women donors, not the
    // family roster.

    // ---- Carry forward unpaid receivables into the new year ----
    // Preferred source: the previous year's Income receivable rows
    // (Expected − Collected), which reflect partial repayments.
    // Fallback: the Loans sheet (for a prior year that predates the
    // partial-payment model, e.g. the first rollover from 2026).
    var carryPrincipal = 0;
    var rec = prevRecvByFamily[fid];
    if (rec && rec.hasAny) {
      var unpaidPrincipal = Math.max(0, rec.cashExpected - rec.cashCollected);
      // Unpaid interest = full-year (12mo) interest on the unpaid principal,
      // minus interest already collected during the year.
      var fullYearInterest = calcTimeInterest(unpaidPrincipal, 9, 12);
      var unpaidInterest = Math.max(0, fullYearInterest - rec.intCollected);
      // Uncollected interest folds into principal (compounds).
      carryPrincipal = unpaidPrincipal + unpaidInterest;
    } else {
      // Fallback: prior year stored only in Loans sheet.
      var prevLoan = prevLoansByFamily[fid];
      if (prevLoan) {
        var lastYrRec = Number(prevLoan.LastYearReceivable) || 0;
        var loanAmt = Number(prevLoan.LoanAmount) || 0;
        var intAmt = Number(prevLoan.InterestAmount) || 0;
        // Legacy carry-forward = principal + stored interest (folds in).
        carryPrincipal = lastYrRec + loanAmt + intAmt;
      }
    }

    if (carryPrincipal > 0) {
      // Cash Receivable (principal owed) — Expected.
      newRows.push(buildIncomeRow(incomeHeaders, {
        TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
        DonorType: "Family", FamilyID: fid, DonorName: fname,
        Category: "Cash Receivable", Amount: carryPrincipal,
        Description: "Carried from " + prevYear + " (principal + unpaid interest)",
        EnteredBy: email, Timestamp: ts, Status: "Expected",
      }));
      counts.cashRec++;

      // Interest Receivable — ONE row per family, starts at ₹0.
      // It accumulates as principal installments are collected (time-based).
      newRows.push(buildIncomeRow(incomeHeaders, {
        TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
        DonorType: "Family", FamilyID: fid, DonorName: fname,
        Category: "Interest Receivable", Amount: 0,
        Description: "Accrues on collection (9% × months)",
        EnteredBy: email, Timestamp: ts, Status: "Expected",
      }));
      counts.intRec++;
    }
  });

  // Pen Pulli Varavu — carry forward each previous-year donor at the default
  // amount, Status="Expected", so returning women donors are ready to confirm.
  prevPenPulliDonors.forEach(function(donor) {
    newRows.push(buildIncomeRow(incomeHeaders, {
      TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
      DonorType: donor.DonorType, FamilyID: donor.FamilyID, DonorName: donor.DonorName,
      Category: "Pen Pulli Varavu", Amount: penPulliDefault, Description: "Carried from " + prevYear,
      EnteredBy: email, Timestamp: ts, Status: "Expected",
    }));
    counts.penPulli++;
  });

  // Bulk write all new income rows in one batch
  if (newRows.length > 0) {
    var startRow = incomeSh.getLastRow() + 1;
    incomeSh.getRange(startRow, 1, newRows.length, incomeHeaders.length).setValues(newRows);
  }

  // Create Ledger row for new year (Day1 blank — admin sets via Event Setup)
  var ledgerNewRow = [];
  for (var k = 0; k < ledgerHeaders.length; k++) {
    ledgerNewRow.push(ledgerHeaders[k] === "Year" ? newYear : "");
  }
  ledgerSh.appendRow(ledgerNewRow);

  writeAudit("CREATE_FY", "Ledger",
    "FY " + newYear + " created from " + prevYear +
    " · Families: " + counts.families +
    " · PulliVari: " + counts.pulli +
    " · PenPulli: " + counts.penPulli +
    " · CashRec: " + counts.cashRec +
    " · IntRec: " + counts.intRec, email);

  return jsonResponse({ success: true, data: { newYear: newYear, prevYear: prevYear, counts: counts } });
}

// ============================================================
// CONFIRM INCOME (flip Expected → Confirmed)
// ============================================================
// Used to mark a pre-populated receivable row as actually
// collected. Sets Status=Confirmed and updates Date/Day to today
// (with Day computed relative to the row's year Day1 from Ledger).
// ============================================================

function handleConfirmIncome(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  if (!sheet) throw new Error("Income sheet not found");

  var txnId = String(data.txnId || "").trim();
  if (!txnId) throw new Error("txnId required");

  // Ensure Status column exists
  ensureColumn(sheet, "Status");

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var txnIdx = headers.indexOf("TxnID");
  var statusIdx = headers.indexOf("Status");
  var dateIdx = headers.indexOf("Date");
  var dayIdx = headers.indexOf("Day");
  var yearIdx = headers.indexOf("Year");
  if (txnIdx === -1) throw new Error("Income sheet missing TxnID column");

  var todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var found = -1;
  var rowYear = "";
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][txnIdx]) === txnId) {
      found = i + 1; // 1-indexed
      rowYear = String(rows[i][yearIdx]);
      break;
    }
  }
  if (found === -1) throw new Error("Income row not found: " + txnId);

  // Set Status=Confirmed
  sheet.getRange(found, statusIdx + 1).setValue("Confirmed");
  // Update Date if blank
  if (dateIdx >= 0) sheet.getRange(found, dateIdx + 1).setNumberFormat("@").setValue(todayDate);
  // Compute and update Day (relative to Ledger Day1 for this year)
  if (dayIdx >= 0 && rowYear) {
    var day = computeDayFromLedger(ss, rowYear, todayDate);
    if (day) sheet.getRange(found, dayIdx + 1).setValue(day);
  }

  writeAudit("CONFIRM", "Income", "Confirmed " + txnId + " on " + todayDate, email);
  return jsonResponse({ success: true, data: { txnId: txnId, date: todayDate } });
}

// ============================================================
// TIME-BASED INTEREST HELPERS
// ============================================================
// Interest = Principal × 9% × (months ÷ 12), where months are counted
// from February of the financial year (month 0) up to and including the
// payment month, capped at 12 (a full year = full 9%).
// Example: event Feb, paid May → Mar,Apr,May = 3 months → 2.25%.
// ============================================================

// Months of accrual up to & including the payment month, 0..12.
// The interest clock for FYn receivables starts at the PREVIOUS event —
// February of (fyYear − 1) — because the loan was given at that event and
// is repaid over the following year. So a payment in Oct for FY2027
// (loan from the Feb 2026 event) = 8 months (Mar..Oct 2026).
function monthsSinceEvent(fyYear, paymentDateStr) {
  if (!paymentDateStr) return 0;
  var pay = (paymentDateStr instanceof Date) ? paymentDateStr : new Date(paymentDateStr);
  if (isNaN(pay.getTime())) return 0;
  var fy = Number(fyYear);
  if (!fy) return 0;
  var anchorYear = fy - 1; // February of the previous year = loan origination
  var months = (pay.getFullYear() - anchorYear) * 12 + (pay.getMonth() - 1);
  if (months < 0) months = 0;
  if (months > 12) months = 12;
  return months;
}

// Normalize any date value to a plain "yyyy-MM-dd" string so Google Sheets
// stores it as text, not an auto-converted Date object (which renders as
// "Sat Oct 17 2026 12:30:00 GMT+0530..."). Accepts Date objects, ISO
// strings, or dd/MM/yyyy.
function toDateStr(v) {
  if (!v) return "";
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;                 // already ISO
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.split("T")[0];   // ISO datetime
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {                        // dd/MM/yyyy
    var p = s.split("/"); return p[2] + "-" + p[1] + "-" + p[0];
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return s;
}

// Interest on a principal amount for a given number of months at a rate.
function calcTimeInterest(principal, annualRatePct, months) {
  var p = Number(principal) || 0;
  var r = Number(annualRatePct);
  if (isNaN(r)) r = 9;
  var m = Number(months) || 0;
  return Math.round(p * (r / 100) * (m / 12));
}

// ============================================================
// INTEREST = REBATE MODEL
// ============================================================
// Each family owes a full year's interest on the full principal:
//   fullInterest = fullPrincipal × rate
// Paying principal early earns a rebate on the UNEXPIRED months:
//   rebate(payment) = payment × rate × (12 − monthsElapsed) / 12
// Interest owed = fullInterest − Σ rebate(payment), floored at 0.
//
// monthsElapsed = months from the previous February (loan origination) to
// the payment month, 0..12 (same as monthsSinceEvent). Paying AT the next
// event (12 months elapsed) gives 0 rebate = full interest.
// Example: 12L owed, pay 3L in Oct (8 months elapsed, 4 remaining):
//   full = 1,08,000; rebate = 3L × 9% × 4/12 = 9,000; owed = 99,000.
function interestOwed(fullPrincipal, installments, fyYear, annualRatePct) {
  var r = Number(annualRatePct); if (isNaN(r)) r = 9;
  var fullInterest = (Number(fullPrincipal) || 0) * (r / 100);
  var rebate = 0;
  (installments || []).forEach(function(ins) {
    var elapsed = monthsSinceEvent(fyYear, ins.date);
    var remaining = 12 - elapsed;
    rebate += (Number(ins.amount) || 0) * (r / 100) * (remaining / 12);
  });
  return Math.max(0, Math.round(fullInterest - rebate));
}

// ============================================================
// COLLECT RECEIVABLE (partial principal payment + time-based interest)
// ============================================================
// Records one principal installment against a family's Cash Receivable
// for a year:
//   1. Writes a new Cash Receivable row (Status="Collected") for the
//      amount paid, stamped with the payment Date.
//   2. Computes interest on THIS installment = amount × 9% × months/12
//      (months from Feb to the payment month), and ADDS it to the
//      family's single Interest Receivable row for the year (its Amount
//      accumulates; Status flips to "Collected" once any interest lands).
//   3. If the family's Cash Receivable is now fully collected
//      (sum of Collected ≥ the Expected obligation), flips the Expected
//      Cash Receivable row to "Confirmed".
//
// data: { familyId, year, amount, date, rate? }
// ============================================================

function handleCollectReceivable(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  if (!sheet) throw new Error("Income sheet not found");

  var familyId = String(data.familyId || "").trim();
  var year = String(data.year || "").trim();
  var amount = Number(data.amount) || 0;
  var payDate = toDateStr(String(data.date || "").trim());
  var rate = (data.rate === undefined || data.rate === null || data.rate === "") ? 9 : Number(data.rate);
  if (!familyId) throw new Error("familyId required");
  if (!year) throw new Error("year required");
  if (amount <= 0) throw new Error("amount must be positive");
  if (!payDate) throw new Error("date required");

  ensureColumn(sheet, "Status");
  SpreadsheetApp.flush();

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var txnIdx = headers.indexOf("TxnID");
  var yearIdx = headers.indexOf("Year");
  var fidIdx = headers.indexOf("FamilyID");
  var nameIdx = headers.indexOf("DonorName");
  var catIdx = headers.indexOf("Category");
  var amtIdx = headers.indexOf("Amount");
  var statusIdx = headers.indexOf("Status");
  var dateIdx = headers.indexOf("Date");
  var dayIdx = headers.indexOf("Day");

  // Scan this family/year's Cash Receivable + Interest Receivable rows.
  // Collect the existing collected installments so we can RECOMPUTE the
  // interest total from scratch (never trust a stale rollover figure).
  var expectedCashRowNo = -1, expectedCashAmt = 0;
  var collectedCashSum = 0;
  var collectedInstallments = []; // {amount, date}
  var interestRowNo = -1;
  var donorName = "";
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][yearIdx]) !== year) continue;
    if (String(rows[i][fidIdx] || "").trim() !== familyId) continue;
    var cat = String(rows[i][catIdx] || "").trim().toLowerCase();
    var st = String(rows[i][statusIdx] || "").trim();
    if (!donorName) donorName = String(rows[i][nameIdx] || "").trim();
    if (cat === "cash receivable") {
      if (st === "Collected") {
        var instAmt = Number(rows[i][amtIdx]) || 0;
        collectedCashSum += instAmt;
        collectedInstallments.push({ amount: instAmt, date: rows[i][dateIdx] });
      } else {
        // Expected (or legacy blank) = the obligation anchor
        expectedCashRowNo = i + 1;
        expectedCashAmt = Number(rows[i][amtIdx]) || 0;
      }
    } else if (cat === "interest receivable") {
      interestRowNo = i + 1;
    }
  }

  if (expectedCashRowNo === -1 && collectedCashSum === 0) {
    throw new Error("No Cash Receivable found for " + familyId + " in " + year);
  }

  var computedDay = computeDayFromLedger(ss, year, payDate) || "";
  var months = monthsSinceEvent(year, payDate);

  // 1. Write the collected principal row. Force the Date cell to plain-text
  //    format BEFORE writing so Sheets doesn't convert it to a Date object.
  var collectedRow = buildIncomeRow(headers, {
    TxnID: generateId("DON"), Year: year, Day: computedDay, Date: payDate,
    DonorType: "Family", FamilyID: familyId, DonorName: donorName,
    Category: "Cash Receivable", Amount: amount,
    Description: "Collected " + payDate + " (" + months + "mo elapsed)",
    EnteredBy: email, Timestamp: now(), Status: "Collected",
  });
  sheet.appendRow(collectedRow);
  var newRowNo = sheet.getLastRow();
  if (dateIdx >= 0) {
    sheet.getRange(newRowNo, dateIdx + 1).setNumberFormat("@").setValue(payDate);
  }

  // 2. REBATE-MODEL interest: full-year interest on the full principal owed,
  //    minus the rebate on every early principal payment (including this one).
  var allInstallments = collectedInstallments.concat([{ amount: amount, date: payDate }]);
  var newInterestTotal = interestOwed(expectedCashAmt, allInstallments, year, rate);
  if (interestRowNo !== -1) {
    sheet.getRange(interestRowNo, amtIdx + 1).setValue(newInterestTotal);
    sheet.getRange(interestRowNo, statusIdx + 1).setValue("Collected");
    sheet.getRange(interestRowNo, dateIdx + 1).setNumberFormat("@").setValue(payDate);
    sheet.getRange(interestRowNo, headers.indexOf("Description") + 1).setValue(
      "Full " + Math.round(expectedCashAmt * rate / 100) + " − early-pay rebate"
    );
  } else {
    var intRow = buildIncomeRow(headers, {
      TxnID: generateId("DON"), Year: year, Day: computedDay, Date: payDate,
      DonorType: "Family", FamilyID: familyId, DonorName: donorName,
      Category: "Interest Receivable", Amount: newInterestTotal,
      Description: "Full " + Math.round(expectedCashAmt * rate / 100) + " − early-pay rebate",
      EnteredBy: email, Timestamp: now(), Status: "Collected",
    });
    sheet.appendRow(intRow);
    if (dateIdx >= 0) {
      sheet.getRange(sheet.getLastRow(), dateIdx + 1).setNumberFormat("@").setValue(payDate);
    }
  }

  // 3. If principal fully collected, flip the Expected row to Confirmed.
  var totalCollected = collectedCashSum + amount;
  var fullyPaid = (expectedCashAmt > 0 && totalCollected >= expectedCashAmt - 0.5);
  if (fullyPaid && expectedCashRowNo !== -1) {
    sheet.getRange(expectedCashRowNo, statusIdx + 1).setValue("Confirmed");
  }

  writeAudit("COLLECT", "Income",
    familyId + " " + year + " collected ₹" + amount + " on " + payDate +
    " (" + months + "mo elapsed · interest now ₹" + newInterestTotal + ")" +
    (fullyPaid ? " · principal fully paid" : ""), email);

  return jsonResponse({ success: true, data: {
    familyId: familyId, year: year, collected: amount, date: payDate,
    months: months, interestTotal: newInterestTotal,
    totalCollected: totalCollected,
    remaining: Math.max(0, expectedCashAmt - totalCollected),
    fullyPaid: fullyPaid,
  }});
}

// ============================================================
// RECALCULATE INTEREST (fix stale rollover values)
// ============================================================
// For a given year, recompute every family's single Interest Receivable
// row as the sum over their collected Cash Receivable installments of
// amount × 9% × months/12. Overwrites any stale value (e.g. a fixed
// figure written by an older rollover). Interest rows with no collections
// are reset to 0.
// ============================================================
function handleRecalcInterest(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  if (!sheet) throw new Error("Income sheet not found");
  var year = String(data.year || "").trim();
  if (!year) throw new Error("year required");

  ensureColumn(sheet, "Status");
  SpreadsheetApp.flush();

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var yearIdx = headers.indexOf("Year");
  var fidIdx = headers.indexOf("FamilyID");
  var catIdx = headers.indexOf("Category");
  var amtIdx = headers.indexOf("Amount");
  var statusIdx = headers.indexOf("Status");
  var dateIdx = headers.indexOf("Date");
  var descIdx = headers.indexOf("Description");

  // Per family: full principal owed (Expected Cash Receivable) + collected
  // installments + the interest row location. Then apply the rebate model.
  var fullPrincipalByFamily = {}; // fid -> Expected Cash Receivable amount
  var installmentsByFamily = {};  // fid -> [{amount,date}]
  var interestRowNoByFamily = {}; // fid -> 1-indexed interest row
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][yearIdx]) !== year) continue;
    var fid = String(rows[i][fidIdx] || "").trim();
    if (!fid) continue;
    var cat = String(rows[i][catIdx] || "").trim().toLowerCase();
    var st = String(rows[i][statusIdx] || "").trim();
    if (cat === "cash receivable") {
      if (st === "Collected") {
        if (!installmentsByFamily[fid]) installmentsByFamily[fid] = [];
        installmentsByFamily[fid].push({ amount: Number(rows[i][amtIdx]) || 0, date: rows[i][dateIdx] });
        // Clean up any ugly full-timestamp date stored on this collected row.
        if (dateIdx >= 0) {
          var clean = toDateStr(rows[i][dateIdx]);
          if (clean && clean !== String(rows[i][dateIdx])) {
            sheet.getRange(i + 1, dateIdx + 1).setNumberFormat("@").setValue(clean);
          }
        }
      } else {
        // Expected / blank / Confirmed obligation anchor = full principal owed
        fullPrincipalByFamily[fid] = Number(rows[i][amtIdx]) || 0;
      }
    } else if (cat === "interest receivable") {
      interestRowNoByFamily[fid] = i + 1;
    }
  }

  var updated = 0;
  for (var fid2 in interestRowNoByFamily) {
    var rowNo = interestRowNoByFamily[fid2];
    var fullP = fullPrincipalByFamily[fid2] || 0;
    var insts = installmentsByFamily[fid2] || [];
    // Rebate model: full-year interest on full principal − early-pay rebate.
    var computed = interestOwed(fullP, insts, year, 9);
    // If nothing collected yet, the row sits as the full obligation but
    // "Expected" (not yet realized income).
    var anyCollected = insts.length > 0;
    sheet.getRange(rowNo, amtIdx + 1).setValue(computed);
    sheet.getRange(rowNo, statusIdx + 1).setValue(anyCollected ? "Collected" : "Expected");
    if (descIdx >= 0) {
      sheet.getRange(rowNo, descIdx + 1).setValue(
        anyCollected ? ("Full " + Math.round(fullP * 0.09) + " − early-pay rebate")
                     : "Accrues as principal is collected"
      );
    }
    updated++;
  }

  writeAudit("RECALC_INT", "Income", "Recalculated interest for " + year + " · " + updated + " families", email);
  return jsonResponse({ success: true, data: { year: year, updated: updated } });
}

// Helpers used by FY rollover and confirmIncome
function ensureColumn(sheet, headerName) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(headerName) === -1) {
    sheet.getRange(1, lastCol + 1).setValue(headerName);
  }
}

function getHeaderIndex(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName);
}

function readSheetRows(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    out.push(obj);
  }
  return out;
}

function buildIncomeRow(headers, dataObj) {
  return headers.map(function(h) {
    return dataObj.hasOwnProperty(h) ? dataObj[h] : "";
  });
}

function computeDayFromLedger(ss, yearStr, dateStr) {
  var lsh = ss.getSheetByName(SHEETS.LEDGER);
  if (!lsh) return "";
  var data = lsh.getDataRange().getValues();
  var headers = data[0];
  var yIdx = headers.indexOf("Year");
  var d1Idx = headers.indexOf("Day1");
  if (yIdx === -1 || d1Idx === -1) return "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][yIdx]) === String(yearStr)) {
      var day1 = data[i][d1Idx];
      if (!day1) return "";
      var day1Date = (day1 instanceof Date) ? day1 : new Date(day1);
      var inputDate = new Date(dateStr);
      var diffDays = Math.round((inputDate - day1Date) / (1000 * 60 * 60 * 24));
      if (diffDays >= 0 && diffDays < 3) return String(diffDays + 1);
      return "";
    }
  }
  return "";
}

// ============================================================
// SEED STANDARD AUCTION ITEMS (48 items per year)
// ============================================================

function handleSeedAuctionItems(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUCTION);
  if (!sheet) throw new Error("Auction sheet not found");

  var year = String(data.year || new Date().getFullYear());

  // Check if items already exist for this year
  var existing = sheet.getDataRange().getValues();
  var headers = existing[0];
  var yearCol = headers.indexOf("Year");
  var count = 0;
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][yearCol]) === year) count++;
  }
  if (count > 0) {
    return jsonResponse({ success: true, data: { count: 0, skipped: true, existing: count, year: year } });
  }

  // Standard 48 items
  var items = [
    { name: "\u0BAE\u0B9E\u0BCD\u0B9A\u0BB3\u0BCD", qty: 5 },
    { name: "\u0B89\u0BAA\u0BCD\u0BAA\u0BC1", qty: 5 },
    { name: "\u0B95\u0BB1\u0BCD\u0B95\u0BA3\u0BCD\u0B9F\u0BC1", qty: 5 },
    { name: "\u0BA4\u0BC7\u0B99\u0BCD\u0B95\u0BBE\u0BAF\u0BCD", qty: 5 },
    { name: "\u0BAA\u0BBE\u0B95\u0BCD\u0B95\u0BC1", qty: 5 },
    { name: "\u0BB5\u0BC6\u0BB2\u0BCD\u0BB2\u0BAE\u0BCD", qty: 5 },
    { name: "\u0B9A\u0BB0\u0BCD\u0B95\u0BCD\u0B95\u0BB0\u0BC8", qty: 5 },
    { name: "\u0B9C\u0BC0\u0BA9\u0BBF", qty: 5 },
    { name: "\u0B8E\u0BB2\u0BC1\u0BAE\u0BBF\u0B9A\u0BCD\u0B9A\u0BAE\u0BCD \u0BAA\u0BB4\u0BAE\u0BCD", qty: 5 },
    { name: "\u0BB5\u0BBE\u0BB4\u0BC8\u0BAA\u0BCD\u0BAA\u0BB4\u0BAE\u0BCD", qty: 3 }
  ];

  var startItemNo = sheet.getLastRow(); // header = row 1
  var itemNo = startItemNo;
  var rows = [];
  for (var j = 0; j < items.length; j++) {
    for (var k = 0; k < items[j].qty; k++) {
      // Row: ItemNo, Year, ItemDescription, DonatedBy, WinnerFamilyID, WinnerName, BidAmount, EnteredBy1, Timestamp1, ConfirmedBy2, Timestamp2, Status
      rows.push([itemNo, year, items[j].name, "", "", "", "", "", "", "", "", ""]);
      itemNo++;
    }
  }

  // Bulk append
  if (rows.length > 0) {
    sheet.getRange(startItemNo + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  writeAudit("CREATE", "Auction", "Seeded " + rows.length + " standard items for " + year, email);
  return jsonResponse({ success: true, data: { count: rows.length, year: year } });
}

// ============================================================
// TEST (run manually to verify setup)
// ============================================================

function testSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var missing = [];
  var sheetNames = Object.keys(SHEETS);
  for (var i = 0; i < sheetNames.length; i++) {
    var name = SHEETS[sheetNames[i]];
    if (!ss.getSheetByName(name)) {
      missing.push(name);
    }
  }
  if (missing.length > 0) {
    Logger.log("⚠️ Missing sheets: " + missing.join(", "));
  } else {
    Logger.log("✅ All 7 sheets found!");
  }

  // Test access control
  var accessSheet = ss.getSheetByName(SHEETS.ACCESS);
  var rows = accessSheet.getDataRange().getValues();
  Logger.log("AccessControl has " + (rows.length - 1) + " users:");
  for (var j = 1; j < rows.length; j++) {
    Logger.log("  " + rows[j][0] + " → " + rows[j][1]);
  }
}

// ============================================================
// LOANS (Pangali Advances / ஈனத்து கணக்கு)
// ============================================================
// Sheet columns: LoanID | Year | FamilyID | FamilyName | LoanAmount | InterestRate | InterestAmount | TotalReceivable | ReceivedAmount | RecStatus | IssuedBy | IssuedDate | ReceivedBy | ReceivedDate | Notes

function handleAddLoan(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LOANS);
  var loanId = generateId("LN");
  var ts = now();
  var year = data.year || String(new Date().getFullYear());
  var loanAmt = Number(data.loanAmount) || 0;
  var rate = Number(data.interestRate) || 9;
  var interestOnLoan = Math.round(loanAmt * rate / 100);

  // Use manually entered lastYearReceivable if provided, else auto-calculate
  // from previous year's principal carry-forward (LastYearReceivable + LoanAmount).
  // Interest is tracked separately as Interest Receivable income line.
  var lastYearRec = 0;
  if (data.lastYearReceivable && Number(data.lastYearReceivable) > 0) {
    lastYearRec = Number(data.lastYearReceivable);
  } else {
    var prevYear = String(Number(year) - 1);
    var allRows = sheet.getDataRange().getValues();
    var headers = allRows[0];
    var yearCol = headers.indexOf("Year");
    var fidCol = headers.indexOf("FamilyID");
    var lyrCol = headers.indexOf("LastYearReceivable");
    var loanAmtCol = headers.indexOf("LoanAmount");
    for (var i = 1; i < allRows.length; i++) {
      if (String(allRows[i][yearCol]) === prevYear && String(allRows[i][fidCol]).trim() === String(data.familyId).trim()) {
        var principal = (Number(allRows[i][lyrCol]) || 0) + (Number(allRows[i][loanAmtCol]) || 0);
        if (principal > 0) lastYearRec += principal;
      }
    }
  }

  var interestOnPrev = Math.round(lastYearRec * rate / 100);
  var interestAmt = interestOnLoan + interestOnPrev;
  var totalRec = lastYearRec + loanAmt + interestAmt;

  // ReceivedAmount, RecStatus, ReceivedBy, ReceivedDate columns are no longer
  // maintained — written as empty so legacy sheets keep their column positions
  // but existing data isn't touched.
  var row = [
    loanId,
    year,
    data.familyId || "",
    data.familyName || "",
    lastYearRec,
    loanAmt,
    rate,
    interestAmt,
    totalRec,
    "",
    "",
    email,
    data.issuedDate || ts,
    "",
    "",
    data.notes || ""
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Loan", loanId + " ₹" + loanAmt + " to " + data.familyName + " @" + rate + "% (prev:" + lastYearRec + ")", email);
  return jsonResponse({ success: true, data: { loanId: loanId, lastYearReceivable: lastYearRec } });
}

function handleEditLoan(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LOANS);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var idCol = headers.indexOf("LoanID");

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(data.loanId).trim()) {
      var loanAmt = (data.loanAmount !== undefined && data.loanAmount !== "") ? Number(data.loanAmount) : Number(rows[i][headers.indexOf("LoanAmount")]) || 0;
      var rate = (data.interestRate !== undefined && data.interestRate !== "") ? Number(data.interestRate) : Number(rows[i][headers.indexOf("InterestRate")]) || 9;
      var lastYearRec = (data.lastYearReceivable !== undefined && data.lastYearReceivable !== "") ? Number(data.lastYearReceivable) : Number(rows[i][headers.indexOf("LastYearReceivable")]) || 0;
      var interestOnLoan = Math.round(loanAmt * rate / 100);
      var interestOnPrev = Math.round(lastYearRec * rate / 100);
      var interestAmt = interestOnLoan + interestOnPrev;
      var totalRec = lastYearRec + loanAmt + interestAmt;

      sheet.getRange(i + 1, headers.indexOf("FamilyID") + 1).setValue(data.familyId || rows[i][headers.indexOf("FamilyID")]);
      sheet.getRange(i + 1, headers.indexOf("FamilyName") + 1).setValue(data.familyName || rows[i][headers.indexOf("FamilyName")]);
      sheet.getRange(i + 1, headers.indexOf("LastYearReceivable") + 1).setValue(lastYearRec);
      sheet.getRange(i + 1, headers.indexOf("LoanAmount") + 1).setValue(loanAmt);
      sheet.getRange(i + 1, headers.indexOf("InterestRate") + 1).setValue(rate);
      sheet.getRange(i + 1, headers.indexOf("InterestAmount") + 1).setValue(interestAmt);
      sheet.getRange(i + 1, headers.indexOf("TotalReceivable") + 1).setValue(totalRec);
      sheet.getRange(i + 1, headers.indexOf("Notes") + 1).setValue(data.notes !== undefined ? data.notes : rows[i][headers.indexOf("Notes")]);

      writeAudit("UPDATE", "Loan", data.loanId + " updated ₹" + loanAmt + " @" + rate + "%", email);
      return jsonResponse({ success: true, data: { loanId: data.loanId } });
    }
  }
  return jsonResponse({ success: false, error: "Loan not found: " + data.loanId });
}

function handleDeleteLoan(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.LOANS);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var idCol = headers.indexOf("LoanID");

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(data.loanId).trim()) {
      sheet.deleteRow(i + 1);
      writeAudit("DELETE", "Loan", data.loanId + " deleted", email);
      return jsonResponse({ success: true, data: { loanId: data.loanId } });
    }
  }
  return jsonResponse({ success: false, error: "Loan not found: " + data.loanId });
}
