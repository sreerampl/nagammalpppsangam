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
  LOANS: "Loans"
};

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

      case "seedAuctionItems":
        return requireRole(userRole, "Admin", function() { return handleSeedAuctionItems(params.data, email); });

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
//   - Pen Pulli Varavu (Amount = penPulliDefault, Status = Expected)
// And for each family with an outstanding loan from previous year:
//   - Cash Receivable (Amount = LastYearReceivable + LoanAmount)
//   - Interest Receivable (Amount = InterestAmount)
//
// All rolled-over rows are marked Status="Expected" so they do NOT
// inflate income totals until an admin confirms collection via the
// confirmIncome action.
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

  // Ensure Status column exists on Income sheet
  ensureColumn(incomeSh, "Status");

  // Default amounts (passed from frontend; falls back to 51)
  var pulliDefault = Number(data.pulliDefault) || 51;
  var penPulliDefault = Number(data.penPulliDefault) || 51;

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

    // Pen Pulli Varavu — Expected
    newRows.push(buildIncomeRow(incomeHeaders, {
      TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
      DonorType: "Family", FamilyID: fid, DonorName: fname,
      Category: "Pen Pulli Varavu", Amount: penPulliDefault, Description: "",
      EnteredBy: email, Timestamp: ts, Status: "Expected",
    }));
    counts.penPulli++;

    // Cash Receivable + Interest Receivable from prev year loan
    var prevLoan = prevLoansByFamily[fid];
    if (prevLoan) {
      var lastYrRec = Number(prevLoan.LastYearReceivable) || 0;
      var loanAmt = Number(prevLoan.LoanAmount) || 0;
      var intAmt = Number(prevLoan.InterestAmount) || 0;
      var cashRecAmt = lastYrRec + loanAmt;

      if (cashRecAmt > 0) {
        newRows.push(buildIncomeRow(incomeHeaders, {
          TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
          DonorType: "Family", FamilyID: fid, DonorName: fname,
          Category: "Cash Receivable", Amount: cashRecAmt, Description: "From " + prevYear + " loan",
          EnteredBy: email, Timestamp: ts, Status: "Expected",
        }));
        counts.cashRec++;
      }
      if (intAmt > 0) {
        newRows.push(buildIncomeRow(incomeHeaders, {
          TxnID: generateId("DON"), Year: String(newYear), Day: "", Date: "",
          DonorType: "Family", FamilyID: fid, DonorName: fname,
          Category: "Interest Receivable", Amount: intAmt, Description: "From " + prevYear + " loan",
          EnteredBy: email, Timestamp: ts, Status: "Expected",
        }));
        counts.intRec++;
      }
    }
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
  if (dateIdx >= 0) sheet.getRange(found, dateIdx + 1).setValue(todayDate);
  // Compute and update Day (relative to Ledger Day1 for this year)
  if (dayIdx >= 0 && rowYear) {
    var day = computeDayFromLedger(ss, rowYear, todayDate);
    if (day) sheet.getRange(found, dayIdx + 1).setValue(day);
  }

  writeAudit("CONFIRM", "Income", "Confirmed " + txnId + " on " + todayDate, email);
  return jsonResponse({ success: true, data: { txnId: txnId, date: todayDate } });
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
