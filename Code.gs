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
  FAMILIES: "Families",
  INCOME: "Income",
  EXPENSES: "Expenses",
  AUCTION: "Auction",
  LEDGER: "Ledger",
  ACCESS: "AccessControl",
  AUDIT: "AuditLog"
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
    var userEmail = (params.userEmail || "").toLowerCase().trim();

    // Validate user
    var userRole = getUserRole(userEmail);

    // Public actions (no role needed)
    if (action === "checkAccess") {
      return jsonResponse({ success: true, role: userRole, email: userEmail });
    }

    // All other actions need at least Viewer role
    if (!userRole) {
      return jsonResponse({ success: false, error: "Access denied. Email not in AccessControl." });
    }

    // Route actions
    switch (action) {
      // --- READ (Viewer+) ---
      case "loadAll":
        return handleLoadAll(userEmail, userRole);

      // --- WRITE (Editor+) ---
      case "addIncome":
        return requireRole(userRole, "Editor", function() { return handleAddIncome(params.data, userEmail); });
      case "addExpense":
        return requireRole(userRole, "Editor", function() { return handleAddExpense(params.data, userEmail); });
      case "addFamily":
        return requireRole(userRole, "Editor", function() { return handleAddFamily(params.data, userEmail); });
      case "bulkAddFamilies":
        return requireRole(userRole, "Admin", function() { return handleBulkAddFamilies(params.data, userEmail); });
      case "addAuctionItem":
        return requireRole(userRole, "Admin", function() { return handleAddAuctionItem(params.data, userEmail); });
      case "enterBid":
        return requireRole(userRole, "Editor", function() { return handleEnterBid(params.data, userEmail); });
      case "uploadAuctionPhoto":
        return requireRole(userRole, "Editor", function() { return handleUploadAuctionPhoto(params.data, userEmail); });

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
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email) {
      return data[i][1]; // Role column (B)
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

function handleLoadAll(userEmail, userRole) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var result = {
    families: readSheet(ss, SHEETS.FAMILIES),
    income: readSheet(ss, SHEETS.INCOME),
    expenses: readSheet(ss, SHEETS.EXPENSES),
    auction: readSheet(ss, SHEETS.AUCTION),
    access: readSheet(ss, SHEETS.ACCESS),
    ledger: readSheet(ss, SHEETS.LEDGER),
    role: userRole,
  };

  // Audit log only for Admin
  if (userRole === "Admin") {
    result.auditLog = readSheet(ss, SHEETS.AUDIT);
  }

  writeAudit("LOGIN", "Auth", userEmail + " loaded data", userEmail);

  return jsonResponse({ success: true, data: result });
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
// INCOME (Donations)
// ============================================================

function handleAddIncome(data, userEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.INCOME);
  var txnId = generateId("DON");
  var ts = now();

  var year = data.year || String(new Date().getFullYear());

  var row = [
    txnId,
    year,
    data.day,
    data.donorType,
    data.familyId || "",
    data.donorName,
    data.category,
    data.amount || 0,
    data.description || "",
    userEmail,
    ts
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Income", txnId + " ₹" + data.amount + " " + data.category + " by " + data.donorName, userEmail);

  var result = { txnId: txnId };

  // Auto-add to auction if "Objects for Auction"
  if (data.category === "Objects for Auction" && data.description) {
    var auctionSheet = ss.getSheetByName(SHEETS.AUCTION);
    var nextItemNo = auctionSheet.getLastRow(); // includes header, so this = item count + 1
    var auctionRow = [
      nextItemNo,
      year,
      data.description,
      data.donorName,
      "", "", "", "", "", "", "", ""
    ];
    auctionSheet.appendRow(auctionRow);
    writeAudit("CREATE", "Auction", "Auto-added \"" + data.description + "\" from income by " + data.donorName, userEmail);
    result.auctionItemAdded = true;
    result.auctionItemNo = nextItemNo;
  }

  return jsonResponse({ success: true, data: result });
}

// ============================================================
// EXPENSES
// ============================================================

function handleAddExpense(data, userEmail) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EXPENSES);
  var txnId = generateId("EXP");
  var ts = now();

  var year = data.year || String(new Date().getFullYear());

  var row = [
    txnId,
    year,
    data.day,
    data.category,
    data.description,
    data.amount,
    data.vendor || "",
    userEmail,
    ts
  ];
  sheet.appendRow(row);
  writeAudit("CREATE", "Expenses", txnId + " ₹" + data.amount + " " + data.category + " - " + data.description, userEmail);

  return jsonResponse({ success: true, data: { txnId: txnId } });
}

// ============================================================
// FAMILIES
// ============================================================

function handleAddFamily(data, userEmail) {
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
  writeAudit("CREATE", "Families", famId + " " + data.familyName, userEmail);

  return jsonResponse({ success: true, data: { familyId: famId } });
}

function handleBulkAddFamilies(data, userEmail) {
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
  writeAudit("BULK_CREATE", "Families", count + " families added", userEmail);
  return jsonResponse({ success: true, data: { count: count } });
}

// ============================================================
// AUCTION — Add Item (Admin only)
// ============================================================

function handleAddAuctionItem(data, userEmail) {
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
  writeAudit("CREATE", "Auction", "Manual add #" + nextItemNo + " \"" + data.itemName + "\"", userEmail);

  return jsonResponse({ success: true, data: { itemNo: nextItemNo } });
}

// ============================================================
// AUCTION — Enter/Confirm Bid
// ============================================================

function handleEnterBid(data, userEmail) {
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
    values[7] = userEmail;               // EnteredBy1
    values[9] = ts;                      // Timestamp1
    values[11] = "Pending";              // Status

    range.setValues([values]);
    writeAudit("BID", "Auction", "#" + values[0] + " \"" + values[2] + "\" ₹" + data.amount + " by " + data.winnerName, userEmail);

    return jsonResponse({ success: true, data: { status: "Pending" } });

  } else if (enteredBy1.toLowerCase() !== userEmail.toLowerCase() && !confirmedBy2) {
    // Second person — confirm
    values[8] = userEmail;               // ConfirmedBy2
    values[10] = ts;                     // Timestamp2
    values[11] = "Confirmed";            // Status

    range.setValues([values]);
    writeAudit("CONFIRM", "Auction", "#" + values[0] + " \"" + values[2] + "\" ₹" + values[6] + " to " + values[5], userEmail);

    return jsonResponse({ success: true, data: { status: "Confirmed" } });

  } else if (enteredBy1.toLowerCase() === userEmail.toLowerCase()) {
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

function handleUploadAuctionPhoto(data, userEmail) {
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
      writeAudit("UPDATE", "Auction", "Photo uploaded for item #" + data.itemNo, userEmail);
      return jsonResponse({ success: true, data: { itemNo: data.itemNo } });
    }
  }
  
  return jsonResponse({ success: false, error: "Item #" + data.itemNo + " not found" });
}

function writeAudit(action, module, details, userEmail) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUDIT);
    if (sheet) {
      sheet.appendRow([now(), userEmail, action, module, details]);
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
