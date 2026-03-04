// ============================================================
//  ORMOC PRINTSHOPPE CRM — WEB APP
//  Google Apps Script Backend (Code.gs)
//  Step 1: Lead Submission Form
// ============================================================

const QUOTATION_FOLDER_ID = "1UAYIKkcc8yb1PQQGvmKhyJZIRGCOoYae";
const PHOTO_FOLDER_ID     = "1UXWFgkjtgmaW6VRZNJF4heTaCKmcF2MY";
const DB_SHEET            = "📋 Daily Field Log";

// ── ENTRY POINT ──────────────────────────────────────────────────────
// This function runs when someone opens the web app URL.
// It serves the Lead Form HTML page.
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'lead';

  if (page === 'search') {
    return HtmlService.createHtmlOutputFromFile('SearchPage')
      .setTitle('Search Leads — Ormoc Printshoppe CRM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('DashboardPage')
      .setTitle('Dashboard — Ormoc Printshoppe CRM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default: lead form
  return HtmlService.createHtmlOutputFromFile('LeadForm')
    .setTitle('New Lead — Ormoc Printshoppe CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── NAVIGATE (called from nav links in the HTML) ─────────────────────
function navigateTo(page) {
  // Navigation is handled client-side via URL params
  // This is a placeholder for future use
}

// ── SUBMIT LEAD FROM WEB FORM ─────────────────────────────────────────
// Called by the Lead Form HTML via google.script.run.submitLeadFromWeb(data)
function submitLeadFromWeb(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const db = ss.getSheetByName(DB_SHEET);

    if (!db) throw new Error('Sheet "' + DB_SHEET + '" not found.');

    const fmtDate = (val) => {
      if (!val) return "";
      try {
        return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch(e) { return String(val); }
    };

    const todayStr    = fmtDate(new Date());
    const visitDate   = fmtDate(data.date);
    const isQuotSent  = String(data.quotationSent).toLowerCase() === 'yes';

    // Auto-generate IDs
    const leadId       = generateLeadId(db);
    const followUpCount = getFollowUpCount(db, data.businessName);
    const quotationNo  = isQuotSent ? generateQuotationNo(db) : "";
    const quotationDate = isQuotSent ? todayStr : "";

    // Build the 28-column row (A through AB)
    // Matches exactly the Daily Field Log column structure
    const row = [
      leadId,                                          // A  (1)  Lead ID
      todayStr,                                        // B  (2)  Date Added
      visitDate,                                       // C  (3)  Date Visited
      data.area        || "",                          // D  (4)  Area
      data.municipality || "",                         // E  (5)  Municipality
      data.businessName || "",                         // F  (6)  Business Name
      data.industry    || "",                          // G  (7)  Industry Type
      data.contactPerson || "",                        // H  (8)  Contact Person
      data.position    || "",                          // I  (9)  Position
      data.mobile      || "",                          // J  (10) Mobile
      data.email       || "",                          // K  (11) Email
      data.decisionMaker || "",                        // L  (12) Decision Maker?
      data.currentSupplier || "",                      // M  (13) Current Supplier
      data.clientStatus || "",                         // N  (14) Client Status
      data.itemsNeeded || "",                          // O  (15) Items Needed
      data.estimatedValue ? Number(data.estimatedValue) : 0, // P  (16) Est. Value
      data.stage       || "",                          // Q  (17) Stage
      visitDate,                                       // R  (18) Last Contact Date
      fmtDate(data.nextContactDate),                   // S  (19) Next Action Date
      data.nextActionType || "",                       // T  (20) Next Action Type
      followUpCount,                                   // U  (21) Follow-up Count
      data.photoLink   || "",                          // V  (22) Visit Photo Link
      quotationNo,                                     // W  (23) Quotation No.
      quotationDate,                                   // X  (24) Quotation Date
      isQuotSent ? (Number(data.quotationAmount) || 0) : 0, // Y  (25) Quotation Amount
      isQuotSent ? (data.quotationFileLink || "") : "", // Z  (26) Quotation File Link
      isQuotSent ? (data.sentVia || "") : "",          // AA (27) Sent Via
      isQuotSent ? (data.sentProofLink || "") : "",    // AB (28) Sent Proof Link
    ];

    // Find next empty row (data starts at row 5)
    const allData = db.getRange(5, 1, Math.max(db.getMaxRows() - 4, 1), 1).getValues();
    let nextRow = 5;
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0].toString().trim() !== "") {
        nextRow = i + 6;
      }
    }

    // Write row
    db.getRange(nextRow, 1, 1, row.length).setValues([row]);

    // Format currency columns
    db.getRange(nextRow, 16).setNumberFormat('₱#,##0.00'); // P  Est. Value
    db.getRange(nextRow, 25).setNumberFormat('₱#,##0.00'); // Y  Quotation Amount

    // Also write the remarks into the correct column
    // Remarks are stored in column X only when no quotation — otherwise they go in a separate notes field
    // Based on your original sheet: Remarks are in column X (24) when quotation date is empty
    // Since your original script maps remarks to col X (Quotation Date position), we handle this:
    if (!isQuotSent && data.remarks) {
      // When no quotation, col X (24) holds remarks
      db.getRange(nextRow, 24).setValue(data.remarks);
    }

    return {
      success: true,
      leadId: leadId,
      message: 'Lead submitted successfully.',
      followUpCount: followUpCount,
      quotationNo: quotationNo
    };

  } catch(err) {
    Logger.log('submitLeadFromWeb error: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ── SAVE FILE TO DRIVE ────────────────────────────────────────────────
// Same as original script — called by HTML sidebar/upload
function saveFileToDrive(base64Data, mimeType, fileName, uploadType) {
  const folderId = uploadType === "photo" ? PHOTO_FOLDER_ID : QUOTATION_FOLDER_ID;
  const folder   = DriveApp.getFolderById(folderId);
  const decoded  = Utilities.base64Decode(base64Data);
  const blob     = Utilities.newBlob(decoded, mimeType, fileName);
  const file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { success: true, url: file.getUrl(), name: fileName };
}

// ── HELPERS ───────────────────────────────────────────────────────────
function generateLeadId(db) {
  const now    = new Date();
  const yr     = now.getFullYear();
  const mo     = String(now.getMonth() + 1).padStart(2, "0");
  const prefix = "LD-" + yr + mo + "-";
  const ids    = db.getRange(5, 1, Math.max(db.getMaxRows() - 4, 1), 1).getValues().flat()
    .map(v => v.toString())
    .filter(v => v.startsWith(prefix))
    .map(v => parseInt(v.replace(prefix, "")) || 0);
  const max = ids.length ? Math.max(...ids) : 0;
  return prefix + String(max + 1).padStart(4, "0");
}

function generateQuotationNo(db) {
  const now    = new Date();
  const yr     = now.getFullYear();
  const mo     = String(now.getMonth() + 1).padStart(2, "0");
  const prefix = "Q-" + yr + "-" + mo + "-";
  const nums   = db.getRange(5, 23, Math.max(db.getMaxRows() - 4, 1), 1).getValues().flat()
    .map(v => v.toString())
    .filter(v => v.startsWith(prefix))
    .map(v => parseInt(v.split("-").pop()) || 0);
  const max = nums.length ? Math.max(...nums) : 0;
  return prefix + String(max + 1).padStart(3, "0");
}

function getFollowUpCount(db, businessName) {
  const names = db.getRange(5, 6, Math.max(db.getMaxRows() - 4, 1), 1).getValues().flat();
  const count = names.filter(n =>
    n.toString().toLowerCase() === businessName.toString().toLowerCase()
  ).length;
  return count + 1;
}