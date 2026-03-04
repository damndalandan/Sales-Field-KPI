// ============================================================
//  CRM LEAD ENTRY FORM — Google Apps Script (v3 - Quotation Yes column added)
// ============================================================

const QUOTATION_FOLDER_ID = "1UAYIKkcc8yb1PQQGvmKhyJZIRGCOoYae";
const PHOTO_FOLDER_ID     = "1UXWFgkjtgmaW6VRZNJF4heTaCKmcF2MY";

// ── CHANGE THESE TO MATCH YOUR EXACT TAB NAMES (case-sensitive) ──
const FORM_SHEET = "CRM Form";        // <-- Tab name of your entry form
const DB_SHEET   = "Daily Field Log"; // <-- Tab name of your database/log

// ── CELL REFERENCES on the Form sheet ────────────────────────
const CELLS = {
  date:              "E5",
  area:              "D7",
  businessName:      "D9",
  industry:          "D11",
  contactPerson:     "D13",
  position:          "D15",
  mobile:            "D17",
  email:             "D19",
  decisionMaker:     "D21",
  currentSupplier:   "D23",
  remarks:           "D25",
  itemsNeeded:       "I7",
  estimatedValue:    "I9",
  stage:             "I11",
  nextContactDate:   "I13",
  quotationSent:     "I15",
  quotationAmount:   "I17",
  quotationFileLink: "I19",
  photoLink:         "I21",
  searchBox:         "I3",
};

// ── Daily Field Log Column Map ────────────────────────────────
// A  Lead ID
// B  Date Added
// C  Date Visited
// D  Area
// E  Business Name
// F  Industry Type
// G  Contact Person
// H  Position
// I  Mobile
// J  Email
// K  Decision Maker
// L  Current Supplier
// M  Items Needed
// N  Est. Value (P)
// O  Stage
// P  Last Contact Date
// Q  Next Action Date
// R  Next Action Type
// S  Follow-up Count
// T  Visit Photo Link
// U  Quotation Yes        ← NEW COLUMN
// V  Quotation Amount (P)
// W  Quotation File Link
// X  Remarks
// (Y, Z, AA = auto-calculated audit columns — not written by script)

// ── DEBUG: Run this first to see all your sheet names ─────────
function debugSheetNames() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map((s, i) => i + ": " + s.getName());
  SpreadsheetApp.getUi().alert(
    "Your sheet tabs are:\n\n" + sheets.join("\n") +
    "\n\n──────────────────\n" +
    "FORM_SHEET is set to: \"" + FORM_SHEET + "\"\n" +
    "DB_SHEET   is set to: \"" + DB_SHEET + "\"\n\n" +
    "Update the constants at the top of the script to match exactly."
  );
}

// ── SAFE sheet getter — shows helpful error if name is wrong ──
function getSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    const all = ss.getSheets().map(s => '"' + s.getName() + '"').join(", ");
    throw new Error(
      'Sheet "' + name + '" not found.\n\n' +
      'Available sheets: ' + all + '\n\n' +
      'Fix the FORM_SHEET / DB_SHEET constants at the top of the script.'
    );
  }
  return sheet;
}

// ── SUBMIT ────────────────────────────────────────────────────
function submitForm() {
  const form = getSheet(FORM_SHEET);
  const db   = getSheet(DB_SHEET);

  const get = (cell) => cell ? form.getRange(cell).getValue() : "";

  // Helper: format any date value to yyyy-MM-dd string (no time)
  const fmtDate = (val) => {
    if (!val) return "";
    try { return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
    catch (e) { return val; }
  };

  // Validate required fields
  const required = {
    "Business Name": get(CELLS.businessName),
    "Contact Person": get(CELLS.contactPerson),
    "Mobile":         get(CELLS.mobile),
    "Remarks":        get(CELLS.remarks),
    "Photo Link":     get(CELLS.photoLink),
    "Est. Value":     get(CELLS.estimatedValue),
  };
  const missing = Object.keys(required).filter(k => !required[k]);
  if (missing.length) {
    SpreadsheetApp.getUi().alert("⚠️ Missing required fields:\n" + missing.join(", "));
    return;
  }

  const leadId        = generateLeadId(db);
  const followUpCount = getFollowUpCount(db, get(CELLS.businessName));
  const quotationSentVal = String(get(CELLS.quotationSent)).toLowerCase();
  let   quotationNo   = "";
  if (quotationSentVal === "yes") {
    quotationNo = generateQuotationNo(db);
  }

  const dateStr = fmtDate(get(CELLS.date)); // clean date string, no time

  const row = [
    leadId,                              // A  Lead ID
    dateStr,                             // B  Date Added
    dateStr,                             // C  Date Visited
    get(CELLS.area),                     // D  Area
    get(CELLS.businessName),             // E  Business Name
    get(CELLS.industry),                 // F  Industry Type
    get(CELLS.contactPerson),            // G  Contact Person
    get(CELLS.position),                 // H  Position
    get(CELLS.mobile),                   // I  Mobile
    get(CELLS.email),                    // J  Email
    get(CELLS.decisionMaker),            // K  Decision Maker
    get(CELLS.currentSupplier),          // L  Current Supplier
    get(CELLS.itemsNeeded),              // M  Items Needed
    get(CELLS.estimatedValue),           // N  Est. Value
    get(CELLS.stage),                    // O  Stage
    dateStr,                             // P  Last Contact Date
    fmtDate(get(CELLS.nextContactDate)), // Q  Next Action Date
    "",                                  // R  Next Action Type (not on form)
    followUpCount,                       // S  Follow-up Count
    get(CELLS.photoLink),                // T  Visit Photo Link
    get(CELLS.quotationSent),            // U  Quotation Yes  ← SAVED HERE
    get(CELLS.quotationAmount),          // V  Quotation Amount
    get(CELLS.quotationFileLink),        // W  Quotation File Link
    get(CELLS.remarks),                  // X  Remarks
  ];

  // Find first truly empty row starting from row 5
  const allData = db.getRange(5, 1, db.getMaxRows() - 4, 1).getValues();
  let   nextRow = 5;
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0].toString().trim() !== "") {
      nextRow = i + 6;
    }
  }
  db.getRange(nextRow, 1, 1, row.length).setValues([row]);

  // Format currency columns only — dates are already clean strings
  db.getRange(nextRow, 14).setNumberFormat('#,##0.00'); // N  Est. Value
  db.getRange(nextRow, 22).setNumberFormat('#,##0.00'); // V  Quotation Amount

  SpreadsheetApp.getUi().alert("✅ Lead submitted!\nLead ID: " + leadId);
  clearForm(form);
}

// ── SEARCH ────────────────────────────────────────────────────
function searchLead() {
  const form  = getSheet(FORM_SHEET);
  const db    = getSheet(DB_SHEET);
  const query = form.getRange(CELLS.searchBox).getValue().toString().trim().toLowerCase();

  if (!query) {
    SpreadsheetApp.getUi().alert("Enter a business name or Lead ID to search.");
    return;
  }

  const data    = db.getDataRange().getValues();
  const matches = data.slice(4).filter(row =>
    row[0].toString().toLowerCase().includes(query) || // Lead ID  (col A = index 0)
    row[4].toString().toLowerCase().includes(query)    // Biz Name (col E = index 4)
  );

  if (!matches.length) {
    SpreadsheetApp.getUi().alert('No records found for: "' + query + '"');
    return;
  }

  const r   = matches[matches.length - 1]; // Load most recent match
  const set = (cell, val) => { if (cell) form.getRange(cell).setValue(val); };

  set(CELLS.date,              r[1]);   // B  Date Added
  set(CELLS.area,              r[3]);   // D  Area
  set(CELLS.businessName,      r[4]);   // E  Business Name
  set(CELLS.industry,          r[5]);   // F  Industry Type
  set(CELLS.contactPerson,     r[6]);   // G  Contact Person
  set(CELLS.position,          r[7]);   // H  Position
  set(CELLS.mobile,            r[8]);   // I  Mobile
  set(CELLS.email,             r[9]);   // J  Email
  set(CELLS.decisionMaker,     r[10]);  // K  Decision Maker
  set(CELLS.currentSupplier,   r[11]);  // L  Current Supplier
  set(CELLS.itemsNeeded,       r[12]);  // M  Items Needed
  set(CELLS.estimatedValue,    r[13]);  // N  Est. Value
  set(CELLS.stage,             r[14]);  // O  Stage
  set(CELLS.nextContactDate,   r[16]);  // Q  Next Action Date
  set(CELLS.photoLink,         r[19]);  // T  Visit Photo Link
  set(CELLS.quotationSent,     r[20]);  // U  Quotation Yes
  set(CELLS.quotationAmount,   r[21]);  // V  Quotation Amount
  set(CELLS.quotationFileLink, r[22]);  // W  Quotation File Link
  set(CELLS.remarks,           r[23]);  // X  Remarks
  form.getRange(CELLS.searchBox).clearContent();

  SpreadsheetApp.getUi().alert("✅ Loaded: " + r[4] + " (" + r[0] + ")\nFollow-up #" + r[18]);
}

// ── UPLOAD PHOTO ──────────────────────────────────────────────
function uploadPhoto() {
  const form    = getSheet(FORM_SHEET);
  const bizName = form.getRange(CELLS.businessName).getValue();
  const dateVal = form.getRange(CELLS.date).getValue();

  if (!bizName || !dateVal) {
    SpreadsheetApp.getUi().alert("Fill in Date and Business Name before uploading.");
    return;
  }

  const dateStr  = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const safeName = bizName.toString().replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = dateStr + "_" + safeName + ".jpg";

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(getSidebarHtml(fileName, "photo"))
      .setTitle("Upload Proof Photo").setWidth(400)
  );
}

// ── UPLOAD QUOTATION ──────────────────────────────────────────
function uploadQuotation() {
  const form    = getSheet(FORM_SHEET);
  const db      = getSheet(DB_SHEET);
  const bizName = form.getRange(CELLS.businessName).getValue();

  if (!bizName) {
    SpreadsheetApp.getUi().alert("Fill in Business Name before uploading.");
    return;
  }

  const qNo      = generateQuotationNo(db);
  const safeName = bizName.toString().replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = qNo + "_" + safeName + ".pdf";

  form.getRange(CELLS.quotationSent).setValue("Yes");

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(getSidebarHtml(fileName, "quotation"))
      .setTitle("Upload Quotation PDF").setWidth(400)
  );
}

// ── RECEIVE FILE FROM SIDEBAR ─────────────────────────────────
function saveFileToDrive(base64Data, mimeType, fileName, uploadType) {
  const folderId = uploadType === "photo" ? PHOTO_FOLDER_ID : QUOTATION_FOLDER_ID;
  const folder   = DriveApp.getFolderById(folderId);
  const decoded  = Utilities.base64Decode(base64Data);
  const blob     = Utilities.newBlob(decoded, mimeType, fileName);
  const file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const form = getSheet(FORM_SHEET);
  const cell = uploadType === "photo" ? CELLS.photoLink : CELLS.quotationFileLink;
  form.getRange(cell).setValue(file.getUrl());

  return { success: true, url: file.getUrl(), name: fileName };
}

// ── SIDEBAR HTML ──────────────────────────────────────────────
function getSidebarHtml(fileName, uploadType) {
  const accept = uploadType === "photo" ? "image/*" : ".pdf,application/pdf";
  return `<!DOCTYPE html>
<html>
<head>
<style>
  body{font-family:Arial,sans-serif;padding:16px;font-size:13px}
  .fname{background:#f0f0f0;padding:6px;border-radius:3px;font-family:monospace;
         word-break:break-all;margin:8px 0}
  input[type=file]{width:100%;margin:10px 0}
  button{background:#1a5276;color:white;border:none;padding:10px 20px;
         border-radius:4px;cursor:pointer;width:100%;font-size:14px}
  button:disabled{background:#aaa}
  #status{margin-top:12px;color:#27ae60;font-weight:bold}
  #error{margin-top:12px;color:#e74c3c}
</style>
</head>
<body>
  <b>Will be saved as:</b>
  <div class="fname">${fileName}</div>
  <input type="file" id="f" accept="${accept}">
  <button id="btn" onclick="upload()">⬆ Upload to Drive</button>
  <div id="status"></div>
  <div id="error"></div>
<script>
function upload(){
  const file=document.getElementById('f').files[0];
  if(!file){document.getElementById('error').innerText='Select a file first.';return;}
  document.getElementById('btn').disabled=true;
  document.getElementById('status').innerText='Uploading… please wait';
  document.getElementById('error').innerText='';
  const reader=new FileReader();
  reader.onload=function(e){
    const b64=e.target.result.split(',')[1];
    google.script.run
      .withSuccessHandler(function(r){
        document.getElementById('status').innerText='✅ Done: '+r.name;
        document.getElementById('btn').disabled=false;
      })
      .withFailureHandler(function(err){
        document.getElementById('error').innerText='❌ '+err.message;
        document.getElementById('btn').disabled=false;
        document.getElementById('status').innerText='';
      })
      .saveFileToDrive(b64,file.type,'${fileName}','${uploadType}');
  };
  reader.readAsDataURL(file);
}
</script>
</body>
</html>`;
}

// ── HELPERS ───────────────────────────────────────────────────
function generateLeadId(db) {
  const ids = db.getRange(5, 1, db.getMaxRows() - 4, 1).getValues().flat()
    .map(v => v.toString())
    .filter(v => /^L-\d+$/.test(v))
    .map(v => parseInt(v.replace("L-", "")));
  const max = ids.length ? Math.max(...ids) : 0;
  return "L-" + String(max + 1).padStart(4, "0");
}

function generateQuotationNo(db) {
  const now  = new Date();
  const yr   = now.getFullYear();
  const mo   = String(now.getMonth() + 1).padStart(2, "0");
  // Quotation File Link is now col W (column number 23)
  const nums = db.getRange(5, 23, db.getMaxRows() - 4, 1).getValues().flat()
    .map(v => v.toString())
    .filter(v => /^Q-\d{4}-\d{2}-\d+$/.test(v))
    .map(v => parseInt(v.split("-").pop()));
  const max = nums.length ? Math.max(...nums) : 0;
  return "Q-" + yr + "-" + mo + "-" + String(max + 1).padStart(3, "0");
}

function getFollowUpCount(db, businessName) {
  // Business Name is col E (column number 5)
  const names = db.getRange(5, 5, db.getMaxRows() - 4, 1).getValues().flat();
  const count = names.filter(n =>
    n.toString().toLowerCase() === businessName.toString().toLowerCase()
  ).length;
  return count + 1;
}

function clearForm(form) {
  [CELLS.area, CELLS.businessName, CELLS.industry,
   CELLS.contactPerson, CELLS.position, CELLS.mobile, CELLS.email,
   CELLS.currentSupplier, CELLS.remarks,
   CELLS.itemsNeeded, CELLS.estimatedValue, CELLS.nextContactDate,
   CELLS.quotationAmount, CELLS.quotationFileLink,
   CELLS.photoLink, CELLS.searchBox
  ].filter(c => c).forEach(c => form.getRange(c).clearContent());
  form.getRange(CELLS.date).setValue(new Date());
  form.getRange(CELLS.stage).setValue("Lead");
  form.getRange(CELLS.decisionMaker).setValue("Yes");
  form.getRange(CELLS.quotationSent).setValue("No");
}

// ── DELETE ────────────────────────────────────────────────────
function deleteRecord() {
  const form    = getSheet(FORM_SHEET);
  const db      = getSheet(DB_SHEET);
  const bizName = form.getRange(CELLS.businessName).getValue().toString().trim();

  if (!bizName) {
    SpreadsheetApp.getUi().alert("Load a record first before deleting.");
    return;
  }

  const ui      = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "⚠️ Confirm Delete",
    'Delete record for "' + bizName + '"?\nThis cannot be undone.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const lastRow = db.getLastRow();
  if (lastRow < 5) {
    ui.alert("No records found in the database.");
    return;
  }

  const data = db.getRange(5, 1, lastRow - 4, 24).getValues();
  let deletedRow = -1;

  // Match by Business Name — col E = index 4 (most recent match)
  for (let i = data.length - 1; i >= 0; i--) {
    const rowBiz = data[i][4].toString().trim().toLowerCase();
    if (rowBiz === bizName.toLowerCase()) {
      deletedRow = i + 5;
      break;
    }
  }

  if (deletedRow === -1) {
    ui.alert('No matching record found for: "' + bizName + '"');
    return;
  }

  db.deleteRow(deletedRow);
  ui.alert("🗑️ Record deleted: " + bizName + " (row " + deletedRow + ")");
  clearForm(form);
}

// ── MENU ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 CRM")
    .addItem("Submit Lead",                    "submitForm")
    .addItem("Search Lead",                    "searchLead")
    .addSeparator()
    .addItem("Upload Proof Photo",             "uploadPhoto")
    .addItem("Upload Quotation PDF",           "uploadQuotation")
    .addSeparator()
    .addItem("🗑️ Delete Current Record",       "deleteRecord")
    .addSeparator()
    .addItem("🔧 Debug: Show Sheet Names",     "debugSheetNames")
    .addToUi();
}
