// ============================================================
//  GRADUATION TRACKER — Code.gs
//  Main entry point & routing for the Apps Script Web App
// ============================================================

// ── Spreadsheet & Sheet Names ────────────────────────────────
const SS_ID        = '1S3NpFDIKsyrcCwbeJJFSMxfJybuiTxViAaVCH8LK4GQ'; // ← Paste your Spreadsheet ID here
const SHEET_STUDENTS  = 'Students';
const SHEET_CATEGORIES = 'Categories';
const SHEET_SEATING   = 'SeatingConfig';
const SHEET_USERS     = 'Users';

// ── Web App Entry Points ─────────────────────────────────────

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Graduation Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Utility: include HTML partials in Index.html
 * Usage: <?!= include('Partials/Sidebar'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Spreadsheet Helper ───────────────────────────────────────

function getSpreadsheet() {
  return SS_ID
    ? SpreadsheetApp.openById(SS_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    SheetSetup.initSheet(name, sheet);
  }
  return sheet;
}

// ── Auth / Session ───────────────────────────────────────────

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const sheet = getSheet(SHEET_USERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return { email, role: data[i][1], categories: data[i][2] ? data[i][2].split(',') : [] };
    }
  }
  // Default: read-only viewer
  return { email, role: 'viewer', categories: [] };
}

function isAdmin() {
  return getCurrentUser().role === 'admin';
}

// ── Public API (called via google.script.run) ────────────────

function getInitialData() {
  return {
    user:       getCurrentUser(),
    students:   StudentService.getAllStudents(),
    categories: CategoryService.getAllCategories(),
    seating:    SeatingService.getSeatingConfig(),
    stats:      StudentService.getStats()
  };
}
