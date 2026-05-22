// ============================================================
//  SheetSetup.gs
//  One-time setup: creates all sheets with headers & formatting
// ============================================================

const SheetSetup = (() => {

  const SHEET_HEADERS = {
    [SHEET_STUDENTS]: [
      'Student ID', 'Last Name', 'First Name', 'Non-Grad', 'Not Walking',
      'Honor Grad', 'CTE', 'AVID', 'ASB', 'Military', 'CSF', 'HPC', 'VAPA', 
      'PA', 'SSA', 'Art', 'FNL', 'Yearbook', 'Norsecrew', 'Sharp-Mem', 
      'Key Club', 'Bi-Literacy', 'Gold State Seal', 'HOSA', 'Academic League', 
      'Not Walking Reason', 'Seat Row', 'Seat Number', 'Notes', 'Added By', 'Timestamp'
    ],
    [SHEET_CATEGORIES]: ['ID','Label','Color','Icon','Sort Priority'],
    [SHEET_SEATING]:    ['Config JSON'],
    [SHEET_USERS]:      ['Email','Role','Categories','Display Name','Added Date']
  };

  const CATEGORY_DEFAULTS = [
    ['honor',    'Honor Grad',         '#F0C040','star',    1],
    ['cte',      'CTE',                '#4A90D9','tool',    2],
    ['avid',     'AVID',               '#7B68EE','books',   3],
    ['asb',      'ASB',                '#50C878','flag',    4],
    ['military', 'Military',           '#8B6914','shield',  5],
    ['special',  'Special Recognition','#E8735A','award',   6]
  ];

  function initSheet(name, sheet) {
    const headers = SHEET_HEADERS[name];
    if (!headers) return;

    // Write headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1A1A2E');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);

    sheet.setFrozenRows(1);

    // Seed category defaults
    if (name === SHEET_CATEGORIES) {
      sheet.getRange(2, 1, CATEGORY_DEFAULTS.length, 5).setValues(CATEGORY_DEFAULTS);
    }

    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
  }

  /**
   * Run this manually from the Script Editor to initialize all sheets.
   * Menu: Extensions > Apps Script > Run > setupAllSheets
   */
  function setupAllSheets() {
    const ss     = getSpreadsheet();
    const names  = [SHEET_STUDENTS, SHEET_CATEGORIES, SHEET_SEATING, SHEET_USERS];

    names.forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
      }
      // Re-init even if sheet exists (safe – only writes headers in row 1)
      initSheet(name, sheet);
    });

    // Ensure current user is admin
    const email = Session.getActiveUser().getEmail();
    const users = ss.getSheetByName(SHEET_USERS);
    users.appendRow([email, 'admin', 'honor,cte,avid,asb,military,special', 'Admin', new Date()]);

    SpreadsheetApp.getUi().alert('✅ All sheets initialized! You are set as Admin.');
  }

  // Expose menu item
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('🎓 Grad Tracker')
      .addItem('Initialize Sheets', 'SheetSetup.setupAllSheets')
      .addItem('Open Web App', 'openWebApp')
      .addToUi();
  }

  return { initSheet, setupAllSheets, onOpen };
})();

// ── Global hooks ─────────────────────────────────────────────
function onOpen()     { SheetSetup.onOpen(); }
function setupAllSheets() { SheetSetup.setupAllSheets(); }
