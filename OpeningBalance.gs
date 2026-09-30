/**
 * ============================================================================
 * OPENING BALANCE CALCULATOR - SAVE TO SPREADSHEET
 * Add this as a NEW file in your existing Apps Script project
 * (Files > + > Script, name it "OpeningBalance"). Leave your existing Code.gs
 * exactly as it is. This file re-uses SPREADSHEET_ID and SPREADSHEET_LABEL
 * from Code.gs, so it saves into the same spreadsheet.
 * ----------------------------------------------------------------------------
 * What it does:
 *   - Receives the figures from the Opening Balance Calculator page (POST).
 *   - Checks the Client Name is filled in (required).
 *   - Re-calculates the results itself so the sheet can't hold wrong maths.
 *   - Adds one new row per calculation to the "Opening Balance Call" tab.
 *   - Creates the tab (with headers) automatically the first time.
 *
 * If you want the tab called something else, change OB_SHEET_NAME below.
 * ============================================================================
 */

const OB_SHEET_NAME = 'Opening Balance Call';
const OB_TIMEZONE = 'Europe/London';

const OB_HEADERS = [
  'Submitted At',
  'Client Name',
  'Go-Live Date',
  'Last Reconciled Date',
  'Insured Deposit Closing Balance',
  'Client Account Closing Balance',
  'Balances in Street',
  'NRL Tax',
  'Unreconciled Payment Groups',
  'Deposits Held in Client Account',
  'Suspense Opening Balance (With Reconciliation)',
  'Suspense Opening Balance (Without Reconciliation)',
  'Client Account Balance (Enter in Company Settings)'
];

/** Receives the calculator's POST request. */
function doPost(e) {
  let result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data was received.');
    }
    const data = JSON.parse(e.postData.contents);
    result = saveOpeningBalance_(data);
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Validates, calculates and saves one submission as a new row. */
function saveOpeningBalance_(data) {
  // ---- Required: client name ----
  const clientName = String(data.clientName || '').trim().substring(0, 200);
  if (!clientName) throw new Error('Client name is required.');

  // ---- Dates (sent as YYYY-MM-DD) ----
  const goLiveDate = obParseDate_(data.goLiveDate, 'Go-live date');
  const lastReconciledDate = obParseDate_(data.lastReconciledDate, 'Last reconciled date');

  // ---- Amounts ----
  const insuredDeposit = obToNumber_(data.insuredDeposit, 'Insured deposit');
  const clientAccount = obToNumber_(data.clientAccount, 'Client account closing balance');
  const balancesInStreet = obToNumber_(data.balancesInStreet, 'Balances in Street');
  const nrlTax = obToNumber_(data.nrlTax, 'NRL tax');
  const unreconciled = obToNumber_(data.unreconciledPayments, 'Unreconciled payment groups');
  const depositsHeld = obToNumber_(data.depositsHeld, 'Deposits held in client account');

  if (clientAccount <= 0) throw new Error('Client account closing balance must be greater than zero.');

  // ---- Calculate (same maths as the page) ----
  const withRecon = obRound_(clientAccount - balancesInStreet - nrlTax - depositsHeld);
  const withoutRecon = obRound_(clientAccount - balancesInStreet - nrlTax - unreconciled - depositsHeld);

  const row = [
    new Date(),
    obSafeText_(clientName),
    goLiveDate,
    lastReconciledDate,
    insuredDeposit,
    clientAccount,
    balancesInStreet,
    nrlTax,
    unreconciled,
    depositsHeld,
    withRecon,
    withoutRecon,
    withoutRecon
  ];

  // Lock so two people submitting at the same moment can't overwrite each other
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = obGetOrCreateSheet_();
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

    sheet.getRange(newRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(newRow, 3, 1, 2).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(newRow, 5, 1, 9).setNumberFormat('£#,##0.00;[Red]-£#,##0.00');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    spreadsheet: SPREADSHEET_LABEL,
    sheet: OB_SHEET_NAME,
    suspenseWithRecon: withRecon,
    suspenseWithoutRecon: withoutRecon
  };
}

/** Finds the "Opening Balance Call" tab, or creates it with a formatted header row. */
function obGetOrCreateSheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(OB_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OB_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, OB_HEADERS.length).setValues([OB_HEADERS]);
    sheet.getRange(1, 1, 1, OB_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#5B4BE7')
      .setFontColor('#FFFFFF')
      .setWrap(true)
      .setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidths(3, 2, 120);
    sheet.setColumnWidths(5, 9, 150);
  }
  return sheet;
}

// ---- small helpers (prefixed "ob" so they never clash with Code.gs) ---------

function obParseDate_(value, label) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
  if (!m) throw new Error(label + ' is required.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function obToNumber_(value, label) {
  const n = Number(value);
  if (value === '' || value === null || value === undefined || !isFinite(n)) {
    throw new Error(label + ' must be a number.');
  }
  return obRound_(n);
}

function obRound_(n) {
  return Math.round(n * 100) / 100;
}

/** Stops text starting with = + - @ from being treated as a spreadsheet formula. */
function obSafeText_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

/**
 * ============================================================================
 * DEPLOYMENT NOTES
 * 1. Open the "Accounting Onboarding / Amy Dawber" spreadsheet.
 * 2. Extensions > Apps Script.
 * 3. Click + next to Files > Script > name it "OpeningBalance" and paste this in.
 *    Do NOT change Code.gs.
 * 4. Deploy > Manage deployments > pencil icon on your EXISTING deployment >
 *    Version: New version > Deploy. The /exec URL stays the same.
 *    (Make sure "Execute as: Me" and "Who has access: Anyone" are set.)
 * 5. The first time you deploy, Google may ask you to authorise the script again.
 *    Click through and allow it.
 * ============================================================================
 */
