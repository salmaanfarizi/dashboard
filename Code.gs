/**
 * CUMULATIVE MONTHLY ACCOUNTING REPORT SYSTEM
 * Version: 3.2 (Improved & Fixed)
 *
 * Fixes:
 * - Dashboard formulas now dynamically reference correct rows
 * - Removed duplicate functions
 * - Added constants for magic numbers
 * - Improved error handling
 * - Batch operations for better performance
 */

/* ==================== CONFIGURATION ==================== */

const CONFIG = {
  VERSION: '3.2',
  SETTINGS_LIST_START_ROW: 15,
  MAX_DATA_ROWS: 502,
  USD_CONVERSION_RATE: 3.75,
  DROPDOWN_REPAIR_COOLDOWN_MS: 1500,
  MONTHS: ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
};

/* ==================== INITIALIZATION ==================== */

function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createMasterComparisonSheets(ss);
  createMasterDashboard(ss);
  createSettingsSheet(ss);
  addNewMonth('MAR-2025', true);
  createCustomMenu();

  SpreadsheetApp.getUi().alert(
    '✅ System Initialized!',
    'Use the "📊 Monthly Reports" menu to add months, update data, and view reports.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ==================== CUSTOM MENU ==================== */

function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Monthly Reports')
    .addItem('➕ Add New Month', 'addNewMonthDialog')
    .addItem('📝 Update Current Month', 'updateCurrentMonth')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📈 View Reports')
      .addItem('Banks Comparison', 'viewBanksComparison')
      .addItem('Advances Comparison', 'viewAdvancesComparison')
      .addItem('Suspense Comparison', 'viewSuspenseComparison')
      .addItem('Outstanding Comparison', 'viewOutstandingComparison')
      .addItem('Dashboard', 'viewDashboard'))
    .addSeparator()
    .addItem('🔧 Fix Dropdowns', 'repairDropdownsForAllMonths')
    .addItem('👁️ Show/Hide Sheets', 'toggleSheetVisibility')
    .addItem('📧 Email Monthly Report', 'emailMonthlyReport')
    .addItem('📄 Export to PDF', 'exportToPDF')
    .addSeparator()
    .addItem('🔄 Refresh All Calculations', 'refreshAllCalculations')
    .addItem('⚙️ Settings', 'openSettings')
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

/* ==================== MONTH HELPERS ==================== */

function splitMonthYear_(monthYear) {
  const s = (monthYear || '').trim().toUpperCase();
  if (!/^[A-Z]{3}-\d{4}$/.test(s)) throw new Error('Invalid monthYear: ' + monthYear);
  const [month, year] = s.split('-');
  return { month, year };
}

function monthYearToDate_(monthYear) {
  const { month, year } = splitMonthYear_(monthYear);
  const i = CONFIG.MONTHS.indexOf(month);
  if (i < 0) throw new Error('Invalid month: ' + monthYear);
  return new Date(Number(year), i, 1);
}

function getPreviousMonthYear_(monthYear) {
  const { month, year } = splitMonthYear_(monthYear);
  const i = CONFIG.MONTHS.indexOf(month);
  if (i === 0) return `DEC-${Number(year) - 1}`;
  return `${CONFIG.MONTHS[i - 1]}-${year}`;
}

function getMonthlySheets_(ss, prefix) {
  const P = prefix + '_';
  return ss.getSheets()
    .filter(s => s.getName().startsWith(P))
    .map(s => {
      const suffix = s.getName().substring(P.length);
      if (!/^[A-Z]{3}-\d{4}$/.test(suffix)) return null;
      return { sheet: s, monthYear: suffix, date: monthYearToDate_(suffix) };
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);
}

function getLatestMonthYear_(ss, prefix) {
  const list = getMonthlySheets_(ss, prefix);
  return list.length ? list[list.length - 1].monthYear : null;
}

/* ==================== UTILITY FUNCTIONS ==================== */

function findLastNonBlankRow_(sheet, colLetter, startRow) {
  const col = colToIndex_(colLetter);
  const vals = sheet.getRange(startRow, col, sheet.getMaxRows() - startRow + 1, 1).getValues();
  let last = startRow - 1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== '') last = startRow + i;
  }
  return last;
}

function colToIndex_(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}

function colToA1_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function setIfEmpty_(rng, val) {
  if (rng.getValue() === '') {
    rng.clearDataValidations();
    rng.setValue(val);
  }
}

function seedListIfEmpty_(sheet, colLetter, startRow, defaults) {
  const last = findLastNonBlankRow_(sheet, colLetter, startRow);
  if (last < startRow) {
    const col = colToIndex_(colLetter);
    sheet.getRange(startRow, col, defaults.length, 1).setValues(defaults.map(v => [v]));
  }
}

/* ==================== ADD NEW MONTH ==================== */

function addNewMonthDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.prompt('Add New Month', 'Enter month (e.g., APR-2025):', ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const monthYear = resp.getResponseText().trim().toUpperCase();

  if (!/^[A-Z]{3}-\d{4}$/.test(monthYear)) {
    ui.alert('❌ Error', 'Please enter in format MMM-YYYY (e.g., APR-2025)', ui.ButtonSet.OK);
    return;
  }

  // Check if month already exists
  if (ss.getSheetByName(`Banks_${monthYear}`)) {
    ui.alert('❌ Error', `Month ${monthYear} already exists!`, ui.ButtonSet.OK);
    return;
  }

  addNewMonth(monthYear, false);
  ui.alert('✅ Success', `Month ${monthYear} created & balances carried forward.`, ui.ButtonSet.OK);
}

function addNewMonth(monthYear, isInitialSetup) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!isInitialSetup) hidePreviousMonthSheets(ss, monthYear);

  createBankSheet(ss, monthYear);
  createAdvanceSheet(ss, monthYear);
  createSuspenseSheet(ss, monthYear);
  createOutstandingSheet(ss, monthYear);

  if (!isInitialSetup) carryForwardBalances(ss, monthYear);

  updateAllComparisons(ss);
}

/* ==================== HIDE PREVIOUS MONTH SHEETS ==================== */

function hidePreviousMonthSheets(ss, currentMonthYear) {
  const prevMonthYear = getPreviousMonthYear_(currentMonthYear);
  ['Banks','Advances','Suspense','Outstanding'].forEach(p => {
    const sh = ss.getSheetByName(`${p}_${prevMonthYear}`);
    if (sh) sh.hideSheet();
  });
}

/* ==================== CARRY FORWARD ==================== */

function carryForwardBalances(ss, monthYear) {
  const prevMY = getPreviousMonthYear_(monthYear);

  // Banks: carry closing into next opening
  const prevBank = ss.getSheetByName(`Banks_${prevMY}`);
  const curBank  = ss.getSheetByName(`Banks_${monthYear}`);
  if (prevBank && curBank) {
    for (let r = 4; r <= 7; r++) {
      const v = prevBank.getRange(`F${r}`).getValue();
      if (v !== '') curBank.getRange(`B${r}`).setValue(v);
    }
    const usd = prevBank.getRange('F13').getValue();
    if (usd !== '') curBank.getRange('B13').setValue(usd);
  }

  // Advances: opening shows previous closing
  const prevAdv = ss.getSheetByName(`Advances_${prevMY}`);
  const curAdv  = ss.getSheetByName(`Advances_${monthYear}`);
  if (prevAdv && curAdv) curAdv.getRange('B3').setValue(prevAdv.getRange('E3').getValue() || 0);

  // Suspense: opening shows previous closing
  const prevSus = ss.getSheetByName(`Suspense_${prevMY}`);
  const curSus  = ss.getSheetByName(`Suspense_${monthYear}`);
  if (prevSus && curSus) curSus.getRange('B3').setValue(prevSus.getRange('E3').getValue() || 0);
}

/* ==================== SHEET CREATION ==================== */

function createBankSheet(ss, monthYear) {
  const name = `Banks_${monthYear}`;
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue(`BANK SUMMARY - ${monthYear}`);
  sheet.getRange('A1:G1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:G1').setBackground('#1a73e8').setFontColor('#ffffff');

  // SAR section
  sheet.getRange('A3:G3').setValues([['Bank Name','Opening Balance','Total Received','Bank Charges','Total Payments','Closing Balance','Notes']])
    .setBackground('#e8f0fe').setFontWeight('bold');

  sheet.getRange('A4:G9').setValues([
    ['Alrajhi-1097','','','','','=B4+C4-D4-E4',''],
    ['Alrajhi-new','','','','','=B5+C5-D5-E5',''],
    ['SNB Al Ahsa Branch','','','','','=B6+C6-D6-E6',''],
    ['Albilad','','','','','=B7+C7-D7-E7',''],
    ['','','','','','',''],
    ['SUB-TOTAL SAR','=SUM(B4:B7)','=SUM(C4:C7)','=SUM(D4:D7)','=SUM(E4:E7)','=SUM(F4:F7)','']
  ]);
  sheet.getRange('A9:G9').setFontWeight('bold').setBackground('#fff3cd');

  // USD section
  sheet.getRange('A11').setValue('USD ACCOUNTS');
  sheet.getRange('A11:G11').merge().setBackground('#e8f0fe').setFontWeight('bold');
  sheet.getRange('A12:G12').setValues([['Bank Name','Opening','Amount Received','Bank Charge','Payment','Balance','Notes']])
    .setBackground('#f8f9fa').setFontWeight('bold');
  sheet.getRange('A13:G15').setValues([
    ['Albilad USD','','','','','=B13+C13-D13-E13',''],
    ['','','','','','',''],
    ['SUB-TOTAL USD','=SUM(B13:B13)','=SUM(C13:C13)','=SUM(D13:D13)','=SUM(E13:E13)','=SUM(F13:F13)','']
  ]);
  sheet.getRange('A15:G15').setFontWeight('bold').setBackground('#fff3cd');

  // Grand total
  sheet.getRange('A17').setValue('GRAND TOTAL (SAR + USD Converted)');
  sheet.getRange('B17').setFormula(`=B9+B15*${CONFIG.USD_CONVERSION_RATE}`);
  sheet.getRange('C17').setFormula(`=C9+C15*${CONFIG.USD_CONVERSION_RATE}`);
  sheet.getRange('D17').setFormula(`=D9+D15*${CONFIG.USD_CONVERSION_RATE}`);
  sheet.getRange('E17').setFormula(`=E9+E15*${CONFIG.USD_CONVERSION_RATE}`);
  sheet.getRange('F17').setFormula(`=F9+F15*${CONFIG.USD_CONVERSION_RATE}`);
  sheet.getRange('A17:G17').setFontWeight('bold').setBackground('#d4edda');

  // Summary
  sheet.getRange('A19').setValue('SUMMARY');
  sheet.getRange('A19:C19').merge().setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A20').setValue('Total Bank Balance (SAR):');
  sheet.getRange('C20').setFormula('=F17').setFontWeight('bold').setFontSize(12);

  // Formatting
  sheet.getRange('B4:F17').setNumberFormat('#,##0.00');
  sheet.getRange('C20').setNumberFormat('#,##0.00');
  sheet.getRange('A3:G9').setBorder(true,true,true,true,true,true);
  sheet.getRange('A12:G15').setBorder(true,true,true,true,true,true);
  sheet.getRange('A17:G17').setBorder(true,true,true,true,false,false);
  sheet.getRange('A19:C20').setBorder(true,true,true,true,true,true);

  sheet.setColumnWidth(1,180);
  sheet.setColumnWidths(2,5,120);
  sheet.setColumnWidth(7,150);
  sheet.setFrozenRows(3);

  // Dropdowns
  applyDropdownFromSettings_(sheet.getRange('A4:A7'),'A', CONFIG.SETTINGS_LIST_START_ROW);
  applyDropdownFromSettings_(sheet.getRange('A13'),'A', CONFIG.SETTINGS_LIST_START_ROW);
}

function createAdvanceSheet(ss, monthYear) {
  const name = `Advances_${monthYear}`;
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue(`ADVANCE RECONCILIATION - ${monthYear}`);
  sheet.getRange('A1:G1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:G1').setBackground('#28a745').setFontColor('#ffffff');

  // Header row
  sheet.getRange('A3').setValue('OPENING BALANCE');
  sheet.getRange('B3').setValue(0).setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('D3').setValue('CLOSING BALANCE');

  const advancesBase = getAdvancesBase_();
  sheet.getRange('E3')
    .setFormula(`=${advancesBase}+SUM(E7:E1000)`)
    .setNumberFormat('#,##0.00').setFontWeight('bold');

  sheet.getRange('A3:B3').setBackground('#e7f3e7');
  sheet.getRange('D3:E3').setBackground('#e7f3e7');

  // Transactions section
  sheet.getRange('A5').setValue('TRANSACTIONS');
  sheet.getRange('A5:G5').merge().setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A6:G6').setValues([['Date','Voucher No','Description','Person/Party','Advance Given','Amount Settled','Running Balance']])
    .setBackground('#f8f9fa').setFontWeight('bold');

  // Running balance formulas
  for (let i = 7; i <= 100; i++) {
    sheet.getRange(`G${i}`).setFormula(i === 7
      ? `=IF(E${i}="","",B3+E${i}-F${i})`
      : `=IF(E${i}="","",G${i-1}+E${i}-F${i})`);
  }

  // Total row
  sheet.getRange('A102').setValue('TOTAL');
  sheet.getRange('E102').setFormula('=SUM(E7:E1000)');
  sheet.getRange('F102').setFormula('=SUM(F7:F1000)');
  sheet.getRange('A102:G102').setFontWeight('bold').setBackground('#fff3cd');

  // Formatting
  sheet.getRange('E7:G102').setNumberFormat('#,##0.00');
  sheet.getRange('A6:G101').setBorder(true,true,true,true,true,true);
  sheet.getRange('A102:G102').setBorder(true,true,true,true,false,false);

  sheet.setColumnWidth(1,100);
  sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,250);
  sheet.setColumnWidth(4,150);
  sheet.setColumnWidths(5,3,120);
  sheet.setFrozenRows(6);
}

function createSuspenseSheet(ss, monthYear) {
  const name = `Suspense_${monthYear}`;
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue(`SUSPENSE RECONCILIATION - ${monthYear}`);
  sheet.getRange('A1:G1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:G1').setBackground('#ffc107').setFontColor('#000000');

  // Header row
  sheet.getRange('A3').setValue('OPENING BALANCE');
  sheet.getRange('B3').setValue(0).setNumberFormat('#,##0.00').setFontWeight('bold');
  sheet.getRange('D3').setValue('CLOSING BALANCE');

  const suspenseBase = getSuspenseBase_();
  sheet.getRange('E3').setFormula(`=${suspenseBase}+SUM(E7:E1000)-SUM(F7:F1000)`)
    .setNumberFormat('#,##0.00').setFontWeight('bold');

  sheet.getRange('A3:B3').setBackground('#fff3cd');
  sheet.getRange('D3:E3').setBackground('#fff3cd');

  // Transactions section
  sheet.getRange('A5').setValue('UNIDENTIFIED TRANSACTIONS');
  sheet.getRange('A5:G5').merge().setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A6:G6').setValues([['Date','Journal No','Description','Reference','Debit','Credit','Running Balance']])
    .setBackground('#f8f9fa').setFontWeight('bold');

  // Running balance formulas
  for (let i = 7; i <= 100; i++) {
    sheet.getRange(`G${i}`).setFormula(i === 7
      ? `=IF(AND(E${i}="",F${i}=""),"",B3+E${i}-F${i})`
      : `=IF(AND(E${i}="",F${i}=""),"",G${i-1}+E${i}-F${i})`);
  }

  // Total row
  sheet.getRange('A102').setValue('TOTAL');
  sheet.getRange('E102').setFormula('=SUM(E7:E1000)');
  sheet.getRange('F102').setFormula('=SUM(F7:F1000)');
  sheet.getRange('A102:G102').setFontWeight('bold').setBackground('#fff3cd');

  // Formatting
  sheet.getRange('E7:G102').setNumberFormat('#,##0.00');
  sheet.getRange('A6:G101').setBorder(true,true,true,true,true,true);
  sheet.getRange('A102:G102').setBorder(true,true,true,true,false,false);

  sheet.setColumnWidth(1,100);
  sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,250);
  sheet.setColumnWidth(4,120);
  sheet.setColumnWidths(5,3,120);
  sheet.setFrozenRows(6);
}

function createOutstandingSheet(ss, monthYear) {
  const name = `Outstanding_${monthYear}`;
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue(`CUSTOMER OUTSTANDING - ${monthYear}`);
  sheet.getRange('A1:H1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:H1').setBackground('#dc3545').setFontColor('#ffffff');

  // Get salesmen from Settings
  const settings = ss.getSheetByName('Settings');
  const lastRow = findLastNonBlankRow_(settings, 'C', CONFIG.SETTINGS_LIST_START_ROW);
  const salesmen = settings.getRange(`C${CONFIG.SETTINGS_LIST_START_ROW}:C${lastRow}`).getValues().flat().filter(String);

  // Summary headers
  sheet.getRange('A3:D3').setValues([['Salesman','Total Outstanding','No. of Customers','Average']])
    .setBackground('#f8f9fa').setFontWeight('bold');

  // Summary rows for each salesman
  const detailsStartRow = 4 + salesmen.length + 2; // After summary + gap
  salesmen.forEach((s, i) => {
    const row = 4 + i;
    sheet.getRange(row, 1).setValue(s);
    sheet.getRange(row, 2).setFormula(`=SUMIF($D$${detailsStartRow}:$D$${CONFIG.MAX_DATA_ROWS},A${row},$G$${detailsStartRow}:$G$${CONFIG.MAX_DATA_ROWS})`);
    sheet.getRange(row, 3).setFormula(`=COUNTIF($D$${detailsStartRow}:$D$${CONFIG.MAX_DATA_ROWS},A${row})`);
    sheet.getRange(row, 4).setFormula(`=IFERROR(B${row}/C${row},0)`);
  });

  // Total row
  const totalRow = 4 + salesmen.length;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold').setBackground('#fff3cd');
  sheet.getRange(totalRow, 2).setFormula(`=SUM(B4:B${totalRow-1})`);
  sheet.getRange(totalRow, 3).setFormula(`=SUM(C4:C${totalRow-1})`);
  sheet.getRange(totalRow, 4).setFormula(`=IFERROR(B${totalRow}/C${totalRow},0)`);

  // Details section header
  const detailsHeaderRow = totalRow + 2;
  sheet.getRange(detailsHeaderRow, 1, 1, 8)
    .setValues([['Customer Code','Customer Name','Area','Salesman','Invoice Amount','Paid Amount','Balance','Days']])
    .setBackground('#f8f9fa').setFontWeight('bold');

  // Balance formulas for detail rows
  for (let r = detailsHeaderRow + 1; r <= CONFIG.MAX_DATA_ROWS; r++) {
    sheet.getRange(`G${r}`).setFormula(`=IF(E${r}="","",E${r}-F${r})`);
  }

  // Grand total row
  const grandTotalRow = CONFIG.MAX_DATA_ROWS + 1;
  sheet.getRange(`A${grandTotalRow}`).setValue('GRAND TOTAL');
  sheet.getRange(`E${grandTotalRow}`).setFormula(`=SUM(E${detailsHeaderRow + 1}:E${CONFIG.MAX_DATA_ROWS})`);
  sheet.getRange(`F${grandTotalRow}`).setFormula(`=SUM(F${detailsHeaderRow + 1}:F${CONFIG.MAX_DATA_ROWS})`);
  sheet.getRange(`G${grandTotalRow}`).setFormula(`=SUM(G${detailsHeaderRow + 1}:G${CONFIG.MAX_DATA_ROWS})`);
  sheet.getRange(`A${grandTotalRow}:H${grandTotalRow}`).setFontWeight('bold').setBackground('#fff3cd');

  // Formatting
  sheet.getRange(`B4:D${totalRow}`).setNumberFormat('#,##0.00');
  sheet.setFrozenRows(detailsHeaderRow);

  // Dropdowns
  applyDropdownFromSettings_(sheet.getRange(`C${detailsHeaderRow + 1}:C${CONFIG.MAX_DATA_ROWS}`), 'E', CONFIG.SETTINGS_LIST_START_ROW);
  applyDropdownFromSettings_(sheet.getRange(`D${detailsHeaderRow + 1}:D${CONFIG.MAX_DATA_ROWS}`), 'C', CONFIG.SETTINGS_LIST_START_ROW);
}

/* ==================== COMPARISON SHEETS ==================== */

function createMasterComparisonSheets(ss) {
  createBankComparisonSheet(ss);
  createAdvanceComparisonSheet(ss);
  createSuspenseComparisonSheet(ss);
  createOutstandingComparisonSheet(ss);
}

function createBankComparisonSheet(ss) {
  const name = 'Banks_Comparison';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  sheet.getRange('A1').setValue('BANKS - MONTHLY COMPARISON');
  sheet.getRange('A1:N1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:N1').setBackground('#1a73e8').setFontColor('#ffffff');

  sheet.getRange('A3').setValue('Metric').setBackground('#f8f9fa').setFontWeight('bold');
  const metrics = ['Opening Balance','Total Received','Bank Charges','Total Payments','Closing Balance','Net Cash Flow','Month-over-Month %'];
  sheet.getRange(4, 1, metrics.length, 1).setValues(metrics.map(m => [m]));
  sheet.getRange('A3:A10').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
}

function createAdvanceComparisonSheet(ss) {
  const name = 'Advances_Comparison';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  sheet.getRange('A1').setValue('ADVANCES - MONTHLY COMPARISON');
  sheet.getRange('A1:N1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:N1').setBackground('#28a745').setFontColor('#ffffff');

  sheet.getRange('A3').setValue('Metric').setBackground('#f8f9fa').setFontWeight('bold');
  const metrics = ['Opening Balance','Advances Given','Advances Settled','Closing Balance','Settlement Rate %','Average Advance'];
  sheet.getRange(4, 1, metrics.length, 1).setValues(metrics.map(m => [m]));
  sheet.getRange('A3:A9').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
}

function createSuspenseComparisonSheet(ss) {
  const name = 'Suspense_Comparison';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  sheet.getRange('A1').setValue('SUSPENSE - MONTHLY COMPARISON');
  sheet.getRange('A1:N1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:N1').setBackground('#ffc107').setFontColor('#000000');

  sheet.getRange('A3').setValue('Metric').setBackground('#f8f9fa').setFontWeight('bold');
  const metrics = ['Opening Balance','Total Debits','Total Credits','Closing Balance','Unidentified Items','Resolution Rate %'];
  sheet.getRange(4, 1, metrics.length, 1).setValues(metrics.map(m => [m]));
  sheet.getRange('A3:A9').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
}

function createOutstandingComparisonSheet(ss) {
  const name = 'Outstanding_Comparison';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  sheet.getRange('A1').setValue('CUSTOMER OUTSTANDING - MONTHLY COMPARISON');
  sheet.getRange('A1:N1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:N1').setBackground('#dc3545').setFontColor('#ffffff');

  sheet.getRange('A3').setValue('Salesman');
  sheet.getRange('B3').setValue('Trend');
  sheet.getRange('A3:B3').setBackground('#f8f9fa').setFontWeight('bold');

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 110);
}

/* ==================== COMPARISON UPDATERS ==================== */

function updateAllComparisons(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  try {
    updateBankComparison(ss);
    updateAdvanceComparison(ss);
    updateSuspenseComparison(ss);
    updateOutstandingComparison(ss);
    updateDashboard(ss);
  } catch (error) {
    console.error('Error in updateAllComparisons:', error);
    throw error;
  }
}

function updateBankComparison(ss) {
  const sheet = ss.getSheetByName('Banks_Comparison');
  if (!sheet) return;

  try {
    sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearDataValidations();

    sheet.getRange('A3').setValue('Metric').setBackground('#f8f9fa').setFontWeight('bold');
    const metrics = ['Opening Balance','Total Received','Bank Charges','Total Payments','Closing Balance','Net Cash Flow','Month-over-Month %'];
    sheet.getRange(4, 1, metrics.length, 1).setValues(metrics.map(m => [m]));
    sheet.getRange('A3:A10').setFontWeight('bold');

    const list = getMonthlySheets_(ss,'Banks');
    const lastCol = sheet.getLastColumn();
    if (lastCol > 1) sheet.deleteColumns(2, lastCol - 1);

    list.forEach(({sheet:src, monthYear}, idx) => {
      const col = idx + 2;
      sheet.insertColumnAfter(col - 1);
      sheet.getRange(3, col).setValue(monthYear).setBackground('#f8f9fa').setFontWeight('bold');
      sheet.getRange(4, col).setFormula(`=IFERROR('${src.getName()}'!B9,0)`);
      sheet.getRange(5, col).setFormula(`=IFERROR('${src.getName()}'!C9,0)`);
      sheet.getRange(6, col).setFormula(`=IFERROR('${src.getName()}'!D9,0)`);
      sheet.getRange(7, col).setFormula(`=IFERROR('${src.getName()}'!E9,0)`);
      sheet.getRange(8, col).setFormula(`=IFERROR('${src.getName()}'!F17,0)`);
      sheet.getRange(9, col).setFormula(`=${colToA1_(col)}5-${colToA1_(col)}7`);
      if (idx > 0) sheet.getRange(10, col).setFormula(
        `=IFERROR((${colToA1_(col)}8-${colToA1_(col-1)}8)/${colToA1_(col-1)}8,0)`
      ).setNumberFormat('#,##0.00%');
      sheet.setColumnWidth(col, 120);
    });

    if (list.length) sheet.getRange(4, 2, 6, list.length).setNumberFormat('#,##0.00');
  } catch (error) {
    console.error('Error updating bank comparison:', error);
  }
}

function updateAdvanceComparison(ss) {
  const sheet = ss.getSheetByName('Advances_Comparison');
  if (!sheet) return;

  sheet.getDataRange().clearDataValidations();

  const list = getMonthlySheets_(ss,'Advances');
  const lastCol = sheet.getLastColumn();
  if (lastCol > 1) sheet.deleteColumns(2, lastCol - 1);

  list.forEach(({sheet:src, monthYear}, idx) => {
    const col = idx + 2;
    sheet.insertColumnAfter(col - 1);
    sheet.getRange(3, col).setValue(monthYear).setBackground('#f8f9fa').setFontWeight('bold');

    sheet.getRange(4, col).setFormula(`=IFERROR('${src.getName()}'!B3,0)`);
    sheet.getRange(5, col).setFormula(`=IFERROR(SUM('${src.getName()}'!E7:E1000),0)`);
    sheet.getRange(6, col).setFormula(`=IFERROR(SUM('${src.getName()}'!F7:F1000),0)`);
    sheet.getRange(7, col).setFormula(`=IFERROR('${src.getName()}'!E3,0)`);
    sheet.getRange(8, col).setValue('');
    sheet.getRange(9, col).setValue('');
    sheet.setColumnWidth(col, 120);
  });

  if (list.length) sheet.getRange(4, 2, 4, list.length).setNumberFormat('#,##0.00');
}

function updateSuspenseComparison(ss) {
  const sheet = ss.getSheetByName('Suspense_Comparison');
  if (!sheet) return;

  sheet.getDataRange().clearDataValidations();

  const list = getMonthlySheets_(ss,'Suspense');
  const lastCol = sheet.getLastColumn();
  if (lastCol > 1) sheet.deleteColumns(2, lastCol - 1);

  list.forEach(({sheet:src, monthYear}, idx) => {
    const col = idx + 2;
    sheet.insertColumnAfter(col - 1);
    sheet.getRange(3, col).setValue(monthYear).setBackground('#f8f9fa').setFontWeight('bold');

    sheet.getRange(4, col).setFormula(`=IFERROR('${src.getName()}'!B3,0)`);
    sheet.getRange(5, col).setFormula(`=IFERROR(SUM('${src.getName()}'!E7:E1000),0)`);
    sheet.getRange(6, col).setFormula(`=IFERROR(SUM('${src.getName()}'!F7:F1000),0)`);
    sheet.getRange(7, col).setFormula(`=IFERROR('${src.getName()}'!E3,0)`);
    sheet.setColumnWidth(col, 120);
  });

  if (list.length) sheet.getRange(4, 2, 4, list.length).setNumberFormat('#,##0.00');
}

function updateOutstandingComparison(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Outstanding_Comparison');
  if (!sheet) return;

  const salesmen = getSalesmenList_();
  const rows = salesmen.length;

  // Clear below title
  sheet.getRange(3, 1, sheet.getMaxRows() - 2, sheet.getMaxColumns()).clear();

  // Headers
  sheet.getRange('A3').setValue('Salesman');
  sheet.getRange('B3').setValue('Trend');
  sheet.getRange('A3:B3').setBackground('#f8f9fa').setFontWeight('bold');

  // Row labels
  if (rows) sheet.getRange(4, 1, rows, 1).setValues(salesmen.map(s => [s]));
  const totalRow = 4 + rows;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold').setBackground('#fff3cd');

  // Get months
  const months = getMonthlySheets_(ss, 'Outstanding');

  // Fill month columns
  months.forEach(({sheet:src, monthYear}, idx) => {
    const col = 3 + idx;
    sheet.insertColumnAfter(col - 1);
    sheet.getRange(3, col).setValue(monthYear).setBackground('#f8f9fa').setFontWeight('bold');

    // Dynamic row detection for outstanding sheets
    const srcSalesmenCount = getSalesmenList_().length;
    const detailsStartRow = 4 + srcSalesmenCount + 2;

    const rngData = `'${src.getName()}'!$D$${detailsStartRow}:$D$${CONFIG.MAX_DATA_ROWS}`;
    const rngBal  = `'${src.getName()}'!$G$${detailsStartRow}:$G$${CONFIG.MAX_DATA_ROWS}`;

    for (let r = 0; r < rows; r++) {
      const row = 4 + r;
      sheet.getRange(row, col).setFormula(`=IFERROR(SUMIF(${rngData},$A${row},${rngBal}),0)`);
    }
    sheet.getRange(totalRow, col).setFormula(`=SUM(${colToA1_(col)}4:${colToA1_(col)}${totalRow-1})`);
    sheet.setColumnWidth(col, 120);
  });

  // Number format
  if (months.length) sheet.getRange(4, 3, rows + 1, months.length).setNumberFormat('#,##0.00');

  // Sparklines
  if (months.length) {
    const endCol = 2 + months.length;
    const endL = colToA1_(endCol);
    for (let r = 4; r <= totalRow; r++) {
      const f = `=IF(COUNTA(C${r}:${endL}${r})=0,"",SPARKLINE(C${r}:${endL}${r},{"charttype","line";"linewidth",2}))`;
      sheet.getRange(r, 2).setFormula(f);
    }
  } else {
    sheet.getRange(`B4:B${totalRow}`).clearContent();
  }

  // MoM row
  const momRow = totalRow + 1;
  sheet.getRange(momRow, 1).setValue('MoM Δ (TOTAL)').setFontWeight('bold');
  for (let i = 0; i < months.length; i++) {
    const col = 3 + i;
    if (i === 0) {
      sheet.getRange(momRow, col).setValue('');
    } else {
      const curr = `${colToA1_(col)}${totalRow}`;
      const prev = `${colToA1_(col-1)}${totalRow}`;
      const f = `=IF(${prev}="","",IFERROR(IF(${curr}=${prev},"—",IF(${curr}>${prev},"▲ +"&TEXT((${curr}-${prev})/${prev},"0.0%"),"▼ -"&TEXT(ABS((${curr}-${prev})/${prev}),"0.0%"))),""))`;
      sheet.getRange(momRow, col).setFormula(f);
    }
  }

  styleOutstandingTrends_Dynamic_(sheet, totalRow, momRow);
}

/* ==================== DASHBOARD ==================== */

function createMasterDashboard(ss) {
  const name = 'Dashboard';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('FINANCIAL DASHBOARD - CUMULATIVE SUMMARY');
  sheet.getRange('A1:H1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A1:H1').setBackground('#343a40').setFontColor('#ffffff');

  // Current Status section
  sheet.getRange('A3').setValue('CURRENT STATUS');
  sheet.getRange('A3:D3').merge().setBackground('#007bff').setFontColor('#ffffff').setFontWeight('bold');

  sheet.getRange('A4').setValue('Latest Month:');
  // Fixed: Get the actual month text, not a date serial
  sheet.getRange('B4').setFormula('=IFERROR(INDEX(Banks_Comparison!3:3,1,COUNTA(Banks_Comparison!3:3)),"N/A")');

  sheet.getRange('A5').setValue('Bank Balance:');
  sheet.getRange('B5').setFormula('=IFERROR(INDEX(Banks_Comparison!8:8,1,COUNTA(Banks_Comparison!3:3)),0)');

  sheet.getRange('A6').setValue('Total Outstanding:');
  // Fixed: Dynamically find TOTAL row in Outstanding_Comparison
  sheet.getRange('B6').setFormula('=IFERROR(SUMIF(Outstanding_Comparison!A:A,"TOTAL",INDEX(Outstanding_Comparison!3:3,1,COUNTA(Outstanding_Comparison!3:3)):INDEX(Outstanding_Comparison!1000:1000,1,COUNTA(Outstanding_Comparison!3:3))),IFERROR(INDEX(INDIRECT("Outstanding_Comparison!"&ADDRESS(MATCH(\"TOTAL\",Outstanding_Comparison!A:A,0),COUNTA(Outstanding_Comparison!3:3))),1,1),0))');

  sheet.getRange('A7').setValue('Advance Balance:');
  sheet.getRange('B7').setFormula('=IFERROR(INDEX(Advances_Comparison!7:7,1,COUNTA(Advances_Comparison!3:3)),0)');

  sheet.getRange('A8').setValue('Suspense Balance:');
  sheet.getRange('B8').setFormula('=IFERROR(INDEX(Suspense_Comparison!7:7,1,COUNTA(Suspense_Comparison!3:3)),0)');

  // YTD section
  sheet.getRange('F3').setValue('YEAR-TO-DATE SUMMARY');
  sheet.getRange('F3:H3').merge().setBackground('#28a745').setFontColor('#ffffff').setFontWeight('bold');

  sheet.getRange('F4').setValue('Total Received:');
  sheet.getRange('H4').setFormula('=SUM(Banks_Comparison!5:5)');
  sheet.getRange('F5').setValue('Total Payments:');
  sheet.getRange('H5').setFormula('=SUM(Banks_Comparison!7:7)');
  sheet.getRange('F6').setValue('Net Cash Flow:');
  sheet.getRange('H6').setFormula('=H4-H5');
  sheet.getRange('F7').setValue('Avg Outstanding:');
  // Fixed: Calculate average from TOTAL row dynamically
  sheet.getRange('H7').setFormula('=IFERROR(AVERAGE(FILTER(OFFSET(Outstanding_Comparison!A1,MATCH("TOTAL",Outstanding_Comparison!A:A,0)-1,2,1,100),OFFSET(Outstanding_Comparison!A1,MATCH("TOTAL",Outstanding_Comparison!A:A,0)-1,2,1,100)<>"")),0)');

  // KPI section
  sheet.getRange('A10').setValue('KEY PERFORMANCE INDICATORS');
  sheet.getRange('A10:H10').merge().setBackground('#6c757d').setFontColor('#ffffff').setFontWeight('bold');

  const kpis = [
    ['Months Tracked','=COUNTA(Banks_Comparison!3:3)-1'],
    ['Highest Bank Balance','=MAX(Banks_Comparison!8:8)'],
    ['Lowest Bank Balance','=MIN(IF(Banks_Comparison!8:8>0,Banks_Comparison!8:8))'],
    ['Average Monthly Revenue','=IFERROR(AVERAGE(Banks_Comparison!5:5),0)'],
    ['Total Advances Given','=SUM(Advances_Comparison!5:5)'],
    ['Total Advances Settled','=SUM(Advances_Comparison!6:6)'],
    ['Outstanding Growth Rate','=IFERROR(TEXT((INDEX(OFFSET(Outstanding_Comparison!A1,MATCH("TOTAL",Outstanding_Comparison!A:A,0)-1,0,1,100),1,COUNTA(Outstanding_Comparison!3:3))-INDEX(OFFSET(Outstanding_Comparison!A1,MATCH("TOTAL",Outstanding_Comparison!A:A,0)-1,0,1,100),1,2))/INDEX(OFFSET(Outstanding_Comparison!A1,MATCH("TOTAL",Outstanding_Comparison!A:A,0)-1,0,1,100),1,2),"0.0%"),"0%")'],
    ['Cash Position','=B5-B6']
  ];

  kpis.forEach((k, i) => {
    const r = 11 + Math.floor(i/4) * 2;
    const c = 1 + (i % 4) * 2;
    sheet.getRange(r, c).setValue(k[0]);
    sheet.getRange(r + 1, c).setFormula(k[1]).setFontWeight('bold').setFontSize(14);
    sheet.getRange(r, c, 2, 2).setBorder(true, true, true, true, false, false);
  });

  // Formatting
  sheet.getRange('B5:B8').setNumberFormat('#,##0.00');
  sheet.getRange('H4:H7').setNumberFormat('#,##0.00');
  sheet.setColumnWidths(1, 8, 100);
}

function updateDashboard(ss) {
  // Force recalculation
  SpreadsheetApp.flush();

  // The dashboard uses live formulas, so just flush is enough
  // But let's also verify the Outstanding TOTAL row reference is correct
  const dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return;

  const outComp = ss.getSheetByName('Outstanding_Comparison');
  if (!outComp) return;

  // Find the actual TOTAL row
  const data = outComp.getRange('A:A').getValues();
  let totalRowNum = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === 'TOTAL') {
      totalRowNum = i + 1;
      break;
    }
  }

  if (totalRowNum > 0) {
    // Update the Outstanding formula to use correct row
    const lastCol = outComp.getLastColumn();
    dashboard.getRange('B6').setFormula(`=IFERROR(INDEX(Outstanding_Comparison!${totalRowNum}:${totalRowNum},1,COUNTA(Outstanding_Comparison!3:3)),0)`);
  }
}

/* ==================== SETTINGS ==================== */

function createSettingsSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Settings';
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);

  // Title
  setIfEmpty_(sheet.getRange('A1'), 'SYSTEM SETTINGS & CONFIGURATION');
  sheet.getRange('A1:F1').merge().setHorizontalAlignment('center');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:F1').setBackground('#6c757d').setFontColor('#ffffff');

  // Instructions
  setIfEmpty_(sheet.getRange('A3'), 'INSTRUCTIONS FOR USE');
  sheet.getRange('A3:F3').merge().setBackground('#d1ecf1').setFontWeight('bold');

  const instructions = [
    'Add new months using the menu: Monthly Reports > Add New Month',
    'Update data in the current month sheets (Banks, Advances, Suspense, Outstanding)',
    'View comparison reports and dashboard to track trends',
    'Edit the master lists below to customize dropdowns in all sheets',
    'Use "Fix Dropdowns" menu item after editing lists to update all sheets'
  ];
  instructions.forEach((instr, i) => {
    setIfEmpty_(sheet.getRange(4 + i, 1), `${i + 1}. ${instr}`);
  });

  // Master Lists header
  setIfEmpty_(sheet.getRange('A12'), 'MASTER LISTS');
  sheet.getRange('A12:F12').merge().setBackground('#d1ecf1').setFontWeight('bold');

  // Lists
  setIfEmpty_(sheet.getRange('A14'), 'BANKS');
  sheet.getRange('A14').setFontWeight('bold').setBackground('#e8f0fe');
  seedListIfEmpty_(sheet, 'A', CONFIG.SETTINGS_LIST_START_ROW, ['Alrajhi-1097','Alrajhi-new','SNB Al Ahsa Branch','Albilad','Albilad USD']);

  setIfEmpty_(sheet.getRange('C14'), 'SALESMEN');
  sheet.getRange('C14').setFontWeight('bold').setBackground('#e8f0fe');
  seedListIfEmpty_(sheet, 'C', CONFIG.SETTINGS_LIST_START_ROW, ['Nidheesh','Jaseel','Bassam','Dhiya','Shareef','Khalid','Samir','Ashik','Arshad','company sales']);

  setIfEmpty_(sheet.getRange('E14'), 'AREAS');
  sheet.getRange('E14').setFontWeight('bold').setBackground('#e8f0fe');
  seedListIfEmpty_(sheet, 'E', CONFIG.SETTINGS_LIST_START_ROW, ['AL AHSA','RIYADH','DAMMAM','QASEEM','WHOLESALE']);

  // System info
  setIfEmpty_(sheet.getRange('A27'), 'SYSTEM INFORMATION');
  sheet.getRange('A27:F27').merge().setBackground('#d1ecf1').setFontWeight('bold');
  setIfEmpty_(sheet.getRange('A28'), 'Version:');
  setIfEmpty_(sheet.getRange('B28'), CONFIG.VERSION);
  setIfEmpty_(sheet.getRange('A29'), 'Last Updated:');
  sheet.getRange('B29').setValue(new Date()).setNumberFormat('yyyy-mm-dd');
  setIfEmpty_(sheet.getRange('A30'), 'Created By:');
  setIfEmpty_(sheet.getRange('B30'), 'Account Department');

  // Parameters
  sheet.getRange('D27').setValue('PARAMETERS').setFontWeight('bold').setBackground('#d1ecf1');
  setIfEmpty_(sheet.getRange('D32'), 'SUSPENSE Base (Jan)');
  if (!sheet.getRange('E32').getValue()) sheet.getRange('E32').setValue(-17267.64).setNumberFormat('#,##0.00');
  setIfEmpty_(sheet.getRange('D34'), 'ADVANCES Base (Jan)');
  if (!sheet.getRange('E34').getValue()) sheet.getRange('E34').setValue(163800.84).setNumberFormat('#,##0.00');

  sheet.setColumnWidths(1, 6, 150);
}

/* ==================== DROPDOWN FUNCTIONS ==================== */

function applyDropdownFromSettings_(targetRange, listColLetter, listStartRow) {
  const settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!settings) return;

  try {
    const last = findLastNonBlankRow_(settings, listColLetter, listStartRow);
    const a1 = `${listColLetter}${listStartRow}:${listColLetter}${Math.max(last, listStartRow)}`;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(settings.getRange(a1), true)
      .build();
    targetRange.setDataValidation(rule);
  } catch (error) {
    console.error('Error applying dropdown:', error);
  }
}

function repairDropdownsForAllMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Banks
    getMonthlySheets_(ss, 'Banks').forEach(({sheet}) => {
      applyDropdownFromSettings_(sheet.getRange('A4:A7'), 'A', CONFIG.SETTINGS_LIST_START_ROW);
      applyDropdownFromSettings_(sheet.getRange('A13'), 'A', CONFIG.SETTINGS_LIST_START_ROW);
    });

    // Outstanding - need to find correct detail start row for each sheet
    getMonthlySheets_(ss, 'Outstanding').forEach(({sheet}) => {
      const salesmenCount = getSalesmenList_().length;
      const detailsStartRow = 4 + salesmenCount + 2;
      applyDropdownFromSettings_(sheet.getRange(`C${detailsStartRow}:C${CONFIG.MAX_DATA_ROWS}`), 'E', CONFIG.SETTINGS_LIST_START_ROW);
      applyDropdownFromSettings_(sheet.getRange(`D${detailsStartRow}:D${CONFIG.MAX_DATA_ROWS}`), 'C', CONFIG.SETTINGS_LIST_START_ROW);
    });

    SpreadsheetApp.getUi().alert('✅ Dropdowns reapplied from Settings to all month sheets.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to repair dropdowns: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/* ==================== HELPER FUNCTIONS ==================== */

function getSalesmenList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Settings');
  if (!sh) return [];

  try {
    const last = findLastNonBlankRow_(sh, 'C', CONFIG.SETTINGS_LIST_START_ROW);
    return sh.getRange(`C${CONFIG.SETTINGS_LIST_START_ROW}:C${Math.max(CONFIG.SETTINGS_LIST_START_ROW, last)}`).getValues()
             .map(r => String(r[0]).trim())
             .filter(Boolean);
  } catch (error) {
    console.error('Error getting salesmen list:', error);
    return [];
  }
}

function getAdvancesBase_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  return sh ? Number(sh.getRange('E34').getValue() || 0) : 0;
}

function getSuspenseBase_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  return sh ? Number(sh.getRange('E32').getValue() || 0) : 0;
}

function styleOutstandingTrends_Dynamic_(sheet, totalRow, momRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return;

  try {
    const momRange = sheet.getRange(momRow, 3, 1, lastCol - 2);
    const totalRange = sheet.getRange(totalRow, 3, 1, lastCol - 2);

    momRange.clearFormat();
    totalRange.clearFormat();

    let rules = sheet.getConditionalFormatRules();
    const A1s = [momRange.getA1Notation(), totalRange.getA1Notation()];
    rules = rules.filter(r => !r.getRanges().some(rr => A1s.includes(rr.getA1Notation())));

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('▲').setFontColor('#ffffff').setBackground('#dc3545')
        .setRanges([momRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('▼').setFontColor('#ffffff').setBackground('#28a745')
        .setRanges([momRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(COLUMN()>3, INDIRECT(ADDRESS(' + totalRow + ',COLUMN()))>INDIRECT(ADDRESS(' + totalRow + ',COLUMN()-1)))')
        .setFontColor('#ffffff').setBackground('#dc3545')
        .setRanges([totalRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(COLUMN()>3, INDIRECT(ADDRESS(' + totalRow + ',COLUMN()))<INDIRECT(ADDRESS(' + totalRow + ',COLUMN()-1)))')
        .setFontColor('#ffffff').setBackground('#28a745')
        .setRanges([totalRange]).build()
    );

    sheet.setConditionalFormatRules(rules);
  } catch (error) {
    console.error('Error applying conditional formatting:', error);
  }
}

/* ==================== MENU FUNCTIONS ==================== */

function viewBanksComparison() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banks_Comparison');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function viewAdvancesComparison() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advances_Comparison');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function viewSuspenseComparison() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Suspense_Comparison');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function viewOutstandingComparison() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Outstanding_Comparison');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function viewDashboard() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function openSettings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function toggleSheetVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const resp = ui.alert('Show/Hide Sheets',
      'YES = Show all hidden sheets\nNO = Hide all but latest month',
      ui.ButtonSet.YES_NO_CANCEL);

    if (resp === ui.Button.YES) {
      ss.getSheets().forEach(sh => {
        if (sh.isSheetHidden()) sh.showSheet();
      });
      ui.alert('✅ All hidden sheets are now visible');
    } else if (resp === ui.Button.NO) {
      const latest = getLatestMonthYear_(ss, 'Banks') ||
                    getLatestMonthYear_(ss, 'Advances') ||
                    getLatestMonthYear_(ss, 'Outstanding') ||
                    getLatestMonthYear_(ss, 'Suspense');
      if (latest) {
        ss.getSheets().forEach(sh => {
          const n = sh.getName();
          if (/_/.test(n) && !n.endsWith(latest) &&
              !/Comparison|Dashboard|Settings/.test(n)) {
            sh.hideSheet();
          }
        });
        ui.alert('✅ Previous month sheets have been hidden');
      }
    }
  } catch (error) {
    ui.alert('❌ Error', 'Could not toggle sheet visibility: ' + error.message, ui.ButtonSet.OK);
  }
}

function refreshAllCalculations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    updateAllComparisons(ss);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ All calculations refreshed successfully!');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error',
      'Failed to refresh: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function updateCurrentMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const latest = getLatestMonthYear_(ss, 'Banks') ||
                  getLatestMonthYear_(ss, 'Advances') ||
                  getLatestMonthYear_(ss, 'Suspense') ||
                  getLatestMonthYear_(ss, 'Outstanding');

    if (!latest) {
      ui.alert('❌ No monthly sheets found', 'Please add a month first.', ui.ButtonSet.OK);
      return;
    }

    const r = ui.prompt('Update Current Month',
      `Current month: ${latest}\nChoose:\n1. Banks\n2. Advances\n3. Suspense\n4. Outstanding`,
      ui.ButtonSet.OK_CANCEL);

    if (r.getSelectedButton() !== ui.Button.OK) return;

    const choice = r.getResponseText().trim();
    const sheets = {
      '1': `Banks_${latest}`,
      '2': `Advances_${latest}`,
      '3': `Suspense_${latest}`,
      '4': `Outstanding_${latest}`
    };

    const sheetName = sheets[choice];
    if (sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.setActiveSheet(sheet);
        ui.alert('✅ Sheet opened', `Now editing: ${sheetName}`, ui.ButtonSet.OK);
      } else {
        ui.alert('❌ Sheet not found', `${sheetName} does not exist.`, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('❌ Invalid choice', 'Please enter 1, 2, 3, or 4.', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('❌ Error', 'Could not update: ' + error.message, ui.ButtonSet.OK);
  }
}

function emailMonthlyReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('Email Monthly Report', 'Enter recipient email address:', ui.ButtonSet.OK_CANCEL);

    if (r.getSelectedButton() !== ui.Button.OK) return;

    const email = r.getResponseText().trim();

    // Basic email validation
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      ui.alert('❌ Invalid email address format');
      return;
    }

    const pdf = ss.getAs('application/pdf').setName('Monthly_Report.pdf');
    MailApp.sendEmail({
      to: email,
      subject: 'Monthly Financial Report - ' + new Date().toLocaleDateString(),
      body: 'Please find attached the monthly financial report.\n\n• Bank Summary\n• Advance Reconciliation\n• Suspense Reconciliation\n• Customer Outstanding\n• Monthly Comparisons\n\nBest regards,\nAccount Department',
      attachments: [pdf]
    });
    ui.alert('✅ Report sent successfully.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to send email: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function exportToPDF() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pdf = ss.getAs('application/pdf').setName('Monthly_Report_' + new Date().getTime() + '.pdf');
    const file = DriveApp.createFile(pdf);
    const html = HtmlService
      .createHtmlOutput('<p>PDF created successfully!</p>' +
                        '<p><a href="' + file.getDownloadUrl() + '" target="_blank">Download PDF</a></p>' +
                        '<p>The file has been saved to your Google Drive.</p>')
      .setWidth(400).setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(html, 'Export to PDF');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to export PDF: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Monthly Accounting Report System - Help</h2>

      <h3>Getting Started:</h3>
      <ul>
        <li><strong>Initialize:</strong> Run initializeSystem() once to set up all sheets</li>
        <li><strong>Add Months:</strong> Use "Add New Month" from the menu</li>
        <li><strong>Enter Data:</strong> Update monthly sheets with actual figures</li>
        <li><strong>View Reports:</strong> Use comparison sheets and dashboard</li>
      </ul>

      <h3>Sheet Types:</h3>
      <ul>
        <li><strong>Banks_[Month]:</strong> Bank account summaries and balances</li>
        <li><strong>Advances_[Month]:</strong> Staff advance tracking</li>
        <li><strong>Suspense_[Month]:</strong> Unidentified transactions</li>
        <li><strong>Outstanding_[Month]:</strong> Customer receivables</li>
        <li><strong>Comparison Sheets:</strong> Month-to-month trend analysis</li>
        <li><strong>Dashboard:</strong> Executive summary and KPIs</li>
        <li><strong>Settings:</strong> Master lists and configuration</li>
      </ul>

      <h3>Key Features:</h3>
      <ul>
        <li><strong>Automatic Carry Forward:</strong> Balances transfer to new months</li>
        <li><strong>Dynamic Dropdowns:</strong> Based on Settings master lists</li>
        <li><strong>Sparkline Trends:</strong> Visual trend indicators</li>
        <li><strong>Conditional Formatting:</strong> Color-coded performance</li>
      </ul>

      <h3>Tips:</h3>
      <ul>
        <li>Edit Settings sheet to customize dropdown options</li>
        <li>Use "Fix Dropdowns" after editing Settings</li>
        <li>Previous month sheets auto-hide when adding new months</li>
        <li>Dashboard formulas update automatically</li>
      </ul>

      <h3>Troubleshooting:</h3>
      <ul>
        <li><strong>Missing dropdowns:</strong> Run "Fix Dropdowns"</li>
        <li><strong>Wrong calculations:</strong> Run "Refresh All Calculations"</li>
        <li><strong>Can't see sheets:</strong> Use "Show/Hide Sheets"</li>
      </ul>

      <p><strong>Version:</strong> ${CONFIG.VERSION} | <strong>Support:</strong> Account Department</p>
    </div>
  `).setWidth(500).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Help - Monthly Accounting System');
}

/* ==================== MIGRATION FUNCTIONS ==================== */

function migrateAdvancesClosingFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const advBase = getAdvancesBase_();

  try {
    getMonthlySheets_(ss, 'Advances').forEach(({sheet}) => {
      sheet.getRange('E3').setFormula(`=${advBase}+SUM(E7:E1000)`)
           .setNumberFormat('#,##0.00').setFontWeight('bold');
    });
    SpreadsheetApp.getUi().alert('✅ Advances closing updated to fixed base + monthly sum.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Migration failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function migrateSuspenseClosingFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const susBase = getSuspenseBase_();

  try {
    getMonthlySheets_(ss, 'Suspense').forEach(({sheet}) => {
      sheet.getRange('E3').setFormula(`=${susBase}+SUM(E7:E1000)-SUM(F7:F1000)`)
           .setNumberFormat('#,##0.00').setFontWeight('bold');
    });
    SpreadsheetApp.getUi().alert('✅ Suspense closing updated to fixed base + monthly sum.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Migration failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function setupExistingSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    createMasterComparisonSheets(ss);
    createMasterDashboard(ss);
    createSettingsSheet(ss);
    updateAllComparisons(ss);
    createCustomMenu();
    SpreadsheetApp.getUi().alert('✅ Setup Complete.\nDashboard, Settings, and Comparison sheets created/refreshed.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Setup failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/* ==================== AUTO-REPAIR ON EDIT ==================== */

function shouldRunDropdownRepair_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const last = Number(props.getProperty('last_dropdown_fix_ts') || 0);
    const now = Date.now();
    if (now - last < CONFIG.DROPDOWN_REPAIR_COOLDOWN_MS) return false;
    props.setProperty('last_dropdown_fix_ts', String(now));
    return true;
  } catch (error) {
    console.error('Error checking dropdown repair cooldown:', error);
    return false;
  }
}

function onEdit(e) {
  try {
    const sh = e && e.range && e.range.getSheet();
    if (!sh || sh.getName() !== 'Settings') return;

    const col = e.range.getColumn();
    if ((col === 1 || col === 3 || col === 5) && shouldRunDropdownRepair_()) {
      repairDropdownsForAllMonths();
    }
  } catch (err) {
    console.log('onEdit dropdown auto-fix skipped:', err);
  }
}

/* ==================== FIX EXISTING DATA ==================== */

/**
 * Run this once to fix the Outstanding_Comparison TOTAL row issue
 * and update Dashboard formulas
 */
function fixOutstandingAndDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Rebuild Outstanding_Comparison with correct structure
    updateOutstandingComparison(ss);

    // Fix Dashboard formulas to reference TOTAL row dynamically
    const dashboard = ss.getSheetByName('Dashboard');
    const outComp = ss.getSheetByName('Outstanding_Comparison');

    if (dashboard && outComp) {
      // Find TOTAL row
      const data = outComp.getRange('A:A').getValues();
      let totalRowNum = -1;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).toUpperCase() === 'TOTAL') {
          totalRowNum = i + 1;
          break;
        }
      }

      if (totalRowNum > 0) {
        // Fix Total Outstanding formula
        dashboard.getRange('B6').setFormula(`=IFERROR(INDEX(Outstanding_Comparison!${totalRowNum}:${totalRowNum},1,COUNTA(Outstanding_Comparison!3:3)),0)`);

        // Fix Avg Outstanding formula
        dashboard.getRange('H7').setFormula(`=IFERROR(AVERAGE(FILTER(Outstanding_Comparison!${totalRowNum}:${totalRowNum},Outstanding_Comparison!${totalRowNum}:${totalRowNum}<>"")),0)`);
      }
    }

    SpreadsheetApp.flush();
    ui.alert('✅ Outstanding comparison and Dashboard fixed!');
  } catch (error) {
    ui.alert('❌ Error', 'Fix failed: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Fix Outstanding summary section for a specific month
 */
function fixOutstandingSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get latest month dynamically instead of hardcoding
  const monthYear = getLatestMonthYear_(ss, 'Outstanding');

  if (!monthYear) {
    ui.alert('No Outstanding sheets found');
    return;
  }

  const sheet = ss.getSheetByName(`Outstanding_${monthYear}`);
  if (!sheet) {
    ui.alert('Sheet not found: Outstanding_' + monthYear);
    return;
  }

  const salesmen = getSalesmenList_();
  if (!salesmen.length) {
    ui.alert('No salesmen found in Settings');
    return;
  }

  // Clear and rebuild summary section
  const detailsStartRow = 4 + salesmen.length + 2;
  sheet.getRange('A4:D' + (salesmen.length + 10)).clear();

  salesmen.forEach((salesman, i) => {
    const row = 4 + i;
    sheet.getRange(row, 1).setValue(salesman);
    sheet.getRange(row, 2).setFormula(`=SUMIF($D$${detailsStartRow}:$D$${CONFIG.MAX_DATA_ROWS},A${row},$G$${detailsStartRow}:$G$${CONFIG.MAX_DATA_ROWS})`);
    sheet.getRange(row, 3).setFormula(`=COUNTIF($D$${detailsStartRow}:$D$${CONFIG.MAX_DATA_ROWS},A${row})`);
    sheet.getRange(row, 4).setFormula(`=IFERROR(B${row}/C${row},0)`);
  });

  // Add TOTAL row
  const totalRow = 4 + salesmen.length;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold').setBackground('#fff3cd');
  sheet.getRange(totalRow, 2).setFormula(`=SUM(B4:B${totalRow-1})`);
  sheet.getRange(totalRow, 3).setFormula(`=SUM(C4:C${totalRow-1})`);
  sheet.getRange(totalRow, 4).setFormula(`=IFERROR(B${totalRow}/C${totalRow},0)`);

  ui.alert(`Outstanding summary section fixed for ${monthYear}!`);
}

/* ==================== WEB APP FUNCTIONS ==================== */

/**
 * Serves the web dashboard
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Financial Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include HTML files (for CSS and JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Opens the dashboard in a sidebar
 */
function openDashboardSidebar() {
  const html = HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Financial Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the dashboard in a modal dialog
 */
function openDashboardModal() {
  const html = HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Financial Dashboard');
}

/**
 * Get all dashboard data for the frontend
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const months = getMonthlySheets_(ss, 'Banks').map(m => m.monthYear);
    const latestMonth = months.length ? months[months.length - 1] : null;

    return {
      latestMonth: latestMonth,
      months: months,
      kpi: getKPIData_(ss, latestMonth),
      ytd: getYTDData_(ss),
      banks: getBanksData_(ss),
      outstanding: getOutstandingData_(ss),
      advances: getAdvancesData_(ss),
      suspense: getSuspenseData_(ss),
      bankAccounts: getBankAccountsData_(ss, latestMonth)
    };
  } catch (error) {
    console.error('Error getting dashboard data:', error);
    throw error;
  }
}

/**
 * Get dashboard data for a specific month
 */
function getDashboardDataForMonth(monthYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    return {
      latestMonth: monthYear,
      months: getMonthlySheets_(ss, 'Banks').map(m => m.monthYear),
      kpi: getKPIData_(ss, monthYear),
      ytd: getYTDData_(ss),
      banks: getBanksData_(ss),
      outstanding: getOutstandingData_(ss),
      advances: getAdvancesData_(ss),
      suspense: getSuspenseData_(ss),
      bankAccounts: getBankAccountsData_(ss, monthYear)
    };
  } catch (error) {
    console.error('Error getting month data:', error);
    throw error;
  }
}

/**
 * Get KPI data
 */
function getKPIData_(ss, monthYear) {
  const banksComp = ss.getSheetByName('Banks_Comparison');
  const outComp = ss.getSheetByName('Outstanding_Comparison');
  const advComp = ss.getSheetByName('Advances_Comparison');
  const susComp = ss.getSheetByName('Suspense_Comparison');

  const months = getMonthlySheets_(ss, 'Banks');
  const monthIdx = months.findIndex(m => m.monthYear === monthYear);
  const col = monthIdx >= 0 ? monthIdx + 2 : months.length + 1;
  const prevCol = col - 1;

  // Get bank balance
  let bankBalance = 0;
  let bankPrev = 0;
  if (banksComp && col > 1) {
    bankBalance = banksComp.getRange(8, col).getValue() || 0;
    if (prevCol > 1) bankPrev = banksComp.getRange(8, prevCol).getValue() || 0;
  }

  // Get outstanding total (find TOTAL row)
  let outstanding = 0;
  let outPrev = 0;
  if (outComp) {
    const outData = outComp.getRange('A:A').getValues();
    for (let i = 0; i < outData.length; i++) {
      if (String(outData[i][0]).toUpperCase() === 'TOTAL') {
        outstanding = outComp.getRange(i + 1, col).getValue() || 0;
        if (prevCol > 1) outPrev = outComp.getRange(i + 1, prevCol).getValue() || 0;
        break;
      }
    }
  }

  // Get advances closing
  let advances = 0;
  let advPrev = 0;
  if (advComp && col > 1) {
    advances = advComp.getRange(7, col).getValue() || 0;
    if (prevCol > 1) advPrev = advComp.getRange(7, prevCol).getValue() || 0;
  }

  // Get suspense closing
  let suspense = 0;
  let susPrev = 0;
  if (susComp && col > 1) {
    suspense = susComp.getRange(7, col).getValue() || 0;
    if (prevCol > 1) susPrev = susComp.getRange(7, prevCol).getValue() || 0;
  }

  return {
    bankBalance: bankBalance,
    bankChange: bankPrev ? ((bankBalance - bankPrev) / Math.abs(bankPrev)) * 100 : 0,
    outstanding: outstanding,
    outstandingChange: outPrev ? ((outstanding - outPrev) / Math.abs(outPrev)) * 100 : 0,
    advances: advances,
    advancesChange: advPrev ? ((advances - advPrev) / Math.abs(advPrev)) * 100 : 0,
    suspense: suspense,
    suspenseChange: susPrev ? ((suspense - susPrev) / Math.abs(susPrev)) * 100 : 0
  };
}

/**
 * Get YTD summary data
 */
function getYTDData_(ss) {
  const banksComp = ss.getSheetByName('Banks_Comparison');

  let received = 0;
  let payments = 0;
  let monthCount = 0;

  if (banksComp) {
    const lastCol = banksComp.getLastColumn();
    for (let col = 2; col <= lastCol; col++) {
      const rec = banksComp.getRange(5, col).getValue();
      const pay = banksComp.getRange(7, col).getValue();
      if (rec || pay) {
        received += Number(rec) || 0;
        payments += Number(pay) || 0;
        monthCount++;
      }
    }
  }

  return {
    received: received,
    payments: payments,
    netFlow: received - payments,
    months: monthCount
  };
}

/**
 * Get banks comparison data for charts
 */
function getBanksData_(ss) {
  const banksComp = ss.getSheetByName('Banks_Comparison');
  const labels = [];
  const balance = [];
  const received = [];
  const payments = [];

  if (banksComp) {
    const lastCol = banksComp.getLastColumn();
    for (let col = 2; col <= lastCol; col++) {
      const month = banksComp.getRange(3, col).getValue();
      if (month) {
        labels.push(formatMonthLabel_(month));
        balance.push(banksComp.getRange(8, col).getValue() || 0);
        received.push(banksComp.getRange(5, col).getValue() || 0);
        payments.push(banksComp.getRange(7, col).getValue() || 0);
      }
    }
  }

  return { labels, balance, received, payments };
}

/**
 * Get outstanding comparison data for charts
 */
function getOutstandingData_(ss) {
  const outComp = ss.getSheetByName('Outstanding_Comparison');
  const labels = [];
  const total = [];
  const salesmen = [];

  if (outComp) {
    const lastCol = outComp.getLastColumn();
    const lastRow = outComp.getLastRow();

    // Get month labels
    for (let col = 3; col <= lastCol; col++) {
      const month = outComp.getRange(3, col).getValue();
      if (month) labels.push(formatMonthLabel_(month));
    }

    // Get salesman data
    let totalRow = -1;
    for (let row = 4; row <= lastRow; row++) {
      const name = outComp.getRange(row, 1).getValue();
      if (!name) continue;

      if (String(name).toUpperCase() === 'TOTAL') {
        totalRow = row;
        // Get total values
        for (let col = 3; col <= lastCol; col++) {
          const val = outComp.getRange(row, col).getValue();
          if (val !== '') total.push(Number(val) || 0);
        }
      } else if (String(name).toUpperCase() !== 'MOM Δ (TOTAL)') {
        // Get salesman's latest value and calculate trend
        const lastVal = outComp.getRange(row, lastCol).getValue() || 0;
        const prevVal = lastCol > 3 ? (outComp.getRange(row, lastCol - 1).getValue() || 0) : 0;
        const trend = prevVal ? ((lastVal - prevVal) / Math.abs(prevVal)) * 100 : 0;

        salesmen.push({
          name: name,
          value: lastVal,
          trend: trend
        });
      }
    }

    // Sort by value descending
    salesmen.sort((a, b) => b.value - a.value);
  }

  return { labels, total, salesmen };
}

/**
 * Get advances comparison data for charts
 */
function getAdvancesData_(ss) {
  const advComp = ss.getSheetByName('Advances_Comparison');
  const labels = [];
  const opening = [];
  const given = [];
  const settled = [];
  const closing = [];

  if (advComp) {
    const lastCol = advComp.getLastColumn();
    for (let col = 2; col <= lastCol; col++) {
      const month = advComp.getRange(3, col).getValue();
      if (month) {
        labels.push(formatMonthLabel_(month));
        opening.push(advComp.getRange(4, col).getValue() || 0);
        given.push(advComp.getRange(5, col).getValue() || 0);
        settled.push(advComp.getRange(6, col).getValue() || 0);
        closing.push(advComp.getRange(7, col).getValue() || 0);
      }
    }
  }

  return { labels, opening, given, settled, closing };
}

/**
 * Get suspense comparison data for charts
 */
function getSuspenseData_(ss) {
  const susComp = ss.getSheetByName('Suspense_Comparison');
  const labels = [];
  const balance = [];

  if (susComp) {
    const lastCol = susComp.getLastColumn();
    for (let col = 2; col <= lastCol; col++) {
      const month = susComp.getRange(3, col).getValue();
      if (month) {
        labels.push(formatMonthLabel_(month));
        balance.push(susComp.getRange(7, col).getValue() || 0);
      }
    }
  }

  return { labels, balance };
}

/**
 * Get bank accounts data for the current month
 */
function getBankAccountsData_(ss, monthYear) {
  const bankSheet = ss.getSheetByName(`Banks_${monthYear}`);
  const accounts = [];

  if (bankSheet) {
    // SAR accounts (rows 4-7)
    for (let row = 4; row <= 7; row++) {
      const name = bankSheet.getRange(row, 1).getValue();
      const balance = bankSheet.getRange(row, 6).getValue();
      if (name) {
        accounts.push({
          name: name,
          balance: balance || 0,
          change: 0 // Would need previous month data to calculate
        });
      }
    }

    // USD accounts (row 13)
    const usdName = bankSheet.getRange(13, 1).getValue();
    const usdBalance = bankSheet.getRange(13, 6).getValue();
    if (usdName) {
      accounts.push({
        name: usdName,
        balance: usdBalance || 0,
        change: 0
      });
    }
  }

  return accounts;
}

/**
 * Format month label for charts
 */
function formatMonthLabel_(monthValue) {
  if (monthValue instanceof Date) {
    const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    return months[monthValue.getMonth()];
  }
  return String(monthValue).split('-')[0] || monthValue;
}

/**
 * Send report via email (for web app)
 */
function sendReportEmail(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pdf = ss.getAs('application/pdf').setName('Monthly_Report.pdf');

    MailApp.sendEmail({
      to: email,
      subject: 'Monthly Financial Report - ' + new Date().toLocaleDateString(),
      body: 'Please find attached the monthly financial report.\n\nBest regards,\nAccount Department',
      attachments: [pdf]
    });

    return true;
  } catch (error) {
    throw new Error('Failed to send email: ' + error.message);
  }
}

/**
 * Generate specific report type
 */
function generateReport(type, monthYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  switch(type) {
    case 'monthly':
      // Navigate to the month's sheets
      const bankSheet = ss.getSheetByName(`Banks_${monthYear}`);
      if (bankSheet) ss.setActiveSheet(bankSheet);
      break;
    case 'comparison':
      const compSheet = ss.getSheetByName('Banks_Comparison');
      if (compSheet) ss.setActiveSheet(compSheet);
      break;
    case 'outstanding':
      const outSheet = ss.getSheetByName('Outstanding_Comparison');
      if (outSheet) ss.setActiveSheet(outSheet);
      break;
  }

  return true;
}
