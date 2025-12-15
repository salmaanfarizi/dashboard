/**
 * =====================================================
 * MONTHLY ACCOUNTING REPORT SYSTEM - UPDATED VERSION
 * =====================================================
 *
 * This version includes:
 * - Fixed Advances closing balance retrieval
 * - Fixed Suspense balance retrieval
 * - JSONP support for Netlify dashboard
 * - Improved error handling
 */

/* ==================== CONFIGURATION ==================== */
const CONFIG = {
  MONTHS: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'],
  CURRENT_YEAR: new Date().getFullYear(),
  SHEET_PREFIXES: {
    BANKS: 'Banks_',
    ADVANCES: 'Advances_',
    SUSPENSE: 'Suspense_',
    OUTSTANDING: 'Outstanding_'
  }
};

/* ==================== WEB APP API ==================== */

/**
 * Main web app entry point - serves HTML dashboard or JSON API
 */
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action || '';
  var callback = params.callback || '';

  // API endpoint - return JSON/JSONP data
  if (action === 'getData') {
    try {
      var data = getDashboardData();
      return jsonResponse_(data, callback);
    } catch (err) {
      Logger.log('API Error: ' + err.message);
      return jsonResponse_({ error: err.message }, callback);
    }
  }

  if (action === 'getMonth') {
    try {
      var monthYear = params.month || '';
      var data = getDashboardDataForMonth(monthYear);
      return jsonResponse_(data, callback);
    } catch (err) {
      return jsonResponse_({ error: err.message }, callback);
    }
  }

  // Default response
  return ContentService.createTextOutput('API is running. Use ?action=getData to get data.')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Returns JSON or JSONP response
 */
function jsonResponse_(data, callback) {
  var jsonStr = JSON.stringify(data);

  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonStr + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==================== MAIN DATA FUNCTIONS ==================== */

/**
 * Get all dashboard data
 */
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var months = getMonthlySheets_(ss, 'Banks').map(function(m) { return m.monthYear; });
  var latestMonth = months.length ? months[months.length - 1] : null;

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
}

/**
 * Get dashboard data for specific month
 */
function getDashboardDataForMonth(monthYear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    latestMonth: monthYear,
    months: getMonthlySheets_(ss, 'Banks').map(function(m) { return m.monthYear; }),
    kpi: getKPIData_(ss, monthYear),
    ytd: getYTDData_(ss),
    banks: getBanksData_(ss),
    outstanding: getOutstandingData_(ss),
    advances: getAdvancesData_(ss),
    suspense: getSuspenseData_(ss),
    bankAccounts: getBankAccountsData_(ss, monthYear)
  };
}

/* ==================== KPI DATA ==================== */

/**
 * Get KPI summary data
 */
function getKPIData_(ss, monthYear) {
  var months = getMonthlySheets_(ss, 'Banks');
  var monthIdx = -1;
  for (var i = 0; i < months.length; i++) {
    if (months[i].monthYear === monthYear) {
      monthIdx = i;
      break;
    }
  }

  // Bank Balance from Banks_Comparison
  var bankBalance = 0, bankPrev = 0;
  var banksComp = ss.getSheetByName('Banks_Comparison');
  if (banksComp) {
    var col = monthIdx >= 0 ? monthIdx + 2 : months.length + 1;
    bankBalance = getNumericValue_(banksComp, 8, col);
    if (col > 2) bankPrev = getNumericValue_(banksComp, 8, col - 1);
  }

  // Outstanding Total
  var outstanding = 0, outPrev = 0;
  var outComp = ss.getSheetByName('Outstanding_Comparison');
  if (outComp) {
    var totalRow = findRowByValue_(outComp, 'A:A', 'TOTAL');
    if (totalRow > 0) {
      var col = monthIdx >= 0 ? monthIdx + 2 : months.length + 1;
      outstanding = getNumericValue_(outComp, totalRow, col);
      if (col > 2) outPrev = getNumericValue_(outComp, totalRow, col - 1);
    }
  }

  // Advances Closing Balance - FIXED
  var advances = 0, advPrev = 0;
  var advData = getAdvancesClosingBalance_(ss, monthYear);
  advances = advData.current;
  advPrev = advData.previous;

  // Suspense Closing Balance - FIXED
  var suspense = 0, susPrev = 0;
  var susData = getSuspenseClosingBalance_(ss, monthYear);
  suspense = susData.current;
  susPrev = susData.previous;

  return {
    bankBalance: bankBalance,
    bankChange: calcChange_(bankBalance, bankPrev),
    outstanding: outstanding,
    outstandingChange: calcChange_(outstanding, outPrev),
    advances: advances,
    advancesChange: calcChange_(advances, advPrev),
    suspense: suspense,
    suspenseChange: calcChange_(suspense, susPrev)
  };
}

/**
 * Get Advances closing balance - reads directly from monthly sheet
 */
function getAdvancesClosingBalance_(ss, monthYear) {
  var result = { current: 0, previous: 0 };

  // Try to get from monthly Advances sheet first
  var advSheet = ss.getSheetByName('Advances_' + monthYear);
  if (advSheet) {
    // E3 contains the closing balance formula
    result.current = getNumericValue_(advSheet, 3, 5); // E3
  }

  // If no data, try comparison sheet
  if (result.current === 0) {
    var advComp = ss.getSheetByName('Advances_Comparison');
    if (advComp) {
      var months = getMonthlySheets_(ss, 'Advances');
      var col = findMonthColumn_(advComp, monthYear, months);
      if (col > 1) {
        result.current = getNumericValue_(advComp, 7, col); // Row 7 = Closing Balance
        if (col > 2) result.previous = getNumericValue_(advComp, 7, col - 1);
      }
    }
  }

  // Get previous month if not yet set
  if (result.previous === 0 && monthYear) {
    var prevMonth = getPreviousMonth_(monthYear);
    var prevSheet = ss.getSheetByName('Advances_' + prevMonth);
    if (prevSheet) {
      result.previous = getNumericValue_(prevSheet, 3, 5); // E3
    }
  }

  return result;
}

/**
 * Get Suspense closing balance - reads directly from monthly sheet
 */
function getSuspenseClosingBalance_(ss, monthYear) {
  var result = { current: 0, previous: 0 };

  // Try to get from monthly Suspense sheet first
  var susSheet = ss.getSheetByName('Suspense_' + monthYear);
  if (susSheet) {
    // E3 contains the closing balance formula
    result.current = getNumericValue_(susSheet, 3, 5); // E3
  }

  // If no data, try comparison sheet
  if (result.current === 0) {
    var susComp = ss.getSheetByName('Suspense_Comparison');
    if (susComp) {
      var months = getMonthlySheets_(ss, 'Suspense');
      var col = findMonthColumn_(susComp, monthYear, months);
      if (col > 1) {
        result.current = getNumericValue_(susComp, 7, col); // Row 7 = Closing Balance
        if (col > 2) result.previous = getNumericValue_(susComp, 7, col - 1);
      }
    }
  }

  // Get previous month if not yet set
  if (result.previous === 0 && monthYear) {
    var prevMonth = getPreviousMonth_(monthYear);
    var prevSheet = ss.getSheetByName('Suspense_' + prevMonth);
    if (prevSheet) {
      result.previous = getNumericValue_(prevSheet, 3, 5); // E3
    }
  }

  return result;
}

/* ==================== YTD DATA ==================== */

/**
 * Get Year-to-Date summary
 */
function getYTDData_(ss) {
  var banksComp = ss.getSheetByName('Banks_Comparison');
  var received = 0, payments = 0, monthCount = 0;

  if (banksComp) {
    var lastCol = banksComp.getLastColumn();
    for (var col = 2; col <= lastCol; col++) {
      var rec = getNumericValue_(banksComp, 5, col);
      var pay = getNumericValue_(banksComp, 7, col);
      if (rec > 0 || pay > 0) {
        received += rec;
        payments += pay;
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

/* ==================== BANKS DATA ==================== */

/**
 * Get banks comparison data for charts
 */
function getBanksData_(ss) {
  var banksComp = ss.getSheetByName('Banks_Comparison');
  var labels = [], balance = [], received = [], payments = [];

  if (banksComp) {
    var lastCol = banksComp.getLastColumn();
    for (var col = 2; col <= lastCol; col++) {
      var month = banksComp.getRange(3, col).getValue();
      if (month) {
        labels.push(formatMonthLabel_(month));
        balance.push(getNumericValue_(banksComp, 8, col));
        received.push(getNumericValue_(banksComp, 5, col));
        payments.push(getNumericValue_(banksComp, 7, col));
      }
    }
  }

  return { labels: labels, balance: balance, received: received, payments: payments };
}

/**
 * Get bank accounts for current month
 */
function getBankAccountsData_(ss, monthYear) {
  var accounts = [];
  var bankSheet = ss.getSheetByName('Banks_' + monthYear);

  if (bankSheet) {
    // SAR accounts (rows 4-7 typically)
    for (var row = 4; row <= 10; row++) {
      var name = bankSheet.getRange(row, 1).getValue();
      var balance = getNumericValue_(bankSheet, row, 6);
      if (name && typeof name === 'string' && name.trim() !== '') {
        accounts.push({ name: name, balance: balance, change: 0 });
      }
    }

    // USD accounts (check row 13)
    var usdName = bankSheet.getRange(13, 1).getValue();
    var usdBalance = getNumericValue_(bankSheet, 13, 6);
    if (usdName && typeof usdName === 'string' && usdName.trim() !== '') {
      accounts.push({ name: usdName, balance: usdBalance, change: 0 });
    }
  }

  return accounts;
}

/* ==================== OUTSTANDING DATA ==================== */

/**
 * Get outstanding comparison data
 */
function getOutstandingData_(ss) {
  var outComp = ss.getSheetByName('Outstanding_Comparison');
  var labels = [], total = [], salesmen = [];

  if (outComp) {
    var lastCol = outComp.getLastColumn();
    var lastRow = outComp.getLastRow();

    // Get monthly totals
    var totalRow = findRowByValue_(outComp, 'A:A', 'TOTAL');

    for (var col = 2; col <= lastCol; col++) {
      var month = outComp.getRange(3, col).getValue();
      if (month) {
        labels.push(formatMonthLabel_(month));
        total.push(totalRow > 0 ? getNumericValue_(outComp, totalRow, col) : 0);
      }
    }

    // Get salesmen data from latest column
    var latestCol = lastCol;
    var prevCol = lastCol > 2 ? lastCol - 1 : 2;

    for (var row = 4; row < lastRow; row++) {
      var name = outComp.getRange(row, 1).getValue();
      if (name && String(name).toUpperCase() !== 'TOTAL') {
        var value = getNumericValue_(outComp, row, latestCol);
        var prevValue = getNumericValue_(outComp, row, prevCol);
        var trend = calcChange_(value, prevValue);

        if (value > 0 || prevValue > 0) {
          salesmen.push({ name: String(name), value: value, trend: trend });
        }
      }
    }

    // Sort by value descending
    salesmen.sort(function(a, b) { return b.value - a.value; });
  }

  return { labels: labels, total: total, salesmen: salesmen };
}

/* ==================== ADVANCES DATA ==================== */

/**
 * Get advances comparison data for charts - FIXED VERSION
 */
function getAdvancesData_(ss) {
  var labels = [], opening = [], given = [], settled = [], closing = [];

  // Get all monthly Advances sheets
  var months = getMonthlySheets_(ss, 'Advances');

  months.forEach(function(m) {
    labels.push(formatMonthLabel_(m.monthYear));

    var sheet = m.sheet;
    // B3 = Opening Balance, E3 = Closing Balance
    // E7:E1000 = Advances Given, F7:F1000 = Advances Settled
    var openingVal = getNumericValue_(sheet, 3, 2); // B3
    var closingVal = getNumericValue_(sheet, 3, 5); // E3
    var givenVal = sumColumn_(sheet, 5, 7, 1000);   // Column E (Given)
    var settledVal = sumColumn_(sheet, 6, 7, 1000); // Column F (Settled)

    opening.push(openingVal);
    given.push(givenVal);
    settled.push(settledVal);
    closing.push(closingVal);
  });

  // If no monthly sheets, try comparison sheet
  if (labels.length === 0) {
    var advComp = ss.getSheetByName('Advances_Comparison');
    if (advComp) {
      var lastCol = advComp.getLastColumn();
      for (var col = 2; col <= lastCol; col++) {
        var month = advComp.getRange(3, col).getValue();
        if (month) {
          labels.push(formatMonthLabel_(month));
          opening.push(getNumericValue_(advComp, 4, col));
          given.push(getNumericValue_(advComp, 5, col));
          settled.push(getNumericValue_(advComp, 6, col));
          closing.push(getNumericValue_(advComp, 7, col));
        }
      }
    }
  }

  return { labels: labels, opening: opening, given: given, settled: settled, closing: closing };
}

/* ==================== SUSPENSE DATA ==================== */

/**
 * Get suspense comparison data for charts - FIXED VERSION
 */
function getSuspenseData_(ss) {
  var labels = [], balance = [];

  // Get all monthly Suspense sheets
  var months = getMonthlySheets_(ss, 'Suspense');

  months.forEach(function(m) {
    labels.push(formatMonthLabel_(m.monthYear));

    var sheet = m.sheet;
    // E3 = Closing Balance
    var closingVal = getNumericValue_(sheet, 3, 5); // E3
    balance.push(closingVal);
  });

  // If no monthly sheets, try comparison sheet
  if (labels.length === 0) {
    var susComp = ss.getSheetByName('Suspense_Comparison');
    if (susComp) {
      var lastCol = susComp.getLastColumn();
      for (var col = 2; col <= lastCol; col++) {
        var month = susComp.getRange(3, col).getValue();
        if (month) {
          labels.push(formatMonthLabel_(month));
          balance.push(getNumericValue_(susComp, 7, col));
        }
      }
    }
  }

  return { labels: labels, balance: balance };
}

/* ==================== HELPER FUNCTIONS ==================== */

/**
 * Get monthly sheets for a given prefix (Banks, Advances, Suspense, Outstanding)
 */
function getMonthlySheets_(ss, prefix) {
  var sheets = ss.getSheets();
  var pattern = new RegExp('^' + prefix + '_([A-Z]{3})-(\\d{4})$');
  var result = [];

  sheets.forEach(function(sheet) {
    var match = sheet.getName().match(pattern);
    if (match) {
      var monthIdx = CONFIG.MONTHS.indexOf(match[1]);
      var year = parseInt(match[2]);
      result.push({
        sheet: sheet,
        monthYear: match[1] + '-' + match[2],
        sortKey: year * 100 + monthIdx
      });
    }
  });

  // Sort by date
  result.sort(function(a, b) { return a.sortKey - b.sortKey; });

  return result;
}

/**
 * Get numeric value from cell, handling errors
 */
function getNumericValue_(sheet, row, col) {
  try {
    var val = sheet.getRange(row, col).getValue();
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      var parsed = parseFloat(val.replace(/[^0-9.-]/g, ''));
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  } catch (e) {
    return 0;
  }
}

/**
 * Sum a column range
 */
function sumColumn_(sheet, col, startRow, endRow) {
  try {
    var range = sheet.getRange(startRow, col, endRow - startRow + 1, 1);
    var values = range.getValues();
    var sum = 0;
    values.forEach(function(row) {
      var val = row[0];
      if (typeof val === 'number') sum += val;
    });
    return sum;
  } catch (e) {
    return 0;
  }
}

/**
 * Find row by value in column
 */
function findRowByValue_(sheet, colRange, searchValue) {
  try {
    var data = sheet.getRange(colRange).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === String(searchValue).toUpperCase()) {
        return i + 1;
      }
    }
  } catch (e) {}
  return -1;
}

/**
 * Find column for a specific month
 */
function findMonthColumn_(sheet, monthYear, monthsList) {
  for (var i = 0; i < monthsList.length; i++) {
    if (monthsList[i].monthYear === monthYear) {
      return i + 2; // Column B = 2
    }
  }
  return monthsList.length + 1;
}

/**
 * Format month label (NOV-2025 -> NOV)
 */
function formatMonthLabel_(monthYear) {
  if (!monthYear) return '';
  var str = String(monthYear);
  var parts = str.split('-');
  return parts[0] || str.substring(0, 3).toUpperCase();
}

/**
 * Calculate percentage change
 */
function calcChange_(current, previous) {
  if (!previous || previous === 0) return 0;
  return ((current - previous) / Math.abs(previous)) * 100;
}

/**
 * Get previous month string
 */
function getPreviousMonth_(monthYear) {
  if (!monthYear) return '';
  var parts = monthYear.split('-');
  var monthIdx = CONFIG.MONTHS.indexOf(parts[0]);
  var year = parseInt(parts[1]);

  if (monthIdx === 0) {
    return 'DEC-' + (year - 1);
  }
  return CONFIG.MONTHS[monthIdx - 1] + '-' + year;
}

/* ==================== MENU & UI FUNCTIONS ==================== */

/**
 * Create custom menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Accounting System')
    .addItem('Create New Month', 'showNewMonthDialog')
    .addItem('Update All Comparisons', 'updateAllComparisons')
    .addSeparator()
    .addItem('Refresh Dashboard Data', 'refreshDashboard')
    .addItem('Test API', 'testAPI')
    .addToUi();
}

/**
 * Update all comparison sheets
 */
function updateAllComparisons() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Update Banks Comparison
  updateBanksComparison_(ss);

  // Update Outstanding Comparison
  updateOutstandingComparison_(ss);

  // Update Advances Comparison
  updateAdvancesComparison_(ss);

  // Update Suspense Comparison
  updateSuspenseComparison_(ss);

  SpreadsheetApp.getUi().alert('All comparison sheets updated!');
}

/**
 * Update Banks Comparison sheet
 */
function updateBanksComparison_(ss) {
  var sheet = ss.getSheetByName('Banks_Comparison');
  if (!sheet) return;

  var months = getMonthlySheets_(ss, 'Banks');

  // Clear existing data columns
  var lastCol = sheet.getLastColumn();
  if (lastCol > 1) {
    sheet.getRange(3, 2, sheet.getLastRow(), lastCol - 1).clear();
  }

  months.forEach(function(m, idx) {
    var col = idx + 2;
    var srcName = m.sheet.getName();

    sheet.getRange(3, col).setValue(m.monthYear);
    sheet.getRange(4, col).setFormula("=IFERROR('" + srcName + "'!B3,0)");  // Opening
    sheet.getRange(5, col).setFormula("=IFERROR('" + srcName + "'!B11,0)"); // Received
    sheet.getRange(6, col).setFormula("=IFERROR('" + srcName + "'!C11,0)"); // Paid Out
    sheet.getRange(7, col).setFormula("=IFERROR('" + srcName + "'!D11,0)"); // Expenses
    sheet.getRange(8, col).setFormula("=IFERROR('" + srcName + "'!F11,0)"); // Closing
  });
}

/**
 * Update Advances Comparison sheet
 */
function updateAdvancesComparison_(ss) {
  var sheet = ss.getSheetByName('Advances_Comparison');
  if (!sheet) return;

  var months = getMonthlySheets_(ss, 'Advances');

  // Clear existing data columns
  var lastCol = sheet.getLastColumn();
  if (lastCol > 1) {
    sheet.getRange(3, 2, sheet.getLastRow(), lastCol - 1).clear();
  }

  months.forEach(function(m, idx) {
    var col = idx + 2;
    var srcName = m.sheet.getName();

    sheet.getRange(3, col).setValue(m.monthYear);
    sheet.getRange(4, col).setFormula("=IFERROR('" + srcName + "'!B3,0)");           // Opening
    sheet.getRange(5, col).setFormula("=IFERROR(SUM('" + srcName + "'!E7:E1000),0)"); // Given
    sheet.getRange(6, col).setFormula("=IFERROR(SUM('" + srcName + "'!F7:F1000),0)"); // Settled
    sheet.getRange(7, col).setFormula("=IFERROR('" + srcName + "'!E3,0)");           // Closing
  });
}

/**
 * Update Suspense Comparison sheet
 */
function updateSuspenseComparison_(ss) {
  var sheet = ss.getSheetByName('Suspense_Comparison');
  if (!sheet) return;

  var months = getMonthlySheets_(ss, 'Suspense');

  // Clear existing data columns
  var lastCol = sheet.getLastColumn();
  if (lastCol > 1) {
    sheet.getRange(3, 2, sheet.getLastRow(), lastCol - 1).clear();
  }

  months.forEach(function(m, idx) {
    var col = idx + 2;
    var srcName = m.sheet.getName();

    sheet.getRange(3, col).setValue(m.monthYear);
    sheet.getRange(4, col).setFormula("=IFERROR('" + srcName + "'!B3,0)");           // Opening
    sheet.getRange(5, col).setFormula("=IFERROR(SUM('" + srcName + "'!E7:E1000),0)"); // Debits
    sheet.getRange(6, col).setFormula("=IFERROR(SUM('" + srcName + "'!F7:F1000),0)"); // Credits
    sheet.getRange(7, col).setFormula("=IFERROR('" + srcName + "'!E3,0)");           // Closing
  });
}

/**
 * Update Outstanding Comparison sheet
 */
function updateOutstandingComparison_(ss) {
  var sheet = ss.getSheetByName('Outstanding_Comparison');
  if (!sheet) return;

  var months = getMonthlySheets_(ss, 'Outstanding');

  // Get all salesmen from all months
  var allSalesmen = {};
  months.forEach(function(m) {
    var lastRow = m.sheet.getLastRow();
    for (var row = 4; row <= lastRow; row++) {
      var name = m.sheet.getRange(row, 1).getValue();
      if (name && String(name).trim() !== '') {
        allSalesmen[String(name).toLowerCase()] = String(name);
      }
    }
  });

  var salesmenList = Object.values(allSalesmen).sort();

  // Clear and rebuild
  sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clear();

  // Header row
  sheet.getRange(3, 1).setValue('Salesman');
  months.forEach(function(m, idx) {
    sheet.getRange(3, idx + 2).setValue(m.monthYear);
  });

  // Data rows
  salesmenList.forEach(function(name, rowIdx) {
    var row = rowIdx + 4;
    sheet.getRange(row, 1).setValue(name);

    months.forEach(function(m, colIdx) {
      var col = colIdx + 2;
      var srcName = m.sheet.getName();
      sheet.getRange(row, col).setFormula(
        "=IFERROR(VLOOKUP(\"" + name + "\",'" + srcName + "'!A:B,2,FALSE),0)"
      );
    });
  });

  // Total row
  var totalRow = salesmenList.length + 4;
  sheet.getRange(totalRow, 1).setValue('TOTAL');
  months.forEach(function(m, colIdx) {
    var col = colIdx + 2;
    sheet.getRange(totalRow, col).setFormula('=SUM(' + sheet.getRange(4, col).getA1Notation() + ':' + sheet.getRange(totalRow - 1, col).getA1Notation() + ')');
  });
}

/**
 * Refresh dashboard data (for testing)
 */
function refreshDashboard() {
  var data = getDashboardData();
  Logger.log(JSON.stringify(data, null, 2));
  SpreadsheetApp.getUi().alert('Dashboard data refreshed! Check View > Logs for details.');
}

/**
 * Test API endpoint
 */
function testAPI() {
  var data = getDashboardData();
  var ui = SpreadsheetApp.getUi();

  var msg = 'API Test Results:\n\n' +
    'Latest Month: ' + data.latestMonth + '\n' +
    'Bank Balance: ' + data.kpi.bankBalance + '\n' +
    'Outstanding: ' + data.kpi.outstanding + '\n' +
    'Advances: ' + data.kpi.advances + '\n' +
    'Suspense: ' + data.kpi.suspense + '\n\n' +
    'Months found: ' + data.months.length;

  ui.alert('API Test', msg, ui.ButtonSet.OK);
}

/**
 * Show dialog to create new month sheets
 */
function showNewMonthDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Create New Month', 'Enter month-year (e.g., DEC-2025):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    var monthYear = response.getResponseText().toUpperCase().trim();
    if (/^[A-Z]{3}-\d{4}$/.test(monthYear)) {
      createMonthSheets(monthYear);
      ui.alert('Created sheets for ' + monthYear);
    } else {
      ui.alert('Invalid format. Please use MMM-YYYY (e.g., DEC-2025)');
    }
  }
}

/**
 * Create all sheets for a new month
 */
function createMonthSheets(monthYear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createBanksSheet_(ss, monthYear);
  createAdvancesSheet_(ss, monthYear);
  createSuspenseSheet_(ss, monthYear);
  createOutstandingSheet_(ss, monthYear);

  // Update comparisons
  updateAllComparisons();
}

/**
 * Create Banks sheet for a month
 */
function createBanksSheet_(ss, monthYear) {
  var name = 'Banks_' + monthYear;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);

  // Title
  sheet.getRange('A1').setValue('BANK RECONCILIATION - ' + monthYear);
  sheet.getRange('A1:F1').merge().setHorizontalAlignment('center').setFontSize(14).setFontWeight('bold');

  // Headers
  sheet.getRange('A3').setValue('Opening Balance:');
  sheet.getRange('B3').setValue(0).setNumberFormat('#,##0.00');

  // Column headers
  sheet.getRange('A4:F4').setValues([['Bank Account', 'Opening', 'Received', 'Paid Out', 'Expenses', 'Closing']]);
  sheet.getRange('A4:F4').setFontWeight('bold').setBackground('#f0f0f0');

  return sheet;
}

/**
 * Create Advances sheet for a month
 */
function createAdvancesSheet_(ss, monthYear) {
  var name = 'Advances_' + monthYear;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);

  // Get previous month closing as opening
  var prevMonth = getPreviousMonth_(monthYear);
  var prevSheet = ss.getSheetByName('Advances_' + prevMonth);
  var openingBalance = prevSheet ? getNumericValue_(prevSheet, 3, 5) : 0;

  // Title
  sheet.getRange('A1').setValue('ADVANCES RECONCILIATION - ' + monthYear);
  sheet.getRange('A1:G1').merge().setHorizontalAlignment('center').setFontSize(14).setFontWeight('bold');

  // Balances
  sheet.getRange('A3').setValue('OPENING BALANCE');
  sheet.getRange('B3').setValue(openingBalance).setNumberFormat('#,##0.00');
  sheet.getRange('D3').setValue('CLOSING BALANCE');
  sheet.getRange('E3').setFormula('=B3+SUM(E7:E1000)-SUM(F7:F1000)').setNumberFormat('#,##0.00');

  // Headers
  sheet.getRange('A6:G6').setValues([['Date', 'Voucher No', 'Description', 'Person', 'Given', 'Settled', 'Balance']]);
  sheet.getRange('A6:G6').setFontWeight('bold').setBackground('#f0f0f0');

  return sheet;
}

/**
 * Create Suspense sheet for a month
 */
function createSuspenseSheet_(ss, monthYear) {
  var name = 'Suspense_' + monthYear;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);

  // Get previous month closing as opening
  var prevMonth = getPreviousMonth_(monthYear);
  var prevSheet = ss.getSheetByName('Suspense_' + prevMonth);
  var openingBalance = prevSheet ? getNumericValue_(prevSheet, 3, 5) : 0;

  // Title
  sheet.getRange('A1').setValue('SUSPENSE RECONCILIATION - ' + monthYear);
  sheet.getRange('A1:G1').merge().setHorizontalAlignment('center').setFontSize(14).setFontWeight('bold');

  // Balances
  sheet.getRange('A3').setValue('OPENING BALANCE');
  sheet.getRange('B3').setValue(openingBalance).setNumberFormat('#,##0.00');
  sheet.getRange('D3').setValue('CLOSING BALANCE');
  sheet.getRange('E3').setFormula('=B3+SUM(E7:E1000)-SUM(F7:F1000)').setNumberFormat('#,##0.00');

  // Headers
  sheet.getRange('A6:G6').setValues([['Date', 'Journal No', 'Description', 'Reference', 'Debit', 'Credit', 'Balance']]);
  sheet.getRange('A6:G6').setFontWeight('bold').setBackground('#f0f0f0');

  return sheet;
}

/**
 * Create Outstanding sheet for a month
 */
function createOutstandingSheet_(ss, monthYear) {
  var name = 'Outstanding_' + monthYear;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);

  // Title
  sheet.getRange('A1').setValue('OUTSTANDING RECEIVABLES - ' + monthYear);
  sheet.getRange('A1:B1').merge().setHorizontalAlignment('center').setFontSize(14).setFontWeight('bold');

  // Headers
  sheet.getRange('A3:B3').setValues([['Salesman', 'Outstanding Amount']]);
  sheet.getRange('A3:B3').setFontWeight('bold').setBackground('#f0f0f0');

  return sheet;
}
