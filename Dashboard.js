/**
 * Initializes all project sheets with headers.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Helper to format headers
  const formatHeaders = (sheet, headers) => {
    sheet.clear();
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
  };

  // Apply to all sheets
  formatHeaders(getSheet(CONFIG.SHEETS.RAW), CONFIG.HEADERS.RAW);
  formatHeaders(getSheet(CONFIG.SHEETS.DAILY), CONFIG.HEADERS.DAILY);
  formatHeaders(getSheet(CONFIG.SHEETS.MONTHLY), CONFIG.HEADERS.MONTHLY);
  formatHeaders(getSheet(CONFIG.SHEETS.LOG), CONFIG.HEADERS.LOG);

  // Clear sheets without headers
  getSheet(CONFIG.SHEETS.STATISTICS).clear();
  getSheet(CONFIG.SHEETS.SYSTEM).clear();

  // Build dashboard
  buildDashboard();

  const sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  sheet.setColumnWidths(1, 1, 190);
  sheet.setColumnWidths(2, 1, 190);
  sheet.setColumnWidths(3, 1, 120);
  sheet.setColumnWidths(4, 5, 110);

  toast("Project initialized successfully");
}

/**
 * Creates or rebuilds the Dashboard sheet.
 */
function buildDashboard() {

  const sheet = getSheet(CONFIG.SHEETS.DASHBOARD);
  const dailySheet = getSheet(CONFIG.SHEETS.DAILY);

  // ===========================
  // Read Daily data
  // ===========================

  const lastRow = dailySheet.getLastRow();

  let importedDays = 0;
  let activeDays = 0;
  let total = 0;
  let average = "";
  let maximum = "";
  let minimum = "";
  let latest = "";

  const zeroDays = [];

  if (lastRow > 1) {

    const rows = dailySheet
      .getRange(2, 1, lastRow - 1, 2)
      .getValues();

    importedDays = rows.length;

    const values = [];

    rows.forEach(([date, consumption]) => {

      const value = Number(consumption);

      if (isNaN(value)) return;

      if (value === 0) {
        zeroDays.push([date]);
      } else {
        values.push(value);
      }

    });

    activeDays = values.length;

    if (activeDays > 0) {
      total = values.reduce((a, b) => a + b, 0);
      average = total / activeDays;
      maximum = Math.max(...values);
      minimum = Math.min(...values);
      latest = values[values.length - 1];
    }
  }

  // ===========================
  // Build dashboard
  // ===========================

  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidths(4, 5, 120);

  sheet.getRange("A1:H1")
    .merge()
    .setValue("⚡ Electricity Consumption Dashboard")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const labels = [
    ["Imported Days", "B3"],
    ["Active Days", "B4"],
    ["Total Consumption (kWh)", "B5"],
    ["Average per Active Day", "B6"],
    ["Highest Day", "B7"],
    ["Lowest Day", "B8"],
    ["Latest Active Day", "B9"]
  ];

  labels.forEach(([text, cell]) => {
    sheet.getRange(cell)
      .setValue(text)
      .setFontWeight("bold");
  });

  sheet.getRange("C3").setValue(importedDays);
  sheet.getRange("C4").setValue(activeDays);
  sheet.getRange("C5").setValue(total);
  sheet.getRange("C6").setValue(average);
  sheet.getRange("C7").setValue(maximum);
  sheet.getRange("C8").setValue(minimum);
  sheet.getRange("C9").setValue(latest);

  sheet.getRange("C3:C4").setNumberFormat("0");
  sheet.getRange("C5:C9").setNumberFormat("0.000");

  // Charts section

  sheet.getRange("A12")
    .setValue("Daily Consumption")
    .setFontWeight("bold")
    .setFontSize(14);

  sheet.getRange("A31")
    .setValue("Monthly Consumption")
    .setFontWeight("bold")
    .setFontSize(14);

  // Zero Consumption Days

  sheet.getRange("J2")
    .setValue("Zero Consumption Days")
    .setFontWeight("bold")
    .setFontSize(14);

  if (zeroDays.length === 0) {

    sheet.getRange("J3").setValue("None");

  } else {

    sheet.getRange(3, 10, zeroDays.length, 1)
      .setValues(zeroDays)
      .setNumberFormat("yyyy-mm-dd");

  }

  return sheet;
}

/**
 * Recreates all dashboard charts.
 */
/**
 * Recreates all dashboard charts.
 */
function refreshDashboard() {

  const dashboard = getSheet(CONFIG.SHEETS.DASHBOARD);

  // Remove existing charts
  dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));

  // Daily chart
  const daily = getSheet(CONFIG.SHEETS.DAILY);

  if (daily.getLastRow() > 1) {

    const chart = dashboard.newChart()
      .asLineChart()
      .addRange(daily.getRange(1, 1, daily.getLastRow(), 2))
      .setPosition(12, 1, 0, 0)
      .setOption("title", "Daily Consumption")
      .setOption("legend", { position: "none" })
      .setOption("pointSize", 3)
      .setOption("height", 350)
      .setOption("width", 900)
      .build();

    dashboard.insertChart(chart);
  }

  // Monthly chart
  const monthly = getSheet(CONFIG.SHEETS.MONTHLY);

  if (monthly.getLastRow() > 1) {

    const chart = dashboard.newChart()
      .asColumnChart()
      .addRange(monthly.getRange(1, 1, monthly.getLastRow(), 2))
      .setPosition(32, 1, 0, 0)
      .setOption("title", "Monthly Consumption")
      .setOption("legend", { position: "none" })
      .setOption("height", 350)
      .setOption("width", 900)
      .build();

    dashboard.insertChart(chart);
  }
}

/**
 * Refreshes all calculations without re-importing
 */
function refreshAll() {
  calculateDailyTotals();
  calculateMonthlyTotals();
  updateStatistics();
  buildDashboard();
  refreshDashboard();
  logInfo("All data refreshed");
  toast("All data refreshed", "Refresh Complete");
}

/**
 * Shows the hidden Statistics sheet
 */
function showStatistics() {
  const sheet = getSheet(CONFIG.SHEETS.STATISTICS);
  sheet.showSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  toast("Statistics sheet is now visible");
}

/**
 * Creates a backup of Raw Data
 */
function backupRawData() {
  const rawSheet = getSheet(CONFIG.SHEETS.RAW);
  const ss = getSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const backupName = `Backup_Raw_${timestamp}`;

  const backupSheet = rawSheet.copyTo(ss);
  backupSheet.setName(backupName);
  backupSheet.hideSheet();

  logInfo(`Backup created: ${backupName}`);
  toast(`Backup created: ${backupName}`);
}

/**
 * Restores from the most recent backup
 */
function restoreFromBackup() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const backups = sheets.filter(s => s.getName().startsWith("Backup_Raw_"));

  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert("No backups found!");
    return;
  }

  backups.sort((a, b) => b.getName().localeCompare(a.getName()));
  const latestBackup = backups[0];

  const rawSheet = getSheet(CONFIG.SHEETS.RAW);
  rawSheet.clear();

  const data = latestBackup.getDataRange().getValues();
  rawSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  logInfo(`Restored from: ${latestBackup.getName()}`);
  toast(`Restored from: ${latestBackup.getName()}`);
}