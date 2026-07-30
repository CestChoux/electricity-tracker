/**
 * Exports Daily data to a CSV file in Drive.
 */
function exportDailyToCsv() {
  const dailySheet = getSheet(CONFIG.SHEETS.DAILY);
  const data = dailySheet.getDataRange().getValues();
  const csvContent = data.map(row => row.join(",")).join("\r\n");

  const folder = getProjectFolder();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );
  const fileName = `Daily_Export_${timestamp}.csv`;

  folder.createFile(fileName, csvContent, MimeType.CSV);
  toast(`Exported Daily data to: ${fileName}`);
  logInfo(`Exported Daily data to Drive as ${fileName}`);
}

/**
 * Exports Raw data to a CSV file in Drive.
 */
function exportRawToCsv() {
  const rawSheet = getSheet(CONFIG.SHEETS.RAW);
  const data = rawSheet.getDataRange().getValues();
  const csvContent = data.map(row => row.join(",")).join("\r\n");

  const folder = getProjectFolder();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );
  const fileName = `Raw_Export_${timestamp}.csv`;

  folder.createFile(fileName, csvContent, MimeType.CSV);
  toast(`Exported Raw data to: ${fileName}`);
  logInfo(`Exported Raw data to Drive as ${fileName}`);
}