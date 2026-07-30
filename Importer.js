/**
 * Imports all new reports.
 */
function importReports() {

  const files = getCsvFiles();
  const imported = getImportedFileIds();

  let importedFiles = 0;
  let importedRows = 0;

  files.forEach(file => {

    if (imported.has(file.getId())) {
      logInfo("Skipping " + file.getName());
      return;
    }

    importedRows += importCsv(file);
    importedFiles++;

  });

  calculateDailyTotals();
  calculateMonthlyTotals();
  updateStatistics();

  buildDashboard();
  refreshDashboard();

  toast(
    "Import complete.",
    "Imported " + importedFiles + " file(s)."
  );

  function importCsv(file) {
    try {
      const rawSheet = getSheet(CONFIG.SHEETS.RAW);
      const csv = Utilities.parseCsv(
        file.getBlob().getDataAsString("windows-1250"),
        ";"
      );

      if (csv.length < 2) {
        logError(`No data rows in ${file.getName()}`);
        return 0;
      }

      const rows = [];
      let skippedRows = 0;

      for (let i = 1; i < csv.length; i++) {
        const row = csv[i];

        try {
          if (row.length < 3) {
            skippedRows++;
            continue;
          }

          const startDate = parseCzechDateTime(row[0]);
          const endDate = parseCzechDateTime(row[1]);
          const consumption = Number(row[2].replace(",", "."));

          if (isNaN(consumption)) {
            logError(`Invalid consumption in ${file.getName()}, row ${i+1}: "${row[2]}"`);
            skippedRows++;
            continue;
          }

          if (consumption < 0) {
            logError(`Negative consumption in ${file.getName()}, row ${i+1}: ${consumption}`);
            skippedRows++;
            continue;
          }

          if (startDate >= endDate) {
            logError(`Invalid date range in ${file.getName()}, row ${i+1}`);
            skippedRows++;
            continue;
          }

          rows.push([startDate, endDate, consumption, file.getId(), file.getName()]);
        } catch (e) {
          logError(`Row ${i+1} in ${file.getName()}: ${e.message}`);
          skippedRows++;
        }
      }

      if (rows.length > 0) {
        rawSheet.getRange(
          rawSheet.getLastRow() + 1, 1, rows.length, rows[0].length
        ).setValues(rows);
      }

      logInfo(`Imported ${file.getName()} (${rows.length} rows, ${skippedRows} skipped)`);
      return rows.length;
    } catch (e) {
      logError(`Failed to import ${file.getName()}: ${e.message}`);
      return 0;
    }
  } 
}

/**
 * Returns every CSV in the project folder.
 */
function getCsvFiles() {
  const folder = getProjectFolder();
  const iterator = folder.getFiles();
  const files = [];

  while (iterator.hasNext()) {
    const file = iterator.next();
    if (
      file.getMimeType() === MimeType.CSV ||
      file.getName().toLowerCase().endsWith(".csv")
    ) {
      files.push(file);
    }
  }

  files.sort((a, b) => a.getName().localeCompare(b.getName()));
  return files;
}



/**
 * Lists all CSV files in the project folder.
 */
function listCsvFiles() {
  const files = getCsvFiles();

  if (files.length === 0) {
    SpreadsheetApp.getUi().alert("No CSV files found in the project folder.");
    return;
  }

  let message = "CSV files found:\n\n";
  files.forEach(file => {
    message += `• ${file.getName()}\n`;
  });

  SpreadsheetApp.getUi().alert(message, "CSV Files");
}

/**
 * Returns imported Google Drive File IDs.
 */
function getImportedFileIds() {
  const sheet = getSheet(CONFIG.SHEETS.RAW);

  if (sheet.getLastRow() <= 1)
    return new Set();

  return new Set(
    sheet
      .getRange(2, 4, sheet.getLastRow()-1, 1)
      .getValues()
      .flat()
      .filter(String)
  );
}