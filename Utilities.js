/**
 * Returns the active spreadsheet.
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Returns a sheet by configured name.
 */
function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);

  if (!sheet) {
    throw new Error(`Sheet "${name}" does not exist.`);
  }

  return sheet;
}

/**
 * Returns the project folder (the folder containing this spreadsheet).
 */
function getProjectFolder() {

  const spreadsheetId = getSpreadsheet().getId();

  const spreadsheetFile = DriveApp.getFileById(spreadsheetId);

  const parents = spreadsheetFile.getParents();

  if (!parents.hasNext()) {
    throw new Error(
      "The spreadsheet must be inside the Electricity Tracking folder."
    );
  }

  return parents.next();

}

/**
 * Writes a message to the Log sheet.
 */
function logInfo(message) {

  const sheet = getSheet(CONFIG.SHEETS.LOG);

  sheet.appendRow([
    new Date(),
    "INFO",
    message
  ]);

}

/**
 * Writes an error to the Log sheet.
 */
function logError(message) {

  const sheet = getSheet(CONFIG.SHEETS.LOG);

  sheet.appendRow([
    new Date(),
    "ERROR",
    message
  ]);

}

/**
 * Small notification in the bottom-right corner.
 */
function toast(message, title = "Electricity Tracker") {
  getSpreadsheet().toast(message, title, 5);
}

/**
 * Converts Czech decimal numbers to JavaScript numbers.
 */
function parseConsumption(value) {

  if (value === "" || value === null)
    return null;

  return Number(
    String(value)
      .replace(",", ".")
      .trim()
  );

}

/**
 * Formats a Date as yyyy-MM-dd.
 */
function formatDate(date) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

}

/**
 * Formats a Date as HH:mm.
 */
function formatTime(date) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "HH:mm"
  );

}

function testUtilities() {

  const folder = getProjectFolder();

  logInfo("Utilities test successful.");

  SpreadsheetApp.getUi().alert(
    "Project folder:\n\n" + folder.getName()
  );

}

/**
 * Parses Czech datetime (dd.MM.yyyy HH:mm) into a Date.
 */
function parseCzechDateTime(value) {
  if (!value || typeof value !== 'string') {
    throw new Error(`Invalid datetime value: ${value}`);
  }

  const trimmed = value.trim();
  const [datePart, timePart] = trimmed.split(" ");

  if (!datePart || !timePart) {
    throw new Error(`Expected format "dd.MM.yyyy HH:mm", got: ${value}`);
  }

  const dateParts = datePart.split(".");
  const timeParts = timePart.split(":");

  if (dateParts.length !== 3 || timeParts.length !== 2) {
    throw new Error(`Invalid format in: ${value}`);
  }

  const day = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10);
  const year = parseInt(dateParts[2], 10);
  const hour = parseInt(timeParts[0], 10);
  const minute = parseInt(timeParts[1], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year) || isNaN(hour) || isNaN(minute)) {
    throw new Error(`Non-numeric date/time in: ${value}`);
  }

  return new Date(year, month - 1, day, hour, minute);
}