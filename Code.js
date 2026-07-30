function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Electricity")
    .addItem("Initialize Project", "initializeProject")
    .addSeparator()
    .addItem("List CSV Files", "listCsvFiles")
    .addItem("Import New Reports", "importReports")
    .addSeparator()
    .addItem("Refresh All Data", "refreshAll")
    .addItem("Show Statistics", "showStatistics")
    .addSeparator()
    .addItem("Backup Raw Data", "backupRawData")
    .addItem("Restore from Backup", "restoreFromBackup")
    .addToUi()
    .addSeparator()
    .addItem("Export Daily Data", "exportDailyToCsv")
    .addItem("Export Raw Data", "exportRawToCsv");
}