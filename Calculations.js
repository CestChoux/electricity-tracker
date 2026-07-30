function calculateDailyTotals() {

  const rawSheet = getSheet(CONFIG.SHEETS.RAW);
  const dailySheet = getSheet(CONFIG.SHEETS.DAILY);

  dailySheet.clear();
  dailySheet.appendRow(CONFIG.HEADERS.DAILY);

  const lastRow = rawSheet.getLastRow();

  if (lastRow <= 1) return;

//  if (totalConsumption <= 0) {continue;}

  const values = rawSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  const totals = {};

  values.forEach(row => {

    const dateKey = Utilities.formatDate(
      row[0],
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    totals[dateKey] = (totals[dateKey] || 0) + Number(row[2]);

  });

  const output = Object.keys(totals)
    .sort()
    .map(key => [new Date(key), parseFloat(totals[key].toFixed(3).replace(",", "."))]);

  if (output.length) {

    dailySheet
      .getRange(2, 1, output.length, 2)
      .setValues(output);

    dailySheet
      .getRange(2, 1, output.length, 1)
      .setNumberFormat("yyyy-MM-dd");

  }

}

function calculateMonthlyTotals() {

  const dailySheet = getSheet(CONFIG.SHEETS.DAILY);
  const monthlySheet = getSheet(CONFIG.SHEETS.MONTHLY);

  monthlySheet.clear();
  monthlySheet.appendRow(CONFIG.HEADERS.MONTHLY);

  const lastRow = dailySheet.getLastRow();

  if (lastRow <= 1) return;

  const values = dailySheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const totals = {};

  values.forEach(row => {

    const month = Utilities.formatDate(
      row[0],
      Session.getScriptTimeZone(),
      "yyyy-MM"
    );

    totals[month] = (totals[month] || 0) + Number(row[1]);

  });

  const output = Object.keys(totals)
    .sort()
    .map(month => [
      month,
      Number(totals[month].toFixed(3))
    ]);

  if (output.length) {
    monthlySheet
      .getRange(2, 1, output.length, 2)
      .setValues(output);
  }

}