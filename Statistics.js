/**
 * Builds the Statistics sheet from Daily.
 */
function updateStatistics() {

  const dailySheet = getSheet(CONFIG.SHEETS.DAILY);
  const statsSheet = getSheet(CONFIG.SHEETS.STATISTICS);

  statsSheet.clear();

  statsSheet.appendRow([
    "Date",
    "Consumption (kWh)",
    "Running Total",
    "Previous Day",
    "Difference",
    "% Change",
    "7-Day Average"
  ]);

  const lastRow = dailySheet.getLastRow();

  if (lastRow <= 1)
    return;

  const values = dailySheet
    .getRange(2,1,lastRow-1,2)
    .getValues();

  let running = 0;

  const output = [];

  for(let i=0;i<values.length;i++){

    const date = values[i][0];
    const consumption = Number(values[i][1]);

    running += consumption;

    const previous =
      i===0
      ? ""
      : Number(values[i-1][1]);

    const difference =
      i===0
      ? ""
      : consumption-previous;

    const percent =
      i===0 || previous===0
      ? ""
      : difference/previous;

    let average="";

    if(i>=6){

      let sum=0;

      for(let j=i-6;j<=i;j++)
        sum+=Number(values[j][1]);

      average=sum/7;

    }

    output.push([
      date,
      consumption,
      running,
      previous,
      difference,
      percent,
      average
    ]);

  }

  statsSheet
    .getRange(
      2,
      1,
      output.length,
      output[0].length
    )
    .setValues(output);

  statsSheet
    .getRange(
      2,
      6,
      output.length,
      1
    )
    .setNumberFormat("0.00%");

// statsSheet.hideSheet();  // Let users view statistics

  logInfo("Statistics updated.");

}