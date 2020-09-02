// TO DO: update the vinAuditId & vins varaibles with your personal vinAudit ID & vehichle identification numbers (vins)
// docs: https://www.vinaudit.com/vehicle-market-value-api#tab-id-1
var vinAuditId = "VA_DEMO_KEY";
var vins = ["1234", "12345"];

function vinAuditInsert() {
  // load the moment library for date/time handling
  eval(
    UrlFetchApp.fetch(
      "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.min.js"
    ).getContentText()
  );

  // get the Balance History sheet
  var balanceHistorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Balance History"
  );
  // get the Balance History header row
  var header = balanceHistorySheet
    .getRange(1, 1, 2, balanceHistorySheet.getMaxColumns())
    .getValues()[0];

  var columns = {};
  var rows = [];

  // create a lookup for Balance History column headers
  for (var columnIndex = 0; columnIndex < header.length; columnIndex++)
    columns[header[columnIndex].trim().toLowerCase()] = columnIndex;

  // step through each vin in the array
  for (var i = 0; i < vins.length; i++) {
    // scrape vinAudit's  information
    const vin = vins[i];
    var estimate = vinAuditFetch(
      vinAuditId,
      vin,
      (mileage = "average"),
      (period = "90")
    );

    if (estimate) {
      var row = [];

      // initialize a new row array
      for (var j = 0; j < header.length; j++) row.push("");

      function safeSetVal(columnName, value) {
        if (columnName.length && columns.hasOwnProperty(columnName))
          row[columns[columnName]] = value;
      }

      // build a balance history row where column header's exist in the active spreadsheet
      const today = moment();
      const time = today.format("h:mm A");
      const { vin, vehicle, amount } = estimate;
      safeSetVal("date", today.format("M/D/YYYY"));
      safeSetVal("week", today.startOf("week").format("YYYY-MM-DD"));
      safeSetVal("month", today.startOf("month").format("YYYY-MM-DD"));
      safeSetVal("time", time);
      safeSetVal("balance", amount);
      safeSetVal("account", "vinAudit: " + vehicle);
      safeSetVal("account #", vin);
      safeSetVal("account id", "vinAudit-" + vin);
      safeSetVal("index", vin);
      safeSetVal("institution", "vinAudit - API");
      safeSetVal("type", "PROPERTY");
      safeSetVal("class", "Asset");

      // push the new row into the array of rows to add
      rows.push(row);
    }
  }

  // if new rows were created, add them to the sheet...
  if (rows.length) {
    // add new rows to the sheet
    balanceHistorySheet
      .getRange(
        balanceHistorySheet.getLastRow() + 1,
        1,
        rows.length,
        rows[0].length
      )
      .setValues(rows);

    // re-sort the Balance History sheet by date
    if (columns.hasOwnProperty("date"))
      balanceHistorySheet.getDataRange().sort([
        {
          column: columns["date"] + 1,
          ascending: false,
        },
      ]);
  }
}

function vinAuditFetch(vinAuditId, vin, mileage = "average", period = "90") {
  // create a url to query the vinAudit web service
  var url =
    "http://marketvalue.vinaudit.com/getmarketvalue.php?key=" +
    vinAuditId +
    "&vin=" +
    vin +
    "&format=json&period=" +
    period +
    "&mileage=" +
    mileage;
  // fetch a property from the vinAudit web service
  var json = UrlFetchApp.fetch(url).getContentText();
  var data = JSON.parse(json);
  const { vehicle, mileage: mileageEstimate, mean: amount } = data;

  // return an object with relevant values
  return {
    vin,
    vehicle,
    amount,
    mileage: mileageEstimate,
  };
}
