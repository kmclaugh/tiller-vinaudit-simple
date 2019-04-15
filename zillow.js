// TO DO: update the zwsid & zpid varaibles with your personal Zillow Web Services ID & Zillow Property ID
// Create your own Zillow Web Services ID here for free: https://www.zillow.com/howto/api/APIOverview.htm
var zwsid = 'X1-11111111111111_11111';
// Search for your property on Zillow's website then find the ZPID in the URL (will look like:"XXXXXXXX_zpid"— just use the X's)
// Note: to update values for multiple properties, set your ZPID like this '[11111111, 22222222, 33333333]'
var zpid = '11111111';

function zestimateInsert() {
  // load the moment library for date/time handling
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.min.js').getContentText());

  // get the Balance History sheet
  var balanceHistorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Balance History");
  // get the Balance History header row
  var header = balanceHistorySheet.getRange(1, 1, 2, balanceHistorySheet.getMaxColumns()).getValues()[0];

  var columns = {};
  var rows = [];

  // create a lookup for Balance History column headers
  for (var columnIndex = 0; columnIndex < header.length; columnIndex++)
    columns[header[columnIndex].trim().toLowerCase()] = columnIndex;

  // parse the zpid global variable in case the user has used an array
  var zpids = JSON.parse(zpid);

  // if the zpid is just a single value, wrap it in an array
  if (!Array.isArray(zpids))
    zpids = [zpids];

  // step through each zpid in the array
  for (var i = 0; i < zpids.length; i++) {
    // scrape Zillow's Xestimate and property information
    var zestimate = zestimateFetch(zwsid, zpids[i]);

    if (zestimate) {
      var row = [];

      // initialize a new row array
      for (var j = 0; j < header.length; j++)
        row.push('');

      function safeSetVal(columnName, value) {
        if (columnName.length && columns.hasOwnProperty(columnName))
          row[columns[columnName]] = value;
      }

      // build a balance history row where column header's exist in the active spreadsheet
      safeSetVal('date', moment(zestimate.lastUpdated).format('M/D/YYYY'));
      safeSetVal('week', moment(zestimate.lastUpdated).startOf('week').format('YYYY-MM-DD'));
      safeSetVal('month', moment(zestimate.lastUpdated).startOf('month').format('YYYY-MM-DD'));
      safeSetVal('time', 0);
      safeSetVal('balance', zestimate.amount);
      safeSetVal('account', 'Zillow: ' + zestimate.street + ', ' + zestimate.city + ', ' + zestimate.state);
      safeSetVal('account #', zestimate.zpid);
      safeSetVal('index', zestimate.zpid);
      safeSetVal('institution', "Zillow - API");
      safeSetVal('type', "Asset");
      safeSetVal('class', "Zestimate®");

      // push the new row into the array of rows to add
      rows.push(row);
    }
  }

  // if new rows were created, add them to the sheet...
  if (rows.length) {
    // add new rows to the sheet
    balanceHistorySheet.getRange(balanceHistorySheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    // re-sort the Balance History sheet by date
    if (columns.hasOwnProperty('date'))
      balanceHistorySheet.getDataRange().sort([{
        column: columns['date'] + 1,
        ascending: false
      }]);
  }
}

function zestimateFetch(zwsid, zpid) {
  // create a url to query the zillow web service
  var url = "https://www.zillow.com/webservice/GetZestimate.htm?zws-id=" + zwsid + "&zpid=" + zpid;
  // fetch a property from the zillow web service
  var xml = UrlFetchApp.fetch(url).getContentText();
  // parse the captured xml
  var document = XmlService.parse(xml);

  // retrieve the relevant data elements
  var root = document.getRootElement();
  var response = root.getChild('response');
  var address = response.getChild('address');
  var street = address.getChild('street').getText();
  var zipcode = address.getChild('zipcode').getText();
  var city = address.getChild('city').getText();
  var state = address.getChild('state').getText();
  var zestimate = response.getChild('zestimate');
  var amount = zestimate.getChild('amount').getText();
  var lastUpdated = zestimate.getChild('last-updated').getText();

  // return an object with relevant values
  return {
    zpid: zpid,
    amount: amount,
    lastUpdated: lastUpdated,
    street: street,
    zip: zipcode,
    city: city,
    state: state
  };
}
