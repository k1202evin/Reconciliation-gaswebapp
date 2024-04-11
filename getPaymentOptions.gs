function getPaymentOptions() {
  var spreadsheetId = "13juebF8-M7QWnjfeWYvGKxP2XJX9Ve7Cio3tLdcDitY";  // 请替换成你的Google表格ID
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var paymentsheets = spreadsheet.getSheets();

  var paymentsheetNames = paymentsheets.map(function(paymentsheet) {
    return paymentsheet.getName();
  });

  return paymentsheetNames;
}
