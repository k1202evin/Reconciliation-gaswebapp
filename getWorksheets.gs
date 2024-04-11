function getWorksheets() {
  var spreadsheetId = "1bU38faGJ49iUsKAVzdqJLfl-Yq7rq9bD5tlzxWRXkFU";  // 请替换成你的Google表格ID
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var worksheets = spreadsheet.getSheets();

  var worksheetNames = worksheets.map(function(worksheet) {
    return worksheet.getName();
  });

  return worksheetNames;
}
