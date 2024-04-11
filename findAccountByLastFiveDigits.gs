function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page');
}

function findAccountByLastFiveDigits(lastFiveDigits, selectedWorksheet) {
  var spreadsheetId = "1bU38faGJ49iUsKAVzdqJLfl-Yq7rq9bD5tlzxWRXkFU"; 
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(selectedWorksheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();
  var acc = -1; // 默认值
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == "對方帳號")
      acc = i;
  }

  // 创建一个用于存储匹配帐号的数组
  var foundAccounts = [];

  // 遍历数据范围，查找匹配的帐号
  for (var i = 1; i < data.length; i++) {
    var accountNumber = data[i][acc]; // 帐号号码在第9列

    if (accountNumber && accountNumber.toString().slice(-5) === lastFiveDigits) {
      foundAccounts.push({ row: i + 1, data: data[i] }); // 存储行号和数据
    }
  }

  // 返回匹配帐号的数组，每个元素包含行号和数据
  return JSON.stringify(foundAccounts);
}
