function findAccountByname(name,selectedWorksheet) {
  var spreadsheetId = "1bU38faGJ49iUsKAVzdqJLfl-Yq7rq9bD5tlzxWRXkFU";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(selectedWorksheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();

  // 查找名为"備註"的列
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == "備註") {
      var nameColumn = i;
      break;
    }
  }
  // 查找名为"交易分行"的列
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == "交易分行") {
      var bankColumn = i;
      break;
    }
  }
  // 创建一个用于存储匹配帐号的数组
  var foundAccounts = [];

  // 遍历数据范围，查找匹配的帐号
  for (var i = 1; i < data.length; i++) {
    var accountName = data[i][nameColumn]; // 帐号名字在備註列
    var bank = data[i][bankColumn];
    // 检查帐号名字是否包含输入的姓名部分
    if (accountName && accountName.toString().indexOf(name) !== -1) {
      foundAccounts.push({ row: i + 1, data: data[i] }); // 存储行号和数据
    }
    else if (bank && bank.toString().indexOf(name) !== -1) {
      foundAccounts.push({ row: i + 1, data: data[i] }); // 存储行号和数据
    }
  }

  // 返回匹配帐号的数组，每个元素包含行号和数据
  return JSON.stringify(foundAccounts);
}
