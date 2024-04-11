function confirmAccountInSheet(index, personnelCode, selectedWorksheet) {
  var spreadsheetId = "1bU38faGJ49iUsKAVzdqJLfl-Yq7rq9bD5tlzxWRXkFU";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(selectedWorksheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();

  var ccount = -1; 
  var crewIndex = -1; // 默认值
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == "確認次數")
      ccount = i;
    else if(headers[i] == "查帳確認人員")
      crewIndex = i;
  }

  var crew = sheet.getRange(index, crewIndex + 1).getValue();
  var today = new Date();
  var formattedDate = (today.getMonth() + 1) + "/" + today.getDate() + ":" + personnelCode + ",";
  var Count = sheet.getRange(index, ccount + 1).getValue();
  crew = crew + formattedDate;

  // 更新电子表格中的确认次数
  sheet.getRange(index, ccount + 1).setValue(Count + 1);
  // 更新电子表格中的人员代号+日期
  sheet.getRange(index, crewIndex + 1).setValue(crew);

  return Count + 1;
}
