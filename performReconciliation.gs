function formatDate(dateStr) {
      if (dateStr.length == 7) {
        var year =  1911 + parseInt(dateStr.substring(0, 3));
        var month = dateStr.substring(3, 5);
        var day = dateStr.substring(5, 7);
        return year + "/" + month + "/" + day;
      }
    }

function performReconciliation(selectedPayment, selectedWorksheet,personnelCode) {

  // 使用 selectedPayment 参数
  // 这里可以根据需要在函数中使用 selectedPayment

  // 你的现有代码，用于读取工作表数据
  var worksheetId = "1bU38faGJ49iUsKAVzdqJLfl-Yq7rq9bD5tlzxWRXkFU";
  var worksheet = SpreadsheetApp.openById(worksheetId).getSheetByName(selectedWorksheet);
  var workheaders = worksheet.getRange(1, 1, 1, worksheet.getLastColumn()).getValues()[0];
  var workdata = worksheet.getDataRange().getValues();

  var paymentsheetId = "13juebF8-M7QWnjfeWYvGKxP2XJX9Ve7Cio3tLdcDitY";
  var paymentsheet = SpreadsheetApp.openById(paymentsheetId).getSheetByName(selectedPayment);
  var paymentheaders = paymentsheet.getRange(1, 1, 1, paymentsheet.getLastColumn()).getValues()[0];
  var paymentdata = paymentsheet.getDataRange().getValues();

  // 在这里可以使用 headers 和 data 进行进一步的处理，以获取已匯款和待確認的帐号信息
  for (var i = 0; i < paymentheaders.length; i++) {
    if (paymentheaders[i] == "付款方式") {
      paymentColumn = i;
    } else if (paymentheaders[i] == "總金額=商品總額-總折扣+運費+信用卡服務費+超商條碼服務費+營業稅") {
      paymentamountColumn = i;
    } else if (paymentheaders[i] == "匯款後五碼") {
      payment5Column = i;
    } else if (paymentheaders[i] == "備註") {
      noteColumn = i;
    } else if (paymentheaders[i] == "Flag") {
      flagColumn = i;
    }
  }
  // 在这里可以使用headers和data进行进一步的处理，以获取已匯款和待確認的帐号信息
  for (var i = 0; i < workheaders.length; i++) {
    if (workheaders[i] == "存入金額") {
      bankpaymentColumn = i;
    } else if (workheaders[i] == "確認次數") {
      confirmColumn = i;
    } else if (workheaders[i] == "對方帳號") {
      work5Column = i;
    } else if (workheaders[i] == "Flag") {
      workflagColumn = i;
    } else if (workheaders[i] == "帳務日") {
      workdateColumn = i;
    }
  }

  var paidAccounts = [];
  var pendingAccounts = [];

  for (var i = 1; i < paymentdata.length; i++) {
    if(paymentdata[i][paymentColumn] == '匯款' &&(paymentdata[i][flagColumn] == '' || paymentdata[i][flagColumn] == 0)&& paymentdata[i][payment5Column]!= ''){
      var plast5 = paymentdata[i][payment5Column].toString().slice(-5); 
      var pamount = paymentdata[i][paymentamountColumn];//付款單金額
      // 检查帐号名字是否包含输入的姓名部分
      for(var j = 1; j < workdata.length; j++){
        if((workdata[j][confirmColumn] == '' || workdata[j][confirmColumn] == 0)&&(workdata[j][workflagColumn] == '' || workdata[j][workflagColumn] == 0)){
          var wlast5 = workdata[j][work5Column].toString().slice(-5);//交易明細末五碼
          var wamount = workdata[j][bankpaymentColumn];//交易明細金額
          if(plast5 == wlast5 && pamount == wamount){
            paidAccounts.push({ row: j + 1, data: workdata[j] }); 
            var flagValue = 1; // 加上flag1
            worksheet.getRange(j + 1, workflagColumn + 1).setValue(flagValue);
            var noteValue =   formatDate(workdata[j][workdateColumn].toString())+ '匯款國泰後五碼' + wlast5+ '$ 金額' + wamount + '*'+ personnelCode;//
            paymentsheet.getRange(i + 1, noteColumn + 1).setValue(noteValue);
          }
        }
      }
    }
     else{
          pendingAccounts.push({ row: i + 1, data: paymentdata[i] });
        }
  }
  // 构建结果对象
  var results = {
    paid: paidAccounts,
    pending: pendingAccounts
  };

  // 返回结果给前端
  return results;
}

// 调用函数并传递参数


