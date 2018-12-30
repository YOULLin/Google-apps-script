
function myFunction() {
  var token = "";
  var message = "測試\n\n";
  sendLineNotify(message, token);
}


function sendTOLine(message,token){
  var options = {
    "method" : "post",
    "payload" : {"message" : message},
    "headers" : {"Authorization" : "Bearer " + token}
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

