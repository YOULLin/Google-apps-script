/*****
* @ google試算表案件自動通報程式 
* @ ver 1.0
* @ brief 通報當日未完成的案件狀態。
* @ functions:
*
*    1. 2018/12/23更新: 通報當日未完成案件或未退件案件
*    2. 2018/12/26更新:
*       (1) 分開通報當日未完成案件、未退件案件
*       (2) 未退件案件通報修正為都通報
*                      
*
* @ update 2018/12/26
*****/

/*------案件通報主程式------*/
//取得通報用戶token
var token = "";  //正式群組token
var Mytoken = ""; //公司內個人通報token
var MyGrptoken = ""; //測試Mytoken
//未完成通報->圖資部
function unFinishedCaseNotify(){
  //通報今天是否有未完成案件
  var m1 = CaseStatus(1);
  console.log("m1到底在衝尛!!!"+m1);
  console.log("未完成案件回傳訊息為: "+m1);
  if(m1!=='0'){
    //sendLineNotify(m1,token);
    sendLineNotify(m1,Mytoken);
  }
  else{
    var msg = '今天無未完成的案件！'
    sendLineNotify(msg,MyGrptoken);
  }
}

//未退件通報->新北
function unReturnCaseNotify(){
  //通報今天是否有未退件案件
  var m2 = CaseStatus(2);
  console.log("未退件案件回傳訊息為: "+m2);
  
  if(m2!=='0'){
    sendLineNotify(m2,Mytoken);
  }
  else{
    m2 = '今天無未退件的案件！'
    sendLineNotify(m2,MyGrptoken);
  }
}


/*---回傳當天應完成的案件狀態---*/
function CaseStatus(c){
  var cs = c;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('12月');
  //查看表單目前案件狀態
  if (sheet != null){
     var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
     var actRng = sheet.getActiveRange();
     var rowIndex = actRng.getRowIndex();
     var dueColName = '應完成日期';
     var finishColName = '是否已完成';
     var returnColName = '已退件';
     var qrtnColName = '提問';
     var noColName = '案號';
     var townColName = '行政區';
     var rdColName = '路段名稱';
     var imgColName = '1提問照片檔名';
     var dueCol = headers[0].indexOf(dueColName)+1;
     var finishCol = headers[0].indexOf(finishColName)+1;
     var returnCol = headers[0].indexOf(returnColName)+1;
     var qrtnCol = headers[0].indexOf(qrtnColName)+1
     var noCol = headers[0].indexOf(noColName)+1;
     var townCol = headers[0].indexOf(townColName)+1;
     var rdCol = headers[0].indexOf(rdColName)+1;
     var imgCol = headers[0].indexOf(imgColName)+1;
     var lastRow = LastRow(sheet,noColName);
     console.log("***案件數*** "+lastRow);
     
     var unFinishAry = [];  //存入未完成的案件
     var unReturnAry = [];  //存入未被退件的案件
     
     var now = new Date();
     now.setHours(0,0,0,0);
     var nowDate = Date.parse(now.toDateString()).valueOf();   //取得今日日期毫秒值
     var tomorrow = new Date((new Date()).setDate(now.getDate()+1));
     tomorrow.setHours(0,0,0,0);
     var tomDate = Date.parse(tomorrow.toDateString()).valueOf();   //取得明日日期毫秒值
     console.log(now,tomorrow);
     
     rowIndex = 1;   //定義actRange第一次執行的序列為第1列
     
     //判斷順序: 是否有填入應完成日期>應完成日期是否為今天>是否已完成>是否有退件
     if(dueCol >= 1 && rowIndex >= 1){
       for (i=1; i<=lastRow; i++){
          rowIndex+=1;
          console.log("*----*第"+i+"件*----*");
          var dueValue = sheet.getRange(rowIndex,dueCol).getValue();   //取得表上的應完成日期
          console.log(" 應完成日期: "+dueValue);
          var caseNo = sheet.getRange(rowIndex,noCol).getValue();
          //判斷是否有填入應完成日期
          if( dueValue !== ""){
            var dueDate = Date.parse(dueValue.toDateString()).valueOf();
            var finishCheck = (sheet.getRange(rowIndex,finishCol).getValue()).toString();
            var returnCheck = (sheet.getRange(rowIndex,returnCol).getValue()).toString();
            var caseASK = (sheet.getRange(rowIndex,qrtnCol).getValue()).toString();      //取得該案提問內容
            var caseTown = sheet.getRange(rowIndex,townCol).getValue();                  //取得行政區
            var caseRd = sheet.getRange(rowIndex,rdCol).getValue();                      //取得路段名稱
            var caseImg = sheet.getRange(rowIndex,imgCol).getValue();                    //取得照片檔名，若檔名為空，顯示無檔案
            if(strIsNull(caseImg)==true)
              caseImg = '無檔案'
            //判斷應完成日期是否為今天
            if (dueDate == nowDate){
              console.log("日期是今天!");
              console.log(" 第"+i+"件狀態: "+finishCheck.indexOf("ok")+"\n 該案是否有退件: "+returnCheck);
              //確認案件未完成、也未被退件
              if (finishCheck.toLowerCase().indexOf("ok")==-1 && returnCheck!='已退件'){         

                //無退件原因->案件未完成
                if (strIsNull(caseASK)==true){
                  unFinishAry.push({caseno:caseNo,casetown:caseTown,caserdname:caseRd});
                  console.log("該案未完成 "+unFinishAry[unFinishAry.length-1].caseno+": "+unFinishAry[unFinishAry.length-1].casetown+" "+unFinishAry[unFinishAry.length-1].caserdname);
                }
                //有退件原因->案件未被退件
                else{
                  //casedate==T->今日應退件而未退件
                  unReturnAry.push({casedate:'T',caseno:caseNo,casetown:caseTown,caserdname:caseRd,caseask:caseASK,caseimg:caseImg});
                  console.log("該案未退件 "+unReturnAry[unReturnAry.length-1].caseno+": "+unReturnAry[unReturnAry.length-1].casetown+" "+unReturnAry[unReturnAry.length-1].caserdname);
                }
              }
              //測試已完成案件用
              else{console.log("第"+i+"件案號"+sheet.getRange(rowIndex,noCol).getValue()+" 為已完成的案件");}
            }
            //明日退件: 判斷應完成日期是否為明日日期(今日送件)
            else if(dueDate == tomDate){
              console.log('日期為明天');
              if(finishCheck.toLowerCase().indexOf("ok")==-1 && returnCheck!='已退件'){
                if (strIsNull(caseASK)==true){
                  //明日未完成案件
                }
                //casedate==M->明日應退件而未退件
                else{
                  unReturnAry.push({casedate:'M',caseno:caseNo,casetown:caseTown,caserdname:caseRd,caseask:caseASK,caseimg:caseImg});
                }
              }
            }
          }
          //處理未填寫應完成日期的案件
          else{
              var skipCase = [];
              skipCase.push(caseNo);
          }
       }
       console.log("沒日期的案件共 "+skipCase.length+" 件，案號: "+skipCase[skipCase.length-1]);
     }
  }
  //LINE通報
  switch(cs){
    //本日未完成案件訊息
    case 1:
      if (unFinishAry.length != 0 || unFinishAry != ""){
        var msg1='';
        msg1= msg1.concat('==未完成案件通報==\n','本日應完成而未完成案件共 ',unFinishAry.length,' 件如下，\n\n');
        //未完成案件資訊
        var caseinfo1='';   
        for(var i=0;i<unFinishAry.length;i++){
          caseinfo1 =caseinfo1.concat('  ',unFinishAry[i].caseno,' ',unFinishAry[i].casetown,' ',unFinishAry[i].caserdname,', \n');
        } 
        msg1 = msg1.concat(caseinfo1,'\n以上案件請於當日完成！\n');
        return msg1;
         break;
      }
      //本日無未完成案件
      else{
        var msg = '0';
        return msg;
        break;
      }
    //退件訊息
    case 2:
      if (unReturnAry.length != 0 || unReturnAry != ""){
        var msg2 = '';
        msg2 = msg2.concat('==未退件案件通報==\n','目前尚未退件案件共 ',unReturnAry.length,' 件如下，\n');
        //未退件案件資訊
        var caseinfoT2=Utilities.formatDate( now, "GMT+8", "yyyy/MM/dd").concat(' 未退件： \n');   //今天下班前
        var caseinfoM2=Utilities.formatDate( tomorrow, "GMT+8", "yyyy/MM/dd").concat(' 未退件： \n');   //明天下班前
        for(var i=0;i<unReturnAry.length;i++){
           if(unReturnAry[i].casedate=='T'){
             caseinfoT2 =caseinfoT2.concat('  * ',unReturnAry[i].caseno,' ',unReturnAry[i].casetown,' ',unReturnAry[i].caserdname,
             '，理由: ',unReturnAry[i].caseask,'，照片: ',unReturnAry[i].caseimg,'\n');
           }
           else{
             caseinfoM2 =caseinfoM2.concat('  * ',unReturnAry[i].caseno,' ',unReturnAry[i].casetown,' ',unReturnAry[i].caserdname,
             '，理由: ',unReturnAry[i].caseask,'，照片: ',unReturnAry[i].caseimg,'\n');
           }
        } 
        msg2 = msg2.concat(caseinfoT2,'\n',caseinfoM2,'\n以上案件請盡速協助退件！\n\n');
        return msg2;
        break;
      }
      else{
        var msg = '0';
        return msg;
        break;
      }
  }
}

/*------判斷空值或空字串------*/
function strIsNull(str) {
  if(typeof(str)=="string"){
    var a = str.trim();
    if(a.length==0) {
       return true;
    }
    else {
       return false;
    }
  }
}

/*----當日應完成案件總件數----*/
function LastRow(colsheet, columnname){
  var hd = colsheet.getRange(1,1,1,colsheet.getLastColumn()).getValues();
  var col = hd[0].indexOf(columnname)+1
  var actrange = colsheet.getActiveRange();
  var ri = actrange.getRowIndex();
  var sheetLastRow = colsheet.getLastRow();
  var countRaws = 0;
  for(var j=1;j<=sheetLastRow;j++){
    ri +=1
    var rawValue = (colsheet.getRange(ri,col).getValue()).toString();
    if(strIsNull(rawValue)===false || rawValue !== ""){
       countRaws+=1;
    }
  }
  return countRaws;
}

/*--------LINE通報API--------*/
function sendLineNotify(message,token){
  var options = {
    "method" : "post",
    "payload" : {"message" : message},
    "headers" : {"Authorization" : "Bearer " + token}
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}