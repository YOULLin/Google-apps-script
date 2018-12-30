/*****
* @ 停營科檢核表表單內容限制
* @ ver 1.0
* @ brief 就。。表單限制
* @ functions:
*
*    1. 2018/12/25更新:
*       (1) 使用者輸入提問時，須清空ok欄位->此限制沒執行會影響未完成通報
*                      
*
* @ update 2018/12/25
*****/

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if( sheet.getName() == "12月" && sheet != null){
    ask();
  }
  else if (sheet.getName() == "NotifySheet") //"NotifySheet" is the name of the sheet where you want to run this script.
  {
    var timeColumn = "自動填入時間"
    var notifyColumn = "問題通報"
    var actRng = sheet.getActiveRange();   //Returns the selected range in the active sheet, or null if there is no active range.
    var editColumn = actRng.getColumn();   //Returns the starting column position for this range.
    var rowIndex = actRng.getRowIndex();   //Returns the row position for this range.Identical to "getRow()".
    //getRange(): Returns the range with the top left cell at the given coordinates with the given number of rows and columns.
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var dateCol = headers[0].indexOf( timeColumn ) + 1;
    var orderCol = headers[0].indexOf( notifyColumn ) + 1;
    if (dateCol > 0 && rowIndex > 1 && editColumn == orderCol)
    {
      sheet.getRange(rowIndex, dateCol).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"));
    }
  }
  else{
    console.log('非以上兩張表');
  }
}

//提問時同時清空欄位
function ask(){
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues();
  var actRng = sheet.getActiveRange();
  var rowIndex = actRng.getRowIndex();
  var editCol = e.range.getColumn();
  var finishColName = '是否已完成';
  var resubmitColName = '提問';
  var finishCol = headers[0].indexOf(finishColName)+1;
  var resubmitCol = headers[0].indexOf(resubmitColName)+1;
  var resubmitVale = sheet.getRange(rowIndex,resubmitCol)
  if( editCol == resubmitCol && strIsNull(e.value)==false){
    sheet.getRange(rowIndex,finishCol).clearContent();
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