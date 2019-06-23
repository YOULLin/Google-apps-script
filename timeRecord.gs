/*****
* @ 試算表控制碼
* @ brief: 限制檢核表部分欄位功能並確保填寫流程。
* @ functions
*    1. 判斷流程正確性(更新人員->是否已完成->是否已上傳)
*    2. 自動記錄案件完成時間及上傳時間
*
* @ update 2019/06/25
*
*****/

function onEdit(e)
{
  
  //取得使用中的試算表
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //判斷使用中的試算表名為"停營科-檢核表"
  if ( ss.getName() == "LineNotify測試2-試算表" )
  {
    
    //取得使用中的頁籤
    var sheet = e.source.getActiveSheet();
    var sheetname = sheet.getName();
    
    //判斷是否取得正確頁籤,若正確就執行以下步驟
    if ( sheetname.match("月") == "月" )
    {
      
      //取得使用者編輯坐標
      var actRng = sheet.getActiveRange();   //Returns the selected range in the active sheet, or null if there is no active range.
      var editCol = actRng.getColumn();      //Returns the starting column position for this range.
      var rowIndex = actRng.getRowIndex();   //Returns the row position for this range.Identical to "getRow()".
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
      
      //使用中試算表頁籤的欄位名稱
      var caseuser = "更新人員";
      var casefinish = "是否已完成";
      var casetime = "案件完成時間";
      var upfinish = "是否已上傳";
      var uptime = "上傳時間";
      
      //取得欄位所在"行"坐標
      var caseuserCol = headers[0].indexOf( caseuser ) + 1;
      var casefinishCol = headers[0].indexOf( casefinish ) + 1;
      var casetimeCol = headers[0].indexOf( casetime ) + 1;
      var upfinishCol = headers[0].indexOf( upfinish ) + 1;
      var uptimeCol = headers[0].indexOf( uptime ) + 1;
      
      //判斷是否有以上欄位,若有則執行記錄步驟
      if ( caseuserCol > 0 && casefinishCol > 0 && casetimeCol > 0 && upfinishCol > 0 && uptimeCol > 0 )
      {
        if (rowIndex > 1)
        {
          var caseuserVal = sheet.getRange(rowIndex,caseuserCol).getValue();
          var casefinishVal = sheet.getRange(rowIndex,casefinishCol).getValue();
          var casetimeVal = sheet.getRange(rowIndex, casetimeCol).getValue();
          var upfinishVal = sheet.getRange(rowIndex, upfinishCol).getValue();
          var uptimeVal = sheet.getRange(rowIndex, uptimeCol).getValue();
          
          //使用者編輯"是否已完成"欄位
          if ( editCol == casefinishCol )
          {
            //強制先填寫更新人員名字
            if ( caseuserVal != "" )
            {
              //資料未上傳時可編輯
              if ( upfinishVal == "0" )
              {
                //完成案件(是否已完成=1)
                if ( casefinishVal == "1" )
                {
                  sheet.getRange(rowIndex, casetimeCol).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss"));
                  Logger.log('casetime success!');
                }
                else if ( casefinishVal == "0" )
                {
                  sheet.getRange(rowIndex, casetimeCol).clearContent(); 
                }
              }
              //資料上傳時,不可取消已完成案件
              else if ( upfinishVal == "1" )
              {
                sheet.getRange(rowIndex, casefinishCol).check();
                SpreadsheetApp.getUi().alert('車格資料已上傳，不可取消已完成案件！');
                Logger.log('upFinish fail!');
              }
            }
            //若沒有寫更新人員,清除原本勾選的完成欄位
            else if ( caseuserVal == "" )
            {
              sheet.getRange(rowIndex, casefinishCol).uncheck();
              SpreadsheetApp.getUi().alert('請先填寫更新人員!');
              Logger.log('userKeyIn Alert!');
            }
          }
          //使用者編輯"是否已上傳"欄位
          else if ( editCol == upfinishCol )
          {
            //有填寫更新人員、案件已完成 才能勾選資料上傳
            if ( caseuserVal != "" && casefinishVal == "1" && casetimeVal != "" )
            {
                if ( upfinishVal == "1" )
                {
                  sheet.getRange(rowIndex, uptimeCol).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd"));
                  Logger.log('uptime success!');
                }
                else if ( upfinishVal == "0" )
                {
                    sheet.getRange(rowIndex, uptimeCol).clearContent(); 
                }
            }
            else if ( caseuserVal == "" || casefinishVal == "0" )
            {
              sheet.getRange(rowIndex, upfinishCol).uncheck();
              SpreadsheetApp.getUi().alert('案件未完成!');
            }
          }
        }
      }
    }
  }
}
