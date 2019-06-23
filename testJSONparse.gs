var requestURL = "http://data.ntpc.gov.tw/api/v1/rest/datastore/382000000A-000077-002";

var thisyear = '2019'   //要取得的資料年份
var calendar = [];



function Work(){
  twCalendar(thisyear);
}

function twCalendar(thisyear){
  var year = thisyear;
  var rawdata = JSON.parse(UrlFetchApp.fetch(requestURL));
  if (rawdata.success == true){
    var day = rawdata.result.records;
    for(var j in day){
      var date = day[j].date;
      if(date.slice(0,4)==year){
        if(day[j].isHoliday=='是'){
          calendar.push({date: date, isholiday: true});
          Logger.log('是假日: '+calendar[calendar.length-1].date);
        }
        else if(day[j].isHoliday=='否'){
          calendar.push({date: date, isholiday: false});
          Logger.log('是平日: '+calendar[calendar.length-1].date);
        }
      }
    }
  }
}

