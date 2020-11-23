var properties = PropertiesService.getScriptProperties();
var calendarId = properties.getProperty("calendarID");
var url = properties.getProperty("slackWebhook");

var ss = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var sheet = ss[0];
var CALENDAR_ID = 1;
  
function getCalendarInfo(row) {
  var starttime = Moment.moment(sheet.getRange(row,4).getValue());
  var endtime   = Moment.moment(sheet.getRange(row,5).getValue());
  
  var txt = "タイトル：" + sheet.getRange(row,2).getValue() + "\n" +
            "場所：" + sheet.getRange(row,3).getValue() + "\n";
  
  if(sheet.getRange(row,5).getValue() === ""){ //開始日付のみ表示なら
    txt = txt + "時間：" + starttime.format("YYYY/MM/DD") 
  }
  else { // 時間も表示されていたら
    if(endtime.diff(starttime, "days") >= 1){ //1日以上離れていたら (ex. start:2020/11/04 9:00, end:2020/11/05 9:00 なら、日付は2日間表示
      txt = txt + "時間：" + starttime.format("YYYY/MM/DD HH:mm") + " - " + endtime.format("YYYY/MM/DD HH:mm") + "\n";
    }
    else {
      txt = txt + "時間：" + starttime.format("YYYY/MM/DD HH:mm") + " - " + endtime.format("HH:mm") + "\n";
    }
  }
  txt = txt +  sheet.getRange(row,6).getValue();
  return txt;  
}

function deleteCalendarInfo(row) {
  sheet.deleteRows(row,1);  
}

function updateCalendarInfo(row, item) {
  sheet.getRange(row, 1).setValue(item.iCalUID);  
  sheet.getRange(row, 2).setValue(item.summary);  
  sheet.getRange(row, 3).setValue(item.location);
  
  sheet.getRange(row, 4).setValue(item.start.dateTime);
  if(sheet.getRange(row,4).getValue() === ""){
    sheet.getRange(row, 4).setValue(item.start.date);
  }
  sheet.getRange(row, 5).setValue(item.end.dateTime);
  if(sheet.getRange(row,5).getValue() === ""){
    sheet.getRange(row, 5).setValue(item.end.date);
  }
  
  if(item.description != null) {
    sheet.getRange(row, 6).setValue(item.description.replace("<br>","\n").replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'').replace(/&nbsp;/g,'')); 
  }
  else {
    sheet.getRange(row, 6).setValue(""); 
  }
  
}

function setCalendarList() {
  /* 時刻情報取得 */
  var date_0 = Moment.moment().local().startOf('day');
  var getRangeDate_0 = Moment.moment().add(3, 'months').local().startOf('day');
  var myCal =CalendarApp.getCalendarById(calendarId); 
  var myEvents = myCal.getEvents(date_0.toDate(),getRangeDate_0.toDate()); 
  var to_row = 2;
  
  sheet.getRange(2,1,100).clear();
  myEvents.forEach(function(evt){
    　sheet.getRange(to_row, 1).setValue(evt.getId());  
    　sheet.getRange(to_row, 2).setValue(evt.getTitle());  
    　sheet.getRange(to_row, 3).setValue(evt.getLocation());  
    　sheet.getRange(to_row, 4).setValue(evt.getStartTime());  
    　sheet.getRange(to_row, 5).setValue(evt.getEndTime()); 
    　sheet.getRange(to_row, 6).setValue(evt.getDescription().replace("<br>","\n").replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,'').replace(/&nbsp;/g,'')); 
    　to_row++;    
  });
}

//カレンダー初期起動用(差分取得用)
function initialSync() {
  var events = Calendar.Events.list(calendarId, 
   {
     timeMin: Moment.moment().toISOString(),
     timeMax: Moment.moment().add(3, 'months').local().toISOString(),
   });
  var nextSyncToken = events.nextSyncToken;
  Logger.log("Sync:"+ nextSyncToken);
  properties.setProperty("nextSyncToken", nextSyncToken);
}

// ret: 行番号　、false:0
function findRow(val,col){
  var dat = sheet.getDataRange().getValues();
  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      return i+1;
    }
  }
  return 0;
}

//カレンダー編集時に自動的に起動するプログラム
function onCalendarEdit(e) {  
  var nextSyncToken = properties.getProperty("nextSyncToken");
  Logger.log("pre_Sync:"+nextSyncToken);
  var events = Calendar.Events.list(calendarId, {
    syncToken: nextSyncToken,
  });
  
  var items = events.items;
  Logger.log(items);

  for (var i = 0; i < items.length; i++) {
    var status = items[i].status;
    var eventId = items[i].id +"@google.com";
    var des_row = findRow(eventId, CALENDAR_ID);
    
    if( status === "confirmed") {
      if(des_row != 0){
        updateCalendarInfo(des_row, items[i]);
        sendToSlack("予定が更新されました", getCalendarInfo(des_row),"#bce2e8");
      }
      else {
        sheet.insertRows(2,1);
        updateCalendarInfo(2, items[i]);       
        sendToSlack("新しく予定が追加されました", getCalendarInfo(2),"#bce2e8");
      }
    }
    else if(status === "cancelled") {
      if(des_row != 0){
        sendToSlack("予定がキャンセルされました", getCalendarInfo(des_row),"#e8d8bc");
        deleteCalendarInfo(des_row)
      }
      else {
        sendToSlack("予定キャンセルでエラー発生", items[i],"#e80055");
      }     
    }
   }
     
  var nextSyncToken = events.nextSyncToken
  properties.setProperty("nextSyncToken", nextSyncToken);
}

function noticeCalendarWillStart() {
  var dat = sheet.getDataRange().getValues();
  for(var i=1;i<dat.length;i++){
    var starttime = dat[i][3];
    var diff = Moment.moment(starttime).diff(Moment.moment(), "minutes")
    
    if(diff < 15) {
      var des_row = findRow(dat[i][0], CALENDAR_ID);
      if(des_row != 0){
        sendToSlack("まもなく開始する予定があります", getCalendarInfo(des_row),"#e8bccc");
        deleteCalendarInfo(des_row)
      }
      else {
        sendToSlack("予定通知でエラー発生", dat[i],"#e80055");
      }   
    }
  }
}

function sendToSlack(title, body, color) {
  var data = { 
    "attachments": [{
      "color": color,
      "title": title,
      "text" : body
    }],
  };
  var payload = JSON.stringify(data);
  var options = {
    "method" : "POST",
    "contentType" : "application/json",
    "payload" : payload
  };
  var response = UrlFetchApp.fetch(url, options);
}