
var EMAIL_SENT = "BIRTHDAY_WISHES_SENT";

// Image to attach with birthday Wish(random image is attached)
var Image_Urls = {
  1:'https://cdn.filestackcontent.com/SjVthipQb8t0PLD2Un3w',
  2:'https://cdn.filestackcontent.com/XCupultPRTiPtromhpSU',
  3:'https://cdn.filestackcontent.com/jzeEE3JRtK0zWYZKKl1n'
};

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(1, 1, lastRow,3);
  var data = dataRange.getValues();
  for (var i = 1; i < data.length; ++i) {
    var row = data[i];
    var date = row[2];
    var todayDate = new Date();
    var todayDay = parseInt(todayDate.getDate());
    var todayMonth = parseInt(todayDate.getMonth())+1;
    var sheetDate = new Date(date);
    var sheetDay = parseInt(sheetDate.getDate());
    var sheetMonth = parseInt(sheetDate.getMonth())+1;
    if((todayDay == sheetDay) && (sheetMonth == todayMonth)) {
      var name = row[1];
      _sendMail(row[0]);
      sheet.getRange(i+1, 4).setValue(EMAIL_SENT);
      SpreadsheetApp.flush();
    }
  }
}
function _getImage() {
  return UrlFetchApp
  .fetch(Image_Urls[parseInt(Math.random() * (Object.keys(Image_Urls).length - 1) + 1)])
  .getBlob()
  .setName("BirthdayWishes"); 
}
function _sendMail(email) {
    MailApp.sendEmail({
     to: email,
     subject: "Happy Birthday!",
     htmlBody:
      "<!DOCTYPE html>"+
      "<html>"+
        "<body>"+
           "Hey, <br/>"+
      "<p style='font-family: cursive;'>Have a wonderful birthday. I wish your every day to be filled with lots of love, laughter,"+
              "happiness and the warmth of sunshine.</p> <br/>"+
          "<img src='cid:img_url' width='80%' height='90%'> <br/><br>"+
          "<p> --Abhishek Singh</p>"+
        "</body>"+
      "</html>",
     inlineImages:
       {
         img_url: _getImage()
       }
   });
}
