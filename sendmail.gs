function sendMail() {
  var SS_ID = "スプレッドシートID";
  var TARGET_SHEET_MAIL = "シート名";

  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(TARGET_SHEET_MAIL);
  var range = sh.getRange('B:B').getValues(); 
  var lastRow = range.filter(String).length;

  var toAdr = [];
  var ccAdr = [];
  var bccAdr = [];
  var title = sh.getRange("L11").getValue();
  var body = sh.getRange("L12").getValue();
  var EnglishGreeting = sh.getRange("L3").getValue();
  var ChineseGreeting = sh.getRange("L4").getValue();
  var deliveryDateTime = "";
  var moment = Moment.moment();

  //メール本文の置き換え
  body = body.replace("ChineseGreeting",ChineseGreeting); 
  body = body.replace("EnglishGreeting",EnglishGreeting); 

  //当月確認
  for(i=1; i<=lastRow; i++){
    var deliveryValue = sh.getRange(i + 3,2).getValue(); 
    var Month = new Date().getMonth() + 1 + "月";
    if (moment.isSame(deliveryValue[i],'hour')) {
      body = body.replace("Month",Month); //本文中に記載した月の書き換え
   }
  }

  //メール宛先振り分け
  for(i=1; i<=lastRow; i++){
    var sendToValue = sh.getRange(i + 3,7).getValue()
    if (sendToValue == "to") {
       toAdr.push(sh.getRange(i + 3,9).getValue()); 
       body = body.replace("sendTo",(sh.getRange(i + 3,8).getValue())); 
    }
    if (sendToValue == "cc") ccAdr.push(sh.getRange(i + 3,9).getValue()); 
    if (sendToValue == "bcc") bccAdr.push(sh.getRange(i + 3,9).getValue()); 
  }

  //メール送付時間になったら送付する
  if (moment.isSame(deliveryDateTime,'hour')) {
    MailApp.sendEmail({to:toAdr[0], cc:ccAdr[0], bcc:bccAdr[0], subject:title, body:body});
    console.log("\nmailto: " + toAdr + "\n" + ccAdr + "\n" + bccAdr + "\n" +
             "title: " + title + "\n" +
             "body: " + body + "\n"
               );
   }
}
