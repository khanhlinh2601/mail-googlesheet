var SEARCH_QUERY = "EXE"; //nhập tên của subject mail cần tải vd "CV_K17"

var AVOID_REPEATED_ADDRESS = true;

// Main function
function getMail() {
    SpreadsheetApp.getActiveSheet().clear();  
    console.log(`Searching for: "${SEARCH_QUERY}"`);
    var start = 0;
    var max = 500;  
    var threads = GmailApp.search(SEARCH_QUERY, start, max);
    appendData(1, [["Date Send Mail","Email", "Subject", "Link Drive", "Reply", "Date Reply"]]);
    
    var totalEmails = 0;
    var emails = [];
    var addresses = [];
    while (threads.length>0){
      for (var i in threads) {
          var thread=threads[i];
          var data = thread.getLastMessageDate();
          var msgs = threads[i].getMessages();
          for (var j in msgs) {
            var msg = msgs[j];
            var data = msg.getDate();          
            var from = msg.getFrom();
            var sub = msg.getSubject()
            var to = msg.getTo()
            var regex = /(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})"/g
            var body = msg.getBody().match(regex)

            var dataLine = [data,from,sub,body];

            if (!AVOID_REPEATED_ADDRESS || (AVOID_REPEATED_ADDRESS && !addresses.includes(from))){
              emails.push(dataLine);
              addresses.push(from);
            }
          }
      }

      totalEmails = totalEmails + emails.length;

      appendData(2, emails);

      start = start + max; 
      threads = GmailApp.search(SEARCH_QUERY, start, max);
    }
    function appendData(line, array2d) {
      var sheet = SpreadsheetApp.getActiveSheet();
      sheet.getRange(line, 1, array2d.length, array2d[0].length).setValues(array2d);
    }
}

function sendMail(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter your id to send mail:'); //"CVK17_SE171717"
  var id = response.getResponseText();
  var values = sheet.getRange(2, 3, sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (var i=0; i<values.length; i++){
    if(values[i][0] == id ){
      i+=2;
      var mail=sheet.getRange(i, 2).getValue();
      var reply=sheet.getRange(i, 5).getValue();  
      var date = new Date().toLocaleTimeString();
      sheet.getRange(i, 6).setValue(date); 
    }
  }
  if(response.getSelectedButton() == ui.Button.OK){
    MailApp.sendEmail(mail, "Rep lại nè", reply); //html mail in here, cái này test
  }
}
function sendAll(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getRange(2, 3, sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (var i=2; i<=values.length; i++){
      var mail=sheet.getRange(i, 2).getValue();
      var reply=sheet.getRange(i, 5).getValue();  
      var date = new Date().toLocaleTimeString();
      sheet.getRange(i, 6).setValue(date); 
  }
  MailApp.sendEmail(mail, "Rep lại nè", reply); //html mail in here, cái này test
}


