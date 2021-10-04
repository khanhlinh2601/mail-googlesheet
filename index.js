let SEARCH_QUERY = []; //nhập tên của subject mail cần tải vd "CV_K17"

let AVOID_REPEATED_ADDRESS = true;

// Main function
function getMail() {
    SpreadsheetApp.getActiveSheet().clear();  
    let start = 0;
    let max = 500;  
    let threads = GmailApp.search(SEARCH_QUERY, start, max);
    appendData(1, [["Date Send Mail","Email", "Subject", "Link Drive", "Reply", "Date Reply"]]);   
    let totalEmails = 0;
    let emails = [];
    let addresses = [];
    while (threads.length>0){
      for (let i in threads) {
          let thread=threads[i];
          let data = thread.getLastMessageDate();
          let msgs = threads[i].getMessages();
          for (let j in msgs) {
            let msg = msgs[j];
            let data = msg.getDate();          
            let from = msg.getFrom();
            let sub = msg.getSubject()
            let to = msg.getTo()
            let regex = /(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})"/g
            let body = msg.getBody().match(regex)
            let dataLine = [data,from,sub,body];
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
      let sheet = SpreadsheetApp.getActiveSheet();
      sheet.getRange(line, 1, array2d.length, array2d[0].length).setValues(array2d);
    }
}

function sendMailById(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Enter your id to send mail:'); //"CVK17_SE171717"
  let id = response.getResponseText();
  let values = sheet.getRange(2, 3, sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (let i=0; i<values.length; i++){
    if(values[i][0] == id ){
      i+=2;
      let mail=sheet.getRange(i, 2).getValue();
      let reply=sheet.getRange(i, 5).getValue();  
      let date = new Date().toLocaleTimeString();
      sheet.getRange(i, 6).setValue(date); 
    }
  }
  if(response.getSelectedButton() == ui.Button.OK){
    MailApp.sendEmail(mail, "Rep lại nè", reply); //html mail in here, cái này test
  }
}


function sendAll(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let values = sheet.getRange(2, 3, sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (let i=2; i<=values.length; i++){
      let mail=sheet.getRange(i, 2).getValue();
      let reply=sheet.getRange(i, 5).getValue();  
      let date = new Date().toLocaleTimeString();
      sheet.getRange(i, 6).setValue(date); 
  }
  MailApp.sendEmail(mail, "Rep lại nè", reply); //html mail in here, cái này test
}


