
const sheetRequestGrading = SpreadsheetApp.getActive().getSheetByName("Request Grading")
const sheetAll = SpreadsheetApp.getActive().getSheetByName("All CVK17")
const sheetSentMailRecord = SpreadsheetApp.getActive().getSheetByName("Sent Mail Record")
const secondsSinceEpoch = (date) => Math.floor(date.getTime() / 1000);
const after = new Date();
const before = new Date(); //time search mail


if (before.getHours() == 6) {
  before.setHours(6, 00, 0, 0);
  after.setHours(23, 00, 00, 00);
  after.setDate(after.getDate() - 1);
} else if (before.getHours() == 10) {
  before.setHours(10, 00, 0, 0);
  after.setHours(06, 00, 00, 00);
} else if (before.getHours() == 14) {
  before.setHours(14, 00, 0, 0);
  after.setHours(10, 00, 00, 00);
} else if (before.getHours() == 18) { //chinh lai gio hien tai thi demo no moi chay
  before.setHours(18, 00, 0, 0);
  after.setHours(14, 00, 00, 00);
}
const time_search = `after:${secondsSinceEpoch(after)} before:${secondsSinceEpoch(before)}`;
//lấy từ mail về sheets
function getMail() {
  const subject_search = `subject:"^K17 THE 1ST CHALLENGE: CV - *"`; //trim() rồi, ko lo vụ space

  const SEARCH_QUERY = time_search + subject_search;
  Logger.log(SEARCH_QUERY);
  const AVOID_REPEATED_ADDRESS = true;
  const AVOID_REPEATED_ID = true;
  let start = 0;
  let max = 500;
  let threads = GmailApp.search(SEARCH_QUERY, start, max);
  // GmailApp.markThreadsRead(threads);
  let totalEmails = 0;
  let emails = [];
  let addresses = [];
  let id_list = [];
  while (threads.length > 0) {
    for (let i in threads) {
      let thread = threads[i];
      // let data = thread.getLastMessageDate();
      let msgs = threads[i].getMessages();
      for (let j in msgs) {
        let msg = msgs[j];
        // let date = msg.getDate();
        let from = msg.getFrom().match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi)
        let id = msg.getSubject().toUpperCase().trim().slice(-8);
        let regex = /(https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,})"/g
        //test đính kèm link trong mail
        let body = (msg.getBody().match(regex))
        // body=body.toString()
        // body=body.slice(0, body.length-1)
        let dataLine = [from, id, body];
        if (!AVOID_REPEATED_ADDRESS || (AVOID_REPEATED_ADDRESS && !addresses.includes(from))) { //kiểm tra mail nào gửi cuối thì thêm vào sheet
          if (AVOID_REPEATED_ID && !id_list.includes(id)) {
            emails.push(dataLine); //check last id, avoid 2 account mail send 1 id
            addresses.push(from);
            id_list.push(id);
          }
        }
      }
    }
    totalEmails = totalEmails + emails.length;
    appendData(sheetRequestGrading.getLastRow() + 1, emails);
    start = start + max;
    threads = GmailApp.search(SEARCH_QUERY, start, max);
  }

  function appendData(line, array2d) {
    sheetRequestGrading.getRange(line, 2, array2d.length, array2d[0].length).setValues(array2d);
  }

}

function sentMailRecord() {
  let column;
  for (let k = 2; k <= 26; k++) {
    if (sheetSentMailRecord.getRange(192, k).getValue() == 0) { //search last column
      column = k;
      break;
    }
  }
  for (i = 3; i < sheetSentMailRecord.getLastRow(); i++) {
    let id_Search = sheetSentMailRecord.getRange(i, 1).getValue();
    let idTime_Search = time_search + id_Search;
    threads = GmailApp.search(idTime_Search, 0, 200);
    let receivedCount = 0;
    for (j = 0; j < threads.length; j++) {
      receivedCount = receivedCount + threads[j].getMessageCount();
    }
    sheetSentMailRecord.getRange(i, column).setValue(receivedCount);
  }
}


function saveReply() {
  let cv_grader; //lưu lại reply mới nhất vào sheet của người chấm
  let sheetCVGrader
  let name = (new Date().getHours()).toString() + ":00-" + (new Date()).toLocaleDateString();
  let sheetDate = SpreadsheetApp.getActive().getSheetByName("10:00-4/11/2021");
  for (let i = 2; i <= sheetDate.getLastRow(); i++) {
    let lastCol = 1
    cv_grader = sheetDate.getRange(i, 1).getValue();
    sheetCVGrader = SpreadsheetApp.getActive().getSheetByName(cv_grader);
    for (let j = 2; j <= sheetCVGrader.getLastRow(); j++) {
      if (sheetDate.getRange(i, 3).getValue() == sheetCVGrader.getRange(j, 1).getValue()) {
        while (sheetCVGrader.getRange(j, lastCol).getValue() != "") {
          lastCol++;
        }
        sheetCVGrader.getRange(j, lastCol).setValue(sheetDate.getRange(i, 5).getValue())
      }
    }
  }
}

function deleteAll2() {
  for (let i = 2; i <= sheetRequestGrading.getLastRow(); i++) {
    //clear về giá trị ban đầu
    if (sheetRequestGrading.getRange(i, 5).getValue() != "") { //chấm r thì clear
      sheetRequestGrading.getRange(i, 2).clear();
    }
  }
  let lastRow = sheetRequestGrading.getLastRow();
  for (let i = lastRow; i > 0; i--) { //delete blank row
    let data = sheetRequestGrading.getRange(i, 2).getValue();
    if (data == "") {
      sheetRequestGrading.deleteRow(i);
    }
  }
}
// function deleteAll2(){
//   let sheetDate = SpreadsheetApp.getActive().getSheetByName("Request Grading");
//   let r=sheetDate.getRange('E:E');
//   let v=r.getValues();
//   for(let i=v.length-1; i>=0; i--){
//     if(v[0,i]==!""){
//       sheetDate.deleteRow(i+1);
//     }
//   }
// }
function setVLOOKUPK16Grader() {
  let val;
  for (let i = 2; i <= sheetRequestGrading.getLastRow(); i++) {
    //clear về giá trị ban đầu
    val = ""
    val = "=VLOOKUP(C";
    val += i;
    val += "; 'All CVK17'!$A$2:$B$325; 2; FALSE)";
    sheetRequestGrading.getRange(i, 1).setValue(val);
  }
}
function myFunc1() {
  let pro = new Promise((res) => {
    res(saveRequestGranding());
  });
  pro.then(() => {
    deleteAll2();
    //  saveReply();
  })
}
function myFunc2() {
  let pro = new Promise((res) => {
    res(getMail());
  });
  pro.then(() => {
    sendReplyNow();
  });
  pro.then(() => {
    setVLOOKUPK16Grader();
    sentMailRecord();
    // saveReply();
  });
}

function saveRequestGranding() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let name1 = (new Date().getHours()).toString() + ":00-" + (new Date()).toLocaleDateString() + "(1)";
  let name = (new Date().getHours()).toString() + ":00-" + (new Date()).toLocaleDateString()
  // Logger.log(name)
  let test = ss.getSheetByName("Test");
  sheetRequestGrading.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(name1);
  let source = ss.getSheetByName(name1).getRange("A200:F400");
  let filter = `=FILTER('Request Grading'!A1:F200; 'Request Grading'!E1:E200<>"")`
  ss.getSheetByName(name1).getRange(200, 1).setValue(filter)
  test.copyTo(ss).setName(name);
  var destSheet = ss.getSheetByName(name);
  var destRange = destSheet.getRange("A1:F199");
  // let filter1="=LEN("+i+")"
  for (let i = 2; i < 100; i++) {
    let filter1 = `=LEN(E${i})`
    ss.getSheetByName(name).getRange(i, 7).setValue(filter1)
  }
  // ss.getSheetByName(name1).getRange(2,7).setValue(filter1)
  source.copyTo(destRange, { contentsOnly: true });
  // source.clear();
}




function sendReplyNow() {
  let htmlTemplate = HtmlService.createTemplateFromFile("Template.html");
  let subject = `FEEDBACK CV - THE 1ST CHALLENGE`;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let name = (new Date().getHours()).toString() + ":00-" + (new Date()).toLocaleDateString();
  let destSheet = ss.getSheetByName(name);
  // Logger.log(destSheet);
  for (let i = 2; i <= destSheet.getLastRow(); i++) {
    let recipient = destSheet.getRange(i, 2).getValue();
    let reply = destSheet.getRange(i, 5).getValue();
    let dateReply = new Date().toLocaleTimeString();
    htmlTemplate.reply = reply;
    if (reply != "") {
      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        body: '',
        name: 'Fcode CLB',
        htmlBody: htmlTemplate.evaluate().getContent()
      })
      destSheet.getRange(i, 8).setValue(dateReply); //ngày giờ reply
    }

  }
}
