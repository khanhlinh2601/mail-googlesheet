function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('page');
}
function updateRecords(id, note, grade) { 
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("Request Grading"); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 3).getValue().trim().toUpperCase() == id)
    {
      dataSheet.getRange(i, 5).setValue(note)
      dataSheet.getRange(i, 6).setValue(grade)
    }
  } 
}

function searchRecords(username) 
{
  let returnRows = [];
  let allRecords = getRecords();
  allRecords.forEach(function(value, index) {
    let evalRows = [];
    if(username != '')
    {
      if(value[1].toUpperCase().trim() == username.toUpperCase().trim()) {
        evalRows.push('true');
      } else {
        evalRows.push('false');
      }
    }
    else
    {
       evalRows.push('true');
    }
    if(evalRows.indexOf("false") == -1)
    {
      returnRows.push(value);    
    }
  });
  // Logger.log(returnRows);
  return returnRows;
}

function searchRecordsRequest(username) 
{
  let returnRows = [];
  let allRecords = getRequestGrading();
  allRecords.forEach(function(value, index) {
    let evalRows = [];
    if(username != '')
    {
      if(value[0].toUpperCase().trim() == username.toUpperCase().trim()) {
        evalRows.push('true');
      } else {
        evalRows.push('false');
      }
    }
    else
    {
       evalRows.push('true');
    }
    if(evalRows.indexOf("false") == -1)
    {
      returnRows.push(value);    
    }
  });
  return returnRows;
}

                    
function getRecords() { 
  let return_Array = [];
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("All CVK17"); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue() != '')
    {
      return_Array.push([dataSheet.getRange(i, 1).getValue(),
      dataSheet.getRange(i, 2).getValue(), 
      dataSheet.getRange(i, 4).getValue(),
      dataSheet.getRange(i, 5).getValue(),
      dataSheet.getRange(i, 6).getValue(), 
      dataSheet.getRange(i, 7).getValue(),
      dataSheet.getRange(i, 8).getValue()]);
    }
  }
  return return_Array;  
}
function getRequestGrading() { 
  let return_Array = [];
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("Request Grading"); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue() != '')
    {
      return_Array.push([dataSheet.getRange(i, 1).getValue(),
      dataSheet.getRange(i, 2).getValue(), 
      dataSheet.getRange(i, 3).getValue(),
      dataSheet.getRange(i, 4).getValue(),
      dataSheet.getRange(i, 5).getValue(), 
      dataSheet.getRange(i, 6).getValue(),
      ]);
    }
  }
  return return_Array;  
}
function getReplyHistory(username) { 
  let return_Array = [];
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(username); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue() != '')
    {
      return_Array.push([dataSheet.getRange(i, 1).getValue(),
      dataSheet.getRange(i, 2).getValue(), 
      dataSheet.getRange(i, 3).getValue(),
      dataSheet.getRange(i, 4).getValue(),
      dataSheet.getRange(i, 5).getValue(), 
      dataSheet.getRange(i, 6).getValue(),
      dataSheet.getRange(i, 7).getValue(),
      dataSheet.getRange(i, 8).getValue(),
      dataSheet.getRange(i, 9).getValue(),
      dataSheet.getRange(i, 10).getValue(),
      ]);
    }
  }
  return return_Array;    
}
function searchReplyHistory(username, studentId) 
{
  let return_Array = [];
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(username); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue() != '')
    {
      return_Array.push([dataSheet.getRange(i, 1).getValue(),
      dataSheet.getRange(i, 2).getValue(), 
      dataSheet.getRange(i, 3).getValue(),
      dataSheet.getRange(i, 4).getValue(),
      dataSheet.getRange(i, 5).getValue(), 
      dataSheet.getRange(i, 6).getValue(),
      dataSheet.getRange(i, 7).getValue(),
      dataSheet.getRange(i, 8).getValue(),
      dataSheet.getRange(i, 9).getValue(),
      dataSheet.getRange(i, 10).getValue(),
      dataSheet.getRange(i, 11).getValue()
      ]);
    }
  }
  let returnRows = [];
  let allRecords = return_Array;
  allRecords.forEach(function(value, index) {
    let evalRows = [];
    if(studentId != '')
    {
      if(value[0].toUpperCase().trim() == studentId.toUpperCase().trim()) {
        evalRows.push('true');
      } else {
        evalRows.push('false');
      }
    }
    else
    {
       evalRows.push('true');
    }
    if(evalRows.indexOf("false") == -1)
    {
      returnRows.push(value);    
    }
  });
  return returnRows;
}
function updateNoteAllCV(id, note) { 
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("All CVK17"); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue().trim().toUpperCase() == id)
    {
    
      dataSheet.getRange(i, 8).setValue(note)
    }
  }
}
function updatePass(id, pass) { 
  let ss= SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("All CVK17"); 
  let getLastRow = dataSheet.getLastRow();
  for(i = 2; i <= getLastRow; i++)
  {
    if(dataSheet.getRange(i, 1).getValue().trim().toUpperCase() == id)
    {
      dataSheet.getRange(i, 7).setValue(pass)
    }
  }
}
