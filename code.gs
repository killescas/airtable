
function createBackup(){
  var folder = DriveApp.getFolderById('IDFOLDER'); 
  let month = Utilities.formatDate(new Date(), "GMT", "MMMMYYYY");
  let nameWeek = Utilities.formatDate(new Date(), "GMT-1", "MM-dd-yyyy");
  let sheetFlag=0;
  let fileFlag=0;
  let fileflagArr=[];
  let activeSpreadsheet;
  let newSheet;
  var files = folder.getFiles();
  var fileidok;
  while (files.hasNext()){
    file = files.next();
    var row = []
    row.push(file.getName(),file.getId(),file.getSize())
    if(row[0]==month){
      fileflagArr.push("Yes");
      activeSpreadsheet = SpreadsheetApp.openById(row[1]);
      fileidok=row[1];
      var allsheets = activeSpreadsheet.getSheets();
      for(var s in allsheets){
          if(allsheets[s].getName()==nameWeek){
              sheetFlag=1;
              break;
          } else {
              sheetFlag=0;
              fileFlag = 0;
          }
      } 
         if(sheetFlag!=1){
                newSheet = activeSpreadsheet.insertSheet();
                newSheet.setName(nameWeek);
                sheetFlag=0;
                fileFlag = 0;
          } else {
                sheetFlag=1;
          }
    } else {
        fileFlag = 1;
        fileflagArr.push("No");
    }
  }
  fileFlag= fileflagArr.indexOf("Yes", 0);

  if(fileFlag<0){
        let fileidnew=createSpreadsheet(month);
        activeSpreadsheet = SpreadsheetApp.openById(fileidnew);
        newSheet = activeSpreadsheet.insertSheet();
        newSheet.setName(nameWeek);
        var firstSheet = activeSpreadsheet.getSheetByName("Sheet1");
        activeSpreadsheet.deleteSheet(firstSheet);
        backup(fileidnew,nameWeek);
        fileFlag=0;
      } else {
        backup(fileidok,nameWeek);
        fileFlag=0;
      }
}

function createSpreadsheet(month){
  var folderId = "IDFOLDER" // Please set the folder ID.
  var fileName = month;
  var newFile = SpreadsheetApp.create(fileName);
  var folder = DriveApp.getFolderById(folderId);
  var file = DriveApp.getFileById(newFile.getId());
  DriveApp.getFolderById("root").removeFile(file);
  folder.addFile(file);
  return (file.getId());
}


function backup(activeSpreadsheet,nameWeek){
  let dateBakcup = new Date();
  let month = Utilities.formatDate(dateBakcup, "GMT", "MMMM");
  
  var url = "https://api.airtable.com/v0/---/----";//URL AIRTABLE BASE API
  let data = getAllRecords(url);

  ss=SpreadsheetApp.openById(activeSpreadsheet);
  var sheet=ss.getSheetByName(nameWeek);
  sheet.getDataRange().clearContent();

  var fieldNames = ["Record ID"];
  for (var i = 0; i<data.length; i++){
    for (var field in data[i].fields){
      fieldNames.push(field);
    }
  }
  fieldNames = fieldNames.filter(function(item, pos){
    return fieldNames.indexOf(item)== pos;
  });
  
  fieldNames.push('backup Date');

  var headerRow = sheet.getRange(1,1,1,fieldNames.length);
  headerRow.setValues([fieldNames]).setFontWeight("bold").setWrap(true);
  sheet.setFrozenRows(1);
  
  //add Airtable record IDs to the first column of each row
  for (var i = 0; i<data.length; i++){
    sheet.getRange(i+2,1).setValue(data[i].id);
  }

  for (var i = 0; i<data.length; i++){
    //iterate through each field in the record
    for (var field in data[i].fields){
      let range = sheet.getRange(i+2,1); 
       range.setValue(data[i].id); 
     range= sheet.getRange(i+2,fieldNames.indexOf(field)+1); //find the cell we want to update
        range.setValue(data[i].fields[field]); //update the cell 
      range=sheet.getRange(i+2,fieldNames.length);
        range.setValue(new Date());
    }
  }

  newSheet = ss.insertSheet();
  var namesheet2 = nameWeek+"HT";
  newSheet.setName(namesheet2);
  sheet= ss.getSheetByName(namesheet2);

  url = "https://api.airtable.com/v0/---/----"; //URL AIRTABLE BASE 2
  let dataHT= getAllRecords(url);
  dateBakcup = new Date();
  
  var fieldNamesHT = ["Record ID"];
  //add every single field name to the array
  for (var i = 0; i<dataHT.length; i++){
    for (var field in dataHT[i].fields){
      fieldNamesHT.push(field);
    }
  }
  fieldNamesHT = fieldNamesHT.filter(function(item, pos){
    return fieldNamesHT.indexOf(item)== pos;
  });
  
  fieldNamesHT.push('backup Date');

  var headerRowHT = sheet.getRange(1,1,1,fieldNamesHT.length);
  headerRowHT.setValues([fieldNamesHT]).setFontWeight("bold").setWrap(true);
  sheet.setFrozenRows(1);
  
  //add Airtable record IDs to the first column of each row
  for (var i = 0; i<dataHT.length; i++){
    sheet.getRange(i+2,1).setValue(dataHT[i].id);
  }

  for (var i = 0; i<dataHT.length; i++){
    //iterate through each field in the record
    for (var field in dataHT[i].fields){
      let range = sheet.getRange(i+2,1); 
       range.setValue(dataHT[i].id); 
     range= sheet.getRange(i+2,fieldNamesHT.indexOf(field)+1); //find the cell we want to update
        range.setValue(dataHT[i].fields[field]); //update the cell 
      range=sheet.getRange(i+2,fieldNamesHT.length);
        range.setValue(new Date());
    }
  }
}

function getAllRecords(url){
  let records = [];
  const initial_response=requestAirTable(url);
  records.push(initial_response.records)
  let offset = initial_response.offset;
  if(offset !== undefined){
    do{
      const response = requestAirTable(url,offset)
      records.push(response.records)
      offset = response.offset
    } while (offset !== undefined)
  }
  records = records.flat();
  console.log('records count:'+records.length);

return records;
}

function requestAirTable(url, offset) {
  const api_key = "----";//API KEY
  if(offset !== undefined){
    url = url + '?offset='+offset

  }
  const headers={
    'Authorization':'Bearer ' + api_key,
    'Content-Type':'application/json'
  }

  const options={
    headers : headers,
    method: 'GET'
  }

  const response = UrlFetchApp.fetch(url,options).getContentText();
  const result = JSON.parse(response);
  return result;
}


