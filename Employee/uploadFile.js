function onOpen(){
  // Add custom menu 
  let ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Delete Record')
    .addItem('Delete a Record that is Not Yet Approved', 'del')
    .addToUi();

  ui.createMenu("Upload")
    .addItem("Upload File", "showSidebar")
    .addToUi();
}


function showSidebar(){
  var template = HtmlService.createTemplateFromFile("sidebar");
  var userInterface = template.evaluate().setTitle("Upload File");
  SpreadsheetApp.getUi().showSidebar(userInterface);
}


function doGet(){
  var template = HtmlService.createTemplateFromFile("sidebar");
  return template.evaluate().setTitle("Upload File");
}


function saveData(obj) {
  var currentSheet=SpreadsheetApp.getActiveSpreadsheet().getName();
  // RegEx to find folder name based on spreadsheet name
  var parentFolder=DriveApp.getFoldersByName(currentSheet.match(/^(.*?)'s\s/)[1]).next();
  var folderName=obj.input1;
  var folders = parentFolder.getFoldersByName(folderName);
  var file;
  var rowData = [
    obj.input1 
  ];

  if (folders.hasNext()){
    var folder=folders.next();
  }else{;
    var folderId= parentFolder.createFolder(folderName).getId();
    var folder=DriveApp.getFolderById(folderId);
  }

  // Check if there is a project folder inside employee's folder
  // If no, create one 
  if (obj.uploadFile) {
      Object.keys(obj.uploadFile).sort().forEach(key => {
        Logger.log(key)
        let files = obj.uploadFile[key]
        let datafile = Utilities.base64Decode(files.data)
        let blob = Utilities.newBlob(datafile, files.type, files.name);
        file = folder.createFile(blob).getUrl()
        // Append file URL and upload time to the rowData array
        rowData.push(file);
        rowData.push(obj.timestamp);
      })
    }
  console.log(rowData);
  addToSheet(rowData[0],rowData[1],rowData[2])
};


function updateProject(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  var row=sheet.getLastRow();
  if (row<=1){
    var value=[];
  }
  else{
    var range = sheet.getRange(2, 1, sheet.getLastRow()-1, 1);
    var value=range.getValues().flat();
  }
  return value;
}


function updateSheet(text){
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  var cell=sheet.getRange(sheet.getLastRow()+1,1);
  cell.setValue(text);
}
