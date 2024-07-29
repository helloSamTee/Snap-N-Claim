function removeFolder() {
  var folderName = Browser.inputBox('Enter the name of the folder you want to delete:');
  
  if (!folderName) {
    Browser.msgBox('No folder name provided. Operation canceled.');
    return;
  }

  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    var folder = folders.next();
    var folderId = folder.getId();

    // Move all files and subfolders to trash before deleting the folder
    moveAllFilesToTrash(folder);
    moveAllSubfoldersToTrash(folder);
    
    // Remove the folder
    DriveApp.getFolderById(folderId).setTrashed(true);
    Browser.msgBox('Folder "' + folderName + '" has been moved to the trash.');

    // Call removeEmpRow function after deleting the folder
    removeEmpRow(folderName);
  } else {
    Browser.msgBox('No folder found with the name "' + folderName + '".');
  }
}


function moveAllFilesToTrash(folder) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
}


function moveAllSubfoldersToTrash(folder) {
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    moveAllFilesToTrash(subfolder);
    moveAllSubfoldersToTrash(subfolder);
    subfolder.setTrashed(true);
  }
}


function removeEmpRow(folderName) {
  var spreadsheetId = '1mmnuZYsTVpDn3vB0RJQMRV643eEHmI7t7ywd85YNtNQ';
  var sheetName = 'Employees'; 
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  
  if (!sheet) {
    Browser.msgBox('Sheet not found.');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] == folderName) { // The folder name is in the second column
      sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
      Browser.msgBox('Row with folder name "' + folderName + '" has been deleted.');
      return;
    }
  }
  
  Browser.msgBox('No row found with folder name "' + folderName + '".');
}
