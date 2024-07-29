function createFolder() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt the user for the folder name
  var folderResponse = ui.prompt('Enter Folder Name', 'Please enter the name for the new folder:', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's folder name response
  if (folderResponse.getSelectedButton() == ui.Button.OK) {
    var folderName = folderResponse.getResponseText();
    
    if (folderName) {

      // Prompt the user for the employee email
      var emailResponse = ui.prompt('Enter Employee Email', 'Please enter the employee email:', ui.ButtonSet.OK_CANCEL);
      
      // Process the user's email response
      if (emailResponse.getSelectedButton() == ui.Button.OK) {
        var employeeEmail = emailResponse.getResponseText();
        
        if (employeeEmail) {

          // Check if the folder already exists
          var folders = DriveApp.getFoldersByName(folderName);
          if (folders.hasNext()) {
            ui.alert('A folder with the name "' + folderName + '" already exists.');
          } else {

            // ID of the existing spreadsheet to be copied
            var existingSpreadsheetId = '1BmKe1mw9uTBaHyyFk1z_ZRIcUbcqtnG_-2VLlw1HCE0';
            
            // Create new folder in Google Drive
            var folder = DriveApp.createFolder(folderName);

            // Copy the existing spreadsheet
            var existingFile = DriveApp.getFileById(existingSpreadsheetId);
            var copiedFile = existingFile.makeCopy(folderName + '\'s Main Sheet', folder);

            // Share the folder with the employee email
            folder.addEditor(employeeEmail);

            // Protect the active sheet, then remove all other users from the list of editors.
            var sheet = SpreadsheetApp.openById(copiedFile.getId()).getSheetByName("Approved");
            var protection = sheet.protect().setDescription('View-only Approved Record(s)');

            // Ensure the admin is an editor before removing others. Otherwise, if the user's edit
            // permission comes from a group, the script throws an exception upon removing the group.
            protection.addEditor("hello0101wworld@gmail.com");
            protection.removeEditors(protection.getEditors());
            if (protection.canDomainEdit()) {
              protection.setDomainEdit(false);
            }

            //Send a customized email notification
            var subject = 'You have been granted access to the your Employee Main Sheet';
            var body = 'Dear ' + folderName + ',\n\n' +
                       'You have been granted access to your Main Sheet where you can now upload new claim requests and track their approval. You can access it using the following link:\n' +
                       copiedFile.getUrl() + '\n\n' +
                       'Thank you.';

            MailApp.sendEmail(employeeEmail, subject, body);

            // Update employee list
            updateEmpList(employeeEmail, folderName, copiedFile.getId());

            SpreadsheetApp.getActiveSpreadsheet().toast("Folder '" + folderName + "' and a copy of the spreadsheet created successfully! Shared with " + employeeEmail);
          }
        } else {
          ui.alert('Employee email cannot be empty.');
        }
      } else {
        ui.alert('Employee email prompt canceled.');
      }
    } else {
      ui.alert('Folder name cannot be empty.');
    }
  } else {
    ui.alert('Folder creation canceled.');
  }
}


function updateEmpList(email, folderName, fileId){
  const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const empRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(`Employees!allEmployees`);
  
  const targetRow = empRange.getRow();
  const lastRow = empRange.getNumRows() + 1;
  const lastCol = empRange.getNumColumns();

  // Add new row under header
  empSheet.insertRowBefore(targetRow);

  // Populate information about the employee into the new row 
  empSheet 
    .getRange(targetRow, 1, 1, 3)
    .setValues([[email, folderName, fileId]]);
  
  // Add one row to the employee range
  empSheet
    .getNamedRanges()
    .find(namedRange => namedRange.getName() === "allEmployees")
    .setRange(
      empSheet.getRange(targetRow, 1, lastRow, lastCol)
    );

  SpreadsheetApp.flush();
}
