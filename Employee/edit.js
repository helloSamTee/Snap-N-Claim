function onEdit(e) {
  // check if it is in defined range
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notApproved");
  const cell = e.range.getCell(1, 1);
  const row = cell.getRow();
  const col = cell.getColumn();

  // Check whether it is in the "notApproved" sheet
  if (e.range.getSheet().getName() == "notApproved"){
    // Check if it is in the range for unapproved records
    if (checkInRange(row, col, "notApproved", "notRecords")){
          
      // "Currency" column
      if (col == 8){
      calculateMYR(row, col, sheet)  
      }

      // "Receipt ID" column
      if (col == 3){
        if(!(checkUniqueReceipt(cell.getValue(), "notApproved", "notRecords", 3, false))){
          cell.setValue("");
        }
      }

      // Change background colour of the modified cell to yellow
      cell.setBackground("#fff2cc");
    }
  }
}


function checkInRange(cellRow, cellCol, sheetName, tableRange) {

  const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(`${sheetName}!${tableRange}`);

  // Check whether the cell is in the named range
  if(range){
    return cellRow >= range.getRow() && cellRow <= range.getLastRow() && cellCol >= range.getColumn() && cellCol <= range.getLastColumn();
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast("Named range not found");
  }
}


function delRow(rowNum, sheetName, tableRange){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const range = ss.getRangeByName(`${sheetName}!${tableRange}`);
  
  const lastRow = range.getLastRow();
  const copyRow = lastRow - rowNum;
  const lastCol = range.getNumColumns();

  // Find the file ID of supprting document and remove it from employee's folder
  var url = sheet.getRange(rowNum, 11, 1, 1).getValue();
  var id = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
  var file = DriveApp.getFileById(id);                                    
  file.setTrashed(true);

  // Move up the records below the deleted record 
  const below = sheet.getRange(rowNum + 1, 1, copyRow, lastCol);
  const above = sheet.getRange(rowNum, 1, copyRow, lastCol);
  below.copyTo(above)
  
  // Update the range for unapproved records
  sheet
    .getNamedRanges()
    .find(namedRange => namedRange.getName() === tableRange)
    .setRange(
      sheet.getRange(range.getRow(), 1, range.getNumRows() - 1, lastCol)
  );

  // Clear content, formats, and data validations of the last row
  sheet.getRange(lastRow, 1, 1, lastCol).clear();
  sheet.getRange(lastRow, 1, 1, lastCol).clearDataValidations();

  SpreadsheetApp.flush();
}


function del(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notApproved");
  const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("notApproved!notRecords");

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter Row Number', 'Which row of record you would like to delete? Please enter a valid integer value. (Hint: Row 6 is the row of latest record)', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var num;

    try {
      // Attempt to parse the string into an integer
      num = parseInt(response.getResponseText(), 10); // 10 specifies base 10 (decimal)
      
      if (isNaN(num)) {
        throw new Error('Conversion failed: Input is not a valid integer. Please try again!');
      } else if (num < 1){
        throw new Error('Invalid input: Input must be an positive integer. Please try again!');
      } else if (!(range.getRow() <= num && num < range.getLastRow())) {
        throw new Error('Invalid input: Input row number does not hold any record. Please try again!');
      }

      // Confirm employee's decision to delete record
      var recordToDel = sheet.getRange(num, 3, 1, 1).getValue();
      var confirm = ui.alert('Delete Confirmation', `Record about receipt ${recordToDel} will be deleted. Are you sure you want to delete row ${num}?`, ui.ButtonSet.YES_NO);

      if (confirm === ui.Button.YES) {
        delRow(num, "notApproved", "notRecords");
        ui.alert("Record and corresponding document deleted successfully.", ui.ButtonSet.OK);
      } else {
        ui.alert('Record deletion canceled.');
      }

    } catch (error) {
      // Handle the error
      ui.alert(error.message);
    }

  } else {
    ui.alert('Record deletion canceled.');
  }
}
