function onOpen(e) {
  // Retrieve the latest records automatically when the file is opened
  refresh();

  // Add custom menu 
  let ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Refresh All Records')
    .addItem('Retrieve the latest records from all employee', 'refresh')
    .addToUi(); 
};


function refresh(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  var notApprovedSheet = ss.getSheetByName("notApproved");
  var notApprovedRange = ss.getRangeByName(`notApproved!allNotRecords`);
  
  // If there are existing records that are not yet approved, clear the records
  if (notApprovedRange.getNumRows() > 1) {
    // Copy formats and data validations of the range for records that are not yet approved
    var notApprovedOri = notApprovedSheet.getRange(notApprovedRange.getLastRow(), 1, 1, notApprovedRange.getNumColumns());
    var notApprovedFirst = notApprovedSheet.getRange(notApprovedRange.getRow(), 1, 1, notApprovedRange.getNumColumns());
    notApprovedOri.copyTo(notApprovedFirst);

    // Clear existing records in admin main sheet 
    var clearRange = notApprovedSheet.getRange(notApprovedRange.getRow()+1, 1, notApprovedRange.getNumRows()-1, notApprovedRange.getNumColumns());
    clearRange.clear();
    clearRange.clearDataValidations();

    // Update the range for records that are not yet approved
    notApprovedSheet
      .getNamedRanges()
      .find(namedRange => namedRange.getName() === `allNotRecords`)
      .setRange(
        notApprovedSheet.getRange(notApprovedRange.getRow(), 1, 1, notApprovedRange.getNumColumns())
      );
  }

  // Repeat the processes performed above (line 15th - 37th) on approved records
  var approvedSheet = ss.getSheetByName("Approved");
  var approvedRange = ss.getRangeByName(`allApprovedRecords`);

  if (approvedRange.getNumRows() > 1) {
    var approvedOri = approvedSheet.getRange(approvedRange.getLastRow(), 1, 1, approvedRange.getNumColumns());
    var approvedFirst = approvedSheet.getRange(approvedRange.getRow(), 1, 1, approvedRange.getNumColumns());
    approvedOri.copyTo(approvedFirst);

    var clearRange = approvedSheet.getRange(approvedRange.getRow()+1, 1, approvedRange.getNumRows()-1, approvedRange.getNumColumns());
    clearRange.clear();
    clearRange.clearDataValidations();
    
    approvedSheet
      .getNamedRanges()
      .find(namedRange => namedRange.getName() === `allApprovedRecords`)
      .setRange(
        approvedSheet.getRange(approvedRange.getRow(), 1, 1, approvedRange.getNumColumns())
      );
  }

  const empRange = ss.getRangeByName(`Employees!allEmployees`);
  const allEmp = empRange.getValues();
  allEmp.pop();

  var allApproved = [];
  var allNotApproved = [];
  var repeatAlert = false;

  // Approved record has higher accountability as compared to the unapproved ones
  // Hence, collect all the approved records then the unapproved ones
  
  // Retrieve approved records of all employees listed
  for (emp of allEmp){
    // Open the employee's claim sheet
    var empSs = SpreadsheetApp.openById(emp[2]);

    var empApproved = empSs.getRangeByName("Approved!approvedRecords").getValues();
    empApproved.pop();
    var repeat, anyEmpty;

    for (approved of empApproved){
      repeat = allApproved.findIndex(cell => cell[3] == approved[2]);
      anyEmpty = approved.findIndex(cell => cell == "");

      // If the receipt has a unique ID
      if ((repeat == -1)){
        // and if the record holds all the required information, 
        if (anyEmpty == -1) {
          // add the record to allApproved array
          allApproved.push([emp[1], ...approved]);
        } else {
          SpreadsheetApp.getUi().alert("Missing Information in Approved Record", `Please resolve the issue on the record about receipt ${approved[2]} with ${emp[0]} (Folder Name: ${emp[1]}). This approved receipt will not be added into admin's latest list of approved records.`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
        }
      } else {
        SpreadsheetApp.getUi().alert("Repeating Receipt ID in Approved Records", `Please resolve the issue on the record about receipt ${approved[2]} with ${emp[0]} (Folder Name: ${emp[1]}). This approved receipt will not be added into admin's latest list of approved records.`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      }
    }
  }

  // Retrieve records of all employees listed that are not yet approved
  for (emp of allEmp){
    var empSs = SpreadsheetApp.openById(emp[2]);

    var repeat1, repeat2, anyEmpty;

    var empNotApproved = empSs.getRangeByName("notApproved!notRecords");
    var empNotSheet = empSs.getSheetByName("notApproved");
    var notFirstRow = empNotApproved.getRow();
    var notNumRow = empNotApproved.getLastRow() - notFirstRow;
    var notNumCol = empNotApproved.getLastColumn();
    var notCurRow;

    for (i=0; i<notNumRow; i++){
      notCurRow = empNotSheet.getRange(notFirstRow + i, 1, 1, notNumCol).getValues()[0];
      console.log(notCurRow);
      repeat1 = allApproved.findIndex(cell => cell[3] == notCurRow[2]);
      repeat2 = allNotApproved.findIndex(cell => cell[3] == notCurRow[2]);
      anyEmpty = notCurRow.findIndex(cell => cell == "");

      if ((repeat1 == -1) && (repeat2 == -1)){
        if (anyEmpty == -1) {
          allNotApproved.push([emp[1], ...notCurRow]);
          // In employee's claim sheet, change the background colour for records updated to admin main sheet
          empNotSheet.getRange(notFirstRow + i, 1, 1, notNumCol).setBackground("white");
        }
      } else {
        // In employee;s claim sheet, change the background colour for record with repeating receipt ID to red 
        empNotSheet.getRange(notFirstRow + i, 1, 1, notNumCol).setBackground("red");
        repeatAlert = true;
      }
    }
  }

  if (allApproved.length > 0) {populateAllRecord(allApproved, "Approved", "allApprovedRecords");}
  if (allNotApproved.length > 0) {populateAllRecord(allNotApproved, "notApproved", "allNotRecords");}

  if (repeatAlert) {
    SpreadsheetApp.getUi().alert("Repeating Receipt ID in Not Approved Records", `The employee(s) is/are informed regarding his/her records with repeating receipt IDs. The receipt(s) is/are not added into admin's latest list of records.`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  }
}


function populateAllRecord(recordArr, sheetName, tableRange){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const range = ss.getRangeByName(`${sheetName}!${tableRange}`);
  
  const targetRow = range.getRow();
  const numRow = recordArr.length;
  var lastCol = range.getNumColumns();

  const oriRow = sheet.getRange(targetRow, 1, 1, lastCol)

  // Copy formats and data validations to the rows which records will be added into
  for (i=1; i <= numRow; i++){
    var copiedRow = sheet.getRange(targetRow + i, 1, 1, lastCol)
    oriRow.copyTo(copiedRow)
  }
  
  // Sort records
  recordArr.sort((a, b) => b[4] - a[4]);

  // Populate records
  sheet 
    .getRange(targetRow, 1, numRow, recordArr[0].length)
    .setValues(recordArr);
  
  // Update named range
  sheet
    .getNamedRanges()
    .find(namedRange => namedRange.getName() === tableRange)
    .setRange(
      sheet.getRange(targetRow, 1, numRow + 1, lastCol)
    );

  SpreadsheetApp.flush();
}

  
function checkApproved(e) {  
  const range = e.range;
  const cell = range.getCell(1, 1);
  const row = cell.getRow();
  const col = cell.getColumn();
  
  // Check whether it is in the "notApproved" sheet
  if (range.getSheet().getName() == "notApproved")
    // Check if it is in the range for unapproved records
    if (checkInRange(row, col, "notApproved", "allNotRecords")){
          
      // If the "Approved" column is checked
      if (col == 15){
        if(cell.getValue()){
          approveRecord(row);
        }
      }
    }
}


function checkInRange(cellRow, cellCol, sheetName, tableRange) {

  const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(`${sheetName}!${tableRange}`);

  // Check whether the cell is in the named range
  if(range){
    return cellRow >= range.getRow() && cellRow < range.getLastRow() && cellCol >= range.getColumn() && cellCol <= range.getLastColumn();
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast("Named range not found");
  }
}


function approveRecord(rowNum){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notApproved");
  var ui = SpreadsheetApp.getUi();
  var record = sheet.getRange(rowNum, 1, 1, 14).getValues()[0];

  // Confirm admin's decision to approve a record
  var comfirmApprove = ui.alert('Approval Confirmation', `Are you sure you want to approve the record about receipt ${record[3]}?`, ui.ButtonSet.YES_NO);

  if (comfirmApprove === ui.Button.YES) {
    createCalendarEvent(rowNum);

    var approvedDate = new Date();
    console.log(record);

    // If the record is found, update the approval status in the in employee's claim sheet  
    if (moveApprovedInEmp(record, approvedDate)){
      // Update the approval status in the in admin main sheet  
      addToApproved(record, approvedDate, null, "Approved", "allApprovedRecords");
      delInNotApproved(rowNum, null, "notApproved", "allNotRecords");
    } else {
      SpreadsheetApp.getUi().alert("Receipt Record Not Found in Employee's Records", `The employee might have modified/removed the receipt record. Please consider to refresh the records and try again.`, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    }
  } else {
    ui.alert('Record approval cancelled.');
  }  
}


function moveApprovedInEmp(record, approvedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empRange = ss.getRangeByName(`Employees!allEmployees`);
  const allEmp = empRange.getValues();
  allEmp.pop();
  // Find the employee that uploaded this newly approved record
  const emp = allEmp.findIndex(cell => cell[1] == record[0]);

  // Open the employee's claim sheet
  var empSs = SpreadsheetApp.openById(allEmp[emp][2]);

  var empNotRange = empSs.getRangeByName("notApproved!notRecords");
  var empNotApproved = empNotRange.getValues();
  empNotApproved.pop();

  // Find the row of this newly approved record in employee's range of unapproved records 
  const rowInNot = empNotApproved.findIndex(cell => cell[2] == record[3]);

  // If found, move the newly approved record from notApproved to Approved sheet
  if (rowInNot != -1) {
    addToApproved(record.slice(1), approvedDate, allEmp[emp][2], "Approved", "approvedRecords");
    delInNotApproved(empNotRange.getRow() + rowInNot, allEmp[emp][2], "notApproved", "notRecords");
    return true;
  } else {
    return false;
  }
}


function delInNotApproved(rowNum, spreadsheetID, sheetName, tableRange){
  var ss;
  // If spreadsheetID is provided, this function will open and use that particular spreadsheet
  // If not, this function will use the current spreadsheet
  if (spreadsheetID) {
    ss = SpreadsheetApp.openById(spreadsheetID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  const sheet = ss.getSheetByName(sheetName);
  const range = ss.getRangeByName(`${sheetName}!${tableRange}`);

  const lastRow = range.getLastRow();
  const copyRow = lastRow - rowNum;
  const lastCol = range.getNumColumns();

  // Move up the records below the newly approved record 
  const below = sheet.getRange(rowNum + 1, 1, copyRow, lastCol);
  const above = sheet.getRange(rowNum, 1, copyRow, lastCol);
  below.copyTo(above)
  
  // Update the range for records that are unapproved 
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


function addToApproved(result, date, spreadsheetID, sheetName, tableRange){
  var ss;
  // If spreadsheetID is provided, this function will open and use that particular spreadsheet
  // If not, this function will use the current spreadsheet
  if (spreadsheetID) {
    console.log(spreadsheetID);
    ss = SpreadsheetApp.openById(spreadsheetID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  const sheet = ss.getSheetByName(sheetName);
  const table = ss.getRangeByName(`${sheetName}!${tableRange}`);
  
  const targetRow = table.getRow();
  const lastRow = table.getNumRows();
  const lastCol = table.getNumColumns();

  // Add new row under header
  sheet.insertRowBefore(targetRow);

  // Copy formats and data validations to the row which new record will be added into
  const rowBelow = sheet.getRange(table.getLastRow(), 1, 1, lastCol)
  const rowAbove = sheet.getRange(targetRow, 1, 1, lastCol)
  rowBelow.copyTo(rowAbove)
  
  // Populate information about the newly approved record into the new row 
  sheet 
    .getRange(targetRow, 1, 1, lastCol)
    .setValues([[...result, date]]);

  // Update the range for approved records 
  sheet
    .getNamedRanges()
    .find(namedRange => namedRange.getName() === tableRange)
    .setRange(
      sheet.getRange(targetRow, 1, lastRow + 1, lastCol)
    );

  SpreadsheetApp.flush();
}
