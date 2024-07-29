async function extractDataFromReceipt(fileArrs) {
  var fileIDs = []

  // Extract only the file IDs of uploaded documents
  for (file of fileArrs){
    fileIDs.push(file[1]);
  } 

  const apiKey = PropertiesService.getScriptProperties().getProperty("API_KEY"); // Please set your API key.  

  // Query passed to Gemini
  const jsonSchema = { 

    title: 

      "Array including JSON object parsed the following images of the invoices", 

    description: 

      "Create an array including JSON object parsed the following images of the invoices.", 

    type: "array", 

    items: { 

      type: "object", 

      properties: { 

        fileid: { 

          description: "Name given as 'Filename'", 

          type: "string", 

        }, 

        invoiceNumber: { 

          description: "Number of the invoice (Invoice ID)", 

          type: "string", 

        }, 

        invoiceDate: { 

          description: "Date of invoice", 

          type: "string", 

        }, 

        business: { 

          description: "Name of business that generated the invoice", 

          type: "string", 

        }, 

        invoiceLocation: { 

          description: "Complete address of the business that generated the invoice. Address usually ends at the country name, such as Malaysia and Singapore. Replace the '\n' characters encountered in the address text with a comma followed by a space (', '). ", 

          type: "string", 

        }, 

        paymentMethod: { 

          description: "Payment method performed by the customer, such as debit card, credit card, and cash.", 

          type: "string", 

        }, 

        totalCost: { 

          description: "Total cost of all costs. Remove the currency (if there is any) and only return the numerical value.", 

          type: "string", 

        }, 

      }, 

      required: [ 

        "fileid", 

        "invoiceNumber", 

        "invoiceDate", 

        "business", 

        "invoiceLocation", 

        "paymentMethod", 

        "totalCost", 

      ], 

      additionalProperties: false, 

    }, 

  }; 

  const g = GeminiWithFiles.geminiWithFiles({ 

    apiKey, 

    doCountToken: true, 

    response_mime_type: "application/json", 

  }); 

  const fileList = await g.setFileIds(fileIDs, true).uploadFiles(); 

  const res = g 

    .withUploadedFilesByGenerateContent(fileList) 

    .generateContent({ jsonSchema }); 

  g.deleteFiles(fileList.map(({ name }) => name)); // If you want to delete the uploaded files, please use this. 

  console.log(res)
  
  for (i in res){
    if (res[i] != "No values."){
      var unique = checkUniqueReceipt(res[i].invoiceNumber, "notApproved", "notRecords", 3, true);

      // If there is no repeating receipt ID in the employee's unapproved records
      if (unique){
        SpreadsheetApp.getActiveSpreadsheet().toast('A file has been uploaded to Google Drive');
        // Add the new record to notApproved sheet
        populateReceiptRow(res[i], fileArrs[i], "notApproved", "notRecords");
      } else {
        // Remove the uploaded document from Google Drive
        var file = DriveApp.getFileById(fileIDs);                                    
        file.setTrashed(true);
        SpreadsheetApp.getActiveSpreadsheet().toast("File removed from Google Drive due to the existence of receipt with same ID."); 
      }
    } else {
      var file = DriveApp.getFileById(fileIDs);                                    

      SpreadsheetApp.getUi().alert(
        "Error Occured",
        `Please try to upload the file ${file.getName()} again.`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );

      file.setTrashed(true);
    }
  }
  
} 


function checkUniqueReceipt(newID, sheet, tableRange, col, upload){
  if (newID == null) { return true;} // Receipt ID is not extracted by Gemini

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const range = ss.getRangeByName(`${sheet}!${tableRange}`);

  if(range){
    var startRow = range.getRow();
    var numRow = range.getNumRows();

    if (numRow == 1) {return true;} // No unapproved record 

    var existingIDs = ss.getSheetByName(sheet).getRange(startRow, col, numRow - 1, 1).getValues().flat();
    console.log(existingIDs);

    if (upload) {
      // When uploading new receipt to the system
      var repeat = existingIDs.findIndex(cell => cell == newID);
      if (repeat == -1) {
        return true;
      } else {
        SpreadsheetApp.getUi().alert("File Upload Rejected", `Receipt with the same ID (${newID}) is found in the claim sheet. Please try to upload another file.`, SpreadsheetApp.getUi().ButtonSet.OK);
        return false;
      }
    } else {
      // When changing the receipt ID of existing record
      var seen = new Set();
      for (elem of existingIDs){
        if (elem != "") {
          if (seen.has(elem)){
            SpreadsheetApp.getUi().alert("Receipt ID Modification Failed", `Receipt with the same ID (${newID}) is found in the claim sheet. Please try entering another receipt ID.`, SpreadsheetApp.getUi().ButtonSet.OK);
            return false;
          } else {
            seen.add(elem);
          }
        }
      }
      return true;
    }    
  }
}


function populateReceiptRow(result, fileArr, sheetName, tableRange){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const table = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(`${sheetName}!${tableRange}`);
  
  const targetRow = table.getRow();
  const lastRow = table.getNumRows()//range.getLastRow() - targetRow + 3;
  const lastCol = table.getNumColumns();

  // Add new row under header
  sheet.insertRowBefore(targetRow);

  // Copy formats and data validations to the row which new unapproved record will be added into
  const rowBelow = sheet.getRange(table.getLastRow(), 1, 1, lastCol)
  const rowAbove = sheet.getRange(targetRow, 1, 1, lastCol)
  rowBelow.copyTo(rowAbove)

  // Populate information about the record into the new row 
  sheet 
    .getRange(targetRow, 1, 1, 1)
    .setValue("");
  sheet 
    .getRange(targetRow, 2, 1, 1)
    .setValue(fileArr[0]);
  sheet 
    .getRange(targetRow, 3, 1, 1)
    .setValue(result.invoiceNumber);
  sheet 
    .getRange(targetRow, 4, 1, 1)
    .setValue(result.invoiceDate);
  sheet 
    .getRange(targetRow, 5, 1, 1)
    .setValue(result.business);
  sheet 
    .getRange(targetRow, 6, 1, 1)
    .setValue(result.invoiceLocation);
  sheet 
    .getRange(targetRow, 7, 1, 1)
    .setValue(result.paymentMethod);
  sheet 
    .getRange(targetRow, 8, 1, 1)
    .setValue("");
  sheet 
    .getRange(targetRow, 9, 1, 1)
    .setValue(result.totalCost);
  sheet 
    .getRange(targetRow, 10, 1, 1)
    .setValue("");
  sheet 
    .getRange(targetRow, 11, 1, 1)
    .setValue(fileArr[2]);
  sheet 
    .getRange(targetRow, 12, 1, 1)
    .setValue(fileArr[3]);
  
  sheet 
    .getRange(targetRow, 1, 1, lastCol)
    .setBackground("#fff2cc");

  // Add one row to the employee's range of unapproved records 
  sheet
    .getNamedRanges()
    .find(namedRange => namedRange.getName() === tableRange)
    .setRange(
      sheet.getRange(targetRow, 1, lastRow + 1, lastCol)
    ); 

  SpreadsheetApp.flush();
}
  

function addToSheet(project, url, date){
  //regex to find the id
  var fileId=url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
 
  // Create an array [project, id, date]
  var newArr = [project, fileId, url, date];
 
  // Please set file IDs of PDF file of invoices.
  const fileIds = [newArr];
  console.log(fileIds)
 
  extractDataFromReceipt(fileIds);
}
