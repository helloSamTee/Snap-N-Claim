# Snap-N-Claim

## Overview
This is source code in Google App Script for implementing Snap N Claim, an automated expense reimbursement add-ons in Google Spreadsheet. 

Snap N Claim is a technological business solution specially crafted for SMEs that have the need of business travels. Integration with other Google Workspace Tools such as Gemini, Google Drive, Google Calendar and Gmail increases the automation and scalability of Snap N Claim. 

Credits are to be given to ‘tanaikech’ on github as his library of ‘GeminiWithFiles’ and ‘PDFApp’ are deployed in our source code. Repository link are attached : https://github.com/tanaikech/GeminiWithFiles?tab=readme-ov-file 

## Feature
Here is what you can achieve through Snap N Claim source code:

### Content Upload
Receipts generated from business trips can be uploaded in pdf and image format through sideBar in Google Spreadsheet

### File Management
Receipts are uploaded into Google Drive, each inside respective folders tagged by Project Options
Each employee have own Google Drive Folder consisting of Employee Claim Spreadsheet, and different project folders that store receipts for better organization purpose

### Information Extraction
Gemini API is called to extract relevant information automatically from receipts uploaded by employee

### Data Validation
To prevent expense reimbursement fraud, data validation is performed on receipt ID to ensure no duplicate receipts are uploaded and claimed more than once
Data validation is done on Project ID and Employee Name as well to ensure no duplicate values exists for these two data

### Data Synchronization
Information extracted from receipts are used to populate rows in Employee Claim Spreadsheet and Admin Approval Spreadsheet for further CRUD operations
Updates made by Employee inside Employee Unapproved Sheet are always reflected on Admin Unapproved Sheet as well upon opening the spreadsheet
Request approved by Admin inside Admin Unapproved Sheet is moved to Admin Approved Sheet and no further edits are allowed. The respective record in Employee Unapproved Sheet will be moved to Approved Sheet as well
Gray boxes in Employee Unapproved Sheet means blank cell, yellow boxes means edited cell that has not updated to Admin Spreadsheet and white boxes means existing and updated cell

### Integration with other Google Workspace Tools
Gmail is used to send folder access link to new employees and inform employees on success of expense reimbursement request
Google Calendar is used to store events on date and location of receipts for record purpose of admin

### Dealing with multi-currency
Since business trips involve traveling internationally, Snap N Claim can deal with over 100 currencies all around the world and convert total cost in receipt into MYR.

## Usage
In order to implement Snap N Claim add-on, please do the following steps:

### Download Template
Copy the Employee Spreadsheet and Admin Spreadsheet Template consisting of app script code into your Google Drive using the link below: 
Employee Spreadsheet[https://docs.google.com/spreadsheets/d/1BmKe1mw9uTBaHyyFk1z_ZRIcUbcqtnG_-2VLlw1HCE0/copy?usp=sharing] 
Admin Spreadsheet:  [https://docs.google.com/spreadsheets/d/1mmnuZYsTVpDn3vB0RJQMRV643eEHmI7t7ywd85YNtNQ/copy?usp=sharing]

### Create API Key
Access https://makersuite.google.com/app/apikey and create your API key. Enable Generative Language API at the API console. After that, replace the API Key in Employee Spreadsheet line 8.

### Create Installable Trigger
On the top menu bar of Google Spreadsheet, click Extensions> App Script, then you will see the source code of the extensions. On left side bar, click Trigger> Add a trigger> Fill in the details> Save


### Give Access
Opening the add-on the first time will generate a pop up asking for permissions stating  the script is from an untrusted source. Click Review Permissions > Advanced > Go to ‘Name’ (unsafe) > Check and allow all the permissions needed. Then, you’re good to go!

 

## Instructions
### Admin:

Upon opening Admin Spreadsheet, three sheets are listed ‘notApproved’, ‘Approved’ and ‘Employees’.

#### Approve Request
As Admin Spreadsheet is opened, the first default sheet showing in “notApproved” sheet which keep records of all unapproved expense reimbursement request made by all employee. Admin check all the details and approve the request by ticking checkboxes in column O. Upon approving, the record is moved to ‘Approved’ sheet, email is sent and calendar event is created.

#### Add Employee
‘Employees’ consist of records of all ‘employee email’, ‘employee folder name’ and ‘employee folder id’. To add new employee, admin click on the button showing ‘Add New Employee’ and fill in valid name and email address. A google drive folder and spreadsheet will be created, email will be sent to the stated email address attached with folder access link. 

### Employee:

Upon opening the Employee Spreadsheet, three sheets are listed ‘notApproved’, ‘Approved’ and ‘Projects’.

#### Upload Receipt
To upload receipt, click ‘Upload’ on menu bar of spreadsheet and a sideBar will be opened. Choose project option in dropdown list, and click ‘Upload’ to choose receipt to be uploaded from device. Once done, click ‘Submit’ and new row of record will be populated. 

#### Edit Receipt
After new record is added, certain columns such as ‘Purpose/Tag’, ‘Description’ and any missing information that does not exist in the receipt has to be filled in by the employee. Gray cells symbolizes empty value, yellow cells symbolizes edited but not updated value while white cells symbolizes valid and updated value.

#### Add Project
In case of needing to add new project handled by employee, click ‘Upload’ on menu bar of spreadsheet, click ‘Add Project’ and fill in non-duplicate and valid Project Number. Then click ‘Done’ and refresh the sideBar, new project options will appear in the dropdown list.

## Scopes
The source code uses the following scopes:
“https://www.googleapis.com/auth/calendar”
“https://www.googleapis.com/auth/script.send_mail”
“https://www.googleapis.com/auth/spreadsheets”
“https://www.googleapis.com/auth/drive”
“https://www.googleapis.com/auth/script.container.ui”
“https://www.googleapis.com/auth/script.external_request”
“https://www.googleapis.com/auth/presentations”


## Methods
Below are the methods used in Employee Spreadsheet:

extractDataFromReceipt(fileArrs)
Calling Gemini API to extract information such as receipt ID, receipt Date, business Name, business Location, payment Method and total Cost by taking in parameters of fileArrs containing multiple fileId
checkUniqueReceipt(newID, sheet, tableRange, col, main)
Check duplication of receipt IDs
populateReceiptRow(result, fileArr, sheetName, tableRange, copyFormula)
Insert values of result generated by Gemini Ai together with input value from sideBar into new row in Spreadsheet. Update named ranges value and copy formula to new row
addToSheet(project, url, date)
Extract url into fileId through regex expression and pass fileIds to function extractDataFromReceipt(fileArrs)
onEdit(e)
To check if the edited cell is within ‘records’ named range and ensure receiptID cannot be changed
delRow(rowNum, sheetName, tableRange)
Delete row of record
del()
Get user’s input on desired row to be deleted
createCurrencyDropdown()
Create currency dropdown list in named ranges
calculateMYR(cellRow, cellCol, sheet)
Calculate MYR equivalent of total cost in specified currency
checkInRange(cellRow, cellCol, sheetName, tableRange)
Check if edits made by user are within the ‘records’ named range
onOpen()
To show new Menu options on menu bar upon opening the spreadsheets
showSidebar()
Show sidebar made up of html template on right side of spreadsheet
doGet()
Access HTML service and create template from html file
saveData(obj)
Upload receipt file onto project folders and create new folders if the project folder has not been created yet. Call addToSheet() function and passing parameter rowData consisting of project options, file url and date 
updateProject()
Getting and return list of project options from ‘Projects’ sheet
updateSheet(text,index)
Set the next row of project list to be value of parameter text
populateOption(projects)
Receive parameter projects, an array consisting of all existing project options and create as option in dropdown list for each value
showFileIcon(fileID, imgID)
Get format of file uploaded and display respective image or icon
addProject()
onClick function for ‘Add Project’ button to create a new text box and ‘Done’ button 
appendList(text)
Receive text box’s value as parameter ‘text’ and check if it consist of value or text
checkDuplicate(projects)
Receive array ‘projects’ and check if it duplicates with the Project Option being entered
done()
Remove the text box, ‘Done’ button and display ‘Project Added’
submitData()
Ensure project option and file uploaded are not null, then run script function saveData(obj) and passing obj which consist of project option, file and time


Below are the methods used in Admin Spreadsheet:
createFolder()
Get employee name and email and check for duplications. If no, then create folder and spreadsheet and send the folder access link to email entered
updateEmpList(email, folderName, fileId)
Update employee email, folder Name and file ID in ‘Employees’ sheet
onOpen(e)
To show new Menu options on menu bar upon opening the spreadsheets
refresh()
To refresh the exchange between ‘Unapproved’ sheet to ‘Approved’ sheet and perform data validation between records
populateAllRecord(recordArr, sheetName, tableRange)
Copy targeted row into ‘Approved’ sheet and update named ranges
approveRecord(rowNum)
Called after checkbox is checked to called further actions such as createCalendarEvent(row) and moveApprovedInEmp(rowNum, record, approvedDate)
moveApprovedInEmp(rowNum, record, approvedDate)
Find the approved records by admin in ‘Unapprove’ sheet in employee spreadsheet and move to ‘Approve’ sheet
delInNotApproved(rowNum, spreadsheetID, sheetName, tableRange)
Delete approved records by admin in ‘Unapprove’ sheet in employee spreadsheet
addToApproved(result, date, spreadsheetID, sheetName, tableRange)
Insert new records into ‘Approve’ sheet in admin spreadsheet and update the named ranges
checkApproved(e)
To call checkInRange(cellRow, cellCol, tableRange) and call approveRecord(row) if the edited column is on checkbox columns
checkInRange(cellRow, cellCol, tableRange)
To ensure that edited cell is within ‘allRecords’ named range and  
checkNoEmpty(row)
Ensure that all columns of the record are filled with values
createCalendarEvent(row)
Create calendar all day event with attachments, location, description in default calendar and send email to employee to inform on the approval of expense reimbursement request
removeFolder()
Delete google drive folder of removed employee
moveAllFilesToTrash(folder)
Delete files inside google drive folder of removed employee
moveAllSubfoldersToTrash(folder)
Delete folders inside google drive folder of removed employee
removeEmpRow(folderName)
Delete row of records of removed employee in ‘Employees’ sheet of admin

