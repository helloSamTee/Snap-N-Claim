function createCalendarEvent(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("notApproved");
  var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0];
  console.log(values);

  var file = values[11];
  var fileId = file.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
  console.log(fileId);

  var calenderId = CalendarApp.getDefaultCalendar().getId();
  
  const startDate=new Date(values[4]);
  startDate.setDate(startDate.getDate()+1);
  var eventObj={
    summary: values[1],
    location: values[6],
    description: values[13],
    start:{
      date: startDate.toISOString().split('T')[0],
    },
    end:{
      date: startDate.toISOString().split('T')[0],
    },
    attachments: [{'fileUrl':'https://drive.google.com/open?id='+fileId}]
  };
  event = Calendar.Events.insert(eventObj,calenderId,{'supportsAttachments':true});

  var subject='Expense Reimbursement Request has been Approved';
  var sheet2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  var range=sheet2.getRange('allEmployees').getValues();
  for (var row = 0; row < range.length; row++) {
    for (var col = 0; col < range[row].length; col++) {
      if (range[row][col] === values[0]) {
          var email = range[row][col-1];
        break;
      }
    }
  };

  var name = email.match(/^([^@]+)@[^@]+$/)[1];
  console.log(email);
  var body=`Dear ${name},\n\nYour Expense Reimbursement Request falling under ${values[2]} on ${Utilities.formatDate(values[4],'Asia/Singapore',"dd-MM-yyyy")} at ${values[5]} with total of MYR${values[10].toFixed(2)} has just been approved. \n\nBest regards, \nYour company`;
  console.log(body);
  MailApp.sendEmail(email,subject,body);
}
