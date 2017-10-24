function onOpen() {
  var submenu = [{name:"Send ALL Invoices!", functionName:"exportAllSheets"},{name:"Send Just Active Invoice!", functionName:"exportSingleSheet"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Project Admin', submenu);
}

function exportAllSheets() {
  //gets your whole spreadsheet
  var sheets = SpreadsheetApp.getActive().getSheets();
  var now = new Date();
  var weekOf = new Date(sheets[0].getRange(1,1).getValue());
  var by = Session.getActiveUser();

  var ui = SpreadsheetApp.getUi();

  var checkI = 20;
  var check = true;
  var runDates = [];

  while(check){
    if(!sheets[0].getRange(checkI, 6).isBlank()){
      runDates.push(new Date(sheets[0].getRange(checkI, 7).getValue()).toLocaleDateString("en-US"));
      checkI++;

    } else {
      check = false;
    }
  }

  if(runDates.indexOf(weekOf.toLocaleDateString("en-US")) != -1){
    var response = ui.alert('uh-oh...','Looks like I was already ran for the week of ' + weekOf.toLocaleDateString("en-US") + '. Do you want to proceed?', ui.ButtonSet.YES_NO)
    if (response == ui.Button.NO) {
      return;
    }
  }

  sheets[0].getRange(1,4).setValue(now); //last ran:
  sheets[0].getRange(2,4).setValue(weekOf); //for week of:
  sheets[0].getRange(3,4).setValue(by); //by

  var historyI = 20; //starts looking at row 20
  var running = true;

  while(running){
    if(sheets[0].getRange(historyI,6).isBlank()){ //if Col E row (starting at) 20 is blank
      sheets[0].getRange(historyI,6).setValue(now); // add some stuff
      sheets[0].getRange(historyI,7).setValue(weekOf);
      sheets[0].getRange(historyI,8).setValue(by);
      running = false;
    } else {
      historyI++; //else move down
    }
  }

  //loop through sheets/drivers (sheet[0] and sheet[1] we will never care about)
  for (var i = 2; i < sheets.length; i++){

    //remembers the original spreadsheet
    var originalSpreadsheet = sheets[i];

    //if they got paid this week
    if (parseInt(originalSpreadsheet.getRange("C42:C42").getValues()) > 0){

      //these variables are really nicely named
      var emailTo = originalSpreadsheet.getRange("A3:A3").getValues();
      var firstName = originalSpreadsheet.getRange("A4:A4").getValues();
      var lastName = originalSpreadsheet.getRange("B4:B4").getValues();
      var datetime = originalSpreadsheet.getRange("B2:B2").getValues().toString().split(' ');
      var subject = "Urbanstems Driver Invoice for " + firstName + " " + lastName + " -- Week of " + datetime[1] + " " + datetime[2] + ", " + datetime[3];
      var message = "Please see attached! Have a nice day :)";

      //creates new blank sheet
      var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");

      //copies static values to new sheet
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet = originalSpreadsheet;
      sheet.copyTo(newSpreadsheet);

      //deletes "Sheet 1" artifact from new sheet
      newSpreadsheet.getSheetByName('Sheet1').activate();
      newSpreadsheet.deleteActiveSheet();

      //gets computed/referenced values
      var sourceRange = sheet.getRange(1, 1, 45, 6);
      var sourcevalues = sourceRange.getValues();

      //sets computed/referenced values to new sheet
      var destRange = newSpreadsheet.getActiveSheet().getRange(1, 1, 45, 6);
      destRange.setValues(sourcevalues);

      //waits for copying to finish before proceeding
      SpreadsheetApp.flush();

      //creates PDF
      var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
      var attach = {
        fileName:'Weekly Status.pdf',
        content:pdf,
        mimeType:'application/pdf'
      };

      //if they have an email


      if(emailTo){
        //email them
        MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
      }

      //delete the new spreadsheet because we dont care about it anymore
      DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
    }
  }

}

function exportSingleSheet() {

  var originalSpreadsheet = SpreadsheetApp.getActive();
  var emailTo = originalSpreadsheet.getRange("A3:A3").getValues();
  var firstName = originalSpreadsheet.getRange("A4:A4").getValues();
  var lastName = originalSpreadsheet.getRange("B4:B4").getValues();
  var datetime = originalSpreadsheet.getRange("B2:B2").getValues().toString().split(' ');
  var subject = "Urbanstems Driver Invoice for " + firstName + " " + lastName + " -- Week of " + datetime[1] + " " + datetime[2] + ", " + datetime[3];
  var message = "Please see attached! Have a nice day :)";

  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet = originalSpreadsheet.getActiveSheet();
  sheet.copyTo(newSpreadsheet);

  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();

  var sourceRange = sheet.getRange(1, 1, 45, 6);
  var sourcevalues = sourceRange.getValues();

  var destRange = newSpreadsheet.getActiveSheet().getRange(1, 1, 45, 6);
  destRange.setValues(sourcevalues);

  SpreadsheetApp.flush();

  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();

  var attach = {
    fileName:'Weekly Status.pdf',
    content:pdf,
    mimeType:'application/pdf'
  };

  if(emailTo){
    MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
  }

  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
}
