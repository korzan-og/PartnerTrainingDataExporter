function CreateBBITrainingData() {
  var partnerName = ''; // Partner Name
  const reportid = ''; // Id for the Master Training report
  const id= ''; // Id for the new spreadsheet that will be shared Id is the after the url https://docs.google.com/spreadsheets/d/<ID>

  // Global Variables
  var spreadsheet = SpreadsheetApp.openById(reportid); 
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextContains(partnerName).build(); // Filterin criteria
  var reportSpreadsheet = SpreadsheetApp.openById(id);

  // This part is for full number of people only
  var completedSheet = spreadsheet.getSheetByName('Completed');
  spreadsheet.setActiveSheet(completedSheet, true);
  completedSheet.getFilter().setColumnFilterCriteria(1, criteria);
  var completedSource = completedSheet.getRange(1, 1, completedSheet.getMaxRows(), completedSheet.getMaxColumns());
  var checkCompletedSheet = spreadsheet.getSheetByName(partnerName + ' by Numbers');
  if(!checkCompletedSheet){
    var newCompletedSheet = spreadsheet.insertSheet(1);
    spreadsheet.getActiveSheet().setName(partnerName +'  by Numbers');
  }
  else {
    newCompletedSheet = spreadsheet.getSheetByName(partnerName + '  by Numbers');
    newCompletedSheet.clearContents();
  }
  completedSource.copyTo(newCompletedSheet.getRange(1,1));
  completedSource.copyTo(newCompletedSheet.getRange(1,1),{contentsOnly: true});
  completedSheet.getFilter().removeColumnFilterCriteria(5);

  var citt = reportSpreadsheet.getSheetByName(partnerName + ' by Numbers');
  
  if (citt) {
    var completedSheetToBeDeleted = reportSpreadsheet.getSheetByName(partnerName + ' by Numbers');
    reportSpreadsheet.deleteSheet(completedSheetToBeDeleted);
  }
  newCompletedSheet.copyTo(reportSpreadsheet).setName(partnerName + ' by Numbers');

  // This part is for full training data
  // var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Training');
  spreadsheet.setActiveSheet(sheet, true);
  sheet.getFilter().setColumnFilterCriteria(5, criteria);
  
  // TODO Filter check
  // var filter = sheet.getFilter();
  // console.log(filter);
  // if (filter) {
  //   sheet.getFilter().setColumnFilterCriteria(5, criteria);
  // }
  // else {
  //   sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter().setColumnFilterCriteria(5, criteria);
  // }

  var source = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

  var checkSheet = spreadsheet.getSheetByName(partnerName + ' Full Data');
  if(!checkSheet){
    var newSheet = spreadsheet.insertSheet(1);
    spreadsheet.getActiveSheet().setName(partnerName +' Full Data');
  }
  else {
    newSheet = spreadsheet.getSheetByName(partnerName + ' Full Data');
    newSheet.clearContents();
  }
  
  source.copyTo(newSheet.getRange(1,1));
  source.copyTo(newSheet.getRange(1,1),{contentsOnly: true});

  sheet.getFilter().removeColumnFilterCriteria(5);

  var itt = reportSpreadsheet.getSheetByName(partnerName + ' Full Data');
  
  if (itt) {
    var sheetToBeDeleted = reportSpreadsheet.getSheetByName(partnerName + ' Full Data');
    reportSpreadsheet.deleteSheet(sheetToBeDeleted);
  }
  newSheet.copyTo(reportSpreadsheet).setName(partnerName + ' Full Data');

  // Cleanup the temporary Sheets
  var secondSheetToBeDeleted = spreadsheet.getSheetByName(partnerName + ' Full Data');
  spreadsheet.deleteSheet(secondSheetToBeDeleted);
  spreadsheet.deleteSheet(newCompletedSheet);
};
