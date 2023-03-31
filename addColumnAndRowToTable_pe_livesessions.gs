function addColumnAndRowToTable_pe_livesessions(documentId, sheetName, company, columnHeaderText, sourceDocumentId, sourceSheetName, sourceColumnName, attendeeColumnName) {

 var columnHeaderText = "Live Sessions";
var sourceDocumentId = "1NqhIzcclsxSmMm1uax7e16nzHoUtTIK-4pXJFlpVspY";
var sourceSheetName = "liveTraining";
var sourceColumnName = "EVENT";
var attendeeColumnName = "company_std";
  var sheet = SpreadsheetApp.openById(documentId).getSheetByName(sheetName);
  var table = sheet.getDataRange();





  var sourceSheet = SpreadsheetApp.openById(sourceDocumentId).getSheetByName(sourceSheetName);
  var sourceData = sourceSheet.getDataRange().getValues();
  var sourceColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(sourceColumnName);
  var sourceAttendeeColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(attendeeColumnName);
  var sourceColumnData = sourceData.slice(1).map(function(row) { return row[sourceColumnIndex]; });
  //var distinctValues = sourceColumnData.filter(function(value, index, self) { return self.indexOf(value) === index; });
  var distinctValues = sourceColumnData.filter(function(value, index, self) {
  return self.indexOf(value) === index && value != "" && value != null;
});
  var headerRange = sheet.createTextFinder(columnHeaderText).findNext();
  if(headerRange) {
    var columnToDelete = headerRange.getColumn();
    var rangeToDelete = sheet.getRange(1, columnToDelete, sheet.getLastRow(), distinctValues.length);
    rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
    var newColumn = columnToDelete;
  } else {
    var lastColumn = table.getLastColumn();
    var newColumn = lastColumn + 1;
  }

  var columnSpan = distinctValues.length;
  sheet.insertColumnsAfter(newColumn - 1, columnSpan - 1);

  var headerRange = sheet.getRange(1, newColumn, 1, columnSpan);
  headerRange.merge();
  headerRange.setHorizontalAlignment("center");
  headerRange.setValue(columnHeaderText);
  var sessionNameRange = sheet.getRange(2, newColumn, sheet.getLastRow()-1, 1);
  sessionNameRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sessionNameRange.setHorizontalAlignment("left");
  sheet.autoResizeColumn(newColumn);

var headers = table.getValues()[0];
for (var i = 0; i < headers.length; i++) {
  if (headers[i] == columnHeaderText) {
    sheet.autoResizeColumn(i+1);
    var cell = sheet.getRange(1, i+1);
    cell.setHorizontalAlignment("center").setBackground("Green").setFontWeight("bold").setFontColor("#ffffff");
  }
}











distinctValues.forEach(function(value, index) {
  var columnIndex = newColumn + index;
  sheet.getRange(2, columnIndex).setValue(value).setBackground("blue").setFontWeight("bold").setFontColor("#ffffff");
});

distinctValues.forEach(function(value, index) {
  var columnIndex = newColumn + index;
  var attendeeCount = sourceData.slice(1).reduce(function(count, row) {
    if (row[sourceColumnIndex] === value && row[sourceAttendeeColumnIndex] === company) {
      return count + 1;
    } else {
      return count;
    }
  }, 0);
  if(attendeeCount == 0){
    sheet.getRange(3, columnIndex).setValue('0').setHorizontalAlignment("center");
  }else{
    sheet.getRange(3, columnIndex).setValue(attendeeCount).setHorizontalAlignment("center");
  }
});
// Check if "bootcamp" sheet exists, and delete it if it does
var bootcampSheet = SpreadsheetApp.openById(documentId).getSheetByName(columnHeaderText);
if (bootcampSheet) {
  SpreadsheetApp.openById(documentId).deleteSheet(bootcampSheet);
}

// Create a new "bootcamp" sheet and add headers
var bootcampSheet = SpreadsheetApp.openById(documentId).insertSheet(columnHeaderText, 5);
var bootcampHeaders = ["EVENT", "Title", "First Name", "Last Name", "company_std",  "Email", "Phone","Country (text only)"];
bootcampSheet.appendRow(bootcampHeaders);
bootcampSheet.getRange(1, 1, 1, bootcampHeaders.length).setBackground("red");

// Copy data to "bootcamp" sheet
for (var i = 1; i < sourceData.length; i++) {
  var rowData = sourceData[i];


 if (rowData[sourceAttendeeColumnIndex] == company) {
    var EVENT = rowData[sourceData[0].indexOf("EVENT")];
    var Title = rowData[sourceData[0].indexOf("Title")];
        var firstName = rowData[sourceData[0].indexOf("First Name")];
    var lastName = rowData[sourceData[0].indexOf("Last Name")];
    var companyStd = rowData[sourceData[0].indexOf("company_std")];
    var email = rowData[sourceData[0].indexOf("Email")];
    var Phone = rowData[sourceData[0].indexOf("Phone")];
    var Country  = rowData[sourceData[0].indexOf("Country (text only)")];


    var bootcampRowData = [EVENT, Title, firstName, lastName, companyStd, email, Phone, Country];
    bootcampSheet.appendRow(bootcampRowData);
  }
}


}
