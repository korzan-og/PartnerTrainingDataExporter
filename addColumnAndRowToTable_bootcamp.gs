function addColumnAndRowToTable_bootcamp(documentId, sheetName, company, columnHeaderText, sourceDocumentId, sourceSheetName, sourceColumnName, attendeeColumnName) {

  var columnHeaderText = "Bootcamp Sessions";
  var sourceDocumentId = "1llLu68WljkD5jO2oe2JZj9_fIoFaA5wd8r4KzZ5-tcg";
  var sourceSheetName = "sessions";
  var sourceColumnName = "Learning Path Name";
  var attendeeColumnName = "Company";
  var sheet = SpreadsheetApp.openById(documentId).getSheetByName(sheetName);
  var table = sheet.getDataRange();

  var sourceSheet = SpreadsheetApp.openById(sourceDocumentId).getSheetByName(sourceSheetName);
  var sourceData = sourceSheet.getDataRange().getValues();
  var sourceColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(sourceColumnName);
  var sourceAttendeeColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(attendeeColumnName);
  var sourceColumnData = sourceData.slice(1).map(function(row) { return row[sourceColumnIndex]; });
  var distinctValues = sourceColumnData.filter(function(value, index, self) { return self.indexOf(value) === index; });
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
    cell.setHorizontalAlignment("center").setBackground("yellow").setFontWeight("bold").setFontColor("#ffffff");
  }
}











distinctValues.forEach(function(value, index) {
  var columnIndex = newColumn + index;
  sheet.getRange(2, columnIndex).setValue(value).setBackground("lightgreen").setFontWeight("bold").setFontColor("#ffffff");
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
  var bootcampSheet = SpreadsheetApp.openById(documentId).getSheetByName("bootcamp");
  if (bootcampSheet) {
    SpreadsheetApp.openById(documentId).deleteSheet(bootcampSheet);
  }

    // Check if "bootcamp" sheet exists, and delete it if it does
  var bootcampSheet = SpreadsheetApp.openById(documentId).getSheetByName("bootcamp");
  if (bootcampSheet) {
    SpreadsheetApp.openById(documentId).deleteSheet(bootcampSheet);
  }

// Create a new "bootcamp" sheet and add headers
  var bootcampSheet = SpreadsheetApp.openById(documentId).insertSheet("Bootcamp", 3);
  var bootcampHeaders = ["Learning Path Name", "First Name", "Last Name", "Company", "Email Id", "Start Date", "Completion Date", "Country"];
  bootcampSheet.appendRow(bootcampHeaders);
  bootcampSheet.getRange(1, 1, 1, bootcampHeaders.length).setBackground("red");

  // Copy data to "bootcamp" sheet
  for (var i = 1; i < sourceData.length; i++) {
    var rowData = sourceData[i];
    if (rowData[sourceAttendeeColumnIndex] == company) {
      var sessionName = rowData[sourceData[0].indexOf("Learning Path Name")];
      var firstName = rowData[sourceData[0].indexOf("First Name")];
      var lastName = rowData[sourceData[0].indexOf("Last Name")];
      var companyName = rowData[sourceData[0].indexOf("Company")];
      var emailId = rowData[sourceData[0].indexOf("Email Id")];
      var startDate = rowData[sourceData[0].indexOf("Start Date")];
      var completionDate = rowData[sourceData[0].indexOf("Completion Date")];
      var country = rowData[sourceData[0].indexOf("Country")];

      var bootcampRowData = [sessionName, firstName, lastName, companyName, emailId, startDate, completionDate, country];
      bootcampSheet.appendRow(bootcampRowData);
    }
  }
}
