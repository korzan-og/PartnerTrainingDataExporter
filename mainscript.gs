function runScript() {
  var lookupDocumentId = "1YSVmDpjZxNyKzh0JnhllKzXVinZ1T5vSQx2encRSIvI";
  var lookupSheetName = "Config";
  var companyNameColumn = "Partner Name";
  var documentUrlColumn = "Document to be shared ( Link to the document that will be shared/careful id should be at the end)";

  var lookupSheet = SpreadsheetApp.openById(lookupDocumentId).getSheetByName(lookupSheetName);
  var lastRow = lookupSheet.getLastRow();

  for (var row = 2; row <= lastRow; row++) {
    var rowValues = lookupSheet.getRange(row, 1, 1, lookupSheet.getLastColumn()).getValues()[0];
    Logger.log("rowValues: " + rowValues);

    var companyName = rowValues[lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0].indexOf(companyNameColumn)];
    Logger.log("companyName: " + companyName);

    var documentUrl = rowValues[lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0].indexOf(documentUrlColumn)];
    Logger.log("documentUrl: " + documentUrl);

    var documentId = documentUrl.match(/[-\w]{25,}/);

    if (documentId) {
      documentId = documentId[0];
      var sheetName = companyName + " Completed";
      addColumnAndRowToTable_bootcamp(documentId, sheetName, companyName);
      addColumnAndRowToTable_pe_enablement(documentId, sheetName, companyName); // Reusing the same sheetName
      addColumnAndRowToTable_pe_livesessions(documentId, sheetName, companyName);
      // Open the sheet and remove empty columns
      var sheet = SpreadsheetApp.openById(documentId).getSheetByName(sheetName);
      var dataRange = sheet.getDataRange();
      var numColumns = dataRange.getNumColumns();
      var emptyColumns = [];

      // Identify empty columns
      var secondRowValues = sheet.getRange(2, 1, 1, numColumns).getValues()[0];
      for (var col = 0; col < secondRowValues.length; col++) {
        if (secondRowValues[col] === "") {
          emptyColumns.push(col+1);
        }
      }

      // Delete empty columns (starting from the end to avoid affecting column indexes)
      for (var i = emptyColumns.length - 1; i >= 0; i--) {
        var columnToDelete = emptyColumns[i];
        sheet.deleteColumn(columnToDelete);
      }
    } else {
      Logger.log("Document ID not found for company: " + companyName);
    }
  }
}
