// This function call all the others
// This function retrieves data from a Google Sheet and updates other Google Sheets with the data.
function runScript() {
    // The ID of the Google Sheet that contains the lookup data.
    var lookupDocumentId = "1YSVmDpjZxNyKzh0JnhllKzXVinZ1T5vSQx2encRSIvI";

    // The name of the sheet in the lookup document that contains the configuration data.
    var lookupSheetName = "Config";

    // The name of the column in the lookup sheet that contains the partner names.
    var companyNameColumn = "Partner Name";

    // The name of the column in the lookup sheet that contains the document URLs to be shared.
    var documentUrlColumn = "Document to be shared ( Link to the document that will be shared/careful id should be at the end)";

    // Get the lookup sheet from the lookup document.
    var lookupSheet = SpreadsheetApp.openById(lookupDocumentId).getSheetByName(lookupSheetName);
    var lastRow = lookupSheet.getLastRow();

    // Loop through the rows in the lookup sheet to process each row.
    for (var row = 2; row <= lastRow; row++) {
        // Get the values in the current row of the lookup sheet.
        var rowValues = lookupSheet.getRange(row, 1, 1, lookupSheet.getLastColumn()).getValues()[0];
        Logger.log("rowValues: " + rowValues);

        // Get the name of the partner from the current row of the lookup sheet.
        var companyName = rowValues[lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0].indexOf(companyNameColumn)];
        Logger.log("companyName: " + companyName);

        // Get the URL of the document to be shared from the current row of the lookup sheet.
        var documentUrl = rowValues[lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0].indexOf(documentUrlColumn)];
        Logger.log("documentUrl: " + documentUrl);

        // Extract the document ID from the document URL.
        var documentId = documentUrl.match(/[-\w]{25,}/);

        // Check if the document ID was successfully extracted.
        if (documentId) {
            documentId = documentId[0];

            // Define the name of the sheet to be updated in the target document.
            var sheetName = companyName + " Completed";

            // Call the following functions to update the target documents:
            updateTargetDocument_elena();
            updateTargetDocument_ozan();
            updateTargetDocument_pse();
            updateTargetDocument_pse_bootcamp();
            addColumnAndRowToTable_bootcamp(documentId, sheetName, companyName);
            addColumnAndRowToTable_pe_enablement(documentId, sheetName, companyName); // Reusing the same sheetName
            addColumnAndRowToTable_pe_livesessions(documentId, sheetName, companyName);

            // Open the target sheet and remove any empty columns.
            var sheet = SpreadsheetApp.openById(documentId).getSheetByName(sheetName);
            var dataRange = sheet.getDataRange();
            var numColumns = dataRange.getNumColumns();
            var emptyColumns = [];

            // Identify empty columns.
            var secondRowValues = sheet.getRange(2, 1, 1, numColumns).getValues()[0];
            for (var col = 0; col < secondRowValues.length; col++) {
                if (secondRowValues[col] === "") {
                    emptyColumns.push(col + 1);
                }


            }

            // Delete empty columns (starting from the end to avoid affecting column indexes).
            for (var i = emptyColumns.length - 1; i >= 0; i--) {
                var columnToDelete = emptyColumns[i];
                sheet.deleteColumn(columnToDelete);
            }
        } else {
            // Log an error message if the document ID cannot be found for the partner.
            Logger.log("Document ID not found for company: " + companyName);
        }
    }

    // Call the runSortSheetScript function to sort the target sheets by company name.
    runSortSheetScript();
}
