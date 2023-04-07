// This function runs the script to sort sheets in multiple Google Sheets based on a specific order
function runSortSheetScript() {
    // Define the variables to be used in the function

    var lookupDocumentId = "1YSVmDpjZxNyKzh0JnhllKzXVinZ1T5vSQx2encRSIvI";
    var lookupSheetName = "Config";
    var documentUrlColumn = "Document to be shared ( Link to the document that will be shared/careful id should be at the end)";
    // Get the lookup sheet and its last row

    var lookupSheet = SpreadsheetApp.openById(lookupDocumentId).getSheetByName(lookupSheetName);
    var lastRow = lookupSheet.getLastRow();
    // Loop through each row in the lookup sheet

    for (var row = 2; row <= lastRow; row++) {
        var rowValues = lookupSheet.getRange(row, 1, 1, lookupSheet.getLastColumn()).getValues()[0];
        Logger.log("rowValues: " + rowValues);
        // Get the URL of the sheet to be sorted from the current row

        var documentUrl = rowValues[lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0].indexOf(documentUrlColumn)];
        Logger.log("documentUrl: " + documentUrl);
        // Extract the ID from the URL

        var documentId = documentUrl.match(/[-\w]{25,}/);

        if (documentId) {
            documentId = documentId[0];

            try {
                // open document
                var ss = SpreadsheetApp.openById(documentId);

                // sort sheets in the order "Completed", "Training", "Certification", "Bootcamp", "PE ENABLEMENT", "Live Sessions", then alphabetically
                sortSheets(ss);

            } catch (e) {
                Logger.log("Error: " + e.message + " - Skipping document with ID " + documentId);
            }
        } else {
            Logger.log("Document ID not found for row " + row);
        }
    }
}

// This function sorts the sheets in a Google Sheet in a specific order

function sortSheets(ss) {
    var sheets = ss.getSheets();
    var sheetOrder = ["Completed", "Training", "Certification", "Bootcamp", "PE ENABLEMENT", "Live Sessions"];
    var sortedSheets = [];

    // loop through each sheet and add it to the sortedSheets array if it matches the sheetOrder
    for (var i = 0; i < sheetOrder.length; i++) {
        for (var j = 0; j < sheets.length; j++) {
            if (sheets[j].getName().indexOf(sheetOrder[i]) > -1) {
                sortedSheets.push(sheets[j]);
            }
        }
    }

    // add any remaining sheets that were not found in the sheetOrder
    for (var k = 0; k < sheets.length; k++) {
        if (sortedSheets.indexOf(sheets[k]) === -1) {
            sortedSheets.push(sheets[k]);
        }
    }

    // reorder sheets
    for (var m = 0; m < sortedSheets.length; m++) {
        var sheet = sortedSheets[m];
        sheet.activate();
        ss.moveActiveSheet(m + 1);
    }

};
