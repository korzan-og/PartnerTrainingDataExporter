//This function adds columns and rows to a Google Sheet based on data from another sheet, and creates a new sheet with filtered data.
function addColumnAndRowToTable_bootcamp(documentId, sheetName, company, columnHeaderText, sourceDocumentId, sourceSheetName, sourceColumnName, attendeeColumnName) {

    // Set default values for parameters
    var columnHeaderText = "Bootcamp Sessions";
    var sourceDocumentId = "1xobRm8N8yoBrC_BSOo-VHLCTF3H_J3iBWxeCaGx0H-w";
    var sourceSheetName = "sessions";
    var sourceColumnName = "enablement_name";
    var attendeeColumnName = "company_std";

    // Get the target sheet and range
    var sheet = SpreadsheetApp.openById(documentId).getSheetByName(sheetName);
    var table = sheet.getDataRange();

    // Get the source data and column index for the source column
    var sourceSheet = SpreadsheetApp.openById(sourceDocumentId).getSheetByName(sourceSheetName);
    var sourceData = sourceSheet.getDataRange().getValues();
    var sourceColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(sourceColumnName);
    var sourceAttendeeColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf(attendeeColumnName);
    var sourceColumnData = sourceData.slice(1).map(function(row) {
        return row[sourceColumnIndex];
    });
    var distinctValues = sourceColumnData.filter(function(value, index, self) {
        return self.indexOf(value) === index;
    });

    // Find or create the column for the new data
    var headerRange = sheet.createTextFinder(columnHeaderText).findNext();
    if (headerRange) {
        var columnToDelete = headerRange.getColumn();
        var rangeToDelete = sheet.getRange(1, columnToDelete, sheet.getLastRow(), distinctValues.length);
        rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
        var newColumn = columnToDelete;
    } else {
        var lastColumn = table.getLastColumn();
        var newColumn = lastColumn + 1;
    }

    // Insert the new column
    var columnSpan = distinctValues.length;
    sheet.insertColumnsAfter(newColumn - 1, columnSpan - 1);

    // Set the header for the new column
    var headerRange = sheet.getRange(1, newColumn, 1, columnSpan);
    headerRange.merge();
    headerRange.setHorizontalAlignment("center");
    headerRange.setValue(columnHeaderText);
    var sessionNameRange = sheet.getRange(2, newColumn, sheet.getLastRow() - 1, 1);
    sessionNameRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sessionNameRange.setHorizontalAlignment("left");
    sheet.autoResizeColumn(newColumn);

    // Highlight the header cell for the new column
    var headers = table.getValues()[0];
    for (var i = 0; i < headers.length; i++) {
        if (headers[i] == columnHeaderText) {
            sheet.autoResizeColumn(i + 1);
            var cell = sheet.getRange(1, i + 1);
            cell.setHorizontalAlignment("center").setBackground("yellow").setFontWeight("bold").setFontColor("#ffffff");
        }
    }

    // Fill in the new column with attendee counts
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
        if (attendeeCount == 0) {
            sheet.getRange(3, columnIndex).setValue('0').setHorizontalAlignment("center");
        } else {
            sheet.getRange(3, columnIndex).setValue(attendeeCount).setHorizontalAlignment("center");
        }
    });

    // Check if "bootcamp" sheet exists, and delete it if it does
    var bootcampSheet = SpreadsheetApp.openById(documentId).getSheetByName("bootcamp");
    if (bootcampSheet) {
        SpreadsheetApp.openById(documentId).deleteSheet(bootcampSheet);
    }

    // Create a new "bootcamp" sheet and add headers
    var bootcampSheet = SpreadsheetApp.openById(documentId).insertSheet("Bootcamp", 3);
    var bootcampHeaders = ["enablement_name", "first_name", "last_name", "partner_name", "email", "enablement_date", "Country"];
    bootcampSheet.appendRow(bootcampHeaders);
    bootcampSheet.getRange(1, 1, 1, bootcampHeaders.length).setBackground("yellow");

    // Copy data to "bootcamp" sheet
    for (var i = 1; i < sourceData.length; i++) {
        var rowData = sourceData[i];
        if (rowData[sourceAttendeeColumnIndex] == company) {
            var sessionName = rowData[sourceData[0].indexOf("enablement_name")];
            var firstName = rowData[sourceData[0].indexOf("first_name")];
            var lastName = rowData[sourceData[0].indexOf("last_name")];
            var companyName = (rowData[sourceData[0].indexOf("partner_name")]).trim();
            var emailId = rowData[sourceData[0].indexOf("email")];
            var startDate = rowData[sourceData[0].indexOf("enablement_date")];
            var country = rowData[sourceData[0].indexOf("Country")];

            var bootcampRowData = [sessionName, firstName, lastName, companyName, emailId, startDate, country];
            bootcampSheet.appendRow(bootcampRowData);
        }
    }
}
