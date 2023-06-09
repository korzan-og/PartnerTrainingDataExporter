// This function updates the target sheets with data from the external sheet
function updateTargetDocument_pse() {
    // Set target document and column names

    const targetDocumentId = "1YZTiXM1FWaM5bWCDa2yNQKUJYOsLgsc1AkS_blzoX1I";
    const targetCountryColumn = "Country (text only)";
    const targetColumnToUpdate = "company_std";
    const targetEmailColumn = "Email";
    // Set external document and column names

    const externalDocumentId = "1_XPDO9SNw54kG7L8iUxCj6_xfTB7W38VfROEwVLr8sM";
    const externalSheetName = "companies";
    const externalStdNameColumn = "Standardized Company Name";
    const externalCountryColumn = "Country";
    const externalEmailPatternColumn = "Email_pattern";
    // Open target and external documents and get external data

    const targetDoc = SpreadsheetApp.openById(targetDocumentId);
    const externalDoc = SpreadsheetApp.openById(externalDocumentId);
    const externalSheet = externalDoc.getSheetByName(externalSheetName);
    const externalData = externalSheet.getDataRange().getValues();
    // Get all sheets in the target document

    const sheets = targetDoc.getSheets();

    // Iterate over all sheets in the target document

    for (const targetSheet of sheets) {
        // Get data from the target sheet

        const targetData = targetSheet.getDataRange().getValues();
        // Get column indices for target sheet

        const targetCountryColIdx = targetData[0].indexOf(targetCountryColumn);
        const targetEmailColIdx = targetData[0].indexOf(targetEmailColumn);
        const targetUpdateColIdx = targetData[0].indexOf(targetColumnToUpdate);
        // Check if target sheet has column to update

        if (targetUpdateColIdx !== -1) {
            // Get column names and indices for external data

            const targetOrgColumn = "Company";
            const externalCompanyNameColumn = "Company Names";
            const targetOrgColIdx = targetData[0].indexOf(targetOrgColumn);
            const externalCompanyNameColIdx = externalData[0].indexOf(externalCompanyNameColumn);
            // Iterate over rows in target sheet

            for (let i = 1; i < targetData.length; i++) {
                // Get relevant data for the row

                const companyName = targetData[i][targetUpdateColIdx];
                const orgName = targetData[i][targetOrgColIdx].toUpperCase();
                const domain = targetData[i][targetEmailColIdx].split('@')[1];
                const country = targetData[i][targetCountryColIdx];
                // Find standardized company name based on organization and country

                let stdName = findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx);
                if (stdName) {
                    // Update company_std column if a match is found

                    targetSheet.getRange(i + 1, targetUpdateColIdx + 1).setValue(stdName);
                } else {
                    // Find standardized company name based on email pattern and country

                    stdName = findCompanyNameInExternal(externalData, domain, country);
                    targetSheet.getRange(i + 1, targetUpdateColIdx + 1).setValue(stdName);
                }
            }
            // Get column name and index for session name

            const targetSessionNameColumn = "EVENT";
            const targetSessionNameColIdx = targetData[0].indexOf(targetSessionNameColumn);
            // Remove duplicate rows based on company_std, email, and session name

            removeDuplicateRows(targetSheet, targetUpdateColIdx, targetEmailColIdx, targetSessionNameColIdx);
        }
    }
}

function findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx) {
    // Iterate over rows in external data

    for (let i = 1; i < externalData.length; i++) {
        // Get relevant data from external data row

        const countryInExternal = externalData[i][2];
        const companyNameList = externalData[i][externalCompanyNameColIdx].toUpperCase().split(";");
        // Check if organization and country match

        if (companyNameList.includes(orgName) && countryInExternal === country) {
            // Return standardized company name if match is found

            return externalData[i][0];
        }
    }
    // Return empty string if no match is found

    return '';
}

function findCompanyNameInExternal(externalData, domain, country) {
    // Iterate over rows in external data

    for (let i = 1; i < externalData.length; i++) {
        // Get relevant data from external data row

        const emailPattern = externalData[i][3];
        const countryInExternal = externalData[i][2];
        // Check if domain ends with email pattern and country matches

        if (domain && domain.endsWith(emailPattern) && countryInExternal === country) {
            // Return standardized company name if match is found

            return externalData[i][0];
        }
    }
    // Return empty string if no match is found

    return '';
}

function removeDuplicateRows(sheet, companyStdIdx, emailIdx, sessionNameIdx) {
    // Get data from the sheet

    const data = sheet.getDataRange().getValues();
    // Create a new array with the header row

    const uniqueRows = [data[0]];
    // Iterate over rows in data

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const companyStd = row[companyStdIdx];
        const email = row[emailIdx];
        const sessionName = row[sessionNameIdx];
        // Check if a row with the same values for company_std, email, and session name has already been added to the uniqueRows array

        if (!uniqueRows.some(uniqueRow => (
                uniqueRow[companyStdIdx] === companyStd &&
                uniqueRow[emailIdx] === email &&
                uniqueRow[sessionNameIdx] === sessionName
            ))) {
            // If the row is unique, add it to the uniqueRows array

            uniqueRows.push(row);
        }
    }
    // Clear contents of sheet

    sheet.clearContents();
    // Set the values of the sheet to the uniqueRows array

    sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
}
