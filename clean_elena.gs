// This function updates the target sheet with data from the external sheet
function updateTargetDocument_elena() {
    // Replace the placeholders with actual values
    const targetDocumentId = "1FJ-LhUrbjNwII98YUh2nC-SGuMZDyuun5y5ohV-tF44";
    const targetSheetName = "emea";
    const targetCountryColumn = "Country_Region_Name";
    const targetColumnToUpdate = "company_std";
    const targetEmailColumn = "Email";

    const externalDocumentId = "1_XPDO9SNw54kG7L8iUxCj6_xfTB7W38VfROEwVLr8sM";
    const externalSheetName = "companies";
    const externalStdNameColumn = "Standardized Company Name";
    const externalCountryColumn = "Country";
    const externalEmailPatternColumn = "Email_pattern";
    // Open the target and external sheets and get their data

    const targetDoc = SpreadsheetApp.openById(targetDocumentId);
    const targetSheet = targetDoc.getSheetByName(targetSheetName);

    const externalDoc = SpreadsheetApp.openById(externalDocumentId);
    const externalSheet = externalDoc.getSheetByName(externalSheetName);

    const targetData = targetSheet.getDataRange().getValues();
    const externalData = externalSheet.getDataRange().getValues();
    // Get the column indices for the target sheet

    const targetCountryColIdx = targetData[0].indexOf(targetCountryColumn);
    const targetEmailColIdx = targetData[0].indexOf(targetEmailColumn);
    const targetUpdateColIdx = targetData[0].indexOf(targetColumnToUpdate);

    const targetOrgColumn = "Organization";

    // Get the column indices for the external sheet

    const externalCompanyNameColumn = "Company Names";

    const targetOrgColIdx = targetData[0].indexOf(targetOrgColumn);
    const externalCompanyNameColIdx = externalData[0].indexOf(externalCompanyNameColumn);
    // Iterate over the rows of the target sheet and update the company_std column

    for (let i = 1; i < targetData.length; i++) {
        const companyName = targetData[i][targetUpdateColIdx];
        const orgName = targetData[i][targetOrgColIdx].toUpperCase();
        const domain = targetData[i][targetEmailColIdx].split('@')[1];
        const country = targetData[i][targetCountryColIdx];

        // Update company_std based on matching organization and country values
        let stdName = findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx);
        if (stdName) {
            targetSheet.getRange(i + 1, targetUpdateColIdx + 1).setValue(stdName);
        } else {
            // Update company_std based on matching email pattern
            stdName = findCompanyNameInExternal(externalData, domain, country);
            targetSheet.getRange(i + 1, targetUpdateColIdx + 1).setValue(stdName);
        }
    }

    const targetSessionNameColumn = "Session Name";
    const targetSessionNameColIdx = targetData[0].indexOf(targetSessionNameColumn);
    // Remove duplicate rows based on company_std, email, and session name

    removeDuplicateRows(targetSheet, targetUpdateColIdx, targetEmailColIdx, targetSessionNameColIdx);
}

// Helper function to find the company name in the external sheet based on organization and country

function findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx) {
      // Iterate over the rows of the external sheet

    for (let i = 1; i < externalData.length; i++) {
        const countryInExternal = externalData[i][2];
        const companyNameList = externalData[i][externalCompanyNameColIdx].toUpperCase().split(";");
        // Check if organization name is included in company name list and country matches

        if (companyNameList.includes(orgName) && countryInExternal === country) {
            return externalData[i][0];
        }
    }
    return '';
}

// Helper function to find the company name in the external sheet based on email pattern and country

function findCompanyNameInExternal(externalData, domain, country) {
    // Iterate over the rows of the external sheet

    for (let i = 1; i < externalData.length; i++) {
        const emailPattern = externalData[i][3];
        const countryInExternal = externalData[i][2];
        // Check if domain ends with email pattern and country matches

        if (domain && domain.endsWith(emailPattern) && countryInExternal === country) {
            return externalData[i][0];
        }
    }
    return '';
}

// Helper function to remove duplicate rows based on company_std, email, and session name

function removeDuplicateRows(sheet, companyStdIdx, emailIdx, sessionNameIdx) {
    const data = sheet.getDataRange().getValues();
    const uniqueRows = [data[0]];
    // Iterate over the rows and add unique rows to a new array

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const companyStd = row[companyStdIdx];
        const email = row[emailIdx];
        const sessionName = row[sessionNameIdx];
        // Check if row is not already in the uniqueRows array

        if (!uniqueRows.some(uniqueRow => (
                uniqueRow[companyStdIdx] === companyStd &&
                uniqueRow[emailIdx] === email &&
                uniqueRow[sessionNameIdx] === sessionName
            ))) {
            uniqueRows.push(row);
        }
    }
    // Clear the sheet and write the unique rows back to it

    sheet.clearContents();
    sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
}
