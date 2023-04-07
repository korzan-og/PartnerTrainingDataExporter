// This function updates the target sheet with data from the external sheet
function updateTargetDocument_pse_bootcamp() {
    // Replace the placeholders with actual values

    const targetDocumentId = "1xobRm8N8yoBrC_BSOo-VHLCTF3H_J3iBWxeCaGx0H-w";
    const targetCountryColumn = "partner_account_country";
    const targetColumnToUpdate = "company_std";
    const targetEmailColumn = "email";

    const externalDocumentId = "1_XPDO9SNw54kG7L8iUxCj6_xfTB7W38VfROEwVLr8sM";
    const externalSheetName = "companies";
    const externalStdNameColumn = "Standardized Company Name";
    const externalCountryColumn = "Country";
    const externalEmailPatternColumn = "Email_pattern";
    // Open the target and external sheets and get their data

    const targetDoc = SpreadsheetApp.openById(targetDocumentId);
    const externalDoc = SpreadsheetApp.openById(externalDocumentId);
    const externalSheet = externalDoc.getSheetByName(externalSheetName);
    const externalData = externalSheet.getDataRange().getValues();

    const sheets = targetDoc.getSheets();

    // Loop through all sheets in the target document

    for (const targetSheet of sheets) {
        const targetData = targetSheet.getDataRange().getValues();

        const targetCountryColIdx = targetData[0].indexOf(targetCountryColumn);
        const targetEmailColIdx = targetData[0].indexOf(targetEmailColumn);
        const targetUpdateColIdx = targetData[0].indexOf(targetColumnToUpdate);
        // Only update the column if it exists in the sheet

        if (targetUpdateColIdx !== -1) {
            const targetOrgColumn = "partner_name";
            const externalCompanyNameColumn = "Company Names";
            const targetOrgColIdx = targetData[0].indexOf(targetOrgColumn);
            const externalCompanyNameColIdx = externalData[0].indexOf(externalCompanyNameColumn);
            // Loop through all rows in the sheet

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

            const targetSessionNameColumn = "enablement_name";
            const targetSessionNameColIdx = targetData[0].indexOf(targetSessionNameColumn);
            // Remove duplicate rows based on company_std, email, and session name

            removeDuplicateRows(targetSheet, targetUpdateColIdx, targetEmailColIdx, targetSessionNameColIdx);
        }
    }
}

function findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx) {
    // Loop through each row in the external sheet data, starting at index 1 to skip the header row.

    for (let i = 1; i < externalData.length; i++) {
        // Get the country and company name list for the current row.

        const countryInExternal = externalData[i][2];
        const companyNameList = externalData[i][externalCompanyNameColIdx].toUpperCase().split(";");
        // If the company name list includes the organization name and the country matches, return the standardized company name.

        if (companyNameList.includes(orgName) && countryInExternal === country) {
            return externalData[i][0];
        }
    }
    // If no match is found, return an empty string.

    return '';
}

function findCompanyNameInExternal(externalData, domain, country) {
    // Loop through each row in the external sheet data, starting at index 1 to skip the header row.

    for (let i = 1; i < externalData.length; i++) {
        // Get the email pattern and country for the current row.

        const emailPattern = externalData[i][3];
        const countryInExternal = externalData[i][2];
        // If the email domain ends with the email pattern and the country matches, return the standardized company name.

        if (domain && domain.endsWith(emailPattern) && countryInExternal === country) {
            return externalData[i][0];
        }
    }
    // If no match is found, return an empty string.

    return '';
}

function removeDuplicateRows(sheet, companyStdIdx, emailIdx, sessionNameIdx) {
    // Get the data from the sheet.

    const data = sheet.getDataRange().getValues();
    const uniqueRows = [data[0]];
    // Loop through each row in the data, starting at index 1 to skip the header row.

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const companyStd = row[companyStdIdx];
        const email = row[emailIdx];
        const sessionName = row[sessionNameIdx];
        // Check if the row is unique based on the company standardization, email, and session name columns.

        if (!uniqueRows.some(uniqueRow => (
                uniqueRow[companyStdIdx] === companyStd &&
                uniqueRow[emailIdx] === email &&
                uniqueRow[sessionNameIdx] === sessionName
            ))) {
            uniqueRows.push(row);
        }
    }
    // Clear the contents of the sheet and set the unique rows as the new values.

    sheet.clearContents();
    sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
}
