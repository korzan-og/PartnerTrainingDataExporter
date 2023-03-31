function updateTargetDocument_ozan() {
  // Replace the placeholders with actual values
  const targetDocumentId = "1prAMKf3J2sCjN_mPsR0KTubyEpQnM3EsFJYlcoCJV6o";
  const targetSheetName = "sessions";
  const targetCountryColumn = "Country/Region Name";
  const targetColumnToUpdate = "company_std";
  const targetEmailColumn = "Email";

  const externalDocumentId = "1_XPDO9SNw54kG7L8iUxCj6_xfTB7W38VfROEwVLr8sM";
  const externalSheetName = "companies";
  const externalStdNameColumn = "Standardized Company Name";
  const externalCountryColumn = "Country";
  const externalEmailPatternColumn = "Email_pattern";

  const targetDoc = SpreadsheetApp.openById(targetDocumentId);
  const targetSheet = targetDoc.getSheetByName(targetSheetName);

  const externalDoc = SpreadsheetApp.openById(externalDocumentId);
  const externalSheet = externalDoc.getSheetByName(externalSheetName);

  const targetData = targetSheet.getDataRange().getValues();
  const externalData = externalSheet.getDataRange().getValues();

  const targetCountryColIdx = targetData[0].indexOf(targetCountryColumn);
  const targetEmailColIdx = targetData[0].indexOf(targetEmailColumn);
  const targetUpdateColIdx = targetData[0].indexOf(targetColumnToUpdate);


  const targetOrgColumn = "Organization";
  const externalCompanyNameColumn = "Company Names";

  const targetOrgColIdx = targetData[0].indexOf(targetOrgColumn);
  const externalCompanyNameColIdx = externalData[0].indexOf(externalCompanyNameColumn);

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


  const targetSessionNameColumn = "EVENT";
  const targetSessionNameColIdx = targetData[0].indexOf(targetSessionNameColumn);

  removeDuplicateRows(targetSheet, targetUpdateColIdx, targetEmailColIdx, targetSessionNameColIdx);
}


function findCompanyNameByOrgAndCountry(externalData, orgName, country, externalCompanyNameColIdx) {
  for (let i = 1; i < externalData.length; i++) {
    const countryInExternal = externalData[i][2];
    const companyNameList = externalData[i][externalCompanyNameColIdx].toUpperCase().split(";");

    if (companyNameList.includes(orgName) && countryInExternal === country) {
      return externalData[i][0];
    }
  }
  return '';
}





function findCompanyNameInExternal(externalData, domain, country) {
  for (let i = 1; i < externalData.length; i++) {
    const emailPattern = externalData[i][3];
    const countryInExternal = externalData[i][2];

    if (domain && domain.endsWith(emailPattern) && countryInExternal === country) {
      return externalData[i][0];
    }
  }
  return '';
}


function removeDuplicateRows(sheet, companyStdIdx, emailIdx, sessionNameIdx) {
  const data = sheet.getDataRange().getValues();
  const uniqueRows = [data[0]];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const companyStd = row[companyStdIdx];
    const email = row[emailIdx];
    const sessionName = row[sessionNameIdx];

    if (!uniqueRows.some(uniqueRow => (
      uniqueRow[companyStdIdx] === companyStd &&
      uniqueRow[emailIdx] === email &&
      uniqueRow[sessionNameIdx] === sessionName
    ))) {
      uniqueRows.push(row);
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
}
