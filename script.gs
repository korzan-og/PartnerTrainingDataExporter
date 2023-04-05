function CreateTrainingDataForPartners() {
  const reportid = '1j80iS_inFjdZhY-c5df9P12lFF_35d2c01t9X08PWTM'; // Id for the Master Training report
  const configSSId = '1YSVmDpjZxNyKzh0JnhllKzXVinZ1T5vSQx2encRSIvI'; // This is the Id for the spreadsheet that keeps track of the focus partners, their reports and name mappings

  // Global Variables
  var spreadsheet = SpreadsheetApp.openById(reportid); 
  var configSpreadsheet = SpreadsheetApp.openById(configSSId);


  // Get the Config
  var configSheet = configSpreadsheet.getSheetByName('Config');
  configSpreadsheet.setActiveSheet(configSheet, true);
  var ss = configSheet;
  var valuesconfig = configSheet.getDataRange().getValues();
  var y = 0;
  for(x = 1;x<ss.getRange(1, 1, ss.getLastRow(),1).getValues().length;x++){
    var partnerName = valuesconfig[x][y++];
    var docUrl = valuesconfig[x][y++];
    var basicUrl = docUrl.split("?")[0];
    var itemsUrl=basicUrl.split("/");
    if(itemsUrl[itemsUrl.length-1]==""||itemsUrl[itemsUrl.length-1]=="edit") var docId = itemsUrl[itemsUrl.length-2];
    else var docId = itemsUrl[itemsUrl.length-1];
    var partnerRegexPattern = valuesconfig[x][y++];
    var emailRegexPattern = valuesconfig[x][y];
    if(partnerName==""&&docUrl==""&&partnerRegexPattern=="") break;
    console.log("partner : " + partnerName + " docUrl : " + docUrl + " docId: " + docId + " pattern : "+partnerRegexPattern + " emailPattern : "+ emailRegexPattern);
    // TO Export More information Just follow the pattern and add values here
    var regexBy = ['company','company','company']; //Adjust here if you need to add another sheet to export
    var sheetsToExport=['Completed','Training','Certification']; //Adjust here if you need to add another sheet to export
    var numberOfColumnsTofilter=[1,5,4]; //Adjust here if you need to add another sheet to export
    var columnsToBeFiltered=['A:A','E:E','D:D']; //Adjust here if you need to add another sheet to export
    // END
    // var numberOfNewSheets=3; // Adjust here if you need to add another sheet to export
    var numberOfNewSheets=sheetsToExport.length;
    var z=0;
    // if(partnerName==""||docUrl==""||partnerRegexPattern==""||emailRegexPattern==""){
    if(partnerName==""||docUrl==""||partnerRegexPattern==""){
      throw new Error( "One of the fields for the partner is missing! Config File URL: https://docs.google.com/spreadsheets/d/" + configSSId + " row: "+x); 
    }
    var reportSpreadsheet = SpreadsheetApp.openById(docId);

    while(z<numberOfNewSheets){
      var existingSheet = spreadsheet.getSheetByName(sheetsToExport[z]);
      spreadsheet.setActiveSheet(existingSheet, true);
      if(regexBy[z]=='company') var selectionCriteria="=REGEXMATCH("+columnsToBeFiltered[z] +",\""+partnerRegexPattern+"\")";
      else if(regexBy[z]=='email') var selectionCriteria="=REGEXMATCH("+columnsToBeFiltered[z] +",\""+emailRegexPattern+"\")";
      var filterForSelection = SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied(selectionCriteria).build(); // Filtering the criteria
      existingSheet.getFilter().setColumnFilterCriteria(numberOfColumnsTofilter[z], filterForSelection);
      var existingData = existingSheet.getRange(1, 1, existingSheet.getMaxRows(), existingSheet.getMaxColumns());
      var checkIfSheetExists = spreadsheet.getSheetByName(partnerName + ' ' + sheetsToExport[z]);
      if(!checkIfSheetExists){
        var newSheet = spreadsheet.insertSheet(1);
        spreadsheet.getActiveSheet().setName(partnerName + ' ' +sheetsToExport[z]);
      }
      else {
        newSheet = spreadsheet.getSheetByName(partnerName + ' ' + sheetsToExport[z]);
        newSheet.clearContents();
      }
      existingData.copyTo(newSheet.getRange(1,1));
      existingData.copyTo(newSheet.getRange(1,1),{contentsOnly: true});
      existingData.getFilter().removeColumnFilterCriteria(numberOfColumnsTofilter[z]);
      var citt = reportSpreadsheet.getSheetByName(partnerName + ' ' + sheetsToExport[z]);
    
      if (citt) {
        var sheetToBeDeleted = reportSpreadsheet.getSheetByName(partnerName + ' ' + sheetsToExport[z]);
        reportSpreadsheet.deleteSheet(sheetToBeDeleted);
      }
      newSheet.copyTo(reportSpreadsheet).setName(partnerName + ' ' + sheetsToExport[z]);
      z++;
      spreadsheet.deleteSheet(newSheet);
    }
    y=0;
  }
}
