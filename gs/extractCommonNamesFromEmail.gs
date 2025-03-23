function extractCommonNamesFromEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var protectedSheet = ss.getSheetByName("Protected");
  var commonNamesSheet = ss.getSheetByName("Common Names") || ss.insertSheet("Common Names");

  if (!protectedSheet) {
    SpreadsheetApp.getUi().alert("Protected sheet is missing!");
    return;
  }

  var headers = protectedSheet.getRange(1, 1, 1, protectedSheet.getLastColumn()).getValues()[0];
  var emailCol = headers.indexOf("Email") + 1;

  if (emailCol === 0) {
    SpreadsheetApp.getUi().alert("Email column is missing in Protected sheet!");
    return;
  }

  var lastRow = protectedSheet.getLastRow();
  var emails = protectedSheet.getRange(2, emailCol, lastRow - 1).getValues().flat();
  
  var extractedNames = [];

  emails.forEach(email => {
    if (!email) return;

    email = email.toLowerCase().trim();
    var match = email.match(/^([a-z]+)\.([a-z]+)@/); // Match "name.surname@" pattern

    if (!match) return;

    var name = match[1];
    var surname = match[2];

    // Ensure both name and surname are at least 2 characters
    if (name.length <= 2 || surname.length <= 2) return;

    extractedNames.push([name, surname]);
  });

  if (extractedNames.length === 0) {
    SpreadsheetApp.getUi().alert("No valid names found in emails!");
    return;
  }

  // Append data to "Common Names" sheet
  var existingData = commonNamesSheet.getRange(2, 1, commonNamesSheet.getLastRow(), 2).getValues();
  var newData = extractedNames.filter(row => !existingData.some(existing => existing[0] === row[0] && existing[1] === row[1]));
  var allData = existingData.concat(newData);

  // Remove duplicates and sort each column independently
  var allNames = [...new Set(allData.map(row => row[0]))].sort();
  var allSurnames = [...new Set(allData.map(row => row[1]))].sort();

  // Ensure the columns align properly (fill empty spaces if one is shorter)
  var maxLength = Math.max(allNames.length, allSurnames.length);
  var sortedData = Array.from({ length: maxLength }, (_, i) => [
    allNames[i] || "",
    allSurnames[i] || ""
  ]);

  commonNamesSheet.clear();
  commonNamesSheet.getRange(1, 1, 1, 2).setValues([["Name", "Surname"]]);
  commonNamesSheet.getRange(2, 1, sortedData.length, 2).setValues(sortedData);
  commonNamesSheet.deleteRow(2);

  SpreadsheetApp.getUi().alert("Common Names extracted from emails and updated successfully!");
}