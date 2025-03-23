function extractCommonNamesFromFullname() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var commonNamesSheet = ss.getSheetByName("Common Names") || ss.insertSheet("Common Names");

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fullNameCol = headers.indexOf("FullName") + 1;

  if (fullNameCol === 0) {
    SpreadsheetApp.getUi().alert("FullName column is missing!");
    return;
  }

  var lastRow = sheet.getLastRow();
  var fullNames = sheet.getRange(3, fullNameCol, lastRow - 2).getValues().flat();

  var titles = ["phd.", "dr.", "prof.", "eng.", "ing.", "mr.", "ms.", "mrs.", "sir"];
  var extractedNames = [];

  fullNames.forEach(fullName => {
    if (!fullName) return;

    fullName = fullName.toLowerCase().trim();
    var words = fullName.split(/\s+/).filter(word => word.length > 1); // Remove single-letter words

    if (titles.includes(words[0])) return; // Ignore if word has titles
    if (words.length <= 1) return; // Ignore if no clear name-surname split

    var name = words[0];
    var surname = words[words.length - 1];

    if (surname.includes(".") || name.includes(".")) return; // Ignore if name or surname has dot

    // Ensure both name and surname are at least 2 characters
    if (name.length <= 2 || surname.length <= 2) return;

    extractedNames.push([name, surname]);
  });

  if (extractedNames.length === 0) {
    SpreadsheetApp.getUi().alert("No valid names found!");
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

  SpreadsheetApp.getUi().alert("Common Names extracted and updated successfully!");
}
