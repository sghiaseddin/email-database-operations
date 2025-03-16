function highlightUndeliveredEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = sheet.getLastRow();
  
  var emailCol = headers.indexOf("Email") + 1;
  var emailGuessCol = headers.indexOf("EmailGuess") + 1;

  if (emailCol === 0 || emailGuessCol === 0) {
    SpreadsheetApp.getUi().alert("Required columns are missing!");
    return;
  }

  var undeliveredSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Undelivered");
  
  if (!undeliveredSheet) {
    SpreadsheetApp.getUi().alert("Sheets not found!");
    return;
  }

  var emails = sheet.getRange(2, emailCol, lastRow - 1);
  var emailsFlat = emails.getValues().flat();
  var emailGuesses = sheet.getRange(2, emailGuessCol, lastRow);
  var emailGuessesFlat = emailGuesses.getValues().flat();
  var undeliveredEmails = new Set(undeliveredSheet.getRange("A2:A" + undeliveredSheet.getLastRow()).getValues().flat());

  var emailBackgrounds = emails.getBackgrounds();
  var emailGuessBackgrounds = emailGuesses.getBackgrounds();

  
  for (var i = 0; i < emailsFlat.length; i++) {
    emailBackgrounds[i][0] = undeliveredEmails.has(emailsFlat[i]) ? "#FFAAAA" : "#FFFFFF";
    emailGuessBackgrounds[i][0] = undeliveredEmails.has(emailGuessesFlat[i]) ? "#FFAAAA" : "#FFFFFF";
    if (undeliveredEmails.has(emailsFlat[i])) Logger.log(emailsFlat[i]);
  }
  
  emails.setBackgrounds(emailBackgrounds);
  emailGuesses.setBackgrounds(emailGuessBackgrounds);
}
