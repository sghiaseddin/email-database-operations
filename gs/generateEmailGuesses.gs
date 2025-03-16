function generateEmailGuesses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var patternCol = headers.indexOf("Pattern") + 1;
  var fullNameCol = headers.indexOf("FullName") + 1;
  var domainCol = headers.indexOf("Domain") + 1;
  var emailGuessCol = headers.indexOf("EmailGuess") + 1;
  
  if (patternCol === 0 || fullNameCol === 0 || domainCol === 0 || emailGuessCol === 0) {
    SpreadsheetApp.getUi().alert("Required columns are missing!");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var patterns = sheet.getRange(2, patternCol, lastRow - 1).getValues().flat();
  var fullNames = sheet.getRange(2, fullNameCol, lastRow - 1).getValues().flat();
  var domains = sheet.getRange(2, domainCol, lastRow - 1).getValues().flat();
  var emailGuesses = [];
  
  for (var i = 0; i < patterns.length; i++) {
    var email = generateEmail(fullNames[i], domains[i], patterns[i]);
    emailGuesses.push([email]);
  }
  
  sheet.getRange(2, emailGuessCol, emailGuesses.length, 1).setValues(emailGuesses);
}

function generateEmail(fullName, domain, pattern) {
  var parts = fullName.trim().split(/\s+/);
  var name = parts[0];
  var surname = parts.slice(1).join("");
  var n = name.charAt(0);
  var s = surname.charAt(0);
  
  var patterns = {
    "Name": name,
    "Surname": surname,
    "Name.Surname": name + "." + surname,
    "NameSurname": name + surname,
    "Name_Surname": name + "_" + surname,
    "N.Surname": n + "." + surname,
    "NSurname": n + surname,
    "Name.S": name + "." + s,
    "NameS": name + s,
    "N.S": n + "." + s,
    "NS": n + s
  };
  
  var emailPrefix = patterns[pattern] ? patterns[pattern].toLowerCase() : "";
  emailPrefix = emailPrefix.normalize("NFD").replace(/[^a-z0-9._-]/g, "");
  return emailPrefix ? emailPrefix + "@" + domain.toLowerCase() : "";
}
