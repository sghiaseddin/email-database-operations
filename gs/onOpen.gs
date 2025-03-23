function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Menu inject
  ui.createMenu('Custom')
    .addItem('Generate Email Guesses', 'generateEmailGuesses')
    .addItem('Highlight Undelivered Emails', 'highlightUndeliveredEmails')
    .addItem('List Same Domain Emails from Protected', 'listEmailsSameDomainProtected')
    .addItem('Analyze Email List for Consistency', 'analyzeEmailPatterns')
    .addItem("Extract Common Names from FullName", "extractCommonNamesFromFullname")
    .addItem("Extract Common Names from Email", "extractCommonNamesFromEmail")
    .addToUi();

  // Get hashed password from environment
  var scriptProperties = PropertiesService.getScriptProperties();
  var storedHash = scriptProperties.getProperty("PASSWORD_HASH");

  while (true) {
    var response = showPasswordPrompt();

    if (response.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
      var inputPassword = response.getResponseText();
      var inputHash = hashPassword(inputPassword);

      if (inputHash === storedHash) {
        SpreadsheetApp.getUi().alert("Access Granted!");
        createCustomMenu();
        break;
      } else {
        SpreadsheetApp.getUi().alert("Incorrect Password! Please try again.");
      }
    }
  }
}

function showPasswordPrompt() {
return SpreadsheetApp.getUi().prompt("Authentication Required", "Enter Password:", SpreadsheetApp.getUi().ButtonSet.OK);
}

function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return rawHash.map(function (byte) {
    return (byte + 256).toString(16).slice(-2);
  }).join("");
}
