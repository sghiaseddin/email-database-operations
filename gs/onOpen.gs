function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
    .addItem('Generate Email Guesses', 'generateEmailGuesses')
    .addItem('Highlight Undelivered Emails', 'highlightUndeliveredEmails')
    .addToUi();
}
