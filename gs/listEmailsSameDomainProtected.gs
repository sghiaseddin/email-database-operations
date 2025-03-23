function listEmailsSameDomainProtected() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var protectedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Protected");

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var domainCol = headers.indexOf("Domain") + 1;
  var lookupCol = headers.indexOf("Lookup Protected Emails") + 1;

  var protectedHeaders = protectedSheet.getRange(1, 1, 1, protectedSheet.getLastColumn()).getValues()[0];
  var protectedDomainCol = protectedHeaders.indexOf("Domain") + 1;
  var protectedEmailCol = protectedHeaders.indexOf("Email") + 1;

  if (domainCol === 0 || lookupCol === 0 || protectedDomainCol === 0 || protectedEmailCol === 0) {
    SpreadsheetApp.getUi().alert("Required columns are missing!");
    return;
  }

  var lastRow = sheet.getLastRow();
  var domains = sheet.getRange(2, domainCol, lastRow - 2).getValues().flat();
  var protectedData = protectedSheet.getRange(2, 1, protectedSheet.getLastRow() - 2, 2).getValues();

  // Create a map of domains to emails
  var domainEmailMap = {};
  protectedData.forEach(function (row) {
    var domain = row[0];
    var email = row[1];

    if (!domainEmailMap[domain]) {
      domainEmailMap[domain] = [];
    }

    if (domainEmailMap[domain].length < 5) {
      domainEmailMap[domain].push(email);
    }
  });

  // Update "Lookup Protected Emails" column
  var updatedEmails = domains.map(domain => {
    if (!domainEmailMap[domain] || domain == "") return "-";

    return domainEmailMap[domain]
      .map((email, index) => (index === 0 ? email : (index + 1) + ") " + email))
      .join(", ");
  });

  sheet.getRange(2, lookupCol, updatedEmails.length, 1).setValues(updatedEmails.map(e => [e]));
  sheet.getRange(2, lookupCol, updatedEmails.length, 1).setBackground("#ffffff");
}

