function analyzeEmailPatterns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var protectedSheet = ss.getSheetByName("Protected");
  var commonNamesSheet = ss.getSheetByName("Common Names");

  if (!protectedSheet || !commonNamesSheet) {
    SpreadsheetApp.getUi().alert("Required sheets are missing!");
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var domainCol = headers.indexOf("Domain") + 1;
  var analyzedCol = headers.indexOf("Analyzed") + 1;
  var lookupCol = headers.indexOf("Lookup Protected Emails") + 1;
  if (lookupCol === 0 || domainCol === 0 || analyzedCol === 0) {
    SpreadsheetApp.getUi().alert("Lookup Protected Emails column is missing!");
    return;
  }

  var lastRow = sheet.getLastRow();
  var emailLists = sheet.getRange(2, lookupCol, lastRow - 2).getValues().flat();
  var commonNames = commonNamesSheet.getRange(2, 1, commonNamesSheet.getLastRow(), 2).getValues();

  var nameList = commonNames.map(row => row[0]).filter(Boolean);
  var surnameList = commonNames.map(row => row[1]).filter(Boolean);

  var processedEmails = [];
  var patternsFound = [];
  var previousDomain = null;

  for (var i = 0; i < emailLists.length; i++) {
    var isAnalyzed = sheet.getRange(i + 2, analyzedCol).getValue();
    if (isAnalyzed) continue; // Skip already processed rows

    var currentDomain = sheet.getRange(i + 2, domainCol).getValue(); // Get domain for this row

    if (currentDomain === previousDomain) {
      processedEmails.push(["-"]); // Skip processing, mark as skipped
      patternsFound.push([]); // Keep index alignment
      continue;
    }

    previousDomain = currentDomain; // Update for next comparison
    var emailList = emailLists[i];
    if (!emailList || emailList === "-") {
      processedEmails.push(["-"]);
      patternsFound.push([]);
      continue;
    }

    var emails = emailList.split(/, \d+\) /).map(e => e.replace(/^\d+\) /, "").trim());
    var analyzedEmails = [];
    var localPatterns = [];

    emails.forEach(email => {
      Logger.log(email);
      var localPart = email.split("@")[0].toLowerCase();
      var identifiedPattern = "Unknown";
      var maxProbability = 0;
      var formattedEmail = email;

      // Detect separator-based patterns (100% sure)
      if (localPart.includes(".")) {
        let parts = localPart.split(".");
        if (parts.length === 2) {
          let [n, s] = parts;
          if (nameList.includes(n) && surnameList.includes(s)) {
            identifiedPattern = "Name.Surname";
            maxProbability = 100;
            formattedEmail = capitalize(n) + "." + capitalize(s) + "@" + email.split("@")[1];
            // Logger.log("Easy Guess: " + formattedEmail);
          }
        }
      } else if (localPart.includes("_")) {
        let parts = localPart.split("_");
        if (parts.length === 2) {
          let [n, s] = parts;
          if (nameList.includes(n) && surnameList.includes(s)) {
            identifiedPattern = "Name_Surname";
            maxProbability = 100;
            formattedEmail = capitalize(n) + "_" + capitalize(s) + "@" + email.split("@")[1];
          }
        }
      } else {
        // Find longest matching name
        let bestMatch = "";
        let rest = "";

        nameList.forEach(name => {
          if (localPart.startsWith(name) && name.length > bestMatch.length) {
            bestMatch = name;
            rest = localPart.slice(name.length);
          }
        });

        if (bestMatch) {
          Logger.log("Name Guess: " + bestMatch + " " + rest);
          let remainingLength = rest.length;

          if (remainingLength === 0) {
            identifiedPattern = "Name";
            maxProbability = 100;
            formattedEmail = capitalize(bestMatch) + "@" + email.split("@")[1];
          } else if (remainingLength === 1) {
            identifiedPattern = "NameS";
            maxProbability = 100;
            formattedEmail = capitalize(bestMatch) + rest.toUpperCase() + "@" + email.split("@")[1];
          } else {
            let guessedSurname = binarySearchStartsWith(surnameList, rest);
            if (guessedSurname) {
              Logger.log("Surname Guess: " + guessedSurname);
              let remainingAfterSurname = rest.slice(guessedSurname.length);
              maxProbability = 100 - remainingAfterSurname.length * 10;
              identifiedPattern = "NameSurname";
              formattedEmail = capitalize(bestMatch) + capitalize(guessedSurname) + "@" + email.split("@")[1];
            }
          }
        }
      }

      analyzedEmails.push(formattedEmail);
      localPatterns.push(identifiedPattern);
    });

    var formatedEmailList = analyzedEmails
      .map((email, index) => (index === 0 ? email : (index + 1) + ") " + email))
      .join(", ");

    var uniquePatterns = [...new Set(localPatterns)];
    var formattedOutput = "[" + uniquePatterns.join(", ") + "]";
    formattedOutput += " Sample: " + formatedEmailList;
    sheet.getRange(i + 2, lookupCol).setValue(formattedOutput);
    sheet.getRange(i + 2, analyzedCol).setValue(true); // Mark row as analyzed

    var backgroundColor = "#ffffff"; // Default white
    if (formattedOutput === "-") {
      backgroundColor = "#ffffff";
    } else if (uniquePatterns.length === 1) {
      backgroundColor = "#eeffee";
    } else if (uniquePatterns.length === 2) {
      backgroundColor = "#ffefef";
    } else {
      backgroundColor = "#ffeded";
    }
    sheet.getRange(i + 2, lookupCol).setBackground(backgroundColor);
  };


  SpreadsheetApp.getUi().alert("Email consistency analysis completed!");
}

// Helper function to capitalize names
function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function binarySearchStartsWith(sortedList, target) {
  let left = 0, right = sortedList.length - 1;
  while (left <= right) {
    let mid = Math.floor((left + right) / 2);
    if (sortedList[mid] === target || target.startsWith(sortedList[mid])) {
      return sortedList[mid]; // Found a match
    } else if (sortedList[mid] < target) {
      left = mid + 1;
    } else {
      right = mid - 1;
    }
  }
  return null; // No match found
}