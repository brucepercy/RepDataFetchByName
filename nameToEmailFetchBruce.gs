function fetchEmailsAndPhoneNumbers() {
  // Open the sheets
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Data - Agent Activation"); // Master Data Sheet
  var lookupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NameToEmailFetchBruce"); // Lookup Sheet
  
  // Get master data
  var masterData = masterSheet.getDataRange().getValues(); // Get all data from Master Sheet
  
  // Get lookup names
  var lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 1).getValues(); // Column A (names)

  // Iterate through the Lookup Sheet names
  for (var j = 0; j < lookupData.length; j++) {
    var lookupName = lookupData[j][0]; // Name from Lookup Sheet (Column A)
    if (!lookupName) continue; // Skip if the name is empty
    
    var emailsFound = []; // Array to store matching emails
    var phoneNumbersFound = []; // Array to store matching phone numbers

    // Search for matches across all rows in the Master Data
    for (var i = 1; i < masterData.length; i++) {
      var row = masterData[i]; // Current row in Master Data
      var phone = row[3]; // Column D (Phone Number)
      var email = row[2]; // Column C (Email Address)
      
      // Check if the name exists in any cell of the row
      for (var k = 0; k < row.length; k++) {
        if (row[k] && row[k].toString().toLowerCase().includes(lookupName.toLowerCase())) {
          if (email && !emailsFound.includes(email)) emailsFound.push(email); // Add email to the list if not already added
          if (phone && !phoneNumbersFound.includes(phone)) phoneNumbersFound.push(phone); // Add phone number to the list if not already added
          break; // Stop searching the row once a match is found
        }
      }
    }

    // Write the first phone number into Column B
    lookupSheet.getRange(j + 2, 2).setValue(phoneNumbersFound.length > 0 ? phoneNumbersFound[0] : "Not Found");

    // Write the found emails into the Lookup Sheet (Columns C, D, E, ...)
    for (var col = 0; col < emailsFound.length; col++) {
      if (col >= 3) break; // Stop if emails exceed the available columns (C, D, E, F)
      lookupSheet.getRange(j + 2, 3 + col).setValue(emailsFound[col]); // Populate Columns C, D, E, ...
    }

    // Clear remaining columns if fewer matches are found
    for (var clearCol = emailsFound.length; clearCol < 4; clearCol++) {
      lookupSheet.getRange(j + 2, 3 + clearCol).clearContent(); // Clear Columns D, E, F if unused
    }
  }
}
