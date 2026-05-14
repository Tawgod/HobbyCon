function doGet(e) {
  // Check the URL parameter to see which page to load
  if (e && e.parameter && e.parameter.mode === 'register') {
    return HtmlService.createHtmlOutputFromFile('Register')
        .setTitle('Hobby Con Registration')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Default to the Check-In page
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Convention Check-In')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processCheckIn(searchValue, searchTypeIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; 
  
  var today = new Date();
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var todayString = monthNames[today.getMonth()] + " " + today.getDate();
  
  var dateColIndex = headers.indexOf(todayString);
  
  if (dateColIndex === -1) {
    return "Error: Today is not an active event day.";
  }

  // Default to Column D (Index 3) if no specific type is passed
  var colIndex = (searchTypeIndex !== undefined) ? parseInt(searchTypeIndex) : 3;
  
  // Convert search to lowercase to handle case-insensitive Name/Email matches
  var normalizedSearch = String(searchValue).trim().toLowerCase();

  // Loop through the rows to find the match
  for (var i = 1; i < data.length; i++) {
    var cellValue = String(data[i][colIndex]).trim().toLowerCase();
    
    // If searching by Phone (Index 1), strip everything but numbers to ensure a clean match
    if (colIndex === 1) {
      cellValue = cellValue.replace(/\D/g, '');
      var tempSearch = normalizedSearch.replace(/\D/g, '');
      if (tempSearch.length > 0) normalizedSearch = tempSearch;
    }

    if (cellValue === normalizedSearch && normalizedSearch !== "") { 
      
      if (data[i][dateColIndex] !== "") {
        return "You are already checked in for today, " + data[i][0] + "!";
      }
      
      var timeString = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      sheet.getRange(i + 1, dateColIndex + 1).setValue("In at " + timeString);
      
      return "Success! Welcome, " + data[i][0] + "!";
    }
  }
  
  return "Record not found. Please try again or check your spelling.";
}

function processRegistration(name, phone, email, enteredCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  
  // Loop through existing records to check if the 6-digit code is already taken
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(enteredCode)) { // Index 3 is Column D
      return "Error: That 6-digit code is already taken. Please choose a different one.";
    }
  }
  
  // If the loop finishes without finding a match, the code is unique.
  // Append a new row to the bottom of the sheet: [Name, Phone, Email, Code]
  sheet.appendRow([name, phone, email, enteredCode]);
  
  return "Registration successful! See you at the convention!";
}