
function clearEntryTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entrySheet = ss.getSheetByName("Entry");
  if (!entrySheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find a tab named 'Entry'.");
    return;
  }
  entrySheet.clear();
  SpreadsheetApp.getUi().alert("Success! The Entry tab is completely clean and ready for a new paste.");
}

// ==========================================
// 2. WEB APP ROUTER
// ==========================================
function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'register') {
    return HtmlService.createHtmlOutputFromFile('Register')
        .setTitle('Hobby Con Registration')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.mode === 'schedule') {
    return HtmlService.createHtmlOutputFromFile('Schedule')
        .setTitle('Hobby Con Event Schedule')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.mode === 'home') {
    return HtmlService.createHtmlOutputFromFile('Home')
        .setTitle('Hobby Con Hub')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Convention Check-In')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 3. REGISTRATION, CHECK-IN, & WEB DATA LOGIC
// ==========================================
function processRegistration(name, phone, email, enteredCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  
  var stringCode = String(enteredCode).trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === stringCode) { 
      return "Error: That 6-digit code is already taken. Please choose a different one.";
    }
  }
  
  sheet.appendRow([name, phone, email, "'" + stringCode]);
  return "Registration successful! See you at the convention!";
}

function processCheckIn(searchValue, searchTypeIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    
    if (!sheet) return "Error: Could not find the 'Attendance' tab. Please check the spelling.";

    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0]; 
    
    var today = new Date();
    var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var todayString = monthNames[today.getMonth()] + " " + today.getDate();
    
    var dateColIndex = headers.indexOf(todayString);
    if (dateColIndex === -1) {
      return "Error: Today (" + todayString + ") is not listed as a column header in the Attendance tab.";
    }

    var colIndex = (searchTypeIndex !== undefined) ? parseInt(searchTypeIndex) : 3;
    var normalizedSearch = String(searchValue).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var cellValue = String(data[i][colIndex]).replace(/^'/, '').trim().toLowerCase();
      
      if (colIndex === 1) { 
        cellValue = cellValue.replace(/\D/g, '');
        var tempSearch = normalizedSearch.replace(/\D/g, '');
        if (tempSearch.length > 0) normalizedSearch = tempSearch;
      }

      if (cellValue === normalizedSearch && normalizedSearch !== "") { 
        
        // --- NEW RULE: Check if the PIN code is missing! ---
        var currentPin = String(data[i][3]).replace(/^'/, '').trim();
        if (currentPin === "") {
          // Tell the website to ask for a PIN, and tell it which Row we are on
          return "REQUIRE_NEW_PIN|" + i;
        }

        if (data[i][dateColIndex] !== "") {
          return "You are already checked in for today, " + data[i][0] + "!";
        }
        var timeString = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        sheet.getRange(i + 1, dateColIndex + 1).setValue("In at " + timeString);
        return "Success! Welcome, " + data[i][0] + "!";
      }
    }
    return "Record not found. Please check your spelling or verify the code.";
  } catch(e) {
    return "System Crash: " + e.message; 
  }
}

// --- NEW FUNCTION TO SAVE THE CREATED PIN AND FINISH CHECK IN ---
function updatePinAndCheckIn(rowIndex, newPin) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Attendance");
    var data = sheet.getDataRange().getValues();

    var stringCode = String(newPin).trim();
    if (stringCode.length !== 6) return "Error: PIN must be exactly 6 digits.";

    // Make sure they didn't pick a code someone else is already using
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).replace(/^'/, '').trim() === stringCode) {
        return "Error: That 6-digit code is already taken. Please choose a different one.";
      }
    }

    // Save the new PIN to their row (with the apostrophe to protect leading zeros)
    sheet.getRange(parseInt(rowIndex) + 1, 4).setValue("'" + stringCode);

    // Finish the Check-In process
    var headers = data[0]; 
    var today = new Date();
    var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var todayString = monthNames[today.getMonth()] + " " + today.getDate();
    var dateColIndex = headers.indexOf(todayString);
    
    if (dateColIndex !== -1) {
      if (data[rowIndex][dateColIndex] !== "") {
         return "Success! Your new PIN was saved, but you were already checked in for today, " + data[rowIndex][0] + "!";
      }
      var timeString = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      sheet.getRange(parseInt(rowIndex) + 1, dateColIndex + 1).setValue("In at " + timeString);
      return "Success! New code saved. Welcome, " + data[rowIndex][0] + "!";
    } else {
       return "Success! Your new code was saved, but today is not an active event day.";
    }
  } catch(e) {
    return "System Crash: " + e.message;
  }
}

function getScheduleData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule");
  if (!sheet) return []; 
  var data = sheet.getDataRange().getDisplayValues();
  var schedule = [];
  for (var i = 1; i < data.length; i++) {
    schedule.push({
      date: data[i][0], time: data[i][1], location: data[i][2],  
      title: data[i][3], slots: data[i][4], status: data[i][5],
      description: data[i][6] || ""
    });
  }
  return schedule;
}

// ==========================================
// 4. UNIVERSAL DATA PARSER
// ==========================================
function parseRawEntryData(lines) {
  var events = [];
  var currentDate = "Unknown Date";

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    var dateMatch = line.match(/^(\d{1,2})\/(\d{1,2})\/2026/);
    if (dateMatch) {
      var month = parseInt(dateMatch[1], 10);
      var day = parseInt(dateMatch[2], 10);
      var d = new Date(2026, month - 1, day);
      var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
      var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      currentDate = daysOfWeek[d.getDay()] + ", " + monthNames[month - 1] + " " + day;
      continue;
    }

    var looseDateMatch = line.match(/^(May)\s+(\d{2})/i);
    if (looseDateMatch) {
      currentDate = looseDateMatch[1] + " " + looseDateMatch[2];
    }

    if (/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i.test(line)) {
      continue;
    }

    var timeMatchNew = line.match(/^(\d{1,2}:\d{2}\s*[a-zA-Z]{2})\s*-\s*(\d{1,2}:\d{2}\s*[a-zA-Z]{2})/i);
    var timeMatchOld = line.match(/^(\d{1,2}:\d{2}[a-zA-Z]{2})-$/i);

    if (timeMatchNew || timeMatchOld) {
      var startTime, endTime, title;
      var forwardOffset = 0;

      if (timeMatchNew) {
        startTime = timeMatchNew[1].toUpperCase();
        endTime = timeMatchNew[2].toUpperCase();
        title = lines[i+1] || "";
        forwardOffset = 1;
      } else {
        startTime = line.replace('-', '').toUpperCase();
        var rawEndTime = lines[i+1] ? lines[i+1] : "";
        var endTimeMatch = rawEndTime.match(/(\d{1,2}:\d{2}\s*(?:AM|PM))/i);
        endTime = endTimeMatch ? endTimeMatch[1].toUpperCase() : rawEndTime.toUpperCase();
        title = lines[i+2] || "";
        forwardOffset = 2;
      }

      var prevLine = lines[i-1] || "";
      var currentLocation = "TBD";
      if (/Table|Hallway|Room/i.test(prevLine)) { currentLocation = prevLine; }

      var description = "";
      var total = 0, filled = 0, available = 0;
      var foundSlots = false;

      for (var k = forwardOffset + 1; k <= forwardOffset + 6; k++) {
        var lookAhead = lines[i+k];
        if (!lookAhead) break;
        if (/^(\d{1,2}:\d{2})/.test(lookAhead) || /Table|Hallway|Room/i.test(lookAhead)) break;

        var slotMatch = lookAhead.match(/(\d+)\s+of\s+(\d+)\s+slots\s+filled/i);
        var allMatch = lookAhead.match(/All Slots Filled/i);

        if (slotMatch) {
          filled = parseInt(slotMatch[1]);
          total = parseInt(slotMatch[2]);
          available = total - filled;
          foundSlots = true;
        } else if (allMatch) {
          var capMatch = lines[i+forwardOffset+1].match(/^-(\d+)$/);
          total = capMatch ? parseInt(capMatch[1]) : 0;
          filled = total;
          available = 0;
          foundSlots = true;
        } else if (!lookAhead.match(/^-/) && !lookAhead.match(/^Sign Up/i) && lookAhead.length > 20) {
          description = lookAhead;
        }
      }

      var status = (available === 0 && foundSlots) ? "FULL" : available + " Available";

      var finalDate = currentDate;
      if (finalDate === "May 22" || finalDate.includes("05/22") || finalDate.includes("Friday")) finalDate = "Friday, May 22";
      if (finalDate === "May 23" || finalDate.includes("05/23") || finalDate.includes("Saturday")) finalDate = "Saturday, May 23";
      if (finalDate === "May 24" || finalDate.includes("05/24") || finalDate.includes("Sunday")) finalDate = "Sunday, May 24";
      if (finalDate === "May 25" || finalDate.includes("05/25") || finalDate.includes("Monday")) finalDate = "Monday, May 25";

      events.push({
        date: finalDate, 
        timeString: startTime + " - " + endTime,
        location: currentLocation, 
        title: title, 
        slots: foundSlots ? (filled + " / " + total) : "?",
        status: status, 
        description: description
      });
    }
  }

  var knownDescriptions = {};
  for (var m = 0; m < events.length; m++) {
    if (events[m].description !== "") knownDescriptions[events[m].title] = events[m].description;
  }
  for (var m = 0; m < events.length; m++) {
    if (events[m].description === "" && knownDescriptions[events[m].title]) {
      events[m].description = knownDescriptions[events[m].title];
    }
  }

  return events;
}

function getFlattenedEntryLines(ss) {
  var entrySheet = ss.getSheetByName("Entry");
  if (!entrySheet) return null;
  var rawData = entrySheet.getDataRange().getDisplayValues();
  var lines = [];
  for (var r = 0; r < rawData.length; r++) {
    for (var c = 0; c < rawData[r].length; c++) {
      var val = String(rawData[r][c]).trim();
      if (val.startsWith(',')) val = val.substring(1).trim();
      if (val.startsWith('"') && val.endsWith('"')) val = val.substring(1, val.length - 1).trim();
      if (val !== "") lines.push(val);
    }
  }
  return lines;
}

// ==========================================
// 5. GENERATE WEB SCHEDULE TAB
// ==========================================
function generateScheduleSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = getFlattenedEntryLines(ss);
  if (!lines) { SpreadsheetApp.getUi().alert("Error: Entry tab not found."); return; }

  var events = parseRawEntryData(lines);
  if (events.length === 0) { SpreadsheetApp.getUi().alert("No events found."); return; }

  var scheduleSheet = ss.getSheetByName("Schedule");
  if (!scheduleSheet) { scheduleSheet = ss.insertSheet("Schedule"); } else { scheduleSheet.clear(); }

  var headers = ["Date", "Time", "Location", "Event", "Slots Filled", "Status", "Description"];
  var scheduleData = events.map(function(ev) {
    return [ev.date, ev.timeString, ev.location, ev.title, ev.slots, ev.status, ev.description];
  });

  scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
  scheduleSheet.getRange(2, 1, scheduleData.length, headers.length).setValues(scheduleData);
  scheduleSheet.autoResizeColumns(1, headers.length);
  scheduleSheet.hideColumns(7); 
  
  var range = scheduleSheet.getRange(2, 6, scheduleSheet.getMaxRows() - 1, 1);
  var rules = scheduleSheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Available").setBackground("#d4edda").setFontColor("#155724").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FULL").setBackground("#f8d7da").setFontColor("#721c24").setRanges([range]).build());
  scheduleSheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert("Success! Master Schedule updated for the Web App.");
}

// ==========================================
// 6. ALL EVENTS CATALOG GENERATOR
// ==========================================
function generateAllEventsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = getFlattenedEntryLines(ss);
  if (!lines) { SpreadsheetApp.getUi().alert("Error: Entry tab not found."); return; }

  var events = parseRawEntryData(lines);
  if (events.length === 0) { SpreadsheetApp.getUi().alert("No events found in the Entry tab."); return; }

  var grouped = {};
  for (var j = 0; j < events.length; j++) {
    var ev = events[j];
    if (!grouped[ev.title]) { grouped[ev.title] = { description: ev.description, dates: {} }; }
    if (!grouped[ev.title].dates[ev.date]) { grouped[ev.title].dates[ev.date] = []; }
    var timeAndLocation = ev.timeString + "\n" + ev.location;
    if (grouped[ev.title].dates[ev.date].indexOf(timeAndLocation) === -1) {
      grouped[ev.title].dates[ev.date].push(timeAndLocation);
    }
  }

  var allEventsOutput = [];
  var rowTypes = []; 
  var maxCols = 2; 

  for (var eventName in grouped) {
    var eventData = grouped[eventName];
    allEventsOutput.push([eventName]); rowTypes.push('title');
    if (eventData.description) { allEventsOutput.push([eventData.description]); rowTypes.push('description'); }
    for (var d in eventData.dates) {
      var times = eventData.dates[d];
      var dateRow = [d].concat(times); 
      if (dateRow.length > maxCols) { maxCols = dateRow.length; }
      allEventsOutput.push(dateRow); rowTypes.push('date');
    }
    allEventsOutput.push([]); rowTypes.push('blank');
  }

  for (var k = 0; k < allEventsOutput.length; k++) {
    while (allEventsOutput[k].length < maxCols) { allEventsOutput[k].push(""); }
  }

  var allEventsSheet = ss.getSheetByName("All Events");
  if (!allEventsSheet) { allEventsSheet = ss.insertSheet("All Events"); } else { allEventsSheet.clear(); }
  allEventsSheet.getRange(1, 1, allEventsOutput.length, maxCols).setValues(allEventsOutput);

  var backgrounds = []; var fontWeights = []; var fontSizes = [];
  for(var r = 0; r < rowTypes.length; r++) {
    var bgRow = []; var fontRow = []; var sizeRow = [];
    for(var c = 0; c < maxCols; c++) {
      if (rowTypes[r] === 'title') { bgRow.push("#d9edf7"); fontRow.push("bold"); sizeRow.push(14); } 
      else if (rowTypes[r] === 'description') { bgRow.push("#f4f8fa"); fontRow.push("normal"); sizeRow.push(10); } 
      else if (rowTypes[r] === 'date') { bgRow.push("#ffffff"); fontRow.push(c === 0 ? "bold" : "normal"); sizeRow.push(11); } 
      else { bgRow.push("#ffffff"); fontRow.push("normal"); sizeRow.push(11); }
    }
    backgrounds.push(bgRow); fontWeights.push(fontRow); fontSizes.push(sizeRow);
  }

  var formatRange = allEventsSheet.getRange(1, 1, allEventsOutput.length, maxCols);
  formatRange.setBackgrounds(backgrounds); formatRange.setFontWeights(fontWeights); formatRange.setFontSizes(fontSizes);
  
  allEventsSheet.setColumnWidth(1, 200); 
  for (var col = 2; col <= maxCols; col++) { allEventsSheet.setColumnWidth(col, 180); }
  
  var currentRow = 1;
  for (var r2 = 0; r2 < rowTypes.length; r2++) {
    if ((rowTypes[r2] === 'title' || rowTypes[r2] === 'description') && maxCols > 1) {
      allEventsSheet.getRange(currentRow, 1, 1, maxCols).mergeAcross();
    }
    currentRow++;
  }
  formatRange.setWrap(true); 

  SpreadsheetApp.getUi().alert("Success! Your 'All Events' catalog has been generated.");
}

// ==========================================
// 7. NEW DAILY SCHEDULE TABS
// ==========================================
function generateDailyTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = getFlattenedEntryLines(ss);
  if (!lines) { SpreadsheetApp.getUi().alert("Error: Entry tab not found."); return; }

  var events = parseRawEntryData(lines);
  if (events.length === 0) { SpreadsheetApp.getUi().alert("No events found in the Entry tab."); return; }

  var daysData = {};
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!daysData[ev.date]) { daysData[ev.date] = []; }
    daysData[ev.date].push(ev);
  }

  var headers = ["Time", "Location", "Event", "Slots Filled", "Status", "Description"];

  for (var dateString in daysData) {
    var shortDayName = dateString.split(',')[0]; 
    var tabName = shortDayName + " Schedule"; 

    var daySheet = ss.getSheetByName(tabName);
    if (!daySheet) {
      daySheet = ss.insertSheet(tabName);
    } else {
      daySheet.clear();
    }

    var dayEvents = daysData[dateString];
    var outputData = [];

    for (var j = 0; j < dayEvents.length; j++) {
      var e = dayEvents[j];
      outputData.push([e.timeString, e.location, e.title, e.slots, e.status, e.description]);
    }

    daySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#e2e3e5");
    daySheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);

    daySheet.setColumnWidth(1, 150); 
    daySheet.setColumnWidth(2, 120); 
    daySheet.setColumnWidth(3, 300); 
    daySheet.setColumnWidth(4, 100); 
    daySheet.setColumnWidth(5, 120); 
    daySheet.setColumnWidth(6, 400); 

    daySheet.getRange(2, 3, outputData.length, 1).setWrap(true);
    daySheet.getRange(2, 6, outputData.length, 1).setWrap(true);

    if(outputData.length > 0) {
      var range = daySheet.getRange(2, 5, outputData.length, 1);
      var rules = daySheet.getConditionalFormatRules();
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Available").setBackground("#d4edda").setFontColor("#155724").setRanges([range]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FULL").setBackground("#f8d7da").setFontColor("#721c24").setRanges([range]).build());
      daySheet.setConditionalFormatRules(rules);
    }
  }

  SpreadsheetApp.getUi().alert("Success! Individual schedule tabs have been generated for each day.");
}