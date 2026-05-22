// ==========================================
// 1. MENU & UI TOOLS
// ==========================================


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
// 3. REGISTRATION & CHECK-IN LOGIC
// ==========================================
function processRegistration(name, phone, email, enteredCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(enteredCode)) { 
      return "Error: That 6-digit code is already taken. Please choose a different one.";
    }
  }
  sheet.appendRow([name, phone, email, enteredCode]);
  return "Registration successful! See you at the convention!";
}

function processCheckIn(searchValue, searchTypeIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; 
  var today = new Date();
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var todayString = monthNames[today.getMonth()] + " " + today.getDate();
  
  var dateColIndex = headers.indexOf(todayString);
  if (dateColIndex === -1) return "Error: Today is not an active event day.";

  var colIndex = (searchTypeIndex !== undefined) ? parseInt(searchTypeIndex) : 3;
  var normalizedSearch = String(searchValue).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var cellValue = String(data[i][colIndex]).trim().toLowerCase();
    if (colIndex === 1) {
      cellValue = cellValue.replace(/\D/g, '');
      var tempSearch = normalizedSearch.replace(/\D/g, '');
      if (tempSearch.length > 0) normalizedSearch = tempSearch;
    }
    if (cellValue === normalizedSearch && normalizedSearch !== "") { 
      if (data[i][dateColIndex] !== "") return "You are already checked in for today, " + data[i][0] + "!";
      var timeString = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      sheet.getRange(i + 1, dateColIndex + 1).setValue("In at " + timeString);
      return "Success! Welcome, " + data[i][0] + "!";
    }
  }
  return "Record not found. Please try again or check your spelling.";
}

// ==========================================
// 4. WEB APP SCHEDULE DATA FETCH
// ==========================================
function getScheduleData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule");
  if (!sheet) return []; 
  
  var data = sheet.getDataRange().getDisplayValues();
  var schedule = [];
  
  for (var i = 1; i < data.length; i++) {
    schedule.push({
      date: data[i][0],      
      time: data[i][1],      
      location: data[i][2],  
      title: data[i][3],     
      slots: data[i][4],     
      status: data[i][5],
      description: data[i][6] || "" // <--- NOW PULLS FROM COLUMN G
    });
  }
  return schedule;
}

// ==========================================
// 5. SIGNUP GENIUS SCHEDULE GENERATOR
// ==========================================
function generateScheduleSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entrySheet = ss.getSheetByName("Entry");
  if (!entrySheet) {
    SpreadsheetApp.getUi().alert("Error: Please create a tab named 'Entry' and paste your SignUpGenius data there.");
    return;
  }

  var rawData = entrySheet.getDataRange().getDisplayValues();
  var lines = [];
  for (var r = 0; r < rawData.length; r++) {
    for (var c = 0; c < rawData[r].length; c++) {
      var val = String(rawData[r][c]).trim();
      if (val.startsWith(',')) { val = val.substring(1).trim(); }
      if (val.startsWith('"') && val.endsWith('"')) { val = val.substring(1, val.length - 1).trim(); }
      if (val !== "") { lines.push(val); }
    }
  }

  var events = [];
  var currentDate = "Unknown Date";
  var tempDatePart = "";
  var currentLocation = "TBD";

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var dateMatch = line.match(/^(\d{2})\/(\d{2})\/2026$/);
    if (dateMatch) {
      var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      tempDatePart = monthNames[parseInt(dateMatch[1], 10) - 1] + " " + parseInt(dateMatch[2], 10);
      currentDate = tempDatePart; 
      continue;
    }

    var dayMatch = line.match(/^(Friday|Saturday|Sunday|Monday)$/i);
    if (dayMatch) {
      if (tempDatePart !== "") { currentDate = dayMatch[1] + ", " + tempDatePart; tempDatePart = ""; } 
      else { currentDate = dayMatch[1]; }
      continue;
    }

    if (/^\d{1,2}:\d{2}[a-zA-Z]{2}-$/.test(line)) {
      var startTime = line.replace('-', '').toUpperCase();
      var rawEndTime = lines[i+1] ? lines[i+1] : "";
      var endTimeMatch = rawEndTime.match(/(\d{1,2}:\d{2}\s*(?:AM|PM))/i);
      var endTime = endTimeMatch ? endTimeMatch[1].toUpperCase() : rawEndTime.toUpperCase();
      var title = lines[i+2] || "";
      var slotsStr = lines[i+3] || "";

      var total = 0, filled = 0, available = 0;
      var matchPartial = slotsStr.match(/(\d+) of (\d+) slots filled/i);
      var matchAll = slotsStr.match(/All (\d+) slots filled/i);

      if (matchPartial) { filled = parseInt(matchPartial[1]); total = parseInt(matchPartial[2]); available = total - filled; } 
      else if (matchAll) { total = parseInt(matchAll[1]); filled = total; available = 0; } 
      else { continue; }

      var status = (available === 0) ? "FULL" : available + " Available";
      var prevLine = lines[i-1] || "";
      if (/Table|Hallway|Room/i.test(prevLine)) { currentLocation = prevLine; }

      var descLine = lines[i+4] || "";
      var description = "";
      if (!/^\d{2}\/\d{2}\/2026$/.test(descLine) && !/^(Friday|Saturday|Sunday|Monday)$/i.test(descLine) &&
          !/^\d{1,2}:\d{2}[a-zA-Z]{2}-$/.test(descLine) && !/Table|Hallway|Room/i.test(descLine) &&
          !/Share|Sign Up|Available Slot/i.test(descLine)) {
        description = descLine;
      }

      events.push({
        date: currentDate, timeString: startTime + " - " + endTime,
        location: currentLocation, title: title, slots: filled + " / " + total,
        status: status, description: description
      });
    }
  }

  if (events.length === 0) { return; }

  // THE FIX: Memorize descriptions so every time slot gets one on the Schedule tab
  var knownDescriptions = {};
  for (var m = 0; m < events.length; m++) {
    if (events[m].description !== "") { knownDescriptions[events[m].title] = events[m].description; }
  }
  for (var m = 0; m < events.length; m++) {
    if (events[m].description === "" && knownDescriptions[events[m].title]) {
      events[m].description = knownDescriptions[events[m].title];
    }
  }

  // --- BUILD THE "ALL EVENTS" TAB (Untouched & Stable) ---
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

  var ssRef = SpreadsheetApp.getActiveSpreadsheet();
  var allEventsSheet = ssRef.getSheetByName("All Events");
  if (!allEventsSheet) { allEventsSheet = ssRef.insertSheet("All Events"); } else { allEventsSheet.clear(); }
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

  // --- BUILD THE "SCHEDULE" TAB (Now with Description in Column G) ---
  var scheduleSheet = ssRef.getSheetByName("Schedule");
  if (!scheduleSheet) { scheduleSheet = ssRef.insertSheet("Schedule"); } else { scheduleSheet.clear(); }

  var headers = ["Date", "Time", "Location", "Event", "Slots Filled", "Status", "Description"];
  var scheduleData = events.map(function(ev) {
    return [ev.date, ev.timeString, ev.location, ev.title, ev.slots, ev.status, ev.description];
  });

  scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
  scheduleSheet.getRange(2, 1, scheduleData.length, headers.length).setValues(scheduleData);
  scheduleSheet.autoResizeColumns(1, headers.length);
  
  var range = scheduleSheet.getRange(2, 6, scheduleSheet.getMaxRows() - 1, 1);
  var rules = scheduleSheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Available").setBackground("#d4edda").setFontColor("#155724").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FULL").setBackground("#f8d7da").setFontColor("#721c24").setRanges([range]).build());
  scheduleSheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert("Success! The descriptions are now written to Column G of the Schedule tab for the website to read.");
}

// ==========================================
// 6. PRINTABLE DAILY SCHEDULE GENERATOR
// ==========================================
function generatePrintableSchedules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  
  if (!scheduleSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find the 'Schedule' tab. Please run step 1 first.");
    return;
  }

  // Grab all the data from the Schedule tab
  var data = scheduleSheet.getDataRange().getDisplayValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("Error: No data found on the Schedule tab.");
    return;
  }

  // 1. Group data by Date, then by Location
  var daysData = {};
  
  // Skip row 0 (headers)
  for (var i = 1; i < data.length; i++) {
    var date = data[i][0];
    var time = data[i][1];
    var location = data[i][2];
    var title = data[i][3];
    var gm = title.match(/\((.*?)\)/) ? title.match(/\((.*?)\)/)[0] : ""; // Tries to extract (GM Name)
    var cleanTitle = title;
    
    if (!date || !location) continue;

    if (!daysData[date]) {
      daysData[date] = {};
    }
    if (!daysData[date][location]) {
      daysData[date][location] = [];
    }
    
    daysData[date][location].push({
      time: time,
      title: cleanTitle
    });
  }

  // 2. Custom Sort Function to handle "Table 1", "Table 2", "Table 10" properly
  function sortLocations(a, b) {
    var numA = parseInt(a.match(/\d+/)) || 0;
    var numB = parseInt(b.match(/\d+/)) || 0;
    var textA = a.replace(/\d+/g, '').trim();
    var textB = b.replace(/\d+/g, '').trim();
    
    if (textA === textB) {
      return numA - numB;
    }
    return textA.localeCompare(textB);
  }

  // 3. Create a Print Tab for each Day
  for (var day in daysData) {
    // Keep tab names clean and short (e.g., "Print - Friday")
    var shortDayName = day.split(',')[0]; 
    var tabName = "Print - " + shortDayName;
    
    var printSheet = ss.getSheetByName(tabName);
    if (!printSheet) {
      printSheet = ss.insertSheet(tabName);
    } else {
      printSheet.clear();
      // Unmerge everything to prevent formatting errors from previous runs
      printSheet.getDataRange().breakApart(); 
    }

    var outputData = [];
    var mergeRanges = []; // Keep track of rows to merge for the "Location" column
    var currentRow = 2;   // Start at row 2 because row 1 is headers

    // Sort the locations (Table 1, Table 2, etc.)
    var locations = Object.keys(daysData[day]).sort(sortLocations);

    for (var locIndex = 0; locIndex < locations.length; locIndex++) {
      var loc = locations[locIndex];
      var events = daysData[day][loc];
      
      var startMergeRow = currentRow;

      for (var evIndex = 0; evIndex < events.length; evIndex++) {
        var ev = events[evIndex];
        // Column A: Location, Column B: Time, Column C: Event
        outputData.push([loc, ev.time, ev.title]);
        currentRow++;
      }
      
      // If a table has multiple events, mark it for merging
      if (events.length > 1) {
        mergeRanges.push({startRow: startMergeRow, numRows: events.length});
      }
    }

    // 4. Write and Format the Sheet
    var headers = ["Table / Location", "Time Slot", "Scheduled Event"];
    printSheet.getRange(1, 1, 1, 3).setValues([headers])
              .setFontWeight("bold")
              .setBackground("#e2e3e5")
              .setFontSize(14)
              .setHorizontalAlignment("center");

    if (outputData.length > 0) {
      var dataRange = printSheet.getRange(2, 1, outputData.length, 3);
      dataRange.setValues(outputData);
      
      // Set nice print fonts
      dataRange.setFontSize(12);
      dataRange.setVerticalAlignment("middle");
      printSheet.getRange(2, 2, outputData.length, 1).setHorizontalAlignment("center"); // Center the time
      printSheet.getRange(2, 1, outputData.length, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(16); // Big bold table numbers
      
      // Merge the Location cells so it looks like a clean schedule
      for (var m = 0; m < mergeRanges.length; m++) {
        var range = printSheet.getRange(mergeRanges[m].startRow, 1, mergeRanges[m].numRows, 1);
        range.merge();
      }

      // Add borders for easy cutting/reading
      printSheet.getRange(1, 1, outputData.length + 1, 3).setBorder(true, true, true, true, true, true);
    }

    // Set column widths for 8.5x11 Portrait Printing
    printSheet.setColumnWidth(1, 150); // Location
    printSheet.setColumnWidth(2, 180); // Time
    printSheet.setColumnWidth(3, 400); // Event
    printSheet.getRange(1, 3, outputData.length + 1, 1).setWrap(true); // Wrap event titles
  }

  SpreadsheetApp.getUi().alert("Success! Your daily printable schedules have been generated as new tabs.");
}

// ==========================================
// 6. GOOGLE DOCS TABLE TENT GENERATOR
// ==========================================
function generateDocsTableTents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Schedule");
  
  if (!scheduleSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find the 'Schedule' tab. Please run step 1 first.");
    return;
  }

  var data = scheduleSheet.getDataRange().getDisplayValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("Error: No data found on the Schedule tab.");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Generating Google Docs... This may take a moment.", "Processing", 10);

  // Group data by Date, then by Location
  var daysData = {};
  
  for (var i = 1; i < data.length; i++) {
    var date = data[i][0];
    var time = data[i][1];
    var location = data[i][2];
    var title = data[i][3];
    
    if (!date || !location) continue;

    if (!daysData[date]) daysData[date] = {};
    if (!daysData[date][location]) daysData[date][location] = [];
    
    daysData[date][location].push({ time: time, title: title });
  }

  function sortLocations(a, b) {
    var numA = parseInt(a.match(/\d+/)) || 0;
    var numB = parseInt(b.match(/\d+/)) || 0;
    var textA = a.replace(/\d+/g, '').trim();
    var textB = b.replace(/\d+/g, '').trim();
    if (textA === textB) return numA - numB;
    return textA.localeCompare(textB);
  }

  var docLinks = [];

  // Build a Google Doc for each Day
  for (var day in daysData) {
    var shortDayName = day.split(',')[0]; 
    
    // Create a new Google Document
    var doc = DocumentApp.create("Hobby Con Table Tents - " + shortDayName);
    var body = doc.getBody();

    // Set Page to US Letter Landscape (11 x 8.5 inches)
    // Google Docs uses 'points' (72 points = 1 inch)
    body.setPageWidth(792).setPageHeight(612);
    
    // Set half-inch margins
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);
    
    // Clear the default empty paragraph
    body.clear(); 

    var locations = Object.keys(daysData[day]).sort(sortLocations);

    for (var locIndex = 0; locIndex < locations.length; locIndex++) {
      var loc = locations[locIndex];
      var events = daysData[day][loc];
      
      var tableNumMatch = loc.match(/\d+/);
      var tableNum = tableNumMatch ? tableNumMatch[0] : loc;
      var tablePrefix = loc.replace(tableNum, '').trim() || "Table";

      // Create a 1-Row, 2-Column Table for the Tent
      var table = body.appendTable();
      var row = table.appendTableRow();
      
      // ==========================================
      // LEFT CELL (FRONT OF TENT)
      // ==========================================
      var leftCell = row.appendTableCell();
      leftCell.setWidth(355); // Roughly half the page
      leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

      var p1 = leftCell.insertParagraph(0, tablePrefix);
      var attrPrefix = {};
      attrPrefix[DocumentApp.Attribute.FONT_SIZE] = 48;
      attrPrefix[DocumentApp.Attribute.BOLD] = true;
      attrPrefix[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      p1.setAttributes(attrPrefix);

      var p2 = leftCell.appendParagraph(tableNum);
      var attrNum = {};
      attrNum[DocumentApp.Attribute.FONT_SIZE] = 160; // MASSIVE NUMBER
      attrNum[DocumentApp.Attribute.BOLD] = true;
      attrNum[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      p2.setAttributes(attrNum);

      // ==========================================
      // RIGHT CELL (BACK OF TENT)
      // ==========================================
      var rightCell = row.appendTableCell();
      rightCell.setWidth(355);
      rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

      var headerPar = rightCell.insertParagraph(0, loc + " Events\n" + shortDayName + "\n");
      var attrHeader = {};
      attrHeader[DocumentApp.Attribute.FONT_SIZE] = 22;
      attrHeader[DocumentApp.Attribute.BOLD] = true;
      attrHeader[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      headerPar.setAttributes(attrHeader);

      for (var e = 0; e < events.length; e++) {
        var timePar = rightCell.appendParagraph(events[e].time);
        var attrTime = {};
        attrTime[DocumentApp.Attribute.FONT_SIZE] = 13;
        attrTime[DocumentApp.Attribute.BOLD] = true;
        timePar.setAttributes(attrTime);

        var titlePar = rightCell.appendParagraph(events[e].title + "\n");
        var attrTitle = {};
        attrTitle[DocumentApp.Attribute.FONT_SIZE] = 13;
        attrTitle[DocumentApp.Attribute.BOLD] = false;
        titlePar.setAttributes(attrTitle);
      }

      // Add a faint light-gray line down the middle so you know exactly where to fold it!
      table.setBorderWidth(1);
      table.setBorderColor('#dddddd');

      // Add a page break after every table (unless it's the very last one)
      if (locIndex < locations.length - 1) {
        body.appendPageBreak();
      }
    }
    
    // Save and grab the link to the document
    doc.saveAndClose();
    docLinks.push({ name: shortDayName, url: doc.getUrl() });
  }

  // ==========================================
  // SHOW THE DOWNLOAD LINKS TO THE USER
  // ==========================================
  var html = '<div style="font-family: Arial; padding: 10px;">';
  html += '<h3 style="color: #28a745;">✅ Your Printables are Ready!</h3>';
  html += '<p>Click the links below to open your Google Docs. They are already set to 8.5x11 Landscape with a faint gray folding line down the middle.</p>';
  html += '<ul style="line-height: 2.5;">';
  for (var j = 0; j < docLinks.length; j++) {
    html += '<li><a href="' + docLinks[j].url + '" target="_blank" style="font-size: 18px; color: #007bff; text-decoration: none; font-weight: bold;">📄 Open ' + docLinks[j].name + ' Tents</a></li>';
  }
  html += '</ul>';
  html += '<p style="color: #666; font-size: 12px; margin-top: 20px;"><em>Tip: If you want to add your logo, just click Insert > Image inside the Google Doc before you hit print!</em></p>';
  html += '</div>';

  var userInterface = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Google Docs Generated');
}