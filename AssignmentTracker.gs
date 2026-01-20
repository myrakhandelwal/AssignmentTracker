/**
 * Assignment Tracker - Google Apps Script
 * Creates quarterly assignment sheets and syncs with Google Calendar
 */

// Run this function when the spreadsheet opens
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Assignment Tracker')
      .addItem('Create New Quarter Sheet', 'createQuarterSheet')
      .addItem('Add Assignment to Calendar', 'addAssignmentToCalendar')
      .addItem('Sync All Assignments to Calendar', 'syncAllToCalendar')
      .addItem('Manage Classes', 'openClassManager')
      .addItem('Setup Classes & Recurring', 'setupClassesAndRecurringAssignments')
      .addToUi();
}

/**
 * Creates a new sheet for the current quarter
 */
function createQuarterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // Get quarter name from user
  var response = ui.prompt('Create New Quarter Sheet', 
                           'Enter quarter name (e.g., "Q1 2026" or "Winter 2026"):', 
                           ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  var quarterName = response.getResponseText().trim();
  
  if (!quarterName) {
    ui.alert('Please enter a valid quarter name.');
    return;
  }
  
  // Check if sheet already exists
  var existingSheet = ss.getSheetByName(quarterName);
  if (existingSheet) {
    ui.alert('A sheet named "' + quarterName + '" already exists.');
    return;
  }
  
  // Create new sheet
  var sheet = ss.insertSheet(quarterName);
  
  // Set up headers
  setupSheetHeaders(sheet);
  
  // Ask how many classes and their names for this term
  var classesCountResp = ui.prompt('Classes This Term', 'How many classes are you taking this term?', ui.ButtonSet.OK_CANCEL);
  if (classesCountResp.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var classesCount = parseInt(classesCountResp.getResponseText(), 10);
  if (!classesCount || classesCount <= 0 || classesCount > 20) {
    ui.alert('Please enter a valid number of classes (1-20).');
    return;
  }
  var classes = [];
  for (var i = 1; i <= classesCount; i++) {
    var classResp = ui.prompt('Class ' + i, 'Enter class name:', ui.ButtonSet.OK_CANCEL);
    if (classResp.getSelectedButton() != ui.Button.OK) {
      return;
    }
    var className = classResp.getResponseText().trim();
    if (!className) {
      ui.alert('Class name cannot be empty. Try again.');
      i--;
      continue;
    }
    classes.push(className);
  }
  
  // Format the sheet
  formatSheet(sheet, classes);
  
  // Store classes in sheet properties
  setSheetClasses(sheet, classes);
  
  // Auto-group by week
  groupByWeek();
  
  ui.alert('Success', 'Sheet "' + quarterName + '" has been created!\nClasses: ' + classes.join(', '), ui.ButtonSet.OK);
}

/**
 * Sets up the header row for the assignment sheet
 */
function setupSheetHeaders(sheet) {
  var headers = [
    'Assignment Name',
    'Course',
    'Due Date',
    'Time',
    'Status',
    'Notes',
    'Calendar Event ID',
    'Quiz',
    'Midterm',
    'Final Exam',
    'iCalendar File URL'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Bold and freeze header row
  sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
}

/**
 * Formats the sheet with proper column widths and data validation
 */
function formatSheet(sheet, classes) {
  // Set column widths
  sheet.setColumnWidth(1, 200); // Assignment Name
  sheet.setColumnWidth(2, 150); // Course
  sheet.setColumnWidth(3, 120); // Due Date
  sheet.setColumnWidth(4, 100); // Time
  sheet.setColumnWidth(5, 120); // Status
  sheet.setColumnWidth(6, 250); // Notes
  sheet.setColumnWidth(7, 150); // Calendar Event ID (hidden)
  sheet.setColumnWidth(8, 90);  // Quiz
  sheet.setColumnWidth(9, 90); // Midterm
  sheet.setColumnWidth(10, 100); // Final Exam
  sheet.setColumnWidth(11, 220); // iCalendar File URL
  
  // Add data validation for Status column (E)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'In Progress', 'Completed', 'Submitted'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(statusRule);
  
  // Set data validation for Course column based on entered classes
  try {
    var courseRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(classes, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('B2:B1000').setDataValidation(courseRule);
  } catch (e) {
    Logger.log('Error setting course validation: ' + e.message);
  }
  
  // Add conditional formatting for Status
  var completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#b7e1cd')
    .setRanges([sheet.getRange('E2:E1000')])
    .build();
  
  var submittedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Submitted')
    .setBackground('#a4c2f4')
    .setRanges([sheet.getRange('E2:E1000')])
    .build();
  
  var rules = sheet.getConditionalFormatRules();
  rules.push(completedRule, submittedRule);
  sheet.setConditionalFormatRules(rules);
  
  // Hide the Calendar Event ID column
  sheet.hideColumns(7);
  
  // Add checkboxes for Quiz/Midterm/Final
  sheet.getRange('H2:J1000').insertCheckboxes();

  // Add alternating row colors
  sheet.getRange('A2:K1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * Store classes in sheet properties
 */
function setSheetClasses(sheet, classes) {
  var props = PropertiesService.getDocumentProperties();
  var key = 'classes_' + sheet.getName();
  props.setProperty(key, JSON.stringify(classes));
}

/**
 * Retrieve classes from sheet properties
 */
function getSheetClasses(sheet) {
  var props = PropertiesService.getDocumentProperties();
  var key = 'classes_' + sheet.getName();
  var stored = props.getProperty(key);
  return stored ? JSON.parse(stored) : [];
}

/**
 * Open a sidebar to manage classes
 */
function openClassManager() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var classes = getSheetClasses(sheet);
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body { font-family: Arial; padding: 10px; margin: 0; }' +
    '.class-item { display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid #ddd; }' +
    'button { padding: 4px 8px; margin-left: 5px; cursor: pointer; }' +
    'input { width: 75%; padding: 5px; }' +
    '.add-section { margin-bottom: 15px; }' +
    '</style>' +
    '<h3>Manage Classes</h3>' +
    '<div class="add-section">' +
    '<input type="text" id="newClass" placeholder="New class name">' +
    '<button onclick="addClass()">Add</button>' +
    '</div>' +
    '<div id="classList"></div>' +
    '<button onclick="save()" style="margin-top: 15px; width: 100%; padding: 8px; background: #4285f4; color: white; border: none; cursor: pointer;">Save & Close</button>' +
    '<script>' +
    'var classes = ' + JSON.stringify(classes) + ';' +
    'function renderClasses() {' +
    '  var html = "";' +
    '  classes.forEach((c, i) => {' +
    '    html += "<div class=\\"class-item\\"><span>" + c + "</span><button onclick=\\"deleteClass(" + i + ")\\"Delete</button></div>";' +
    '  });' +
    '  document.getElementById("classList").innerHTML = html;' +
    '}' +
    'function addClass() {' +
    '  var input = document.getElementById("newClass");' +
    '  if (input.value.trim()) {' +
    '    classes.push(input.value.trim());' +
    '    input.value = "";' +
    '    renderClasses();' +
    '  }' +
    '}' +
    'function deleteClass(i) {' +
    '  classes.splice(i, 1);' +
    '  renderClasses();' +
    '}' +
    'function save() {' +
    '  google.script.run.saveClasses(classes);' +
    '  google.script.host.close();' +
    '}' +
    'renderClasses();' +
    '</script>'
  );
  SpreadsheetApp.getUi().showModelessDialog(html, 'Class Manager');
}

/**
 * Save classes from sidebar
 */
function saveClasses(classes) {
  var sheet = SpreadsheetApp.getActiveSheet();
  setSheetClasses(sheet, classes);
  
  // Update Course column validation
  if (classes.length > 0) {
    var courseRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(classes, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('B2:B1000').setDataValidation(courseRule);
  }
}

/**
 * Group assignments by week (Sunday to Saturday)
 */
function groupByWeek() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var grouped = {};
  var weeks = [];
  
  // Group by week
  for (var i = 0; i < data.length; i++) {
    var dueDate = new Date(data[i][2]);
    if (!dueDate || isNaN(dueDate.getTime())) continue;
    
    // Calculate week start (Sunday)
    var weekStart = new Date(dueDate);
    var day = weekStart.getDay();
    weekStart.setDate(weekStart.getDate() - day);
    weekStart.setHours(0, 0, 0, 0);
    
    var weekKey = weekStart.getTime();
    if (!grouped[weekKey]) {
      grouped[weekKey] = [];
      weeks.push(weekKey);
    }
    grouped[weekKey].push(data[i]);
  }
  
  // Sort weeks
  weeks.sort(function(a, b) { return a - b; });
  
  // Clear data rows
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
  }
  
  // Write grouped data with week headers
  var row = 2;
  for (var w = 0; w < weeks.length; w++) {
    var weekStart = new Date(weeks[w]);
    var weekEnd = new Date(weekStart.getTime() + 6 * 24 * 60 * 60 * 1000);
    var weekHeader = 'Week of ' + Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'MMM dd') + ' - ' + Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MMM dd');
    
    // Insert week header
    var headerRange = sheet.getRange(row, 1, 1, 11);
    headerRange.setValue(weekHeader);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');
    row++;
    
    // Sort assignments by due date within week
    var weekAssignments = grouped[weeks[w]];
    weekAssignments.sort(function(a, b) {
      var dateA = new Date(a[2]) || new Date();
      var dateB = new Date(b[2]) || new Date();
      return dateA - dateB;
    });
    
    // Write assignments
    for (var a = 0; a < weekAssignments.length; a++) {
      sheet.getRange(row, 1, 1, 11).setValues([weekAssignments[a]]);
      row++;
    }
  }
}

/**
 * Adds a single assignment to Google Calendar
 */
function addAssignmentToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var activeRange = sheet.getActiveRange();
  var row = activeRange.getRow();
  
  // Check if a valid row is selected (not header)
  if (row < 2) {
    ui.alert('Please select a row with assignment data (row 2 or below).');
    return;
  }
  
  var data = sheet.getRange(row, 1, 1, 11).getValues()[0];
  var assignmentName = data[0];
  var course = data[1];
  var dueDate = data[2];
  var time = data[3];
  var status = data[4];
  var notes = data[5];
  var eventId = data[6];
  var isQuiz = data[7];
  var isMidterm = data[8];
  var isFinal = data[9];
  var icsUrl = data[10];
  
  if (!assignmentName || !dueDate) {
    ui.alert('Assignment Name and Due Date are required.');
    return;
  }
  
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var eventDate = new Date(dueDate);
    
    // If time is specified, parse and set it
    if (time) {
      var timeMatch = String(time).match(/(\d+):(\d+)/);
      if (timeMatch) {
        eventDate.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]));
      }
    } else {
      // Default to 11:59 PM if no time specified
      eventDate.setHours(23, 59);
    }
    
    var eventTitle = (course ? '[' + course + '] ' : '') + assignmentName;
    var eventDescription = 'Assignment: ' + assignmentName + '\n';
    if (course) eventDescription += 'Course: ' + course + '\n';
    if (status) eventDescription += 'Status: ' + status + '\n';
    if (notes) eventDescription += 'Notes: ' + notes + '\n';
    if (isQuiz === true) eventDescription += 'Type: Quiz\n';
    if (isMidterm === true) eventDescription += 'Type: Midterm\n';
    if (isFinal === true) eventDescription += 'Type: Final Exam\n';
    
    var event;
    
    // Check if event already exists
    if (eventId) {
      try {
        event = calendar.getEventById(eventId);
        if (event) {
          // Update existing event
          event.setTitle(eventTitle);
          event.setTime(eventDate, new Date(eventDate.getTime() + 60*60000)); // 1 hour duration
          event.setDescription(eventDescription);
          // Update ICS file
          var icsLinkUpdated = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
          sheet.getRange(row, 11).setValue(icsLinkUpdated);
          ui.alert('Calendar event updated successfully!');
        } else {
          // Event was deleted, create new one
          event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
          sheet.getRange(row, 7).setValue(event.getId());
          var icsLinkCreated1 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
          sheet.getRange(row, 11).setValue(icsLinkCreated1);
          ui.alert('Calendar event created successfully!');
        }
      } catch (e) {
        // Event ID is invalid, create new one
        event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
        sheet.getRange(row, 7).setValue(event.getId());
        var icsLinkCreated2 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
        sheet.getRange(row, 11).setValue(icsLinkCreated2);
        ui.alert('Calendar event created successfully!');
      }
    } else {
      // Create new event
      event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
      sheet.getRange(row, 7).setValue(event.getId());
      var icsLinkCreated3 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
      sheet.getRange(row, 11).setValue(icsLinkCreated3);
      ui.alert('Calendar event created successfully!');
    }
    
  } catch (e) {
    ui.alert('Error creating calendar event: ' + e.message);
  }
}

/**
 * Helper function to create a new calendar event with reminders
 */
function createNewEvent(calendar, title, startDate, description) {
  // Create event (1 hour duration)
  var endDate = new Date(startDate.getTime() + 60*60000);
  var event = calendar.createEvent(title, startDate, endDate, {
    description: description
  });
  
  // Remove default reminders
  event.removeAllReminders();
  
  // Add custom reminders
  event.addPopupReminder(24 * 60);    // 1 day before
  event.addPopupReminder(60);         // 1 hour before
  event.addEmailReminder(24 * 60);    // 1 day before (email)
  
  return event;
}

/**
 * Syncs all assignments in the current sheet to Google Calendar
 */
function syncAllToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert('Sync All Assignments', 
                          'This will create/update calendar events for all assignments in this sheet. Continue?',
                          ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    return;
  }
  
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert('No assignments found in this sheet.');
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var calendar = CalendarApp.getDefaultCalendar();
  var successCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var assignmentName = data[i][0];
    var course = data[i][1];
    var dueDate = data[i][2];
    var time = data[i][3];
    var status = data[i][4];
    var notes = data[i][5];
    var eventId = data[i][6];
    var isQuiz = data[i][7];
    var isMidterm = data[i][8];
    var isFinal = data[i][9];
    
    // Skip empty rows
    if (!assignmentName || !dueDate) {
      continue;
    }
    
    try {
      var eventDate = new Date(dueDate);
      
      // Parse time if specified
      if (time) {
        var timeMatch = String(time).match(/(\d+):(\d+)/);
        if (timeMatch) {
          eventDate.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]));
        }
      } else {
        eventDate.setHours(23, 59);
      }
      
      var eventTitle = (course ? '[' + course + '] ' : '') + assignmentName;
      var eventDescription = 'Assignment: ' + assignmentName + '\n';
      if (course) eventDescription += 'Course: ' + course + '\n';
      if (status) eventDescription += 'Status: ' + status + '\n';
      if (notes) eventDescription += 'Notes: ' + notes + '\n';
      if (isQuiz === true) eventDescription += 'Type: Quiz\n';
      if (isMidterm === true) eventDescription += 'Type: Midterm\n';
      if (isFinal === true) eventDescription += 'Type: Final Exam\n';
      
      var event;
      
      // Check if event exists
      if (eventId) {
        try {
          event = calendar.getEventById(eventId);
          if (event) {
            event.setTitle(eventTitle);
            event.setTime(eventDate, new Date(eventDate.getTime() + 60*60000));
            event.setDescription(eventDescription);
            var icsLinkU = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
            sheet.getRange(i + 2, 11).setValue(icsLinkU);
          } else {
            event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
            sheet.getRange(i + 2, 7).setValue(event.getId());
            var icsLinkC1 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
            sheet.getRange(i + 2, 11).setValue(icsLinkC1);
          }
        } catch (e) {
          event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
          sheet.getRange(i + 2, 7).setValue(event.getId());
          var icsLinkC2 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
          sheet.getRange(i + 2, 11).setValue(icsLinkC2);
        }
      } else {
        event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
        sheet.getRange(i + 2, 7).setValue(event.getId());
        var icsLinkC3 = createICSFileForEvent(eventTitle, eventDate, new Date(eventDate.getTime() + 60*60000), eventDescription);
        sheet.getRange(i + 2, 11).setValue(icsLinkC3);
      }
      
      successCount++;
    } catch (e) {
      errorCount++;
      Logger.log('Error syncing row ' + (i + 2) + ': ' + e.message);
    }
  }
  
  ui.alert('Sync Complete', 
           successCount + ' assignment(s) synced successfully.\n' + 
           errorCount + ' error(s) occurred.',
           ui.ButtonSet.OK);
  
  // Auto-group by week after sync
  groupByWeek();
}

function formatICSDate(dt) {
  var iso = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyyMMdd\'T\'HHmmss\'Z\'');
  return iso;
}

function escapeICS(text) {
  if (!text) return '';
  return String(text)
    .replace(/\\/g, '\\\\')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;');
}

function getOrCreateICSFolder() {
  var it = DriveApp.getFoldersByName('AssignmentTracker ICS');
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder('AssignmentTracker ICS');
}

function createICSFileForEvent(title, startDate, endDate, description) {
  var content = 'BEGIN:VCALENDAR\r\n' +
    'VERSION:2.0\r\n' +
    'PRODID:-//Assignment Tracker//EN\r\n' +
    'BEGIN:VEVENT\r\n' +
    'UID:' + Utilities.getUuid() + '\r\n' +
    'DTSTAMP:' + formatICSDate(new Date()) + '\r\n' +
    'DTSTART:' + formatICSDate(startDate) + '\r\n' +
    'DTEND:' + formatICSDate(endDate) + '\r\n' +
    'SUMMARY:' + escapeICS(title) + '\r\n' +
    'DESCRIPTION:' + escapeICS(description) + '\r\n' +
    'END:VEVENT\r\n' +
    'END:VCALENDAR';
  var fileName = title + ' - ' + Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.ics';
  var blob = Utilities.newBlob(content, 'text/calendar', fileName);
  var folder = getOrCreateICSFolder();
  var file = folder.createFile(blob);
  return file.getUrl();
}

function setupClassesAndRecurringAssignments() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var classesResp = ui.prompt('Setup Classes', 'Enter class names separated by commas:', ui.ButtonSet.OK_CANCEL);
  if (classesResp.getSelectedButton() !== ui.Button.OK) return;
  var classes = classesResp.getResponseText().split(',').map(function(c){ return c.trim(); }).filter(function(c){ return c.length; });
  if (classes.length === 0) { ui.alert('No classes provided.'); return; }

  var weeksResp = ui.prompt('Recurring Duration', 'Enter number of weeks to generate (e.g., 10):', ui.ButtonSet.OK_CANCEL);
  if (weeksResp.getSelectedButton() !== ui.Button.OK) return;
  var weeks = parseInt(weeksResp.getResponseText(), 10);
  if (!weeks || weeks <= 0) { ui.alert('Invalid number of weeks.'); return; }

  var startDate = new Date();
  var startRow = sheet.getLastRow() + 1;

  for (var ci = 0; ci < classes.length; ci++) {
    var courseName = classes[ci];
    var freqResp = ui.prompt('Recurring for ' + courseName, "Enter frequency: 'none', 'weekly', or 'biweekly':", ui.ButtonSet.OK_CANCEL);
    if (freqResp.getSelectedButton() !== ui.Button.OK) continue;
    var freq = String(freqResp.getResponseText()).toLowerCase();
    if (freq !== 'weekly' && freq !== 'biweekly') continue;

    var baseResp = ui.prompt('Assignment Base Name', 'For ' + courseName + ', enter base name (e.g., Homework):', ui.ButtonSet.OK_CANCEL);
    if (baseResp.getSelectedButton() !== ui.Button.OK) continue;
    var baseName = baseResp.getResponseText().trim();
    if (!baseName) continue;

    var dayResp = ui.prompt('Due Weekday', 'For ' + courseName + ', enter weekday (Mon/Tue/Wed/Thu/Fri/Sat/Sun):', ui.ButtonSet.OK_CANCEL);
    if (dayResp.getSelectedButton() !== ui.Button.OK) continue;
    var dayStr = dayResp.getResponseText().trim().toLowerCase();
    var dayMap = {sun:0, mon:1, tue:2, wed:3, thu:4, fri:5, sat:6};
    var dayIdx = dayMap[dayStr.substr(0,3)];
    if (dayIdx === undefined) continue;

    var timeResp = ui.prompt('Due Time', 'For ' + courseName + ', enter time HH:MM (24h):', ui.ButtonSet.OK_CANCEL);
    if (timeResp.getSelectedButton() !== ui.Button.OK) continue;
    var timeStr = timeResp.getResponseText().trim();
    var tm = timeStr.match(/^(\d{1,2}):(\d{2})$/);
    if (!tm) continue;
    var hour = parseInt(tm[1], 10);
    var min = parseInt(tm[2], 10);

    var occurrences = weeks * (freq === 'weekly' ? 1 : 0.5);
    occurrences = Math.max(1, Math.floor(occurrences));

    var cur = new Date(startDate);
    var delta = (dayIdx - cur.getDay() + 7) % 7;
    cur.setDate(cur.getDate() + delta);

    for (var k = 0; k < occurrences; k++) {
      var due = new Date(cur);
      due.setHours(hour, min, 0, 0);
      var name = baseName + ' ' + (k + 1);
      var rowValues = [
        name,
        courseName,
        due,
        Utilities.formatDate(due, Session.getScriptTimeZone(), 'HH:mm'),
        'Not Started',
        'Auto-generated',
        '',
        false,
        false,
        false,
        ''
      ];
      sheet.getRange(startRow, 1, 1, rowValues.length).setValues([rowValues]);
      startRow++;
      cur.setDate(cur.getDate() + (freq === 'weekly' ? 7 : 14));
    }
  }
  ui.alert('Recurring assignments generated.');
  
  // Auto-group by week
  groupByWeek();
}

/**
 * Quick function to create a sample quarter sheet (for testing)
 */
function createSampleQuarter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quarterName = 'Q1 2026';
  
  // Delete if exists
  var existingSheet = ss.getSheetByName(quarterName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  var sheet = ss.insertSheet(quarterName);
  setupSheetHeaders(sheet);
  var sampleClasses = ['MATH 101', 'ENG 201', 'CHEM 105', 'HIST 150'];
  formatSheet(sheet, sampleClasses);
  setSheetClasses(sheet, sampleClasses);
  
  // Add sample data
  var sampleData = [
    ['Math Problem Set 5', 'MATH 101', new Date('2026-01-25'), '23:59', 'Not Started', 'Chapter 5 exercises', '', false, false, false, ''],
    ['Literature Essay', 'ENG 201', new Date('2026-01-28'), '17:00', 'In Progress', 'Compare two novels', '', false, false, false, ''],
    ['Lab Report', 'CHEM 105', new Date('2026-02-01'), '14:00', 'Not Started', 'Experiment 3 writeup', '', false, false, false, ''],
    ['History Reading', 'HIST 150', new Date('2026-02-05'), '12:00', 'Completed', 'Chapters 8-10', '', false, false, false, '']
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Auto-group by week
  groupByWeek();
}
