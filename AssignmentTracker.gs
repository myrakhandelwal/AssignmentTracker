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
  
  // Format the sheet
  formatSheet(sheet);
  
  ui.alert('Success', 'Sheet "' + quarterName + '" has been created!', ui.ButtonSet.OK);
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
    'Priority',
    'Status',
    'Notes',
    'Calendar Event ID'
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
function formatSheet(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 200); // Assignment Name
  sheet.setColumnWidth(2, 150); // Course
  sheet.setColumnWidth(3, 120); // Due Date
  sheet.setColumnWidth(4, 100); // Time
  sheet.setColumnWidth(5, 100); // Priority
  sheet.setColumnWidth(6, 120); // Status
  sheet.setColumnWidth(7, 250); // Notes
  sheet.setColumnWidth(8, 150); // Calendar Event ID (hidden)
  
  // Add data validation for Priority column (E)
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(priorityRule);
  
  // Add data validation for Status column (F)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'In Progress', 'Completed', 'Submitted'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(statusRule);
  
  // Add conditional formatting for Priority
  var highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('High')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('E2:E1000')])
    .build();
  
  var mediumPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('E2:E1000')])
    .build();
  
  var lowPriorityRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Low')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('E2:E1000')])
    .build();
  
  // Add conditional formatting for Status
  var completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#b7e1cd')
    .setRanges([sheet.getRange('F2:F1000')])
    .build();
  
  var submittedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Submitted')
    .setBackground('#a4c2f4')
    .setRanges([sheet.getRange('F2:F1000')])
    .build();
  
  var rules = sheet.getConditionalFormatRules();
  rules.push(highPriorityRule, mediumPriorityRule, lowPriorityRule, completedRule, submittedRule);
  sheet.setConditionalFormatRules(rules);
  
  // Hide the Calendar Event ID column
  sheet.hideColumns(8);
  
  // Add alternating row colors
  sheet.getRange('A2:H1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
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
  
  var data = sheet.getRange(row, 1, 1, 8).getValues()[0];
  var assignmentName = data[0];
  var course = data[1];
  var dueDate = data[2];
  var time = data[3];
  var priority = data[4];
  var status = data[5];
  var notes = data[6];
  var eventId = data[7];
  
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
    if (priority) eventDescription += 'Priority: ' + priority + '\n';
    if (status) eventDescription += 'Status: ' + status + '\n';
    if (notes) eventDescription += 'Notes: ' + notes + '\n';
    
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
          ui.alert('Calendar event updated successfully!');
        } else {
          // Event was deleted, create new one
          event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
          sheet.getRange(row, 8).setValue(event.getId());
          ui.alert('Calendar event created successfully!');
        }
      } catch (e) {
        // Event ID is invalid, create new one
        event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
        sheet.getRange(row, 8).setValue(event.getId());
        ui.alert('Calendar event created successfully!');
      }
    } else {
      // Create new event
      event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
      sheet.getRange(row, 8).setValue(event.getId());
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
  
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var calendar = CalendarApp.getDefaultCalendar();
  var successCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var assignmentName = data[i][0];
    var course = data[i][1];
    var dueDate = data[i][2];
    var time = data[i][3];
    var priority = data[i][4];
    var status = data[i][5];
    var notes = data[i][6];
    var eventId = data[i][7];
    
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
      if (priority) eventDescription += 'Priority: ' + priority + '\n';
      if (status) eventDescription += 'Status: ' + status + '\n';
      if (notes) eventDescription += 'Notes: ' + notes + '\n';
      
      var event;
      
      // Check if event exists
      if (eventId) {
        try {
          event = calendar.getEventById(eventId);
          if (event) {
            event.setTitle(eventTitle);
            event.setTime(eventDate, new Date(eventDate.getTime() + 60*60000));
            event.setDescription(eventDescription);
          } else {
            event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
            sheet.getRange(i + 2, 8).setValue(event.getId());
          }
        } catch (e) {
          event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
          sheet.getRange(i + 2, 8).setValue(event.getId());
        }
      } else {
        event = createNewEvent(calendar, eventTitle, eventDate, eventDescription);
        sheet.getRange(i + 2, 8).setValue(event.getId());
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
  formatSheet(sheet);
  
  // Add sample data
  var sampleData = [
    ['Math Problem Set 5', 'MATH 101', new Date('2026-01-25'), '23:59', 'High', 'Not Started', 'Chapter 5 exercises', ''],
    ['Literature Essay', 'ENG 201', new Date('2026-01-28'), '17:00', 'High', 'In Progress', 'Compare two novels', ''],
    ['Lab Report', 'CHEM 105', new Date('2026-02-01'), '14:00', 'Medium', 'Not Started', 'Experiment 3 writeup', ''],
    ['History Reading', 'HIST 150', new Date('2026-02-05'), '12:00', 'Low', 'Completed', 'Chapters 8-10', '']
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}
