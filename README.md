# Assignment Tracker - Google Apps Script

A Google Sheets-based assignment tracker that automatically creates quarterly sheets, syncs assignments with Google Calendar, and organizes tasks by week with automatic reminders and iCalendar file generation.

## Features

- Quarterly Sheet Creation - Automatically creates formatted sheets for each quarter with class management
- Google Calendar Integration - Syncs assignments to your calendar with customizable reminders
- Weekly Organization - Automatically groups assignments by week for easy planning
- Class Management - Add, edit, and remove classes via an interactive sidebar
- Quiz/Midterm/Final Tracking - Checkbox columns for exam types
- iCalendar Files - Generates downloadable .ics files for each assignment stored in Google Drive
- Recurring Assignments - Set up weekly or biweekly assignments automatically
- Smart Formatting - Color-coded status tracking with data validation
- Automatic Reminders - Email and popup notifications 1 day and 1 hour before due dates

## Quick Start for Users

### Option 1: Use the Template (Recommended)

1. Open the Assignment Tracker Template (link coming soon)
2. Click File → Make a copy
3. The script is automatically included and ready to use
4. Refresh the sheet to see the Assignment Tracker menu

### Option 2: Manual Installation

1. Create a new Google Sheet at https://sheets.google.com
2. Click Extensions → Apps Script
3. Delete any existing code
4. Copy the entire contents of `AssignmentTracker.gs` from this repository
5. Paste into the Apps Script editor
6. Click Save and name your project "Assignment Tracker"
7. Close the Apps Script tab and refresh your spreadsheet

### First Time Authorization

1. Click Assignment Tracker menu → Create New Quarter Sheet
2. If prompted, click "Review Permissions"
3. Choose your Google account
4. Click "Advanced" → "Go to Assignment Tracker (unsafe)"
5. Click "Allow"
6. The script now has permission to create calendar events and Drive files

## How to Use

### Creating a New Quarter

1. Click Assignment Tracker → Create New Quarter Sheet
2. Enter quarter name (e.g., "Q1 2026", "Winter 2026")
3. Enter number of classes you're taking
4. Enter each class name when prompted
5. A new sheet is created with proper formatting and class validation

### Sheet Columns

- Assignment Name - Name of the assignment
- Course - Class name (dropdown populated from your classes)
- Due Date - When the assignment is due
- Time - Due time in 24-hour format (e.g., 23:59, 14:00)
- Status - Not Started, In Progress, Completed, or Submitted (dropdown)
- Notes - Additional information
- Calendar Event ID - Hidden column for calendar sync
- Quiz - Checkbox for quiz designation
- Midterm - Checkbox for midterm exam
- Final Exam - Checkbox for final exam
- iCalendar File URL - Link to downloadable .ics file in Google Drive

### Adding Assignments

Fill in the rows with your assignment information. The sheet will automatically group assignments by week with bold week headers showing date ranges.

### Managing Classes

1. Click Assignment Tracker → Manage Classes
2. Use the sidebar to add or delete classes
3. Click "Save & Close" to update the Course dropdown validation

### Syncing with Calendar

**Single Assignment:**
1. Click any cell in the assignment row
2. Click Assignment Tracker → Add Assignment to Calendar
3. Calendar event and iCalendar file are created automatically

**All Assignments:**
1. Click Assignment Tracker → Sync All Assignments to Calendar
2. Confirm the action
3. All assignments sync and the sheet reorganizes by week

### Setting Up Recurring Assignments

1. Click Assignment Tracker → Setup Classes & Recurring
2. Enter class names (comma-separated)
3. Enter number of weeks to generate
4. For each class, specify:
   - Frequency (weekly or biweekly)
   - Assignment base name (e.g., "Homework")
   - Due weekday (Mon/Tue/Wed/Thu/Fri/Sat/Sun)
   - Due time (HH:MM in 24-hour format)
5. Assignments are generated and grouped by week automatically

### Calendar Reminders

Each calendar event includes:
- Email reminder - 1 day before due date
- Popup reminder - 1 day before due date
- Popup reminder - 1 hour before due date

### iCalendar Files

Each synced assignment generates an .ics file stored in a Google Drive folder named "AssignmentTracker ICS". Click the URL in the iCalendar File URL column to download or share the file with other calendar applications.

## Visual Features

### Status Color Coding
- Completed - Green background
- Submitted - Blue background

### Weekly Organization
Assignments are automatically grouped by week with bold headers showing the date range (e.g., "Week of Jan 19 - Jan 26"). Within each week, assignments are sorted by due date.

### Additional Formatting
- Frozen header row for easy scrolling
- Alternating row colors for readability
- Auto-sized columns for optimal viewing
- Hidden Calendar Event ID column

## For Developers

### Repository Structure

```
AssignmentTracker/
├── AssignmentTracker.gs    # Main script file
├── README.md               # This file
├── .github/
│   ├── PULL_REQUEST_TEMPLATE.md
│   └── workflows/
│       └── release.yml     # Automatic release workflow
```

### Making Changes

1. Clone the repository
2. Make changes to `AssignmentTracker.gs`
3. Test in a Google Sheet with Apps Script
4. Commit and push to GitHub
5. Create a release tag (e.g., `v1.0.0`) to trigger automatic packaging

### Creating a Release

```bash
git tag v1.0.0
git push origin v1.0.0
```

The GitHub Actions workflow will automatically create a release with a downloadable zip file.

### Testing

Run the `createSampleQuarter` function from the Apps Script editor to create a test sheet with sample data.

## Troubleshooting

### Menu Not Appearing
- Refresh the spreadsheet
- Close and reopen the spreadsheet
- Check that the script is properly saved in Apps Script

### Calendar Events Not Creating
- Verify you authorized calendar access
- Check that Assignment Name and Due Date are filled in
- Ensure date format is correct (use date picker)

### Week Headers Not Showing
- Click Assignment Tracker → Sync All Assignments to Calendar to regroup
- Ensure due dates are valid dates

### Class Dropdown Not Working
- Use Assignment Tracker → Manage Classes to update your class list
- Verify classes were entered during quarter creation

## Customization

### Modifying Reminder Times

Edit the `createNewEvent` function:

```javascript
event.addPopupReminder(24 * 60);    // 1 day before
event.addPopupReminder(60);         // 1 hour before
event.addEmailReminder(24 * 60);    // 1 day before (email)
```

### Changing Week Start Day

Edit the `groupByWeek` function. Currently set to Sunday (day 0). Change the calculation in the week start logic.

## Privacy & Data

- All data stays in your Google account
- Script only accesses your spreadsheet, calendar, and Drive
- iCalendar files are stored in your personal Google Drive
- No data is sent to external servers

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly in Google Sheets
5. Submit a pull request with a clear description

## License

This project is open source and available for personal and educational use.

---

Created for efficient assignment tracking with seamless calendar integration and weekly organization.
