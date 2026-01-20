# Assignment Tracker - Google Apps Script

A powerful Google Sheets-based assignment tracker that automatically creates quarterly sheets and syncs assignments with Google Calendar, complete with automatic reminders.

## Features

✅ **Quarterly Sheet Creation** - Automatically creates formatted sheets for each quarter  
✅ **Google Calendar Integration** - Syncs assignments to your calendar with reminders  
✅ **Smart Formatting** - Color-coded priorities and status tracking  
✅ **Dropdown Menus** - Easy data entry with validation  
✅ **Automatic Reminders** - Email and popup notifications 1 day and 1 hour before due dates  
✅ **Conditional Formatting** - Visual feedback for priorities and completion status

## Setup Instructions

### Step 1: Create Your Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Name it "Assignments" or whatever you prefer

### Step 2: Add the Script
1. In your spreadsheet, click **Extensions** → **Apps Script**
2. Delete any code in the editor
3. Copy the entire contents of `AssignmentTracker.gs` and paste it into the Apps Script editor
4. Click the **Save** icon (💾) and name your project "Assignment Tracker"

### Step 3: Authorize the Script
1. In the Apps Script editor, select the `onOpen` function from the dropdown
2. Click **Run** (▶️)
3. You'll see an "Authorization required" dialog - click **Review Permissions**
4. Choose your Google account
5. Click **Advanced** → **Go to Assignment Tracker (unsafe)**
6. Click **Allow**

### Step 4: Return to Your Spreadsheet
1. Close the Apps Script tab
2. Refresh your spreadsheet
3. You should now see a new menu: **Assignment Tracker**

## How to Use

### Creating a New Quarter Sheet

1. Click **Assignment Tracker** → **Create New Quarter Sheet**
2. Enter the quarter name (e.g., "Q1 2026", "Winter 2026", "Spring 2026")
3. Click **OK**

The script will create a new sheet with the following columns:
- **Assignment Name** - Name of the assignment
- **Course** - Course code or name
- **Due Date** - When the assignment is due
- **Time** - Due time (e.g., 23:59, 17:00)
- **Priority** - High, Medium, or Low (dropdown)
- **Status** - Not Started, In Progress, Completed, or Submitted (dropdown)
- **Notes** - Additional information
- **Calendar Event ID** - Hidden column for calendar sync

### Adding Assignments

Simply fill in the rows with your assignment information:

| Assignment Name | Course | Due Date | Time | Priority | Status | Notes |
|----------------|--------|----------|------|----------|---------|-------|
| Math Problem Set 5 | MATH 101 | 1/25/2026 | 23:59 | High | Not Started | Chapter 5 exercises |
| Literature Essay | ENG 201 | 1/28/2026 | 17:00 | High | In Progress | Compare two novels |

### Syncing with Google Calendar

#### Option 1: Sync Individual Assignment
1. Click on any cell in the row of the assignment you want to sync
2. Click **Assignment Tracker** → **Add Assignment to Calendar**
3. A calendar event will be created with:
   - **1 day before reminder** (email + popup)
   - **1 hour before reminder** (popup)

#### Option 2: Sync All Assignments
1. Click **Assignment Tracker** → **Sync All Assignments to Calendar**
2. Confirm the action
3. All assignments in the current sheet will be synced

### Calendar Reminders

Each calendar event automatically includes:
- 📧 **Email reminder** - 1 day before the due date
- 🔔 **Popup reminder** - 1 day before the due date
- 🔔 **Popup reminder** - 1 hour before the due date

## Visual Features

### Priority Color Coding
- 🔴 **High Priority** - Red background
- 🟡 **Medium Priority** - Yellow background
- 🟢 **Low Priority** - Green background

### Status Color Coding
- ✅ **Completed** - Green background
- 📘 **Submitted** - Blue background

### Additional Formatting
- Frozen header row for easy scrolling
- Alternating row colors for readability
- Auto-sized columns for optimal viewing

## Tips & Best Practices

1. **Date Format** - Use standard date format (MM/DD/YYYY or click the cell and use the date picker)
2. **Time Format** - Use 24-hour format (e.g., 14:00 for 2 PM) or 12-hour format with AM/PM
3. **Update Calendar** - If you change assignment details, click "Add Assignment to Calendar" again to update
4. **Multiple Quarters** - Create separate sheets for each quarter to keep organized
5. **Status Updates** - Update the Status column as you progress through assignments

## Troubleshooting

### Menu Not Appearing
- Refresh the spreadsheet
- Make sure you ran the `onOpen` function and authorized the script
- Try closing and reopening the spreadsheet

### Calendar Events Not Creating
- Ensure you've authorized the script to access Google Calendar
- Check that Assignment Name and Due Date are filled in
- Verify the date format is correct

### Reminders Not Working
- Check your Google Calendar notification settings
- Make sure you're using your default Google Calendar
- Email reminders go to your Google account email

## Advanced: Customizing Reminders

To modify reminder times, edit the `createNewEvent` function in the script:

```javascript
// Current settings:
event.addPopupReminder(24 * 60);    // 1 day before
event.addPopupReminder(60);         // 1 hour before
event.addEmailReminder(24 * 60);    // 1 day before (email)

// Examples of other options:
event.addPopupReminder(2 * 24 * 60);  // 2 days before
event.addPopupReminder(30);           // 30 minutes before
event.addEmailReminder(3 * 24 * 60);  // 3 days before
```

## Sample Data

Want to test the tracker? Run the `createSampleQuarter` function from the Apps Script editor to create a sample quarter with demo assignments.

## Questions or Issues?

The script includes error handling and user-friendly alerts. If you encounter any issues:
1. Check the spreadsheet data is complete
2. Verify date/time formats are correct
3. Ensure calendar permissions are granted
4. Check the Apps Script execution log (View → Logs in Apps Script editor)

---

**Created for efficient assignment tracking with seamless calendar integration!** 📚✨
