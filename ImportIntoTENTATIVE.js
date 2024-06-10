/* importDataToDestination completes two processes:
1. importDataToDestination calls the registrationsData function in the "sourceScript" file of this project and an object is
returned that includes all of the students from Registrations SY 23.24.

2. importDataToDestination references TENTATIVE (2024-2025 NAHS Student Transition Notes) and cross references the student IDs from both lists.
If a student ID does not exist, it will:
  1. append the students to TENTATIVE;
  2. it then calls function importEmails which references the ContactInfo sheet and adds the student emails, guardian name and guardian emails into column BY, BZ, and CA;
  3. it calls function copySchedulesToTentative to get all of the course titles, teacher names, and case manager names from the Schedules sheet;
  4. it calls function matchAndCopyValues to get the teacher comments from the Form Responses 1 sheet into their designated spots;
  5. it sorts the list in alphabetical order by last name then first; then
  6. it adds checkboxes if they're missing in column BX.
  
Point of contact: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
Last update: 06/10/24 */

function importDataToDestination() {
  SpreadsheetApp.getActiveSpreadsheet().toast('The script is currently running. 🏃🏻‍♀️ It is updating the roster, contact info., and teacher input. Please stand by...', 'Updating Sheet', -1);
  let data = registrationsData();
  let destinationSpreadsheetId = '15ItOl4G-F5F9w-r8m3uojNZcLe4QmOpdRKrHiSUQKdA';
  let destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
  let destinationSheet = destinationSpreadsheet.getSheetByName('TENTATIVE');

  // Get existing Student IDs from the destination sheet
  let existingStudentIds = destinationSheet.getRange(2, 4, destinationSheet.getLastRow()).getValues(); // 'Student ID' is in column D (index 4)

  for (let i = 0; i < data.length; i++) {
    let rowData = data[i];

    let studentId = rowData['Student ID'];

    // Check if the Student ID exists in the destination sheet
    if (existingStudentIds.flat().indexOf(studentId) === -1) {
      // Extract specific values from rowData
      let valuesToImport = [
        new Date(),
        rowData['Student Last Name'],
        rowData['Student First Name'],
        rowData['Student ID'],
        rowData['Grade'],
        rowData['Home Campus'],
        rowData['Start Date'],
        rowData['Projected Exit'],
        rowData['Eligibility'],
        rowData['Behavior Contract']
      ];

      // Specify the target columns using targetColumnIndices
      let targetRow = [
        valuesToImport[0], valuesToImport[1], valuesToImport[2], valuesToImport[3], valuesToImport[4],
        '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        valuesToImport[5], valuesToImport[6], valuesToImport[7], '', '', '',
        valuesToImport[8], '', valuesToImport[9]
      ];

      destinationSheet.appendRow(targetRow);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('📧 Importing the contact info now. Please stand by...', 'Updating Contact Info.', -1);
  importEmails();

  SpreadsheetApp.getActiveSpreadsheet().toast('⏰ Copying the schedules over. Please stand by...', 'Updating Schedules', -1);
  copySchedulesToTentative();

  SpreadsheetApp.getActiveSpreadsheet().toast('👩🏻‍🏫 Adding the teacher comments. This is the step that takes a long time. Please stand by...', 'Updating Teacher Comments', -1);
  matchAndCopyValues();

  SpreadsheetApp.getActiveSpreadsheet().toast('😅 Almost done, sorting the sheet and adding the finishing touches. Please stand by...', 'Sorting Sheet', -1);
  // Sort the rows based on columns B and C
  var rangeToSort = destinationSheet.getRange('A2:CE' + destinationSheet.getLastRow());
  
  // Get values and trim leading/trailing spaces
  var valuesToSort = rangeToSort.getValues().map(row => row.map(cell => cell.toString().trim()));
  
  // Set the values back to the range
  rangeToSort.setValues(valuesToSort);
  
  // Sort the range
  rangeToSort.sort([{column: 2, ascending: true}, {column: 3, ascending: true}]);

  // Get the range of data after sorting
  let sortedDataRange = destinationSheet.getRange('A2:CE' + destinationSheet.getLastRow());

  // Insert checkboxes in column BX (index 75 after sorting) if missing
  let checkBoxColumn = sortedDataRange.offset(0, 77 - 2, sortedDataRange.getNumRows(), 1);
  
  // Get existing values in the checkbox column
  let existingCheckboxes = checkBoxColumn.getValues();

  // Loop through the existing checkboxes and insert if missing
  for (let i = 0; i < existingCheckboxes.length; i++) {
    if (existingCheckboxes[i][0] !== true) {
      checkBoxColumn.getCell(i + 1, 1).insertCheckboxes();
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Script finished updating! 👍🏼 You are free to work on it now.', 'Finished', 5);
  insertNoteWithTimestamp();
}

function insertNoteWithTimestamp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TENTATIVE");
  
  // Get the current timestamp
  var currentTime = new Date();

  // Get the spreadsheet's time zone
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  // Format the timestamp using the spreadsheet's time zone
  var formattedTime = Utilities.formatDate(currentTime, timeZone, "yyyy-MM-dd HH:mm:ss a");
  
  // Construct the note
  var note = "The TENTATIVE Sheet was updated on: " + formattedTime;
  
  // Set the note to cell A1, overwriting any existing note
  sheet.getRange('A1').setNote(note);
}

