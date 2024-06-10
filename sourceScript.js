/**
The registrationsData function is called by the importIntoTENTATIVE file in this project.

The registrationsData function references Irma Lopez's (Attendance Secretary, NAHS) tracking form (Form Responses 1;
Registrations SY 23.24).

The registrationsData function creates an object (dataObject) with the data from the students who are listed on Irma's Form Responses 1 sheet that
aren't listed in this project's Withdrawn sheet. If a student is in the Withdrawn sheet then they will not be added to dataObject.

The registrationsData function also adds a key to dataObject called '10 Days Mark' that provides the 10th day after placement at NAHS.
The '10 Days Mark' date is used to send a reminder (see the "reminderEmails") to teachers to fill out student progress.

Point of contact: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
Latest update: 12/4/23
*/

function registrationsData() {
  let externalSpreadsheetId = '1kAWRpWO4xDtRShLB5YtTtWxTbVg800fuU2RvAlYhrfA';
  let externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  let externalSpreadsheetId2 = '1MTg2DdLGRKtdb2KuVwU-vmn-L_4dIUzW7uKp_AYSVI4';
  let externalSpreadsheet2 = SpreadsheetApp.openById(externalSpreadsheetId2);

  // Retrieve data from "Form Responses 1" sheet in Irma's spreadsheet
  let sheet1 = externalSpreadsheet.getSheetByName('Form Responses 1');
  let dataRange = sheet1.getDataRange();
  let dataValues1 = dataRange.getValues();
  
  // Iterate through the dataValues1 array and capitalize the values in index 5
  let capitalizedDataValues1 = dataValues1.map((row) => {
    // Check if the value in index 5 is a string
    if (typeof row[5] === 'string') {
      // Capitalize the first letter and concatenate the rest of the string
      row[5] = row[5].charAt(0).toUpperCase() + row[5].slice(1);
    }

    return row;
  });

  // Create an object to store the latest date for each unique value in index 3
  let latestDates = {};

  // Iterate through the dataValues1 array and finds the latest start date (index 6) for each unique student ID (index 3).
  capitalizedDataValues1.forEach((row) => {
    let valueInIndex3 = row[3];
    let dateInIndex6 = new Date(row[6]);

    if (!latestDates[valueInIndex3] || dateInIndex6 > latestDates[valueInIndex3]) {
      latestDates[valueInIndex3] = dateInIndex6;
    }
  });

  // Filter the capitalizedDataValues1 array based on the latest dates for each unique value in index 3 (the ID).
  // In other words, it will filter out repeating IDs and keep the row with the most recent start date (index 6).
  let filteredDataValues1 = capitalizedDataValues1.filter((row) => {
    let valueInIndex3 = row[3];
    let dateInIndex6 = new Date(row[6]);

    return dateInIndex6.getTime() === latestDates[valueInIndex3].getTime();
  });
  
  // Retrieve data from "Students not on Registration Doc" sheet
  let sheet2 = externalSpreadsheet2.getSheetByName('Students not on Registration Doc');
  let dataRange2 = sheet2.getRange("A2:I");
  // let dataValues2 = dataRange2.getValues().filter(row => row.some(cell => cell !== '')); // Filter out empty rows
  let dataValues2 = dataRange2.getValues().filter(row => {
  // Check if any cell in the row is not empty
  if (row.some(cell => cell !== '')) {
    // If non-empty, only keep rows where the value in index 8 is empty
    return row[8] === '';
  }

  // If the row is totally empty, exclude it
  return false;
});

  // Add blank elements to each array in dataValues2 so they match up to the same number of elements as filteredDataValues1
  dataValues2.forEach(row => {
    // Add a blank element at the beginning
    row.unshift('');

    // Add 8 blank elements at the end
    for (let i = 0; i < 7; i++) {
      row.push('');
    }
  });

  // Merge data from both sheets
  let dataValues = filteredDataValues1.concat(dataValues2);

  let dataObjects = [];
  let holidayDates = [
    '2023-09-04', '2023-10-09', '2023-11-07', '2023-11-20', '2023-11-21',
    '2023-11-22', '2023-11-23', '2023-11-24', '2023-12-18', '2023-12-19',
    '2023-12-20', '2023-12-21', '2023-12-22', '2023-12-25', '2023-12-26',
    '2023-12-27', '2023-12-28', '2023-12-29', '2024-01-01', '2024-01-02',
    '2024-01-15', '2024-02-19', '2024-03-11', '2024-03-12', '2024-03-13',
    '2024-03-14', '2024-03-15', '2024-03-29', '2024-04-26', '2024-05-27'
  ];

  // Open the "Withdrawn" spreadsheet
  let withdrawnSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Withdrawn');
  let withdrawnSheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('W/D Other');
  let studentsWithSchedules = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedules');
  let withdrawnData = withdrawnSheet.getRange("D:D").getValues().flat(); // Get all values in column D of 'Withdrawn'
  let nonBlankWithdrawnData = withdrawnData.filter(value => value !== "");
  nonBlankWithdrawnData.shift();
  let withdrawnData2 = withdrawnSheet2.getRange("D:D").getValues().flat(); // Get all values in column D of 'W/D Other'
  let nonBlankWithdrawnData2 = withdrawnData2.filter(value => value !== "");
  nonBlankWithdrawnData2.shift()
  let stusWScheds = studentsWithSchedules.getRange("C:C").getValues().flat(); // Get all values in column C of 'Schedules'
  let nonBlankStusWScheds = stusWScheds.filter(value => value !=="");
  nonBlankStusWScheds.shift();
  let uniqueNonBlankStusWScheds = [... new Set(nonBlankStusWScheds)];

  let allWithdrawnData = nonBlankWithdrawnData.concat(nonBlankWithdrawnData2);
  let uniqueAllWithdrawnData = [...new Set(allWithdrawnData)];

  let superFiltered = uniqueAllWithdrawnData.filter(value => !uniqueNonBlankStusWScheds.includes(value));

  for (let i = 1; i < dataValues.length; i++) {
    let rowData = dataValues[i];
    let studentID = rowData[3]; // Assuming 'Student ID' is in the 4th column (index 3)

    // Check if studentID exists in the withdrawnData array
    if (!superFiltered.includes(studentID)) {
      // Student ID doesn't exist in both the "Withdrawn" and "W/D Other" sheet, so add it to dataObjects
      let startDate = new Date(rowData[6]); // 'Start Date' is in the 7th column (index 6)
      let placementDays = parseInt(rowData[7], 10); // Convert Placement Days to integer
      let newDate = addWorkdays(startDate, 10, holidayDates);
      let projectedExit = calculateProjectedExit(startDate, placementDays, holidayDates);
      let daysLeft = calculateDaysLeft(startDate, placementDays, holidayDates);
    
      let dataObject = {
        'Timestamp': rowData[0],
        'Student Last Name': rowData[1],
        'Student First Name': rowData[2],
        'Student ID': studentID, // Use the extracted student ID
        'Grade': rowData[4],
        'Home Campus': rowData[5],
        'Start Date': rowData[6],
        'Placement Days': rowData[7],
        'Placement Offense': rowData[8],
        'Eligibility': rowData[9],
        'Behavior Contract': rowData[10],
        'Email Address': rowData[11],
        '10 Days Mark': formatDate(newDate), // Format the new date
        'Projected Exit': formatDate(projectedExit), // Format projected exit date
        'Days Left': daysLeft
      };
    
      dataObjects.push(dataObject);
    }
  }

let studentIDs = [];

for (let i = 0; i < dataObjects.length; i++) {
  studentIDs.push(dataObjects[i]['Student ID']);
}

  // The function below is used for testing. Sort the studentIDs array numerically
  // studentIDs.sort(function(a, b) {
  //   return a - b;
  // });
  // let numberOfStudentIDs = studentIDs.length;
  // Logger.log(numberOfStudentIDs);
  // Logger.log("List of Student IDs (Numerical Order): " + studentIDs.join(", "));

  // Optionally, return the objects or do something else with them
  return dataObjects;
}

function addWorkdays(startDate, numWorkdays, holidays) {
  let currentDate = new Date(startDate);
  let workdaysAdded = 1;

  while (workdaysAdded < numWorkdays) {
    currentDate.setDate(currentDate.getDate() + 1);

    if (!isWeekend(currentDate) && !isHoliday(currentDate, holidays)) {
      workdaysAdded++;
    }
  }

  return currentDate;
}

function isWeekend(date) {
  return date.getDay() === 0 || date.getDay() === 6;
}

function isHoliday(date, holidays) {
  return holidays.includes(formatDate(date));
}

function formatDate(date) {
  let year = date.getFullYear();
  let month = (date.getMonth() + 1).toString().padStart(2, '0');
  let day = date.getDate().toString().padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

function calculateProjectedExit(startDate, placementDays, holidays) {
  let currentDate = new Date(startDate);
  let workdaysAdded = 1;

  while (workdaysAdded < placementDays) {
    currentDate.setDate(currentDate.getDate() + 1);

    if (!isWeekend(currentDate) && !isHoliday(currentDate, holidays)) {
      workdaysAdded++;
    }
  }

  return currentDate;
}

function calculateDaysLeft(startDate, placementDays, holidays) {
  let currentDate = new Date(startDate);
  let workdaysPassed = 0;

  while (currentDate <= new Date()) {
    if (!isWeekend(currentDate) && !isHoliday(currentDate, holidays)) {
      workdaysPassed++;
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }

  return placementDays - workdaysPassed;
}

