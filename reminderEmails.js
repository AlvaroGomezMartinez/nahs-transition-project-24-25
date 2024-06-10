/**
The sendEmailsForToday function below references the object returned by registrationsData which is found in the sourceScript file in this project.

sendEmailsForToday looks at the tenDaysMark date for each student and if the date is today, then it sends an email to the teachers in
emailRecipients with a list of the students at the 10 Day Mark.

sendEmailsForToday has five triggers set to run it every Monday, Tuesday, Wednesday, Thursday, and Friday on Zina Gonzales' (social worker at NAHS) gmail account so the emails to the teachers are recieved from her.

Point of Contact: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
*/

function sendEmailsForToday() {
  let today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  // let today = '2023-09-25' // This is for testing and sending emails out manually. Set a date in the variable. When using this, don't forget to turn it back off.
  let dataObjects = registrationsData();
  let studentsForToday = [];

  for (let i = 0; i < dataObjects.length; i++) {
    let dataObject = dataObjects[i];
    let tenDaysMark = dataObject['10 Days Mark'];
    
    if (tenDaysMark === today) {
      let lastName = dataObject['Student Last Name'];
      let firstName = dataObject['Student First Name'];
      let studentID = dataObject['Student ID'];
      let studentGrade = dataObject['Grade'];

      studentsForToday.push(`${lastName}, ${firstName} (${studentID}), Grade: ${studentGrade}`);
    }
  }

  if (studentsForToday.length > 0) {
    let holidayDates = [
    '2023-09-04', '2023-10-09', '2023-11-07', '2023-11-20', '2023-11-21',
    '2023-11-22', '2023-11-23', '2023-11-24', '2023-12-18', '2023-12-19',
    '2023-12-20', '2023-12-21', '2023-12-22', '2023-12-25', '2023-12-26',
    '2023-12-27', '2023-12-28', '2023-12-29', '2024-01-01', '2024-01-02',
    '2024-01-15', '2024-02-19', '2024-03-11', '2024-03-12', '2024-03-13',
    '2024-03-14', '2024-03-15', '2024-03-29', '2024-04-26', '2024-05-27'
    ];
    
    // Function to check if a date is a weekend (Saturday or Sunday)
    function isWeekend(date) {
      return date.getDay() === 0 || date.getDay() === 6; // 0 represents Sunday, 6 represents Saturday
    }

    // Calculate the due date
    let dueDate = new Date(today);

    // Loop to add two workdays (excluding weekends and holidays)
    let workdaysAdded = 0;
    while (workdaysAdded < 2) {
      dueDate.setDate(dueDate.getDate()+ 1);

      // Check if the new date is not a weekend and not in the holidayDates array
      if (!isWeekend(dueDate) && !holidayDates.includes(Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'))) {
        workdaysAdded++;
      }
    }

    // Format the due date
    let formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MM-dd-yyyy');
    let subject = 'Transition Reminder: Today\'s List of Students with 10 Days at NAHS';
    let formLink = 'https://forms.gle/vwQqtcgboJidojGi7'
    let body = `NAHS Teachers,\n\nBelow is today's list of students that have been enrolled for 10 days at NAHS:\n\n${studentsForToday.join('\n')}\n\nACTION ITEM (Due by end of day, ${formattedDueDate}): If you have one of these students on your roster, please go to: ${formLink} and provide your input on their academic growth and behavioral progress.\n\nThank you`;

    // The email below is used for testing.
    // let emailRecipients = ['alvaro.gomez@nisd.net'];
    let emailRecipients = [
      'veronica.altamirano@nisd.net',
      'marco.ayala@nisd.net',
      'alita.barrera@nisd.net',
      'gabriela.chavarria-medina@nisd.net',
      'staci.cunningham@nisd.net',
      'samantha.daywood@nisd.net',
      'richard.delarosa@nisd.net',
      'ramon.duran@nisd.net',
      'janice.flores@nisd.net',
      'lauren.flores@nisd.net',
      'roslyn.francis@nisd.net',
      'nancy-1.garcia@nisd.net',
      'cierra.gibson@nisd.net',
      'zina.gonzales@nisd.net',
      'alvaro.gomez@nisd.net',
      'teressa.hensley@nisd.net',
      'catherine.huff@nisd.net',
      'erin.knippa@nisd.net',
      'joshua.lacour@nisd.net',
      'thalia.mendez@nisd.net',
      'alexandria.murphy@nisd.net',
      'dennis.olivares@nisd.net',
      'loretta.owens@nisd.net',
      'denisse.perez@nisd.net',
      'jessica.poladelcastillo@nisd.net',
      'angela.rodriguez@nisd.net',
      'linda.rodriguez@nisd.net',
      'jessica-1.vela@nisd.net',
      'miranda.wenzlaff@nisd.net']
    
    GmailApp.sendEmail(emailRecipients.join(','), subject, body);
  }
}
