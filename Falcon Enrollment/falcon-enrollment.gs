/** Falcon Enrollment - Web App v3.1 **/
/** Falcon EDU © 2023-2025 All Rights Reserved **/
/** Created by: Nick Zagorin **/

//////////////////////
// GLOBAL CONSTANTS //
//////////////////////

const ACTIVE_DATA_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Active Data');
const ARCHIVE_DATA_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive Data');
const CONSOLE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Console');

///////////////////////////
// PAGE RENDER FUNCTIONS //
///////////////////////////

/** Render the web app in the browser **/
function doGet(e) {
  const userSettings = getUserProperties();
  const page = e.parameter.page || "dashboard";
  const htmlTemplate = HtmlService.createTemplateFromFile(page);

  // Inject the user properties into the HTML
  htmlTemplate.userSettings = JSON.stringify(userSettings);

  // Evaluate and prepare the HTML content
  const htmlContent = htmlTemplate.evaluate().getContent();
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent);

  //Replace {{NAVBAR}} in HTML with the navigation bar content
  htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}",getNavbar(page)));
  
  // Set the tab favicon
  htmlOutput.setFaviconUrl("https://meesterzee.github.io/FalconEDU/images/Falcon%20EDU%20Favicon%2032x32.png");
  
  // Set the tab title
  htmlOutput.setTitle("Falcon Enrollment");
  
  return htmlOutput;
}

/** Create navigation/menu bar **/
function getNavbar(activePage) {
  const dashboardURL = getScriptURL();
  const scheduleURL = getScriptURL("page=schedule");
  const settingsURL = getScriptURL("page=settings");
  const enrollmentYear = CONSOLE_SHEET.getRange('B3').getDisplayValue();
  const headerText = "Falcon Enrollment - " + enrollmentYear;

  let navbar = 
    `<div class="menu-bar">
      <button class="menu-button" onclick="showNav()">
        <div id="menu-icon">
          <span></span>
          <span></span>
          <span></span>
          <span></span>
        </div>
      </button>
      <h1 id="header-text">` + headerText + `</h1>
    </div>
    <div class="nav-bar" id="nav-bar-links">
      <a href="${dashboardURL}" class="nav-link ${activePage === 'dashboard' ? 'active' : ''}">
        <i class="bi bi-person-circle"></i>Dashboard
      </a>
      <a href="${scheduleURL}" class="nav-link ${activePage === 'schedule' ? 'active' : ''}">
        <i class="bi bi-calendar"></i>Schedule
      </a>
      <a href="${settingsURL}" class="nav-link ${activePage === 'settings' ? 'active' : ''}">
        <i class="bi bi-gear-wide-connected"></i>Settings
      </a>
      <button class="nav-button" onclick="showAbout()">
        <i class="bi bi-info-circle"></i>About
      </button>
    </div>
    <div class="javascript-code">
    <script>
      function showNav() {
        const icon = document.getElementById('menu-icon');
        const navbar = document.querySelector('.nav-bar');
        icon.classList.toggle('open');
        navbar.classList.toggle('show');
      }

      function showAbout() {
        const title = "<i class='bi bi-info-circle'></i>About Falcon Enrollment";
        const message = "Web App Version: 3.1<br>Build: 25100124 <br><br>Created by: Nick Zagorin<br>© 2023-2025 - All rights reserved";
        showModal(title, message, "Close");
      }
    </script>
    </div>`;

  return navbar;
}

/** Get URL of the Google Apps Script web app **/
function getScriptURL(qs = null) {
  let url = ScriptApp.getService().getUrl();
  if(qs){
    if (qs.indexOf("?") === -1) {
      qs = "?" + qs;
    }
    url = url + qs;
  }

  return url;
}

/** Include additional files in HTML **/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/////////////////////////
// DASHBOARD FUNCTIONS //
/////////////////////////

/** Get active data */
function getActiveData() {
  const lastRow = ACTIVE_DATA_SHEET.getLastRow();
  
  // Return empty array if there are no data rows
  if (lastRow <= 1) {
    return [];
  }

  const dataRange = ACTIVE_DATA_SHEET.getRange(2, 1, lastRow - 1, ACTIVE_DATA_SHEET.getLastColumn());
  const data = dataRange.getValues();
  
  const headers = ACTIVE_DATA_SHEET.getRange(1, 1, 1, ACTIVE_DATA_SHEET.getLastColumn()).getValues()[0];
  let objects = [];
  
  for (let i = 0; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = ACTIVE_DATA_SHEET.getRange(i + 2, j + 1).getDisplayValue();
    }
    objects.push(obj);
  }

  return objects;
}

/** Get archive data */
function getArchiveData() {
  const lastRow = ARCHIVE_DATA_SHEET.getLastRow();
  
  // Return empty array if there are no data rows
  if (lastRow <= 1) {
    return [];
  }

  const dataRange = ARCHIVE_DATA_SHEET.getRange(2, 1, lastRow - 1, ARCHIVE_DATA_SHEET.getLastColumn());
  const data = dataRange.getValues();
  
  const headers = ARCHIVE_DATA_SHEET.getRange(1, 1, 1, ARCHIVE_DATA_SHEET.getLastColumn()).getValues()[0];
  let objects = [];
  
  for (let i = 0; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = ARCHIVE_DATA_SHEET.getRange(i + 2, j + 1).getDisplayValue();
    }
    objects.push(obj);
  }

  return objects;
}

/** Save student data */
function saveStudentData(dataSet, studentData) {
  let sheet;

  if (dataSet === "active") {
    sheet = ACTIVE_DATA_SHEET;
  } else {
    sheet = ARCHIVE_DATA_SHEET;
  }
  
  const sheetLastRow = sheet.getLastRow();
  const studentName = studentData[0][0];

  if (sheetLastRow > 1) {
    let dataRange = sheet.getRange(2, 1, sheetLastRow - 1, 1).getDisplayValues();
    let foundIndex = -1;
    let duplicate = false;

    // First pass: check for duplicates
    for (let i = 0; i < dataRange.length; i++) {
      if (dataRange[i][0] === studentName) {
        if (foundIndex === -1) {
          foundIndex = i;
        } else {
          duplicate = true;
          break;
        }
      }
    }

    // If duplicates found, return error
    if (duplicate) {
      return "duplicateDatabaseEntry";
    }
    
    // If no duplicates found but student is found, update the data
    if (foundIndex !== -1) {
      const range = sheet.getRange(foundIndex + 2, 1, 1, studentData[0].length); // Assuming studentData[0].length is the number of columns
      range.setValues(studentData);
      return "saveChangesSuccess";
    }
  }

  // If the student is not found, return error
  return "missingDatabaseEntry";
}

/** Add student data */
function addStudentData(studentData) {
  const activeLastRow = ACTIVE_DATA_SHEET.getLastRow();
  const studentName = studentData[0];

  // Check for duplicate student
  if (activeLastRow > 1) {
    const activeRange = ACTIVE_DATA_SHEET.getRange(2, 1, activeLastRow - 1, 1).getDisplayValues();
    const duplicate = activeRange.some(row => row[0] === studentName);
    if (duplicate) {
      return false;
    }
  }

  // If duplicate student not found, proceed with adding the student
  ACTIVE_DATA_SHEET.appendRow(studentData); 
  ACTIVE_DATA_SHEET.getRange(2, 1, ACTIVE_DATA_SHEET.getLastRow() - 1, ACTIVE_DATA_SHEET.getLastColumn()).sort({ column: 1, ascending: true });
  return true;
}

/** Remove student from active data and add to archive data */
function removeStudentData(student) {
  const activeLastRow = ACTIVE_DATA_SHEET.getLastRow();
  const archiveLastRow = ARCHIVE_DATA_SHEET.getLastRow();

  // Check for duplicate student in the archive
  if (archiveLastRow > 1) {
    const archiveDataRange = ARCHIVE_DATA_SHEET.getRange(2, 1, archiveLastRow - 1, 1).getDisplayValues();
    const duplicate = archiveDataRange.some(row => row[0] === student);
    if (duplicate) {
      return "duplicateDatabaseEntry";
    }
  }

  // Check if active data is empty
  if (activeLastRow <= 1) {
    return "missingDatabaseEntry";
  }
  
  // Get all data from the active sheet
  const activeDataRange = ACTIVE_DATA_SHEET.getRange(2, 1, activeLastRow - 1, ACTIVE_DATA_SHEET.getLastColumn()).getDisplayValues();

  let studentData = null;
  let studentRowIndex = -1;

  // Find the student in the active data
  for (let i = 0; i < activeDataRange.length; i++) {
    if (activeDataRange[i][0] === student) {
      studentData = activeDataRange[i];
      studentRowIndex = i + 2;  // Account for header row
      break;
    }
  }
  
  if (studentData) {
    // Remove the student from the active sheet
    ACTIVE_DATA_SHEET.deleteRow(studentRowIndex);

    // Append the student data to the archive sheet
    ARCHIVE_DATA_SHEET.appendRow(studentData);
    
    // Re-check archive last row since it has been updated
    const newArchiveLastRow = ARCHIVE_DATA_SHEET.getLastRow();
    
    // Sort the archive sheet if there are more than one rows of data
    if (newArchiveLastRow > 2) {
      const archiveDataToSort = ARCHIVE_DATA_SHEET.getRange(2, 1, newArchiveLastRow - 1, ARCHIVE_DATA_SHEET.getLastColumn()).getDisplayValues();
      archiveDataToSort.sort((a, b) => a[0].localeCompare(b[0]));
      ARCHIVE_DATA_SHEET.getRange(2, 1, newArchiveLastRow - 1, ARCHIVE_DATA_SHEET.getLastColumn()).setValues(archiveDataToSort);
    }
    
    return "archiveSuccess"; // Exit the loop once the student is found and processed
  } else {
    // If the student is not found, return error
    return "missingDatabaseEntry";
  }
}

/** Rename student in data */
function renameStudent(dataSet, oldStudentName, newStudentName) {
  let sheet;

  if (dataSet === "active") {
    sheet = ACTIVE_DATA_SHEET;
  } else {
    sheet = ARCHIVE_DATA_SHEET;
  }
  
  // Check for duplicate student
  const sheetLastRow = sheet.getLastRow();
  let dataRange;

  if (sheetLastRow > 1) {
    dataRange = sheet.getRange(2, 1, sheetLastRow - 1, 1).getDisplayValues();
    const duplicate = dataRange.some(row => row[0] === newStudentName);
    if (duplicate) {
      return "duplicateDatabaseEntry";
    }
  } else {
    return "missingDatabaseEntry";
  }
  
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0] === oldStudentName) {
      sheet.getRange(i + 2, 1).setValue(newStudentName);
      sheet.getRange(2, 1, sheetLastRow - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
      return "renameSuccess";
    }
  }

  // If the student is not found, return error
  return "missingDatabaseEntry";
}

/** Restore student from archive data and add to active data */
function restoreStudentData(student) {
  const activeLastRow = ACTIVE_DATA_SHEET.getLastRow();
  const archiveLastRow = ARCHIVE_DATA_SHEET.getLastRow();

  // Check for duplicate in active data if more than one row
  if (activeLastRow > 1) {
    activeDataRange = ACTIVE_DATA_SHEET.getRange(2, 1, activeLastRow - 1, 1).getDisplayValues();
    const duplicate = activeDataRange.some(row => row[0] === student);
    if (duplicate) {
      return "duplicateDatabaseEntry"; // Return error response if duplicate found
    }
  }

  // Check if archive data is empty
  if (archiveLastRow <= 1) {
    return "missingDatabaseEntry";
  }

  // Get all data from the archive sheet
  const archiveDataRange = ARCHIVE_DATA_SHEET.getRange(2, 1, archiveLastRow - 1, ARCHIVE_DATA_SHEET.getLastColumn()).getDisplayValues();

  let studentData = null;
  let studentRowIndex = -1;

  for (let i = 0; i < archiveDataRange.length; i++) {
    if (archiveDataRange[i][0] === student) {
      studentData = archiveDataRange[i];
      studentRowIndex = i + 2;
      break;
    }
  }

  if (studentData) {
    // Remove the student from the archive sheet
    ARCHIVE_DATA_SHEET.deleteRow(studentRowIndex);

    // Append the student data to the active sheet
    ACTIVE_DATA_SHEET.appendRow(studentData);
  
    // Re-check active last row since it has been updated
    const newActiveLastRow = ACTIVE_DATA_SHEET.getLastRow();
    
    // Sort the archive sheet if there are more than one rows of data
    if (newActiveLastRow > 2) {
      const activeDataToSort = ACTIVE_DATA_SHEET.getRange(2, 1, newActiveLastRow - 1, ACTIVE_DATA_SHEET.getLastColumn()).getDisplayValues();
      activeDataToSort.sort((a, b) => a[0].localeCompare(b[0]));
      ACTIVE_DATA_SHEET.getRange(2, 1, newActiveLastRow - 1, ACTIVE_DATA_SHEET.getLastColumn()).setValues(activeDataToSort);
    }
    
    return "restoreSuccess"; // Exit the loop once the student is found and processed
  } else {
    // If the student is not found, return error
    return "missingDatabaseEntry";
  }
}

/** Delete student from archive data */
function deleteStudentData(student) {
  const archiveLastRow = ARCHIVE_DATA_SHEET.getLastRow();
  let dataRange;

  // Check for duplicate in active data if more than one row
  if (archiveLastRow > 1) {
    dataRange = ARCHIVE_DATA_SHEET.getRange(2, 1, archiveLastRow - 1, 1).getDisplayValues();
    const studentPresent = dataRange.some(row => row[0] === student);
    if (!studentPresent) {
      return "missingDatabaseEntry"; // Return error response if student is not found
    }
  } else {
    return "missingDatabaseEntry";
  }
  
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0] === student) {
      ARCHIVE_DATA_SHEET.deleteRow(i + 2);
      return true; // Exit the loop once the row is deleted
    }
  }

  // If the student is not found, return false to display error
  return "missingDatabaseEntry";
}

/////////////////////
// EMAIL FUNCTIONS //
/////////////////////

/** Create and send email */
function createEmail(recipient, subject, body, attachments) {
  const emailQuota = MailApp.getRemainingDailyQuota();

  // Check user's email quota and warn if it's too low to send emails
  if (emailQuota <= 10) {
    return "emailQuotaLimit";
  }

  const currentUserEmail = Session.getActiveUser().getEmail();
  const emailSettings = CONSOLE_SHEET.getRange('A12:B12').getDisplayValues().flat();
  
  const emailMessage = {
    to: recipient,
    bcc: currentUserEmail,
    replyTo: emailSettings[1],
    subject: subject,
    htmlBody: body,
    name: emailSettings[0],
    attachments: []
  };

  // Add attachments if provided
  if (attachments) {
    const uint8Array = new Uint8Array(attachments);
    const blob = Utilities.newBlob(uint8Array, 'application/pdf', 'First Lutheran School - Enrollment Packet.pdf');
    emailMessage.attachments.push(blob);
  }

  // Send the email
  try {
    MailApp.sendEmail(emailMessage);
    return "emailSuccess";
  }
  catch (e) {
    return "emailFailure";
  }
}

////////////////////////
// SCHEDULE FUNCTIONS //
////////////////////////

function getAllDates() {
  const evaluationDates = getEvaluationDates();
  const screeningDates = getScreeningDates();
  const submissionDates = getSubmissionDates();
  const acceptanceDates = getAcceptanceDates();

  return {
    evaluationDates: evaluationDates,
    screeningDates: screeningDates,
    submissionDates: submissionDates,
    acceptanceDates: acceptanceDates
  };
}

/** Get student evaluation dates **/
function getEvaluationDates() {
  const data = ACTIVE_DATA_SHEET.getDataRange().getDisplayValues();
  let evaluationDates = [];

  // Loop through the data
  for (let i = 1; i < data.length; i++) { // Assuming the first row contains headers
    let studentName = data[i][0];
    let dateValue = data[i][13];

    // Check the document status
    let documentStatus = data[i][15]; // Extract the relevant column
    if (documentStatus === "") {
      documentStatus = "Status missing";
    }

    // Skip processing if dateValue is empty
    if (!dateValue) continue;
    
    evaluationDates.push({ student: studentName, date: dateValue, status: documentStatus});
  }

  // Sort the evaluationDates array by the 'date' property
  evaluationDates.sort(function (a, b) {
    let dateA = new Date(a.date);
    let dateB = new Date(b.date);
    return dateB - dateA;
  });
  return evaluationDates;
}

/** Get student screening dates **/
function getScreeningDates() {
  const data = ACTIVE_DATA_SHEET.getDataRange().getDisplayValues();
  let screeningDates = [];

  // Loop through the data
  for (let i = 1; i < data.length; i++) { // Assuming the first row contains headers
    let studentName = data[i][0]; // Set to correct column number
    let dateValue = data[i][17]; // Set to correct column number
    let screeningTime = data[i][18]; // Set to correct column number

    // Skip processing if dateValue is empty
    if (!dateValue) continue;

    if (!screeningTime) {
      screeningTime = "false";
    }

    screeningDates.push({
      student: studentName, 
      date: dateValue, 
      time: screeningTime
    });
  }

  // Sort the screeningDates array by the 'date' property
  screeningDates.sort(function (a, b) {
    let dateA = new Date(a.date);
    let dateB = new Date(b.date);
    return dateB - dateA;
  });
  return screeningDates;
}

/** Get admin submission dates **/
function getSubmissionDates() {
  const data = ACTIVE_DATA_SHEET.getDataRange().getDisplayValues();
  let submissionDates = [];

  // Loop through the data
  for (let i = 1; i < data.length; i++) { // Assuming the first row contains headers
    let studentName = data[i][0]; // Set to correct column number
    let dateValue = data[i][23]; // Set to correct column number

    // Skip processing if dateValue is empty
    if (!dateValue) continue;

    submissionDates.push({
      student: studentName, 
      date: dateValue, 
    });
  }

  // Sort the screeningDates array by the 'date' property
  submissionDates.sort(function (a, b) {
    let dateA = new Date(a.date);
    let dateB = new Date(b.date);
    return dateB - dateA;
  });

  return submissionDates;
}

/** Get student acceptance dates **/
function getAcceptanceDates() {
  const data = ACTIVE_DATA_SHEET.getDataRange().getDisplayValues();
  let acceptanceDates = [];

  // Loop through the data
  for (let i = 1; i < data.length; i++) { // Assuming the first row contains headers
    let studentName = data[i][0]; // Set to correct column number
    let dateValue = data[i][25]; // Set to correct column number
    let documentStatus = data[i].slice(28, 36); // Extract the relevant columns

    // Skip processing if dateValue or all documents are empty
    if (!dateValue) continue;
    
    // Inside the loop before pushing to acceptanceDates array
    let modifiedStatus = documentStatus.map(function(status) {
      return status === "" ? "Status missing" : status;
    });

    acceptanceDates.push({ 
      student: studentName, 
      date: dateValue, 
      blackbaudAccount: modifiedStatus[0], 
      admissionContract: modifiedStatus[1], 
      tuitionPayment: modifiedStatus[2], 
      medicalConsent: modifiedStatus[3], 
      emergencyContacts: modifiedStatus[4], 
      registrationFee: modifiedStatus[5]
    });
  }

  // Sort the acceptanceDates array by the 'date' property
  acceptanceDates.sort(function (a, b) {
    let dateA = new Date(a.date);
    let dateB = new Date(b.date);
    return dateB - dateA;
  });
  return acceptanceDates;
}

////////////////////////
// SETTINGS FUNCTIONS //
////////////////////////

/** Get user settings from user properties service **/
function getUserProperties() {
  const userProperties = PropertiesService.getUserProperties();

  return {
    theme: userProperties.getProperty('theme') || 'falconLight',
    customThemeType: userProperties.getProperty('customThemeType'),
    customThemePrimaryColor: userProperties.getProperty('customThemePrimaryColor'),
    customThemeAccentColor: userProperties.getProperty('customThemeAccentColor'),
    alertSound: userProperties.getProperty('alertSound') || 'alert01',
    emailSound: userProperties.getProperty('emailSound') || 'email01',
    removeSound: userProperties.getProperty('removeSound') || 'remove01',
    successSound: userProperties.getProperty('successSound') || 'success01',
    silentMode: userProperties.getProperty('silentMode') || 'false'
  };
}

/** Get settings from 'Console' sheet **/
function getSettings() {
  const schoolSettings = CONSOLE_SHEET.getRange('A3:B3').getDisplayValues().flat();
  const managerSettings = CONSOLE_SHEET.getRange('A6:E6').getDisplayValues().flat();
  const feeSettings = CONSOLE_SHEET.getRange('A9:H9').getDisplayValues().flat();
  const emailTemplateSettings = CONSOLE_SHEET.getRange('C12:J13').getDisplayValues().flat();
  const allSettings = {
    'School Name': schoolSettings[0],
    'School Year': schoolSettings[1],
    'Enrollment Manager 1': managerSettings[0],
    'Enrollment Manager 2': managerSettings[1],
    'Enrollment Manager 3': managerSettings[2],
    'Enrollment Manager 4': managerSettings[3],
    'Enrollment Manager 5': managerSettings[4],
    'Developmental Screening Fee (EEC)': feeSettings[0],
    'Developmental Screening Fee (TK/K)': feeSettings[1],
    'Academic Screening Fee': feeSettings[2],
    'Registration Fee': feeSettings[3],
    'HUG Fee': feeSettings[4],
    'Family Commitment Fee': feeSettings[5],
    'FLASH Processing Fee': feeSettings[6],
    'Withdrawal Processing Fee': feeSettings[7],
    'Waitlist Subject': emailTemplateSettings[0],
    'Waitlist Body': emailTemplateSettings[8],
    'Evaluation Subject': emailTemplateSettings[1],
    'Evaluation Body': emailTemplateSettings[9],
    'Screening (EEC) Subject': emailTemplateSettings[2],
    'Screening (EEC) Body': emailTemplateSettings[10],
    'Screening (School) Subject': emailTemplateSettings[3],
    'Screening (School) Body': emailTemplateSettings[11],
    'Acceptance Subject': emailTemplateSettings[4],
    'Acceptance Body': emailTemplateSettings[12],
    'Acceptance (Conditional) Subject': emailTemplateSettings[5],
    'Acceptance (Conditional) Body': emailTemplateSettings[13],
    'Rejection Subject': emailTemplateSettings[6],
    'Rejection Body': emailTemplateSettings[14],
    'Completion Subject': emailTemplateSettings[7],
    'Completion Body': emailTemplateSettings[15]
  };
  
  return allSettings;
}

/** Write settings to 'Console' sheet **/
function writeSettings(userSettings, schoolSettings, managerSettings, feeSettings, emailTemplateSubject, emailTemplateBody) {
  const userProperties = PropertiesService.getUserProperties();
  const properties = {
    theme: userSettings.theme,
    customThemeType: userSettings.customThemeType,
    customThemePrimaryColor: userSettings.customThemePrimaryColor,
    customThemeAccentColor: userSettings.customThemeAccentColor,
    alertSound: userSettings.alertSound,
    emailSound: userSettings.emailSound,
    removeSound: userSettings.removeSound,
    successSound: userSettings.successSound,
    syncSound: userSettings.syncSound,
    silentMode: userSettings.silentMode
  };

  // Sets multiple user properties at once while deleting all other user properties to maintain store
  userProperties.setProperties(properties, true); 
  
  // Write global settings to the 'Console' sheet
  CONSOLE_SHEET.getRange('A3:B3').setValues(schoolSettings);
  CONSOLE_SHEET.getRange('A6:E6').setValues(managerSettings);
  CONSOLE_SHEET.getRange('A9:H9').setValues(feeSettings);
  CONSOLE_SHEET.getRange('C12:J12').setValues(emailTemplateSubject);
  CONSOLE_SHEET.getRange('C13:J13').setValues(emailTemplateBody);
}

////////////////////
// FILE FUNCTIONS //
////////////////////

/** Export data as a .csv file **/
function getCsv(dataType) {
  let data;
  let csvContent = '';
  
  if (dataType === 'activeData') {
    data = ACTIVE_DATA_SHEET.getDataRange().getValues();
  } else {
    data = ARCHIVE_DATA_SHEET.getDataRange().getValues();
  }
    
  data.forEach(function(rowArray) {
    // Wrap fields containing commas in double quotes
    var row = rowArray.map(function(field) {
      if (typeof field === 'string' && field.includes(',')) {
        return `"${field}"`; // Wrap the field with double quotes if it contains a comma
      }
      return field;
    }).join(',');
    
    csvContent += row + '\r\n';
  });
  
  return csvContent;
}

/** Export data as a .xlsx file **/
function getXlsx(dataType) {
  const spreadsheetId = SpreadsheetApp.getActive().getId();
  let sheetId;
  
  if (dataType === 'activeData') {
    sheetId = ACTIVE_DATA_SHEET.getSheetId();
  } else {
    sheetId = ARCHIVE_DATA_SHEET.getSheetId();
  }

  // Construct the export URL
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx&gid=" + sheetId;
  
  // Fetch the xlsx file as a blob
  const blob = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getBlob();
  
  // Return blob as binary
  return blob.getBytes();
}
