const ss = SpreadsheetApp.getActiveSpreadsheet();

function main() {
  // Check to see if today is a weekday.
  if(isWeekday()) {
    // MBI's inactive until Shine Roll-Over
    // Create and store the Participant Level ED DMH MBI into our google sheet project.
    let sheet = ss.getSheetByName(`PL ED DMH MBI`); // Select a tab
    storeFileOnDrive(getAttachment(`Automated - PL ED DMH MBI (.xlsx)`), `PL ED DMH MBI`);
    convertExceltoGoogleSpreadsheet(`PL ED DMH MBI`); // Live Links
    writeFileToSpreadsheet(sheet, `PL ED DMH MBI`); // Store data on selected tab.
    removeMetaData(sheet, 9, 3);

    // MBI's inactive until Shine Roll-Over
    // Create and store the Center Level ED DMH MBI into our google sheet project.
    sheet = ss.getSheetByName(`CL ED DMH MBI`);
    storeFileOnDrive(getAttachment(`Automated - CL ED MBI (.xslx)`), `CL ED DMH MBI`);
    convertExceltoGoogleSpreadsheet(`CL ED DMH MBI`);
    writeFileToSpreadsheet(sheet, `CL ED DMH MBI`);
    removeMetaData(sheet, 5, 3);

    // Create and store the Education Screening (ED105) report into our google sheet project.
    sheet = ss.getSheetByName(`Education Screenings`);
    storeFileOnDrive(getAttachment(`Automated - ED105`), `Education Screenings`);
    convertExceltoGoogleSpreadsheet(`Education Screenings`)
    writeFileToSpreadsheet(sheet, `Education Screenings`, 1);
    removeMetaData(sheet, 20, 2);
    splitHyperlink(1, sheet);

    // Create and store the Home Visits (ED103) report into our google sheet project.
    sheet = ss.getSheetByName(`Home Visits`);
    storeFileOnDrive(getAttachment(`Automated - ED103 - Home Visits`), `Home Visits`);
    convertExceltoGoogleSpreadsheet(`Home Visits`);
    writeFileToSpreadsheet(sheet, `Home Visits`, 1);
    removeMetaData(sheet, 12, 2);
    splitHyperlink(1, sheet);

    // Create and store the Enrollment Report (E201) into our google sheet project.
    sheet = ss.getSheetByName(`E201`);
    storeFileOnDrive(getAttachment(`Automated - E201`), `E201`);
    convertExceltoGoogleSpreadsheet(`E201`)
    writeFileToSpreadsheet(sheet, `E201`);
    removeMetaData(sheet, 14, 2);
    editEnrollmentFields(sheet);
    createFundedEnrolled();
    editHistoricalEnrollment();

    // Create and store advocate reports (h301 and advocate mbi) into our google sheet project if today is Wednesday.
    const today = new Date().getDay();
    // Check to see if today is Monday.
    if(today === 1) {
      // MBI's inactive until Shine Roll-Over
      // Add the MBI By Advocate report to the dashboard data sheet.
      sheet = ss.getSheetByName(`MBI By Advocate`);
      deletePreviousData(sheet);
      storeFileOnDrive(getAttachment(`Automated - MBI By Advocate`), `MBI By Advocate`);
      convertExceltoGoogleSpreadsheet(`MBI By Advocate`);
      writeFileToEndOfSpreadsheet(sheet, 'MBI By Advocate', 8, 3);

      // Add the Immunization data to the dashboard data sheet.
      sheet = ss.getSheetByName(`Immunizations`);
      deletePreviousData(sheet);
      storeFileOnDrive(getAttachment(`Automated - H301 - Immunizations`), `Immunizations`);
      convertExceltoGoogleSpreadsheet(`Immunizations`);
      writeFileToEndOfSpreadsheet(sheet, 'Immunizations', 12, 2);
    }

    // Create and store the Waitlist Report () into our google sheet project.
    sheet = ss.getSheetByName(`Waitlist`);
    storeFileOnDrive(getAttachment(`Automated - E206 - Waitlist`), `Waitlist`);
    convertExceltoGoogleSpreadsheet(`Waitlist`)
    writeFileToSpreadsheet(sheet, `Waitlist`);
    removeMetaData(sheet, 9, 3);
    createLiveLinks(sheet);

    // Create and store the OFA Notes report (F101) into our google sheet project.
    sheet = ss.getSheetByName(`OFA Notes`);
    storeFileOnDrive(getAttachment(`Automated - F101 - OFA Notes`), `OFA Notes`);
    convertExceltoGoogleSpreadsheet(`OFA Notes`);
    writeFileToSpreadsheet(sheet, `OFA Notes`);
    removeMetaData(sheet, 17, 2);
    createContactColumns(sheet);

    // Create and store the Active IEP report (D104) into our google sheet project.
    sheet = ss.getSheetByName(`Active IEP`);
    storeFileOnDrive(getAttachment(`Automated - Active IEP`), `Active IEP`);
    convertExceltoGoogleSpreadsheet(`Active IEP`);
    writeFileToSpreadsheet(sheet, `Active IEP`);
    removeMetaData(sheet, 15, 3);
    editActiveIEP(sheet);

    // Create and store the Referral to LEA Notes report (D109) into our google sheet project.
    sheet = ss.getSheetByName(`Referral to LEA`);
    storeFileOnDrive(getAttachment(`Automated - Referral to LEA`), `Referral to LEA`);
    convertExceltoGoogleSpreadsheet(`Referral to LEA`);
    writeFileToSpreadsheet(sheet, `Referral to LEA`);
    removeMetaData(sheet, 13, 2);
    editReferralToLEA(sheet);
  }
}

/**
 * A function to create a user menu when the google sheet first loads.
 */
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Absence Reason Reports')
  .addItem(`Create New Report`, `getAbsenceData`)
  .addToUi();
}

/**
 *  Adds a date row in the first column to track historical changes in the google studio report.
 * @param sheet - The sheet you want to add the data to.
 * @param startRow - The row you want to begin pasting the data to.
 */
function addDateRows(sheet, startRow, length, column) {
  const todayString = getStrDate(0);
  const data = sheet.getRange(startRow, column, length).getValues();

  const dates = data.map((_) => [todayString]);
  sheet.getRange(startRow, column, dates.length).setValues(dates);
}

/**
 * Returns the calculated date as a string with the form MM-DD-YYYY
 * @param daysOffset - The number of days in the past you want the date to represent. (Use 0 for today's date).
 */
function getStrDate(daysOffset) {
  // Gets the current date given the offset of days.
  const date = new Date(Date.now() - (daysOffset * 24 * 60 * 60 *1000));
  const day = date.getDate() > 9 ? date.getDate() : date.getDate().toString().padStart(2, `0`)
  const month = date.getMonth() + 1 > 9 ? date.getMonth() + 1: (date.getMonth() + 1).toString().padStart(2, `0`)
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

/**
 * Deletes all data outside a 90 day rolling window.
 * @param - The sheet in which you want to delete historical data outside the 90 day window.
 */
function deletePreviousData(sheet) {
  // Get the date column and store it in the data variable. (My case date column is in row one.)
  const data = sheet.getRange(2, 1, sheet.getLastRow()).getValues().flat(1);
  const cutOffTimestamp = new Date() - 83 * 24 * 60 * 60 * 1000; // Keep 90 days of rolling data. One week will be added after deleting.
  const cutOffDate = new Date(cutOffTimestamp);
  const deleteRange = data.filter(date => date < cutOffDate); // Includes the last row that is empty. Will be accounted for by subtracting 1.
  
  // Make sure there is data to delete, if so delete up to 83 days worth of data.
  deleteRange.length - 1 > 0 ? sheet.deleteRows(2, deleteRange.length - 1)  : null;
}

/**
 * Edits the enrollment status for the E201 by combining and renaming fields.
 * @param sheet - The sheet in which you want to edit enrollment data.
 */
function editEnrollmentFields(sheet) {
  let enrollmentValues = sheet.getRange(2, 7, sheet.getLastRow() - 1).getValues().flat(1);

  // Change all enrollment statuses containing 3rd Year to 3rd Year
  enrollmentValues = enrollmentValues.map(enrollmentStatus => {
    if(enrollmentStatus.includes('3rd Year')) return [`3rd Year`]
    if(enrollmentStatus.includes('In Process:')) return [`In Process: OFA`];
    return [enrollmentStatus];
  });

  // Paste the new enrollment statuses over the old ones. Column G.
  sheet.getRange(2, 7, enrollmentValues.length).setValues(enrollmentValues);
 
  // Get the newly pasted enrollment statuses so we can set up the dummy values.
  const enrollmentStatuses = sheet.getRange(2, 7, sheet.getLastRow() -1).getValues().flat(1);
  sheet.getRange(1, 20).setValue(`Dummy Header`);

  // Get the dummy values.
  const dummyStatuses = enrollmentStatuses.map(enrollmentStatus => {
    switch(enrollmentStatus) {
      case `Enrolled`:
        return `8 - Enrolled`;
      case `Waitlisted`:
        return `7 - Waitlisted`;
      case 'In Process: OFA':
        return `2 - In Process: OFA`;
      case 'Denied-Action Required':
        return `5.1 - Denied Action Required`;
      case `Accepted`:
        return `6 - Accepted`;
      case 'In Process':
        return `3 - In Process`;
      case 'Ready for Review':
        return `4 - Ready for Review`;
      case 'New':
        return `1 - New`;
      case `Waitlisted-Once Age Eligible`:
        return `7 - Waitlisted`;
      case `Denied-Ineligible`:
        return `5.2 - Denied-Ineligible`;
    }   
  });

  // Paste the dummy values.
  sheet.getRange(2, 20, dummyStatuses.length).setValues(dummyStatuses.map(dummyStatus => [dummyStatus]));  
}

/**
 * Creates additional columns on the OFA Notes tab. Used to display more information on the Enrollment Dashboard.
 * @param sheet - The sheet in which you want to create the contact columns. (OFA Notes)
 */
function createContactColumns(sheet) {
  // Sort by child name ascending then sort by case note created date ascending.
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort([{column: 4, ascending: true}, {column: 10, ascending: false}]);

  // For any child with a duplicate record, filter out any note that isn't the most recent case note.
  const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();

  const names = [];
  const filtered = allData.filter(row => {
    if(names.indexOf(row[3]) === - 1) {
      names.push(row[3])
      return true
    }
    return false
  })

  // Clear the sheet and paste the filtered data case note data.
  sheet.clear();
  sheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);

  let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  /////////////////////////////////////////////////// CONTACT TYPE ///////////////////////////////////////////////////

  // Define and set the contact type column for the spreadsheet.
  const contactTypeColumn = 16
  sheet.getRange(1, contactTypeColumn).setValue(`Contact Type`); // Set the column header.

  // Extract the contact type by comparing the "Case Note Created By" to the "Online Application" status.
  const contactTypes = data.map(row => row[10] !== 'Online Application' ? ['Advocate'] : ['OFA - Automated']);

  // Paste the data into the spreadsheet.
  sheet.getRange(2, contactTypeColumn, contactTypes.length).setValues(contactTypes);




  /////////////////////////////////////////////////// DAYS BETWEEN CONTACT ///////////////////////////////////////////////////

  // Define and set the days between last contact column for the spreadsheet.
  const daysBetweenContactColumn = 17;
  sheet.getRange(1, daysBetweenContactColumn).setValue(`Workdays Between Last Contact`); // Set column header.

  // Get and set the days between last contact data. (FORMULA --> NETWORKDAYS())
  const daysBetweenLastContact = data.map((_, index) => [`=NETWORKDAYS(J${index + 2}, TODAY()) - 1`]);
  sheet.getRange(2, daysBetweenContactColumn, daysBetweenLastContact.length).setValues(daysBetweenLastContact).setNumberFormat('0.00');





  /////////////////////////////////////////////////// CENTER ///////////////////////////////////////////////////

  // Define and set the center column for the spreadsheet.
  let centerColumn = 18;
  sheet.getRange(1, centerColumn).setValue(`Center`); // Set column header.

  // Get and set the center formula.
  const centerFormulas = data.map((_, index) => [`=VLOOKUP(A${index + 2},'E201'!A:E, 5, FALSE)`]);
  sheet.getRange(2, centerColumn, centerFormulas.length).setValues(centerFormulas);





  /////////////////////////////////////////////////// PRE-ENROLLMENT LINK ///////////////////////////////////////////////////

  // Define and set the pre-enrollment link column
  const preEnrollmentLinkColumn = 19;
  sheet.getRange(1, preEnrollmentLinkColumn).setValue(`Pre Enrollment Link`);

  // Get the particpantsID and participantRecordIDs needed for the Shine link.
  const participantIDs = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat(1);
  const participantRecordIDs = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues().flat(1);
  const links = participantIDs.map((participantId, index) => [
    `https://shine.acelero.com/PreEnrollment/Participant/Records?participantId=${participantId}&participantRecordId=${participantRecordIDs[index]}`
  ]);

  // Set the link data on the spreadsheet.
  sheet.getRange(2, preEnrollmentLinkColumn, links.length).setValues(links);





  /////////////////////////////////////////////////// FAMILY ADVOCATE ///////////////////////////////////////////////////

  // Define and set the center column for the spreadsheet.
  const familyAdvocateColumn = 20;
  sheet.getRange(1, familyAdvocateColumn).setValue(`Family Advocate Assigned`); // Set column header.

  // Get the family advocate assignments
  const familyAdvocates = sheet.getRange(2, 5, sheet.getLastRow() - 1).getValues().flat(1) // Get each family advocate for each child.
  const familyAdvocateAssignments = familyAdvocates.map(familyAdvocate => familyAdvocate !== 'none' ? [1]: [0]);

  // Paste the family advocate assignements.
  sheet.getRange(2, familyAdvocateColumn, familyAdvocateAssignments.length).setValues(familyAdvocateAssignments);
}

/**
 * Splits the hyperlink from a column into two columns. The first column is the hyperlinks text and the second is the hyperlinks URL.
 * @param sheet - The sheet where the hyperlink lives.
 * @param column - The column of the sheet where the hyperlink exists.
 */
function splitHyperlink(column, sheet) {
  const data = sheet.getRange(2, column, sheet.getLastRow() - 1).getRichTextValues().flat();
  sheet.insertColumnBefore(column + 1);
  
  // Get the live link column data.
  let liveLinkColumnData = [[sheet.getRange(1, column).getValue(), 'Live Link']];
  liveLinkColumnData = liveLinkColumnData.concat(data.map(rtv => [rtv.getText(), rtv.getLinkUrl()]));

  // Paste the live link column data.
  sheet.getRange(1, column, liveLinkColumnData.length, liveLinkColumnData[0].length).setValues(liveLinkColumnData);
}

const getClassroomEnrollments = function() {
  const e201 = ss.getSheetByName(`E201`);
  const data = e201.getRange(2, 5, e201.getLastRow(), 3).getValues();
  let classrooms = [];
  let centers = [];
  data.forEach(row => {
    if(row[2] === `Enrolled`) {
      centers.push(row[0]);
      classrooms.push(row[1]) 
    }
  })
  const distinctClassrooms = [...new Set(classrooms)];

  let classroomEnrollment = [];

  distinctClassrooms.forEach(distinctClassroom => {
    classroomEnrollment.push({
      center: getCenter(data, distinctClassroom),
      classroom: distinctClassroom,
      enrolledChildren: classrooms.filter(classroom => classroom === distinctClassroom).length
    })
  })

  return classroomEnrollment;
}

const getCenter = function(data, distinctClassroom) {
  let center = ``;
  data.forEach(row => {
    if(row[1] === distinctClassroom) {
      return center = row[0];
    }
  })
  return center;
}

const getClassroomType = function(classroom) {
  if(classroom.includes(`ED`)) {
    return `ED`
  }
  else if(classroom.includes(`AM`) || classroom.includes(`PM`)) {
    return `DS`
  }
  else if(classroom.includes(`State Pre K`) || classroom.includes('State PREK')) {
    return `State Pre K`
  }
  else if(classroom.includes(`EHS`)) {
    return `EHS`
  }
  else {
    return `FD`
  }
}

const createFundedEnrolled = function() {
  const classroomEnrollment = getClassroomEnrollments();
  const fundedEnrollmentSheet = ss.getSheetByName(`Enrollment Report`);
  fundedEnrollmentSheet.clear();

  // Set Headers
  fundedEnrollmentSheet.getRange(1, 1).setValue(`Center`);
  fundedEnrollmentSheet.getRange(1, 2).setValue(`Classroom`);
  fundedEnrollmentSheet.getRange(1, 3).setValue(`Classroom Type`);
  fundedEnrollmentSheet.getRange(1, 4).setValue(`Enrolled Children`);
  fundedEnrollmentSheet.getRange(1, 5).setValue(`Funded Enrollment`);

  // Get the data.
  const data = classroomEnrollment.map((dataPoint, index) => [
      dataPoint.center,
      dataPoint.classroom,
      getClassroomType(dataPoint.classroom),
      dataPoint.enrolledChildren,
      `=vlookup(B${index + 2},'Funded Enrollment'!B:D, 3, FALSE)`
    ]
  )

  // Paste the data.
  fundedEnrollmentSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}


const createLiveLinks = function(sheet) {
  // Change the value type of the participantID column to plain text
  ss.getSheetByName(`E201`).getRange(1, 1, ss.getSheetByName(`E201`).getLastRow(), 2).setNumberFormat(`0`);
  sheet.getRange(1, 3, sheet.getLastRow()).setNumberFormat(`0`);

  // Create the participant ID Column
  let contactTypeColumn = 12
  // Set the column header and column format.
  sheet.getRange(1, contactTypeColumn).setValue(`Participant ID`);
  sheet.getRange(1, contactTypeColumn, sheet.getLastRow()).setNumberFormat(`0`);

  // Set a formula in each column to get the participantsid from the e201 spreadsheet.
  let participantIDs = sheet.getRange(2, contactTypeColumn, sheet.getLastRow() - 1).getValues().flat(1);
  participantIDs = participantIDs.map((_, index) => [`=INDEX('E201'!A:A, (MATCH(C${index + 2},'E201'!B:B, 0)))`]);

  // Paste the participant id vlookup formula into each cell in column 12.
  sheet.getRange(2, contactTypeColumn, participantIDs.length).setValues(participantIDs);


  // Retrieve the participant ids as a text value after setting their values with the formula.
  participantIDs = sheet.getRange(2, contactTypeColumn, sheet.getLastRow() - 1).getValues().flat(1);

  // Create the live link column
  contactTypeColumn = 13;
  sheet.getRange(1, contactTypeColumn).setValue(`Live Link`);

  // Construct the live links and store them in an array.
  const participantRecordIDs = sheet.getRange(2, 3, sheet.getLastRow() - 1).getValues().flat(1);
  const liveLinks = participantIDs.map((id, index) => [
    `https://shine.acelero.com/PreEnrollment/Participant/Records?participantId=${id}&participantRecordId=${participantRecordIDs[index]}`
  ]);

  // Paste the live links to the waitlist spreadsheet.
  sheet.getRange(2, contactTypeColumn, liveLinks.length).setValues(liveLinks);


  contactTypeColumn = 14

  // Set the column header.
  sheet.getRange(1, contactTypeColumn).setValue(`Program Option Complete`);

  // Get the program options
  const programOptions = sheet.getRange(2, 5, sheet.getLastRow() - 1).getValues().flat(1);

  // Loop through the program options and return a 1 if 'No Preference' is given else 0.
  const completedOptions = programOptions.map(programOptions => programOptions === 'No Preference' ? [1]: [0]);

  // Paste the completedOptions into the spreadsheet.
  sheet.getRange(2, contactTypeColumn, completedOptions.length).setValues(completedOptions);
}

const editHistoricalEnrollment = function() {
  const fundedEnrollment = ss.getSheetByName(`Funded Enrollment`);
  // Get enrolled total for today.
  const classroomTypes = fundedEnrollment.getRange(2, 3, fundedEnrollment.getLastRow() - 1).getValues().flat(1);
  const fundedEnrollmentData = fundedEnrollment.getRange(2, 4, fundedEnrollment.getLastRow() - 1).getValues().flat(1)
  .filter((_, index) => classroomTypes[index] !== 'EHS');

  const enrollmentData = fundedEnrollment.getRange(2, 5, fundedEnrollment.getLastRow() - 1).getValues().flat(1)
  .filter((_, index) => classroomTypes[index] !== `EHS`);

  // Get the values for enrollment and funded enrollment
  const fundedEnrolledValue = fundedEnrollmentData.reduce((a, b) => a + b, 0)
  const enrolledValue = enrollmentData.reduce((a, b) => a + b, 0)

  const historicalEnrollmentSheet = ss.getSheetByName(`Enrollment (Historical)`);
  const editRow = historicalEnrollmentSheet.getLastRow() + 1;

  historicalEnrollmentSheet.getRange(editRow, 1).setValue(getStrDate(0));
  historicalEnrollmentSheet.getRange(editRow, 2).setValue(enrolledValue);
  historicalEnrollmentSheet.getRange(editRow, 3).setValue(fundedEnrolledValue);
  historicalEnrollmentSheet.getRange(editRow, 4).setValue(`=B${editRow} / C${editRow}`);
}


/**
 * Adds all other children to the spreadsheet with the iep status of no IEP. This will help us to graph IEP / Non IEP
 * @param activeIEPSheet - A sheet object that represents the "Active IEP" tab
 */
function editActiveIEP(activeIEPSheet) {
  const e201Sheet = ss.getSheetByName('E201');
  const iepParticipantIDs = activeIEPSheet.getRange(2, 1, activeIEPSheet.getLastRow() - 1).getValues().flat(1);
  const iepParticipantRecordIDs = activeIEPSheet.getRange(2, 2, activeIEPSheet.getLastRow() - 1).getValues().flat(1);

  // Create Live Links for any child with an IEP/IFSP
  const liveLinks = iepParticipantRecordIDs.map(id => [`https://shine.acelero.com/ParticipantRecord/DisabilitiesMentalHealth/Details/${id}`]);

  // Set the header and paste the data.
  const lastColumn = activeIEPSheet.getLastColumn() + 1;
  activeIEPSheet.getRange(1, lastColumn).setValue('Live Link')
  activeIEPSheet.getRange(2, lastColumn, activeIEPSheet.getLastRow() - 1).setValues(liveLinks);

  // Place all the children without and IEP into the spreadsheet. 
  // This is used by Google Data Studio to preform the 10% required calculation
  const enrolledChildrenWithoutIEP = e201Sheet.getRange(2, 1, e201Sheet.getLastRow() - 1, e201Sheet.getLastColumn())
                                    .getValues()
                                    .filter(row => row[6] === 'Enrolled')
                                    .filter(row => !iepParticipantIDs.includes(row[0]))
                                    .map(row => row.slice(0, 6));
  
  activeIEPSheet.getRange(activeIEPSheet.getLastRow() + 1, 1, enrolledChildrenWithoutIEP.length, enrolledChildrenWithoutIEP[0].length).setValues(enrolledChildrenWithoutIEP);
}

/**
 * Adds the "Request for Evaluation to LEA" notes to the "Referral to LEA" tab.
 * @param referralToLEASheet - A sheet object that represents the "Active IEP" tab
 */
function editReferralToLEA(referralToLEASheet) {

  /////////////////////// IMPORT THE PARTICIPANT RECORD ID FIELD FROM THE "E201" SHEET ///////////////////////////////////////
  const e201Sheet = ss.getSheetByName('E201');
  
  // Get the name of each child in the "Referral to LEA" sheet. This will be our unique reference.
  const fullNames = referralToLEASheet.getRange(2, 1, referralToLEASheet.getLastRow() -1).getValues().flat(1);

  // A two dimensional array consisting of [participantRecordID, Child Name]
  const enrolledData = e201Sheet.getRange(2, 2, e201Sheet.getLastRow() - 1, 2).getValues();

  // Loop through each name in the "Referral to LEA" sheet and store there participant record id from the e201 tab.
  const participantRecordIDs = fullNames.map(fullName => {
    // Filtered the enrolled data for the given name. Returns a two dimension array the value we want is at [0][0] (record id)
    return enrolledData.filter(row => row[1] === fullName)[0][0];
  })

    // The values given to setValues should be wrapped in an array of length 1. // map function will do that.
  const liveLinks = participantRecordIDs.map(id => [`https://shine.acelero.com/ParticipantRecord/DisabilitiesMentalHealth/Details/${id}`])

  // Set the header and paste the data.
  let lastColumn = referralToLEASheet.getLastColumn() + 1;
  referralToLEASheet.getRange(1, lastColumn).setValue('Live Link');
  referralToLEASheet.getRange(2, lastColumn, referralToLEASheet.getLastRow() - 1).setValues(liveLinks);







  /////////////////////// IMPORT THE IEP/IFSP FIELD FROM THE "ACTIVE IEP" SHEET ///////////////////////////////////////
  const activeIEPSheet = ss.getSheetByName('Active IEP')

  // A two dimensional array consisting of [participantRecordID, Child Name]
  const iepData = activeIEPSheet.getRange(2, 3, activeIEPSheet.getLastRow() - 1, 7).getValues();

  // Loop through each name in the "Referral to LEA" sheet and store there participant record id from the e201 tab.
  const planTypes = fullNames.map(fullName => {
    // Filtered the enrolled data for the given name. Returns a two dimension array the value we want is at [0][0] (record id)
    const planData = iepData.find(row => row[0] === fullName);

    // Check to see if the plan exists, if it doesn't return no plan.
    return planData ? planData[6] : "No Plan";
  })

  // Set the header and paste the data.
  lastColumn = referralToLEASheet.getLastColumn() + 1;
  referralToLEASheet.getRange(1, lastColumn).setValue('Type (IEP or IFSP)'); // Name must match "Active IEP" for filtering

  // setValues needs a two dimension array. Map will handle that.
  referralToLEASheet.getRange(2, lastColumn, referralToLEASheet.getLastRow() - 1).setValues(planTypes.map(planType => [planType]));
  
  
}



