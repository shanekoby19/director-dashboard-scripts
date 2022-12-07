function getAbsenceData () {
    // // Define the data sheets and remove metadata.
    const dataInSheet = ss.getSheetByName('Data In');
    const normalizedDataInSheet = ss.getSheetByName('Normalized - Data In');
    
    init(dataInSheet);
    normalizeData(dataInSheet, normalizedDataInSheet);
  
    // Clear the "'Other' Notes" sheet before starting to synthesize notes.
    const otherNotesSheet = ss.getSheetByName('"Other" Notes')
    otherNotesSheet.getRange(2, 1, otherNotesSheet.getLastRow() - 1).clear();
  
    synthesizeNotes(normalizedDataInSheet);
  }
  
  /**
   * Cleans up data in to a useable state for the aggregation function.
   * @params dataInSheet - A sheet object representing the "Data In" tab.
   */
  function init(dataInSheet) {
    // This will need to updated each run based on the number of days open in a week.
    removeMetaData(dataInSheet, 16, 1);
  
    // Get the data that needs to be merged. Find each column and each value in that column.
    const mergeData = dataInSheet.getRange(2, 1, 1, dataInSheet.getLastColumn()).getValues().flat(1);
    const columnIndexes = mergeData.map((value, index) => value !== '' ? index + 1 : undefined).filter(value => value !== undefined);
  
    // Define and inital date and initial remainder so we can track any number of days.
    let strDate = ''
    const initialRemainder = columnIndexes[0] % 3;
    columnIndexes.forEach(column => {
      if(column % 3 === initialRemainder) {
        const date = dataInSheet.getRange(1, column).getValue();
        strDate = `${date.getMonth() + 1}-${date.getDate()}-${date.getFullYear()}`
        dataInSheet.getRange(1, column).setValue(`Attendance Code - ${strDate}`)
      }
      else if(column % 3 === initialRemainder + 1) {
        dataInSheet.getRange(1, column).setValue(`Absense Reason - ${strDate}`)
      }
      else {  
        dataInSheet.getRange(1, column).setValue(`Absense Reason Notes - ${strDate}`)
      }
    })
  
    // Delete row 2 as it has already been merged with row 1.
    dataInSheet.deleteRow(2);
  
    // Delete any unused columns
    const unusedColumns = ['Enrollment Date', 'Entry Date',	'Termination Date', 'Present', 'Late', 'Present All', 'Monthly ADA']
    let allColumns = dataInSheet.getRange(1, 1, 1, dataInSheet.getLastColumn()).getValues().flat(1);
    let deletedColumns = 0;
    allColumns.forEach((columnName, columnIndex) => {
      if(unusedColumns.indexOf(columnName) > -1) {
        dataInSheet.deleteColumn(columnIndex + 1 - deletedColumns);
        deletedColumns += 1;
      }
    })
  
    // Filter out days that have less than 50% of attendance entered. These should represent days where we were closed but the closure hasn't been entered into Shine yet.
    allColumns = dataInSheet.getRange(1, 1, 1, dataInSheet.getLastColumn()).getValues().flat(1);
    const attendanceCodeColumns = allColumns.map((value, index) => value.includes('Attendance Code') ? index + 1 : undefined)
                                            .filter(value => value !== undefined);
    let columnsToDelete = [];
    attendanceCodeColumns.forEach(columnIndex => {
      const attendanceCodeData = dataInSheet.getRange(2, columnIndex, dataInSheet.getLastRow() - 1).getValues().flat(1);
      const notYetEnteredAttendance = attendanceCodeData.filter(value => value === '');
      if((notYetEnteredAttendance.length / attendanceCodeData.length) > .50) {
        columnsToDelete.push(columnIndex, columnIndex + 1, columnIndex + 2);
      }
    });
  
    deletedColumns = 0;
    columnsToDelete.forEach(column => {
      ss.deleteColumn(column - deletedColumns)
      deletedColumns += 1;
    })
  
  }
  
  /**
   * Creates a record for every day that the child was in attendance.
   * @params dataInSheet - A sheet object representing the "Data In" tab.
   * @params normalizedDataInSheet - A sheet object representing the "Normalized - Data In" tab.
   */
  function normalizeData(dataInSheet, normalizedDataInSheet) {
    // Clear the data sheet if any data already exists.
    normalizedDataInSheet.getRange(2, 1, normalizedDataInSheet.getLastRow() - 1, normalizedDataInSheet.getLastColumn()).clear();
  
    // Define the data
    const data = dataInSheet.getRange(2, 1, dataInSheet.getLastRow() - 1, dataInSheet.getLastColumn()).getValues();
    const allColumns = dataInSheet.getRange(1, 1, 1, dataInSheet.getLastColumn()).getValues().flat(1);
    const absenceCodeColumns = allColumns.map((value, index) => value.includes('Attendance Code') ? index + 1 : undefined)
                                         .filter(value => value !== undefined);
  
    let finalData = [];
    data.forEach(row => {
      const staticValues = row.slice(0, 4)
      absenceCodeColumns.forEach((columnIndex) => {
        const date = dataInSheet.getRange(1, columnIndex).getValue().replace('Attendance Code - ', '')
        const dynamicValues = row.slice(columnIndex - 1, columnIndex + 2);
  
        if(dynamicValues[0] === "Absent") {
          finalData.push(staticValues.concat(date).concat(dynamicValues));
        }
      });
    });
  
    // Paste the final data
    normalizedDataInSheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
  }
  
  /**
   * 
   */
  function synthesizeNotes(normalizedDataInSheet) {
    // Define the data sheets needed to synthesize the notes
    const keywordsSheet = ss.getSheetByName('Keywords');
  
    // Get all the keyword data and create an array of objects containing the final phrase and any match phrases.
    const postedKeywords = keywordsSheet.getRange(2, 1, keywordsSheet.getLastRow() - 1).getValues().flat(1);
    const keywordData = keywordsSheet.getRange(2, 2, keywordsSheet.getLastRow() - 1, keywordsSheet.getLastColumn() - 1).getValues();
    const keywordObjs = postedKeywords.map((postedKeyword, index) => ({
      postedKeyword,
      keywordMatches: keywordData[index].filter(match => match !== '').map(match => match),
    }));
  
    // Get all the entire week of attendance data.
    // We will write any "A" with the synthesized absent reason.
    // We will write any "PO" with virtual.
    const absentNotes = normalizedDataInSheet.getRange(2, 8, normalizedDataInSheet.getLastRow() - 1).getValues().flat(1);
    const synthesizedAbsentNotes = absentNotes.map(absentNote => [getAbsenceReason(keywordObjs, absentNote.toLowerCase())]);
  
    // Place the synthesized data
    normalizedDataInSheet.getRange(2, 9, synthesizedAbsentNotes.length, 1).setValues(synthesizedAbsentNotes);
  
    // Get the e201 sheet and all it's data
    const e201Sheet = ss.getSheetByName('e201');
    const e201Data = e201Sheet.getRange(2, 1, e201Sheet.getLastRow() - 1, e201Sheet.getLastColumn()).getValues();
  
    // Link to the family advocate (uses PARTICIPANT RECORD ID)
    const participantRecordIDs = normalizedDataInSheet.getRange(2, 4, normalizedDataInSheet.getLastRow() - 1).getValues().flat(1);
    const advocates = participantRecordIDs.map(participantRecordID => [e201Data.filter(row => row[1] === participantRecordID)[0][13]]);
    const participantIDs = participantRecordIDs.map(participantRecordID => [e201Data.filter(row => row[1] === participantRecordID)[0][0]]);
  
    // Place the advocates into the spreadsheet.
    normalizedDataInSheet.getRange(2, 10, advocates.length, 1).setValues(advocates);
    normalizedDataInSheet.getRange(2, 11, participantIDs.length, 1).setValues(participantIDs);
  
    // Create links to the family tab for each child and then paste the data to the "Normalized - Data In" sheet
    const links = participantRecordIDs.map(participantRecordID => [`https://shine.acelero.com/ParticipantRecord/Family/Details/${participantRecordID}`]);
    normalizedDataInSheet.getRange(2, 12, links.length, 1).setValues(links);
  }
  
  
  /**
   * Returns an a synthesized absence reason based on an absence note.
   * @param keywordObjs - A list of objects that represent valid keywords. Created from "Keywords" sheet.
   * @param absenceNote - The note for the given absence day.
   */
  function getAbsenceReason(keywordObjs, absenceNote) {
    const otherNotesSheet = ss.getSheetByName('"Other" Notes');
  
    if(absenceNote === '') {
      return 'No Note'
    }
    else {
      const filteredObjs = keywordObjs.filter(keywordObj => {
        return keywordObj.keywordMatches.filter(match => absenceNote.includes(match)).length
      });
      if(filteredObjs.length === 0) {
        // Store the unmatched note on the "Other Notes" tab.
        otherNotesSheet.getRange(otherNotesSheet.getLastRow() + 1, 1).setValue(absenceNote);
        return 'Other'
      }
      else {
        return filteredObjs[0].postedKeyword;
      }
    }
  }
  
  
  /**
   * Creates a new historical record for the given week
   * @param programSheet - A sheet object that points to the current "Program Sheet".
   */
  function addHistoricalRecord(programSheet) {
    // Define the data sheet.
    const historicalDataSheet = ss.getSheetByName('Historical Data');
  
    // Get the current week value and the row to paste the data
    const currentWeek = programSheet.getRange(1, 1).getValue();
    const currentRow = historicalDataSheet.getLastRow() + 1;
  
    // Get all the reasons in total and also get just the unique reasons
    const allReasons = programSheet.getRange(2, 3, programSheet.getLastRow() - 1, 5).getValues().flat(1);
    const uniqueReasons = [...new Set(allReasons)];
  
    // Loop through each unique reason and create an object with the reason name, count and percentage
    const uniqueReasonObjs = uniqueReasons.map(uniqueReason => {
      const count = allReasons.filter(reason => reason === uniqueReason).length; 
      return {
        name: uniqueReason,
        count,
        percentage: ((count / allReasons.length) * 100).toFixed(2) + '%',
      }
    });
  
  
    // Sort the reason object by their count.
    uniqueReasonObjs.sort((objA, objB) => objB.count - objA.count);
  
    // Get the data to be pasted into the google sheet.
    const dataRow = uniqueReasonObjs.map(uniqueReasonObj => `${uniqueReasonObj.count} ${uniqueReasonObj.name} - ${uniqueReasonObj.percentage}`);
  
    // Paste the data.
    historicalDataSheet.getRange(currentRow, 1).setValue(currentWeek);
    historicalDataSheet.getRange(currentRow, 2, 1, dataRow.length).setValues([dataRow]);
  }