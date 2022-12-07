/**
 * Returns the FIRST attachment of the message with the same subject being passed in.
 * @param messageSubject - A string representing the subject of the message you want to retrieve the attachment from.
 * @param daysOffset (optional) - The number of days you want to look in the past. Defaults to yesterday if not specified.
 * #param startPos (optional) - The starting position to begin searching your email threads. Defaults to the first email in your inbox.
 * @param endPos (optional) - The position at which you want to stop searching your email threads. Defaults to the last email thread on the first page of your inbox.
 */
 function getAttachment(messageSubject, daysOffset=7, startPos=0, endPos=150) {
    // determine the timestamp for today.
    const cutOffTimeStamp = Date.now() - (daysOffset * 24 * 60 * 60 * 1000);
  
    // Get the specified number of threads from your inbox.
    const threads = GmailApp.getInboxThreads(startPos, endPos);
    return findMessageAttachment(threads, messageSubject, cutOffTimeStamp);
  }
  
  /**
   * Finds the first message wihtin the given set of threads. This should only by called by getAttachment above.
   * @param threads - An array of email threads from the users inbox.
   * @param messageSubject - A string representing the subject of the message you want to retrieve.
   * @param cutOffTimeStamp - The cut-off time for the message you want to retrieve.
   */
  function findMessageAttachment(threads, messageSubject, cutOffTimeStamp) {
    return threads
            // Get all messages for each thread.
            .find(thread => thread.getMessages()
            // Return the first message that has the same subject.
            .find(message => message.getSubject() === messageSubject && message.getDate().getTime() >= cutOffTimeStamp))
            // The two find methods return a thread.
            // Return the first attachment of the first message in the found thread. 
            .getMessages()[0].getAttachments()[0];
  }
  
  
  /**
   * This function stores a file in your email on the drive.
   * @param msgSubject - The subject of the message containing the .xlsx attachment
   * @param fileName - The name you want the file to have on the drive. No extension is necessary.
   */
  function storeFileOnDrive(attachment, fileName) {
    // Location the Automated - MBI folder on the users drive or create it if it doesn't exists.
    const convertedSheetFolders = DriveApp.getFoldersByName(`Converted Sheets`);
    const folderFound = convertedSheetFolders.hasNext();
    let convertedSheetFolder = folderFound ? convertedSheetFolders.next() : DriveApp.createFolder(`Converted Sheets`);
  
    // Create a new .xlsx file on your drive.
    const xlsxFile = DriveApp.createFile(attachment.copyBlob());
    xlsxFile.setName(fileName);
    xlsxFile.moveTo(convertedSheetFolder);
  }
  
  /**
   * Written by Amit Agarwal
   * A function to convert an .xlsx file on your drive to a google spreadsheet.
   * @param fileName - The name you want the file to have on the drive. No extension is necessary.
   */
  function convertExceltoGoogleSpreadsheet(fileName) {
    try {
  
      // Written by Amit Agarwal
      // www.ctrlq.org;
  
      // Find and store the excel file from your drive with the given name.
      const excelFile = DriveApp.getFilesByName(fileName).next();
      const fileId = excelFile.getId();
      const folderId = Drive.Files.get(fileId).parents[0].id;
      const blob = excelFile.getBlob();
      // Build a resource object with the parent folder to pass to Drive.Files.insert();
      const resource = {
        title: excelFile.getName(),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{id: folderId}],
      };
  
      Drive.Files.insert(resource, blob);
  
    } catch (f) {
      Logger.log(`Error: f.toString()`); // Logs an error message if it occurs.
    }
  }
  
  /**
   * Writes a file from your drive to the spreadsheet.
   * @param fileName - The name of the file you want to find on the drive. No extension is necessary.
   */
  function writeFileToSpreadsheet(sheet, fileName, linkColumn) {
    const googleSheet = DriveApp.getFilesByName(fileName).next();
    const fileId = googleSheet.getId();
  
    sheet.clear();
  
    const sourceSheet = SpreadsheetApp.openById(fileId).getSheets()[0];
    const sourceValues = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).getValues();
  
    sheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).setValues(sourceValues);
  
    copyPasteFormatting(sourceSheet, sheet, linkColumn);
  } 
  
  /**
   * Writes a file from your drive to the end of a spreadsheet without clearing it first. Allows for historical reporting.
   * @param fileName - The name of the file you want to find on the drive. No extension is necessary.
   */
  function writeFileToEndOfSpreadsheet(sheet, fileName, metaTop, metaBottom) {
    const googleSheet = DriveApp.getFilesByName(fileName).next();
    const fileId = googleSheet.getId();
  
    const sourceSheet = SpreadsheetApp.openById(fileId).getSheets()[0];
    const sourceValues = sourceSheet.getRange(1 + metaTop, 1, sourceSheet.getLastRow() - metaTop - metaBottom, sourceSheet.getLastColumn()).getValues();
  
    const startRow = sheet.getLastRow();
    sheet.getRange(startRow + 1, 2, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
    addDateRows(sheet, startRow + 1, sourceValues.length, 1);
  } 
  
  /**
   * Cleans a data sheet by deleting meta data from top and bottom.
   * @param sheetName - A string with the same name as the google sheet to clean.
   * @param top (optional) - The number of rows to delete from the top of the sheet, if not specified no rows will be deleted from the top.
   * @param bottom (optional) - The number of rows to delete from the bottom of the sheet, if not specified no rows will be deleted from the bottom.
   */
  function removeMetaData(sheet, top=0, bottom=0) {
    if(sheet.getLastRow() !== top + bottom - 1) {
      top !== 0 ? sheet.deleteRows(1, top) : null;
      bottom !== 0 ? sheet.deleteRows(sheet.getLastRow() - bottom + 1, bottom) : null;
    }
  }
  
  /**
   * Copy and paste formatting of data from of sheet to another. Takes all rows and columns of source sheet and copies there formats to all rows and columns of dest sheet.
   * @param sourceSheet - The sheet you want to copy the formatting from.
   * @param destSheet - The sheet you want to paste the formatting to.
   * @param linkColumn - Optional - A column that has a link you would like to preserve.
   */
  function copyPasteFormatting(sourceSheet, destSheet, linkColumn) {
    // Get data and formatting from the source sheet
    const range = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
    const background = range.getBackgrounds();
  
    // Put data and formatting in the destination sheet
    const destRange = destSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
    destRange.setBackgrounds(background);
    
    // If the link range was not given by the writeFileToSpreadsheet then skip this part.
    if(!linkColumn) return;
  
    // Get the link values from the sourceSheet.
    sourceSheet.getRange(1, linkColumn, sourceSheet.getLastRow()).setNumberFormat(`@`);
    const linkValues = sourceSheet.getRange(1, linkColumn, sourceSheet.getLastRow()).getRichTextValues()
  
    // Paste the link values into the destination sheet.
    destSheet.getRange(1, linkColumn, linkValues.length).setRichTextValues(linkValues);
  
  }
  
  /**
   * Determines if the current day is a weekday or weekend.
   */
  const isWeekday = function() {
    const day = new Date().getDay();
    return (day !==0 && day !==6);
  }
  
  
  
  
  
  
  
  
  
  
  
  