// gsheet-sheet.js

// ***********************************************************************************************************************************************************************************************

/**
 * Extracts all URLs from the display value of a cell.
 *
 * @param cell - The cell object containing the text to search.
 * @returns {string[] | null} Array of URLs if any were found, otherwise null.
 */
function extractURLs(cell) {
  
 let text = cell.getDisplayValue();
 let urlRegex = /(https?:\/\/[^\s]+)/g;
 //return urls
 return text.match(urlRegex);

}


/**
 * Updates the value of the specified cell in the active sheet of the current spreadsheet.
 *
 * @param {string} cellAddress - The address of the cell to update (e.g., "A1").
 * @return {void} This function does not return a value.
 */
function updateCells(cellAddress) {
// Open the active sheet
 let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// Update cell A1
sheet.getRange(cellAddress).setValue("Hello world!");
}

// ***********************************************************************************************************************************************************************************************
/**
 * Compares values in the first range against the given target value and
 * appends the corresponding values from the second range as quoted text.
 *
 * @param {Array<Array<*>>} range1 - First range of values to compare.
 * @param {Array<Array<*>>} range2 - Second range of values to append from.
 * @param {*} iValue - Value to match against entries in range1.
 * @returns {string} A comma-separated list of quoted matching values, or an empty string if no match is found.
 */
function COMPAREANDAPPENDTEXT(range1, range2, iValue) {
  //=COMPAREANDAPPENDTEXT(C4:BN4,C1:BN1)  
  if (!range1 || !range2) return '';

  // Flatten in case the ranges are 1 row × N or N × 1
  let flat1 = range1.flat();
  let flat2 = range2.flat();

  if (flat1.length !== flat2.length) {
  throw new Error('Both ranges must have the same number of cells');
  }

  let output = [];

  for (let i = 0; i < flat1.length; i++) {
    if (flat1[i] === iValue || flat1[i] === iValue.toString()) {
    output.push('"' + flat2[i] + '"');
    }
  }

  return output.join(',');
}

/**
 * Hides or unhides columns in a Google Sheets spreadsheet based on the values in a specified row range.
 * If a cell in the specified row contains the value `0`, the corresponding column will be hidden.
 * If a cell contains any other value, the corresponding column will be unhidden.
 *
 * @return {void} This method does not return a value.
 */
function hideOrUnhideColumnsBasedOnRowRange() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // --- Define the range you want to check ---
  const range = sheet.getRange('C2:ci2'); // <<== you can change this
  
  const values = range.getValues()[0]; // Get first (and only) row
  
  const startColumn = range.getColumn(); // Column number of C (e.g., 3)

  // Loop through each cell in the range
  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i];
    const currentColumn = startColumn + i;
    
    if (cellValue === 0) {
      sheet.hideColumns(currentColumn);
    } else {
      sheet.showColumns(currentColumn);
    }
  }
}

/**
 * Unhides all columns within a specified range in the active sheet of the active spreadsheet.
 *
 * This function retrieves a predefined range of cells, iterates over its columns,
 * and unhides any hidden columns within that range.
 *
 * @return {void} Does not return a value.
 */
function UnhideAllColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // --- Define the range you want to check ---
  const range = sheet.getRange('C2:ci2'); // <<== you can change this
  
  const values = range.getValues()[0]; // Get first (and only) row
  
  const startColumn = range.getColumn(); // Column number of C (e.g., 3)

  // Loop through each cell in the range
  for (let i = 0; i < values.length; i++) {
      const currentColumn = startColumn + i
      sheet.showColumns(currentColumn);
    }
}

/**
 * Appends the extracted row data to a specific tab within a Google Sheet.
 *
 * @param {Array<Array<string>>} extractedRows - The array of rows to write.* @param {string} sheetID - The ID of the target Google Sheet.
 * @param {string} targetSheetName - The name of the specific tab (worksheet) to append to.
 * @returns {Result}
 */
const appendTableRowsToSheet = ( extractedRows, sheetID, targetSheetName ) => {
  const functionName = 'appendTableRowsToSheet';
  Logger.log(`${functionName}. Started.`);


  if (extractedRows.length === 0 || extractedRows[0].length === 0) {
    return theResults(false, 'No valid data rows to append.', functionName);
  }
  Logger.log(`appendTableRowsToSheet. sheetID = ${sheetID}. targetSheetName = ${targetSheetName}.  extractedRows.length = ${extractedRows.length}`);
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetID);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    if (!sheet) return theResults(false, `Could not find sheet tab named: ${targetSheetName}`, functionName);

    // Determine the starting row for the new data
    // ➡️ NEW: Create modified rows with prepended row numbers
    let startRow = sheet.getLastRow() + 1;
    const modifiedRows = extractedRows.map((row, index) => {
      const rowNumber = startRow + index;
      return [rowNumber, ...row]; // Prepend row number to each row
    });

    const numRows = modifiedRows.length;
    const numColumns = modifiedRows[0].length; // Now includes the extra column
    
    Logger.log(`appendTableRowsToSheet: numRows: ${numRows}, numColumns: ${numColumns}, startRow: ${startRow}, modifiedRows[0]: ${modifiedRows[0]}`);
    
    // Define the target range (starting at column 1, now with extra column)
    const range = sheet.getRange(startRow, 1, numRows, numColumns);

    // Append all rows at once for optimal performance
    range.setValues(modifiedRows);

    return theResults(true, `Successfully appended ${numRows} rows to tab: ${targetSheetName}`, functionName);

  } catch (error) {
    return theResults(false, `Failed to open sheet or append data: ${error}`, functionName);
  }
}

/**
 * Reads rows from a named tab in a Google Sheet and returns a hash of inputs.
 *
 * @param requiredInitInfo - Accessed via getRequiredInitInfo().sheetTabTitle
 * @param requiredAttribsHash - defined within config.js. Accessed via getConfigHash().
 *        This is the initialization hash that defines the required attributes.
 *        It never gets values for it's keys.
 * @param validatorFunctions  - List of functions that validate the input data.
 * @returns {Result} - An object containing the result of the operation.
 */
const readInput = (requiredInitInfo, requiredAttribsHash, validatorFunctions) => {
  const functionName = 'readInput';
  Logger.log(`${functionName}. Started.`);
  let returnResult;


  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(requiredInitInfo.sheetID);
  } catch (err) {
    return theResults(false, ` Could not open spreadsheet with id: ${requiredInitInfo.sheetID} - ${err.message}`, functionName);
  }

  // --- Find the tab ---
  const sheet = spreadsheet.getSheetByName(requiredInitInfo.sheetInputTabTitle);
  if (!sheet) {
    return theResults(false, `Could not find tab named: ${requiredInitInfo.sheetInputTabTitle} in spreadsheet: ${requiredInitInfo.sheetID}`, functionName);
  }

  // --- Read the rows ---
  let rawRows;
  try {
    rawRows = sheet.getRange(1, 1, requiredInitInfo.inputRows, 2).getValues();
  } catch (err) {
    return theResults(false, ` Could not read rows from tab: ${requiredInitInfo.sheetInputTabTitle} - ${err.message}`, functionName);
  }

  // --- Convert to tuples ---
  // getValues() returns an array of arrays already
  // We just need to make sure values are strings
  const tuples = rawRows.map(row => [
    row[0].toString().trim(),
    row[1].toString().trim(),
  ]);

    // --- Step Validate and load config ---
  returnResult = populateInputValues(tuples,requiredAttribsHash,validatorFunctions);
  if (!returnResult.ok) {
    return theResults(false, returnResult.message, functionName);
  }
  const combinedData =   { ...requiredInitInfo, ...returnResult.data };
  return theResults(true, 'Completed.', functionName, combinedData);
};

const createCurrentSheetTabSnapshot = (sourceSheetName,newSheetName,dateHeader,topicHeader) => {
  const functionName = 'createCurrentSheetTabSnapshot';
  Logger.log(`${functionName}. Started.`);


  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    return theResults(false, `Error. Source sheet '${sourceSheetName}' not found. Please check the sheet name.`, functionName);
  }

  // Get the data range from the source sheet
  const sourceRange = sourceSheet.getDataRange();
  const valuesToCopy = sourceRange.getValues();

  // Determine the dimensions of the copied data
  const numRows = valuesToCopy.length;
  const numCols = valuesToCopy[0] ? valuesToCopy[0].length : 0; // Handle empty source sheet

  // If there's no data, just create the sheet and exit
  if (numRows === 0 || numCols === 0) {
    return theResults(false, `Warning. Could not create a new sheet with the name '${newSheetName}'`, functionName);
  }

  // Create the new sheet
  let destinationSheet;
  try {
    destinationSheet = spreadsheet.insertSheet(newSheetName);
  } catch (e) {
    return theResults(false, `Error. Could not create a new sheet with the name '${newSheetName}'`, functionName);
  }
  destinationSheet.setFrozenRows(1);
  Logger.log(`${functionName}. Created new empty sheet with the name '${newSheetName}'`);

  // Move the new sheet all the way to the right.
  spreadsheet.setActiveSheet(destinationSheet);
  spreadsheet.moveActiveSheet(spreadsheet.getSheets().length);
  Logger.log(`${functionName}. moved '${newSheetName}' to the rightmost position`);

  // Paste values only to the new sheet
  const destinationRange = destinationSheet.getRange(1, 1, numRows, numCols);
  destinationRange.setValues(valuesToCopy);
  destinationRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  Logger.log(`${functionName}. '${newSheetName}'. New values copied into the new sheet.`);

  // --- Copy Column Widths ---
  for (let col = 1; col <= numCols; col++) {
    const width = sourceSheet.getColumnWidth(col);
    destinationSheet.setColumnWidth(col, width);
  }
  Logger.log(`${functionName}. '${newSheetName}'. Replicate the column widths from the source sheet.`);


  // --- Copy Row Heights ---
  //  for (let row = 1; row <= numRows; row++) {
  //    const height = sourceSheet.getRowHeight(row);
  //    destinationSheet.setRowHeight(row, height);
  //  }
  // ui.alert("'${newSheetName}' row heights updated");

  // --- NEW COMMANDS HERE: Sorting ---
  // Find column indices for sorting
  Logger.log(`${functionName}. '${newSheetName}'. Starting column sorting by '${dateHeader}' and '${topicHeader}'.`);
  const headers = valuesToCopy[0]; // Assuming headers are in the first row
  let dateColIndex = -1;
  let topicColIndex = -1;

  for (let i = 0; i < headers.length; i++) {
    const headerValue = headers[i] != null ? String(headers[i]).trim().toLowerCase() : '';
    if (headerValue === dateHeader.trim().toLowerCase()) {
      dateColIndex = i + 1; // Column index is 1-based
    }
    if (headerValue === topicHeader.trim().toLowerCase()) {
      topicColIndex = i + 1; // Column index is 1-based
    }
  }

  if (dateColIndex === -1) {
    Logger.log(`${functionName}. '${newSheetName}'. Warning: Could not find column with header '${dateHeader}'. Skipping date sort.`);
  }
  if (topicColIndex === -1) {
    Logger.log(`${functionName}. '${newSheetName}'. Warning: Could not find column with header '${topicHeader}'. Skipping topic sort.`);
  }

  // Only attempt to sort if at least one column was found and there's more than just the header row
  if ((dateColIndex !== -1 || topicColIndex !== -1) && numRows > 1) {
    const sortRange = destinationSheet.getRange(2, 1, numRows - 1, numCols); // Sorts data AFTER the header row

    const sortCriteria = [];
    if (dateColIndex !== -1) {
      sortCriteria.push({column: dateColIndex, ascending: true}); // Sort by Date, ascending
    }
    if (topicColIndex !== -1) {
      sortCriteria.push({column: topicColIndex, ascending: true}); // Then by Topic, ascending
    }

    if (sortCriteria.length > 0) {
      sortRange.sort(sortCriteria);
      Logger.log(`Contents sorted by ${dateColIndex !== -1 ? 'Date' : ''}${dateColIndex !== -1 && topicColIndex !== -1 ? ' then ' : ''}${topicColIndex !== -1 ? 'Topic' : ''}.`);
    }
  } else {
    Logger.log(`${functionName}. '${newSheetName}'. No data to sort (or only header row) or specified sort columns not found.`);
  }
  // ------------------------------------
  // Optional: Move the newly created sheet to be the active sheet and at the end
  spreadsheet.setActiveSheet(destinationSheet);
  //spreadsheet.moveActiveSheet(spreadsheet.getSheets().length);

  return theResults(true, `Values and dimensions from '${sourceSheetName}' have been copied to a new sheet: '${newSheetName}'!`, functionName);

};


/**
 * @function deleteSnapshotTabs
 * @description Scans all sheets in the active spreadsheet and permanently deletes any
 *              visible, unprotected sheet whose name begins with the prefix "Snapshot-".
 *              Iterates in reverse order to safely handle index shifting during deletion.
 *              Will not attempt to delete the last remaining visible sheet in the spreadsheet.
 *              Note: The prefix match is case-insensitive and ignores leading spaces.
 * @param snapshotPrefix
 * @returns {Result}
 */
const deleteSnapshotTabs = (snapshotPrefix) => {
  const functionName = 'deleteSnapshotTabs';
  Logger.log(`${functionName}. Started.`);

  // Validate snapshotPrefix parameter
  const validationResult = validateNonEmptyString(snapshotPrefix, 'snapshotPrefix');
  if (!validationResult.ok) {
    return theResults(false, validationResult.message, functionName);
  }
  // Use the validated and trimmed prefix value
  const validatedPrefix = validationResult.data;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Fetch the full list of sheets once at the start of the operation
  const sheets = spreadsheet.getSheets();


  // Counter to track how many sheets were successfully removed
  let sheetsDeletedCount = 0;

  // Iterate in reverse order to prevent index-shifting issues as sheets are removed
  for (let i = sheets.length - 1; i >= 0; i--) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    const sheetId = sheet.getSheetId(); // Unique, stable numeric identifier for this sheet

    // Skip any sheet that is currently hidden from the user
    if (sheet.isSheetHidden()) {
      Logger.log(`${functionName}. Skipping hidden sheet: '${sheetName}'.`);
      continue;
    }

    // Skip any sheet that has sheet-level protection applied
    if (sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0) {
      Logger.log(`${functionName}. Skipping protected sheet: '${sheetName}'.`);
      continue;
    }

    // Only process sheets whose names begin with the defined snapshot prefix (case-insensitive, ignoring leading spaces)
    Logger.log(`${functionName}. Process: '${sheetName}'.`);
    if (sheetName.trim().toLowerCase().startsWith(validatedPrefix.toLowerCase())) {

      // Guard: A spreadsheet must always retain at least one visible sheet.
      // Re-query visible sheets on each iteration to reflect any deletions already made.
      const visibleSheets = spreadsheet.getSheets().filter(s => !s.isSheetHidden());
      if (visibleSheets.length === 1 && visibleSheets[0].getSheetId() === sheetId) {
        Logger.log(`${functionName}. Skipping '${sheetName}': it is the last visible sheet and cannot be deleted.`);
        continue;
      }

      // Attempt to delete the qualifying sheet and report any failure back to the caller
      try {
        spreadsheet.deleteSheet(sheet);
        sheetsDeletedCount++;
        Logger.log(`${functionName}. Successfully deleted sheet: '${sheetName}' (id: ${sheetId}, original index: ${i}).`);
      } catch (e) {
        return theResults(false, `Error. Failed to delete sheet '${sheetName}' (id: ${sheetId}). ${e.message}`, functionName);
      }
    }
  }

  // All eligible sheets have been processed — return a success result with a summary
  Logger.log(`${functionName}. Finished. ${sheetsDeletedCount} sheet(s) were deleted.`);
  return theResults(true, `Deletions are complete. ${sheetsDeletedCount} sheet(s) were deleted.`, functionName);
};
