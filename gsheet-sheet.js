// gsheet-sheet.js

// ***********************************************************************************************************************************************************************************************
function extractURLs(cell) {
  
 var text = cell.getDisplayValue();
 var urlRegex = /(https?:\/\/[^\s]+)/g;
 var urls = text.match(urlRegex);
 return urls;
}

// ***********************************************************************************************************************************************************************************************
function updateCells(cellAddress) {
// Open the active sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// Update cell A1
sheet.getRange(cellAddress).setValue("Hello world!");
}

// ***********************************************************************************************************************************************************************************************
function COMPAREANDAPPENDTEXT(range1, range2, iValue) {
  //=COMPAREANDAPPENDTEXT(C4:BN4,C1:BN1)  
  if (!range1 || !range2) return '';

  // Flatten in case the ranges are 1 row × N or N × 1
  var flat1 = range1.flat();
  var flat2 = range2.flat();

  if (flat1.length !== flat2.length) {
  throw new Error('Both ranges must have the same number of cells');
  }

  var output = [];

  for (var i = 0; i < flat1.length; i++) {
    if (flat1[i] === iValue || flat1[i] === iValue.toString()) {
    output.push('"' + flat2[i] + '"');
    }
  }

  return output.join(',');
}

// ***********************************************************************************************************************************************************************************************
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

// ***********************************************************************************************************************************************************************************************
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
 * * @param {Array<Array<string>>} extractedRows - The array of rows to write.
 * @param {string} sheetID - The ID of the target Google Sheet.
 * @param {string} targetSheetName - The name of the specific tab (worksheet) to append to.
 * @returns {boolean} True if the append was successful, false otherwise.
 */
const appendTableRowsToSheet = ( extractedRows, sheetID, targetSheetName ) => {
  const functionName = 'appendTableRowsToSheet';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

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
    var startRow = sheet.getLastRow() + 1;
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
 * @param sheetID defined within config.js. Accessed via getRequiredInitInfo().sheetID
 * @param tabName defined within config.js. Accessed via getRequiredInitInfo().sheetTabTitle
 * @param numRows defined within config.js. Accessed via getRequiredInitInfo().maxInputRows
 * @param requiredAttribsHash defined within config.js. Accessed via getConfigHash().
 *        This is the initialization hash that defines the required attributes.
 *        It never gets values for it's keys.
 * @param validatorFunctions - List of functions that validate the input data.
 * @returns {Result} - An object containing the result of the operation.
 */
const readInput = (sheetID, tabName, numRows, requiredAttribsHash, validatorFunctions) => {
  const functionName = 'readInput';
  Logger.log(`${functionName}. Started.`);
  let returnResult;


  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetID);
  } catch (err) {
    return theResults(false, ` Could not open spreadsheet with id: ${sheetID} - ${err.message}`, functionName);
  }

  // --- Find the tab ---
  const sheet = spreadsheet.getSheetByName(tabName);
  if (!sheet) {
    return theResults(false, `Could not find tab named: ${tabName} in spreadsheet: ${sheetID}`, functionName);
  }

  // --- Read the rows ---
  let rawRows;
  try {
    rawRows = sheet.getRange(1, 1, numRows, 2).getValues();
  } catch (err) {
    return theResults(false, ` Could not read rows from tab: ${tabName} - ${err.message}`, functionName);
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
  //TODO: remove all ui.alerts from anything but workflow
  if (!returnResult.ok) {
    return theResults(false, returnResult.message, functionName);
  }

  return theResults(true, 'Completed.', functionName, returnResult.data);
};

