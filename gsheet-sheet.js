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
 * @param {string} sheetId - The ID of the target Google Sheet.
 * @param {string} targetSheetName - The name of the specific tab (worksheet) to append to.
 * @returns {boolean} True if the append was successful, false otherwise.
 */
const appendTableRowsToSheet = ( extractedRows, sheetId, targetSheetName ) => {
  if (extractedRows.length === 0 || extractedRows[0].length === 0) {
    Logger.log('No valid data rows to append.');
      return failResult(
        `No valid data rows to append.`
      );
  }


  Logger.log(`appendTableRowsToSheet: extractedRows.length: ${extractedRows.length}`);

  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    
    // ➡️ Change is here: Use getSheetByName() instead of getActiveSheet()
    const sheet = spreadsheet.getSheetByName(targetSheetName); 

    if (!sheet) {
        Logger.log(`appendTableRowsToSheet: ERROR: Could not find sheet tab named: ${targetSheetName}`);
        return failResult(
          `appendTableRowsToSheet: ERROR: Could not find sheet tab named: ${targetSheetName}`
        );
    }

    // Determine the starting row for the new data
    var startRow = sheet.getLastRow() + 1;
    
    // ➡️ NEW: Create modified rows with prepended row numbers
    const modifiedRows = extractedRows.map((row, index) => {
      const rowNumber = startRow + index;
      return [rowNumber, ...row]; // Prepend row number to each row
    });

    const numRows = modifiedRows.length;
    const numColumns = modifiedRows[0].length; // Now includes the extra column
    
    Logger.log(`appendTableRowsToSheet: numRows: ${numRows}, numColumns: ${numColumns}, startRow: ${startRow}`);
    Logger.log(`appendTableRowsToSheet: modifiedRows[0]: ${modifiedRows[0]}`);
    
    // Define the target range (starting at column 1, now with extra column)
    const range = sheet.getRange(startRow, 1, numRows, numColumns);

    // Append all rows at once for optimal performance
    range.setValues(modifiedRows);

    Logger.log(`appendTableRowsToSheet: Successfully appended ${numRows} rows to tab: ${targetSheetName}`);
    return okResult(
          `appendTableRowsToSheet: Successfully appended ${numRows} rows to tab: ${targetSheetName}`
        );

  } catch (error) {
    Logger.log(`appendTableRowsToSheet: ERROR: Failed to open sheet or append data: ${error}`);
    return failResult(
        `appendTableRowsToSheet: ERROR: Failed to open sheet or append data: ${error}`
      );
  }
}

/**
 * Reads up to maxRows rows from a named tab in a Google Sheet and returns
 * the contents as an array of tuples.
 * Each tuple is [columnA, columnB] representing one row.
 *
 * Pre-conditions:
 *   - sheetId is a valid Google Sheet ID
 *   - tabName is the name of an existing tab in that sheet
 *   - maxRows is a positive integer
 *
 * @param {string} sheetId  - The ID of the Google Sheet
 * @param {string} tabName  - The name of the tab to read from
 * @param {number} maxRows  - Maximum number of rows to read
 *
 * @returns {{ ok: boolean, message: string, data: Array<Array<string>>|null }}
 *   data = array of tuples [ [attributeName, value], ... ]
 */
const readTabAsTuples = (
  sheetId,
  tabName,
  maxRows
) => {
  Logger.log(`Entered: readTabAsTuples sheetId=${sheetId} tabName=${tabName} maxRows=${maxRows}`);

  // --- Find the spreadsheet ---
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    return failResult(`readTabAsTuples: Could not open spreadsheet with id: ${sheetId} - ${err.message}`);
  }

  // --- Find the tab ---
  const sheet = spreadsheet.getSheetByName(tabName);
  if (!sheet) {
    return failResult(`readTabAsTuples: Could not find tab named: ${tabName} in spreadsheet: ${sheetId}`);
  }

  // --- Read the rows ---
  let rawRows;
  try {
    rawRows = sheet.getRange(1, 1, maxRows, 2).getValues();
  } catch (err) {
    return failResult(`readTabAsTuples: Could not read rows from tab: ${tabName} - ${err.message}`);
  }

  // --- Convert to tuples ---
  // getValues() returns an array of arrays already
  // We just need to make sure values are strings
  const tuples = rawRows.map(row => [
    row[0].toString(),
    row[1].toString(),
  ]);



  Logger.log(`readTabAsTuples: read ${tuples.length} tuples from ${tabName}`);
  return okResult(`Successfully read ${tuples.length} rows from ${tabName}`, tuples);
};

