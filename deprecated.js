// Test comment
// ***********************************************************************************************************************************************************************************************
function createTimestampedSnapshot() {
  // --- Configuration Variables ---
  const SOURCE_SHEET_NAME = "Current Journal Snapshot";      // Set your source sheet name here
  const NEW_SHEET_NAME_PREFIX = "Snapshot-";
  const DATE_HEADER = "Date";    // Exact header name for your date column in tabX
  const TOPIC_HEADER = "Topic";  // Exact header name for your topic column in tabX

  // -------------------------------

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
  const ui = SpreadsheetApp.getUi(); // Get the UI object for alerts

  // Error handling for missing source sheet
  if (!sourceSheet) {
    ui.alert("Error", `Source sheet '${SOURCE_SHEET_NAME}' not found. Please check the sheet name.`, ui.ButtonSet.OK);
    return;
  }

  // Generate the timestamp for the new sheet name (YYYYMMDD_hhmmss)
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');

  const timestamp = `${year}${month}${day}_${hours}${minutes}${seconds}`;
  const newSheetName = NEW_SHEET_NAME_PREFIX + timestamp;

  // Create the new sheet
  let destinationSheet;
  try {
    destinationSheet = spreadsheet.insertSheet(newSheetName);
  } catch (e) {
    ui.alert("Error", `Could not create a new sheet with the name '${newSheetName}'. Error: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  destinationSheet.setFrozenRows(1);
  //  ui.alert("Success", `Created new empty sheet with the name '${newSheetName}'`, ui.ButtonSet.OK);

  // Get the data range from the source sheet
  const sourceRange = sourceSheet.getDataRange();
  const valuesToCopy = sourceRange.getValues();

  // Determine the dimensions of the copied data
  const numRows = valuesToCopy.length;
  const numCols = valuesToCopy[0] ? valuesToCopy[0].length : 0; // Handle empty source sheet

  // If there's no data, just create the sheet and exit
  if (numRows === 0 || numCols === 0) {
    ui.alert("Success", `Created empty snapshot sheet: '${newSheetName}' (Source sheet was empty).`, ui.ButtonSet.OK);
    spreadsheet.setActiveSheet(destinationSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getSheets().length);
    return;
  }

  // Paste values only to the new sheet
  const destinationRange = destinationSheet.getRange(1, 1, numRows, numCols);
  destinationRange.setValues(valuesToCopy);
  destinationRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);


  // --- Copy Column Widths ---
  for (let col = 1; col <= numCols; col++) {
    const width = sourceSheet.getColumnWidth(col);
    destinationSheet.setColumnWidth(col, width);
  }


  // --- Copy Row Heights ---
  //  for (let row = 1; row <= numRows; row++) {
  //    const height = sourceSheet.getRowHeight(row);
  //    destinationSheet.setRowHeight(row, height);
  //  }
  // ui.alert("'${newSheetName}' row heights updated");

  // --- NEW COMMANDS HERE: Sorting ---
  // Find column indices for sorting
  const headers = valuesToCopy[0]; // Assuming headers are in the first row
  let dateColIndex = -1;
  let topicColIndex = -1;

  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === DATE_HEADER.toLowerCase()) {
      dateColIndex = i + 1; // Column index is 1-based
    }
    if (headers[i].toString().trim().toLowerCase() === TOPIC_HEADER.toLowerCase()) {
      topicColIndex = i + 1; // Column index is 1-based
    }
  }

  if (dateColIndex === -1) {
    Logger.log(`Warning: Could not find column with header '${DATE_HEADER}'. Skipping date sort.`);
  }
  if (topicColIndex === -1) {
    Logger.log(`Warning: Could not find column with header '${TOPIC_HEADER}'. Skipping topic sort.`);
  }

  // Only attempt to sort if at least one column was found and there's more than just the header row
  if ((dateColIndex !== -1 || topicColIndex !== -1) && numRows > 1) {
    const sortRange = destinationSheet.getRange(2, 1, numRows - 1, numCols); // Sorts data AFTER the header row

    const sortCriteria = [];
    if (dateColIndex !== -1) {
      sortCriteria.push({ column: dateColIndex, ascending: true }); // Sort by Date, ascending
    }
    if (topicColIndex !== -1) {
      sortCriteria.push({ column: topicColIndex, ascending: true }); // Then by Topic, ascending
    }

    if (sortCriteria.length > 0) {
      sortRange.sort(sortCriteria);
      Logger.log(`Contents sorted by ${dateColIndex !== -1 ? 'Date' : ''}${dateColIndex !== -1 && topicColIndex !== -1 ? ' then ' : ''}${topicColIndex !== -1 ? 'Topic' : ''}.`);
    }
  } else {
      Logger.log("No data to sort (or only header row) or specified sort columns not found.");
  }
  // ------------------------------------
  // Optional: Move the newly created sheet to be the active sheet and at the end
  spreadsheet.setActiveSheet(destinationSheet);
  //spreadsheet.moveActiveSheet(spreadsheet.getSheets().length);

  ui.alert("Success", `Values and dimensions from '${SOURCE_SHEET_NAME}' have been copied to a new sheet: '${newSheetName}'!`, ui.ButtonSet.OK);
}

/**
 * Fails if the attribute already has a defined value in the hash.
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeUnique_OLD = (tuple, hash) => {
  const attributeName = tuple[0];
  Logger.log(`checkIsAttributeUnique. testing  attributeName: ${attributeName} against hash contents: ${hash[attributeName]}`);

  if (hash[attributeName] !== undefined) {
    return failResult(`Duplicate attribute found: ${attributeName}`);
  }
  return okResult('no duplicate found', { ...hash });
};

/**
 * Processes an array of tuples through a pipeline of validator functions.
 * Each validator function receives the current attribute string and the
 * current state of the hash, and returns a result object with an updated
 * copy of the hash as data.
 *
 * Pre-conditions:
 *   - tuples is an array of [attributeString, valueString] pairs
 *   - hash has predefined keys with undefined values
 *   - validators is an array of functions with signature:
 *       (attributeString: string, hash: Object) => { ok, message, data: updatedHash }
 *
 * Skip conditions (handled before validator pipeline):
 *   - attributeString is empty or blank
 *   - attributeString is the word 'comment' (case insensitive)
 *
 * @param {Array<Array<string>>} tuples      - Array of [attributeString, valueString] pairs
 * @param {Object}               hash        - Predefined keys with undefined values
 * @param {Array<Function>}      validators  - Array of validator functions
 *
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = final state of hash after all tuples processed
 */
const processTuplesThroughValidators_OLD = (tuples, hash, validators) => {
  Logger.log(`Entered: processTuplesThroughValidators`);
  // --- Guard: validate inputs ---
  if (!tuples || !Array.isArray(tuples)) {
    return failResult('tuples must be an array');
  }
  if (!hash || typeof hash !== 'object') {
    return failResult('hash must be an object');
  }
  if (!validators || !Array.isArray(validators)) {
    return failResult('validators must be an array');
  }

  // --- Start with a clean copy of the incoming hash ---
  let currentHash = {...hash};

  // --- Outer loop: each tuple ---
  Logger.log(`>> Start Outer Loop`);
  for (const tuple of tuples) {

    Logger.log(`>> OuterLoop: tuple[0]=${tuple[0]}, tuple[1]=${tuple[1]}`);
    // Guard: make sure this tuple is usable
    if (!Array.isArray(tuple) || tuple.length < 2) {
      return failResult(`Invalid tuple encountered: ${JSON.stringify(tuple)}`);
    }

    // --- Skip check: runs before the validator pipeline ---
    const skipCheck = okToSkip(tuple, currentHash);
    Logger.log(`SkipOK? ${skipCheck.ok}`);
    if (skipCheck.ok) {
      continue;  // move to next tuple cleanly
    }

    // --- Inner loop: each validator ---
    for (const validator of validators) {

      Logger.log(`>> >> Inner Loop, Running validator: ${validator.name} for tuple: ${JSON.stringify(tuple)}`);
      const result = validator(tuple, currentHash);
      Logger.log(`>> >> Inner Loop, ${validator.name} result:${result.ok} - ${result.message}`);

      // Any validator failure stops everything
      if (!result.ok) {
        return failResult(result.message);
      }

      // Carry the updated hash forward to the next validator
      currentHash = {...result.data};
    }
  }
};

