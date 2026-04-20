// deprecated.js
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
    //return failResult(`Duplicate attribute found: ${attributeName}`);
  }
  //return okResult('no duplicate found', { ...hash });
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
//return failResult('tuples must be an array');
  }
  if (!hash || typeof hash !== 'object') {
    //return failResult('hash must be an object');
  }
  if (!validators || !Array.isArray(validators)) {
   // return failResult('validators must be an array');
  }

  // --- Start with a clean copy of the incoming hash ---
  let currentHash = {...hash};

  // --- Outer loop: each tuple ---
  Logger.log(`>> Start Outer Loop`);
  for (const tuple of tuples) {

    Logger.log(`>> OuterLoop: tuple[0]=${tuple[0]}, tuple[1]=${tuple[1]}`);
    // Guard: make sure this tuple is usable
    if (!Array.isArray(tuple) || tuple.length < 2) {
     //return failResult(`Invalid tuple encountered: ${JSON.stringify(tuple)}`);
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
        //return failResult(result.message);
      }

      // Carry the updated hash forward to the next validator
      currentHash = {...result.data};
    }
  }
};

/* 1️⃣
 ***************************************************************************************************************************
 * Return document ID
 */
const createNewGoogleDoc = (docName = "tmpDoc202510081748356")  => {
  Logger.log(`createNewGoogleDoc: ${docName}`);
  try {
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    // Remove first child if present.
    if (body.getNumChildren() > 1) {
      body.removeChild(body.getChild(0));
    }

    //body.appendPageBreak();
    doc.saveAndClose();

    return {
      ok: true,
      id: doc.getId(),
      name: doc.getName(),
      message: `✅ Created document: ${doc.getName()} (ID: ${doc.getId()})`
    };

  } catch (err) {
    return {
      ok: false,
      error: err,
      id: null,
      name: docName,
      message: `❌ Failed to create Google Doc: ${docName} - ${err}`
    };
  }
};



/* 1️⃣
 * Converts a Google Doc to a DOCX file in Google Drive.
 * Functional idiom: Returns a result object with status and outcome.
 *
 * @param {string} docId - The ID of the Google Docs file to convert.
 * @returns {Object} An explicit result object:
 *   {
 *     ok: boolean,
 *     docxFile: DriveApp.File|null,
 *     docxId: string|null,
 *     docxUrl: string|null,
 *     docxName: string|null,
 *     message: string,
 *     error: any|null
 *   }
 */
const convertGoogleDocToDocx = (docId = '1fVUBmOKkUET3pEQgoZGhqXb26CLncSE7FQnjZtC4ak8') => {
  try {
    // Get the original file and its information
    //const docFile = DriveApp.getFileById(docId);
    const docFile = DriveApp.getFilesByName(docId);

    // Use Drive API to export as DOCX (Microsoft Word) format
    const apiUrl = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`;

    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    // Check for problems (like not a Google Doc, or bad permissions)
    if (response.getResponseCode() !== 200) {
      const message = `${convertGoogleDocToDocx.name}: Export failed: HTTP ${response.getResponseCode()} - ${response.getContentText()}`;
      Logger.log(message);
      return {
        ok: false,
        docxFile: null,
        docxId: null,
        docxUrl: null,
        docxName: null,
        message: message,
        error: response
      };
    }

    // Name and create .docx file in the same folder (or root, if no parent)
    const docxBlob = response.getBlob();
    const docxName = docFile.getName() + '.docx';
    const parentFolders = docFile.getParents();
    const parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
    const docxFile = parentFolder.createFile(docxBlob).setName(docxName);
    const message = `${convertGoogleDocToDocx.name}: DOCX file created successfully in Drive: ${docxFile.getName()}`;
    Logger.log(message);
    return {
      ok: true,
      docxFile,
      docxId: docxFile.getId(),
      docxUrl: docxFile.getUrl(),
      docxName,
      message: message,
      error: null
    };
  } catch (err) {
    const message = `${convertGoogleDocToDocx.name}: Error during conversion: ${err}`;
    Logger.log(message);
    return {
      ok: false,
      docxFile: null,
      docxId: null,
      docxUrl: null,
      docxName: null,
      message: message,
      error: err
    };
  }
};



/* 1️⃣
 * Creates a copy of a Google Doc by ID.
 *
 * @param {string} docId - The ID of the source Google Doc.
 * @returns {Object} An explicit result object:
 *   {
 *     ok: boolean,
 *     newDocFile: DriveApp.File|null,
 *     newDocId: string|null,
 *     newDocUrl: string|null,
 *     newDocName: string|null,
 *     message: string,
 *     error: any|null
 *   }
 */
const copyGoogleDocById = (docId = '1fVUBmOKkUET3pEQgoZGhqXb26CLncSE7FQnjZtC4ak8') => {
  try {
    const srcFile = DriveApp.getFileById(docId);
    const newFileName = srcFile.getName() + ' (Copy)';
    const parentFolders = srcFile.getParents();
    const parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

    // Create the copy
    const newDocFile = srcFile.makeCopy(newFileName, parentFolder);
    Logger.log(`Copy created: ${newDocFile.getName()}`);

    return {
      ok: true,
      newDocFile,
      newDocId: newDocFile.getId(),
      newDocUrl: newDocFile.getUrl(),
      newDocName: newDocFile.getName(),
      message: `Copy created: ${newDocFile.getName()}`,
      error: null
    };
  } catch (err) {
    Logger.log(`Error creating document copy: ${err}`);
    return {
      ok: false,
      newDocFile: null,
      newDocId: null,
      newDocUrl: null,
      newDocName: null,
      message: `Error creating document copy: ${err}`,
      error: err
    };
  }
};



/* 1️⃣ Experiment
 * Replaces in the top-left cell of every table in a document body (in-place).
 * Returns an object describing what was changed.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @returns {Object} result
 *   {number[]} changedIndices - Array of table indices that were changed.
 *   {number} changedCount - How many tables were updated.
 *   {number} tableCount - Total tables processed.
 *   {boolean} ok - True if any changes were made.
 */
const replaceCharInTablesInPlace = ( body, findChar = '🟪', replaceWithChar = '✔️') => {
  const tables = body.getTables();
  const changedIndices = [];
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (table.getNumRows() > 0 && table.getRow(0).getNumCells() > 0) {
      const cell = table.getCell(0, 0);
      const text = cell.getText();
      if (text.includes(findChar)) {
        // Simplistic replace. Will replace all instances
        const regex = new RegExp(findChar, 'g');
        cell.setText(text.replace(regex, replaceWithChar));
        changedIndices.push(i);
      }
    }
  }
  return {
    ok: changedIndices.length > 0,
    changedIndices,
    changedCount: changedIndices.length,
    tableCount: tables.length,
    message: changedIndices.length > 0
      ? `Changed ${changedIndices.length} table(s).`
      : `No checkmarks found in any table(s).`
  }
};





  // Logger.log(`extractTableRows; rowsData: ${rowsData.length}`);
  //Logger.log(`extractTableRows; Extracted Rows: ${JSON.stringify(rowsData, null, 2)}`);
  // Logger.log('extractTableRows;  ===================================================================');
  // Logger.log(`extractTableRows; rowsData[0,0]: ${rowsData[0,0]}`);
  // Logger.log(`extractTableRows; rowsData[0,1]: ${rowsData[0,1]}`);
  // Logger.log(`extractTableRows; rowsData[0,1]: ${rowsData[0,2]}`);
  // Logger.log('extractTableRows;  ===================================================================');





/* 1️⃣
 ******************************************************************************************************************
 *
 */
const shouldCopyTable = table => {
  const checkedCheckboxChar = '✔';
  let ok = table.getType() === DocumentApp.ElementType.TABLE;
  if (ok) {
    const hasFirstRowCell = ok && table.getRow(0).getNumCells() > 0;
    const hasRows = table.getNumRows() > 0;
    const topLeft = hasRows ? table.getCell(0, 0).getText().trim() : '';
    ok = hasRows && hasFirstRowCell && topLeft.includes(checkedCheckboxChar);
  }

  return {
    ok,
    reason: ok ? 'Checked checkbox found' : 'No checked checkbox or no rows',
    table
  };
};


/* 1️⃣
 ******************************************************************************************************************
 *
 */
const copyCheckedTables = (tables, body) => {
  const checkResults = tables.map(shouldCopyTable);
  const copied = [];
  const skipped = [];

  checkResults.forEach(result => {
    if (result.ok) {
      try {
        // Already validated by shouldCopyTable, but could check again for robustness
        const copy = result.table.copy();
        body.appendTable(copy);
        body.appendParagraph('');
        Logger.log(`Appended table: ${copy.getText()}`);
        copied.push(copy);
      } catch (error) {
        Logger.log(`Failed to append a checked table: ${error}`);
        skipped.push(result.table); // Optional: could mark as error-skipped
      }
    } else {
      skipped.push(result.table);
    }
  });

  return {
    ok: copied.length > 0,
    copied,
    skipped,
    message: `${copied.length} table(s) copied, ${skipped.length} table(s) skipped.`
  };
};

/* 1️⃣
 * Appends tables to the first tab (default) of the specified Google Doc.
 * Returns an explicit result object.
 *
 * @param {Table[]} tablesToCopy - Tables to append.
 * @param {string} googleDocID - ID of the target Doc.
 * @returns {Object} Functional idiom result object.
 */
const saveBodyTablesToFirstTabInNewDoc = (tablesToCopy, googleDocID) => {
  try {
    const doc = DocumentApp.openById(googleDocID);

    // Find the tabs in this doc
    const tabs = doc.getTabs();
    if (!tabs || tabs.length === 0) {
      return {
        ok: false,
        copied: [],
        skipped: tablesToCopy,
        googleDocID,
        docUrl: `https://docs.google.com/document/d/${googleDocID}`,
        message: `${saveBodyTablesToFirstTabInNewDoc.name}: No tabs found in doc id: ${googleDocID}.`,
      };
    }

    // Use the first tab (default created tab)
    const firstTab = tabs[0];
    const tabBody = firstTab.asDocumentTab().getBody();

    // Actually copy and append tables
    const copyResult = copyCheckedTables(tablesToCopy, tabBody);
    doc.saveAndClose();

    return {
      ...copyResult,
      googleDocID,
      docUrl: `https://docs.google.com/document/d/${googleDocID}`,
      message: copyResult.ok
        ? `${saveBodyTablesToFirstTabInNewDoc.name}: Successfully copied tables to the first tab of doc id: ${googleDocID}`
        : `${saveBodyTablesToFirstTabInNewDoc.name}: No tables copied to first tab in doc id: ${googleDocID}.`
    };
  } catch (err) {
    return {
      ok: false,
      copied: [],
      skipped: tablesToCopy,
      googleDocID,
      docUrl: `https://docs.google.com/document/d/${googleDocID}`,
      message: `Exception occurred: ${err}`,
      error: err
    };
  }
};


/**
 * Extracts text from each cell of a table (returns 2D array of cell texts)
 * @param {GoogleAppsScript.Document.Table} table
 * @returns {Object} { ok, tableTexts, messages }
 */
const extractTableTexts = table => {
  try {
    const numRows = table.getNumRows();
    const tableTexts = [];
    for (let r = 0; r < numRows; r++) {
      const row = table.getRow(r);
      const numCells = row.getNumCells();
      const rowTexts = [];
      for (let c = 0; c < numCells; c++) {
        const cell = table.getCell(r, c);
        const result = extractTableCellText(cell);
        rowTexts.push(result.ok ? result.text : null);
      }
      tableTexts.push(rowTexts);
    }
    return {
      ok: true,
      tableTexts,
      message: 'Extracted all cell texts from table'
    };
  } catch (err) {
    return {
      ok: false,
      tableTexts: null,
      message: `Error extracting table text: ${err}`,
      error: err
    };
  }
};

/**
 * Extracts text from all tables (list of 2D arrays of cell texts)
 * @param {GoogleAppsScript.Document.Table[]} tables
 * @returns {Object} { ok, allTablesTexts, messages }
 */
const extractAllTablesTexts = tables => {
  try {
    const allResults = tables.map(extractTableTexts);
    const ok = allResults.every(res => res.ok);
    const messages = allResults.map((r, i) => `Table ${i}: ${r.message || ''}`);
    return {
      ok,
      allTablesTexts: allResults.map(res => res.tableTexts),
      messages,
    };
  } catch (err) {
    return {
      ok: false,
      allTablesTexts: null,
      messages: [`Error extracting texts for tables: ${err}`],
      error: err
    };
  }
};





/*
Extracts cell text from all rows (skipping the header) of a Google Doc table.
Fully functional implementation using map and slice.

@param {DocumentApp.Table} table The Google Doc Table element to process.
@returns {Array<Array<string>>} An array of rows, where each row is an array of cell text strings.
*/
const extractTableRows_OLD = table => {
  const numRows = table.getNumRows();

  Logger.log(`extractTableRows; numTables: ${numRows}`);
  Logger.log(`extractTableRows; numTables: ${numRows}`);

  // 1. Create an array of row indices starting from 1 (to skip the header at index 0).
  //    The 'slice(1)' pattern is the most functional way to skip the first element.
  // 2. Filter: Skip the header row (index 0).
  // 3. Map: Transform each row index into an array of cell strings.
  const rowIndices = Array.from({length: numRows}, (_, i) => i); // [0, 1, 2, 3, ...]
  const dataRowIndices = rowIndices.slice(1); // [1, 2, 3, ...]
  Logger.log(`extractTableRows; rowIndices: ${rowIndices.length}, dataRowIndices: ${dataRowIndices.length}`);

  const rowsData = dataRowIndices.map(r => {
    const row = table.getRow(r);
    const numCells = row.getNumCells();

    // Create an array of cell indices for the current row
    const cellIndices = Array.from({length: numCells}, (_, c) => c);
    Logger.log(`extractTableRows.dataRowIndices.map; numCells: ${numCells}`);
    Logger.log(`extractTableRows.dataRowIndices.map; cellIndices: ${cellIndices}`);

    // Map: Transform each cell index into the cell's trimmed text content
    const rowData = cellIndices.map(c => {
      const cell = row.getCell(c);
      //Logger.log(`extractTableRows; cell: ${cell.getText().trim()}`);
      return cell.getText().trim();
    });

    // Logger.log('extractTableRows.dataRowIndices.map;  --------------------------------------------------------------------');
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[0] ${rowData[0]}`);
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[1] ${rowData[1]}`);
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[2] ${rowData[2]}`);
    // Logger.log('extractTableRows.dataRowIndices.map;  --------------------------------------------------------------------');

    return rowData;
  });

};


/**
 * Takes in an array of tables, each representing a dailies table entry for a particular day.
 * Called from menu.js when the user selects 'Import Completed Dailies'.
 *
 */
const extractTablesRows_OLD = (tablesToProcessAsArray) => {
  //
  Logger.log(`extractTablesRows; num tablesToProcess: ${tablesToProcessAsArray.length}`);

  // The input to this step is an array of google Table objects
  // the result of this step is an arrayX of arraysY where
  // each of the arrayY's in that description is the result of converting a
  //   google doc table into rows of data
  const allTablesRowsNested = tablesToProcessAsArray.map(extractTableRows);

  // the result of this step is
  //   An array of rows.  This array is a concat of each of the array of rows from table in the array of tables.
  //   each entry in the array of rows continas
  //      An array of cells from that row.
  const allTablesRowsFlat = allTablesRowsNested.flat();

  Logger.log(`extractTablesRows; Rows[]: ${allTablesRowsFlat.length}`);
  Logger.log('extractTablesRows; --------------------------------------------------------------------');
  Logger.log(`extractTablesRows; allTablesRowsFlat[0]`);
  Logger.log(`extractTablesRows; ${allTablesRowsFlat[0]}`);
  Logger.log('extractTablesRows; --------------------------------------------------------------------');
  Logger.log(`extractTablesRows; allTablesRowsFlat[1]`);
  Logger.log(`extractTablesRows; ${allTablesRowsFlat[1]}`);
  Logger.log('extractTablesRows; --------------------------------------------------------------------');

  //return okResult(
    `extractTablesRows: Successfully extracted`,
    allTablesRowsFlat
  );
};

/**
 * Filters tables by the unchecked checkbox character, extracts all data
 * rows from each marked table, and returns them as a single flat array.
 *
 * @param {Array<DocumentApp.Table>} tables               - All tables from the doc tab
 * @param {string}                   unCheckedCheckboxChar - Character marking tables for transfer
 *
 * @returns {{ ok: boolean, message: string, data: Array<Array<string>>|null }}
 *   data = flat array of all rows extracted from all marked tables

 */
const convertDocTablestoData = (tables) => {
  Logger.log(`Entered: convertDocTablestoData`);

  // --- Guard: validate inputs ---
  if (!tables || !Array.isArray(tables)) {
    //return failResult('extractTablesRows: tables must be an array');
  }
    // --- Step 2: Collect all rows from marked tables ---
  const collectResult = collectRowsFromTables(tables);
  if (!collectResult.ok) {
    //return failResult(collectResult.message);
  }

  Logger.log(`extractTablesRows: total rows=${collectResult.data.length}`);
  //return okResult(
    `extractTablesRows: Successfully extracted ${collectResult.data.length} row(s) from ${tables.length} table(s)`,
    collectResult.data
  );
};


/**
 * Extracts cell text from all rows (skipping the header) of a Google Doc table.
 * Fully functional implementation using map and slice.
 *
 * @param {DocumentApp.Table} table The Google Doc Table element to process.
 * @returns {Array<Array<string>>} An array of rows, where each row is an array of cell text strings.
 */
const extractTableRows_OLD = table => {

  const numRows = table.getNumRows();
  Logger.log(`extractTableRows; numTables: ${numRows}`);

  // Let's say the daily table has 5 rows so numRows = 5.  The first row is the header.  The other 4 rows are the notes,
  //  one row per person
  // `Array.from({length: 5}, (_, i) => i)` is just a way of saying **"give me an array of numbers from 0 to 4"**.
  // The `_` means "I don't care about this argument, I only want the index `i`".
  const rowIndices = Array.from({length: numRows}, (_, i) => i); // [0, 1, 2, 3, ...]

  //`slice(1)` means **"give me everything from position 1 onwards"** — which drops Row 0 (the 🟪 header row).
  const dataRowIndices = rowIndices.slice(1); // [1, 2, 3, ...]

  Logger.log(`extractTableRows; rowIndices: ${rowIndices.length}, dataRowIndices: ${dataRowIndices.length}`);

  // `map` loops over `[1, 2, 3, 4]` and transforms each index into an array of cell strings.
  const rowsData = dataRowIndices.map(r => {
    const row = table.getRow(r);
    const numCells = row.getNumCells();

    // Create an array of cell indices for the current row
    const cellIndices = Array.from({length: numCells}, (_, c) => c);
    Logger.log(`extractTableRows.dataRowIndices.map; numCells: ${numCells}`);
    Logger.log(`extractTableRows.dataRowIndices.map; cellIndices: ${cellIndices}`);

    // Map: Transform each cell index into the cell's trimmed text content
    const rowData = cellIndices.map(c => {
      const cell = row.getCell(c);
      //Logger.log(`extractTableRows; cell: ${cell.getText().trim()}`);
      return cell.getText().trim();
    });

    // Logger.log('extractTableRows.dataRowIndices.map;  --------------------------------------------------------------------');
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[0] ${rowData[0]}`);
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[1] ${rowData[1]}`);
    // Logger.log(`extractTableRows.dataRowIndices.map; rowData[2] ${rowData[2]}`);
    // Logger.log('extractTableRows.dataRowIndices.map;  --------------------------------------------------------------------');

    return rowData;
  });

};


/*
  // const firstTable = tablesSubset(tables, checkedCheckboxChar )[0];
  //const rows =   extractTableRows(firstTable);
  if ( appendTableRowsToSheet
       (extractTablesRows
          (tables,unCheckedCheckboxChar),sheetID,sheetTabTitle
        )
      )
  {
    const replacedChar = replaceCharInTablesInPlace(sub.tab.asDocumentTab().getBody(),unCheckedCheckboxChar,checkedCheckboxChar)


    Logger.log(`${replacedChar.message}`);
  } else {
    Logger.log(`No rows were appended and no tables were marked as complete}`);
  }
*/
