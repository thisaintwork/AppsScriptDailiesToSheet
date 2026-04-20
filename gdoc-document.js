// gdoc-document.js
/* 1️⃣
 *************************************************************************************************************************** 
 * Searches an array of DocumentApp.Tab objects for a tab with a specific title.
 * @param {DocumentApp.Tab[]} tabs The array of tabs from the Google Doc.
 * @param {string} targetTitle The title of the tab to find.
 * @returns {DocumentApp.Tab|null} The DocumentApp.Tab object if found, otherwise null.
 */
const getTabByTitle = (tabs, targetTitle) => {
  const functionName = 'getTabByTitle';
  Logger.log(`${functionName} Started.`);

  // Use Array.prototype.find: declarative, no mutation
  const tab = tabs.find(tab => tab.getTitle() === targetTitle);

  // Return a result object for clarity (functional idiom)
  return tab
    ? { ok: true, tab }
    : { ok: false, tab: null };
};

/* 1️⃣
 ******************************************************************************************************************
 * Returns a filtered subset of table elements whose top-left cell contains a given checkbox character.
 *
 * @param {GoogleAppsScript.Document.Body[]|GoogleAppsScript.Document.Table[]} tables
 *   An array of DocumentApp elements expected to contain table elements.
 * @param {string} checkboxChar
 *   The character or string that identifies a “selected” or “target” table
 *   when present in the top-left cell (row 0, column 0).
 *
 * @returns {GoogleAppsScript.Document.Table[]}
 *   An array of table elements that:
 *   - Are actually tables,
 *   - Contain at least one row and one cell,
 *   - And whose top-left cell text includes `checkboxChar`.
 */
const tablesSubset = (tables, checkboxChar) =>
  tables.filter(table => {
    // Guard: skip non-table elements
    if (table.getType() !== DocumentApp.ElementType.TABLE) return false;

    // Guard: skip empty tables (no rows)
    if (table.getNumRows() === 0) return false;

    // Guard: skip tables whose first row has no cells
    if (table.getRow(0).getNumCells() === 0) return false;

    // Get the text from the top-left cell (row 0, column 0)
    const topLeft = table.getCell(0, 0).getText().trim();

    // A table belongs to the subset if its top-left text contains the checkbox character
    return topLeft.includes(checkboxChar);
  });

/**
 * Extracts trimmed text from each cell in a single table row.
 *
 * @param {DocumentApp.TableRow} row - A single row from a Google Doc table
 * @returns {Array<string>} Array of trimmed cell text strings
 */
const extractCellsFromTableRow = (row) => {
  const numCells = row.getNumCells();
  const cellIndices = Array.from({length: numCells}, (_, c) => c);
  return cellIndices.map(c => row.getCell(c).getText().trim());
};

/**
 * Extracts cell text from all rows (skipping the header) of a Google Doc table.
 *
 * @param {DocumentApp.Table} table - The Google Doc Table element to process.
 * @returns {Array<Array<string>>} Array of rows, where each row is an array of cell text strings
 */
const extractRowDataFromOneTable = (table) => {
  const numRows = table.getNumRows();

  // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/from#element
  // e.g Array.from([11, 2, 5], (n,x) => n + x);
  //     Expected output: Array [11, 3, 7]
  // e.g Array.from([1, 2, 3], (i,x) => x + i + 5);
  //     Array [6, 8, 10]
  // e.g Array.from([1, 2, 3], (i) => i + 5);  Note- this time I am only giving the element
  // e.g Array.from({length: 5}, (_, i) => i);  Note- this time even though I am giving the element and the index,
  //     The element variable is ignored since there aren't any elements
  // Array.from() never creates a sparse array. If the items object is missing some index properties, they become undefined in the new array.
  const rowIndices = Array.from({length: numRows}, (_, i) => i);  // rowIndices[0]=0
  const dataRowIndices = rowIndices.slice(1);                       // dataRowIndices[0]=1
  // Logger.log(`extractTableRows: numRows=${numRows}, dataRows=${dataRowIndices.length}`);

  // take a table and map each of its row data as an entry in rowsFromOneTable
  const rowsFromOneTable = dataRowIndices.map(r => extractCellsFromTableRow(table.getRow(r)));
  return rowsFromOneTable;
  };

/**
 * Extracts rows from each table in the array and flattens
 * them into a single array of rows.
 * Each table is processed by extractTableRows.
 *
 * @param {Array<DocumentApp.Table>} tables - Array of Table objects to process
 *
 * @returns {{ ok: boolean, message: string, data: Array<Array<string>>|null }}
 *   data = flat array of all rows from all tables, header rows excluded
 *   e.g. data[0] = ["Alice", "30", "Engineer"];
 */
const collectRowsFromTables = (tables) => {
  const functionName = 'collectRowsFromTables';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  // !tables1 catches:  null— was never set, undefined— was never passed in, any other falsy value
  // !Array.isArray(tables) catches: Catches cases where tables exists but is the wrong type, for example: a single
  //    Table object instead of an array of them,- a string, - a number, - an objec {}
  // finally: tables.length === 0 catchs if the array is empty
  if (!tables || !Array.isArray(tables) || tables.length === 0) {
    return theResults(false, ` tables must be a non-empty array'`, functionName);
  }

  // --- Extract rows from each table ---
  // We started with an array of tables, now we need to extract rows from each table
  // extractedTableRows is an array of arrays of strings, each representing a row of data
  // The outer array represents the original table
  // The inner array represents all the rows of data from the table
  const extractedTableRows = tables.map(extractRowDataFromOneTable);

  // --- Unwrap and flatten ---
  const allRowsFlat = extractedTableRows.flat();
  if (allRowsFlat.length === 0) {
    return theResults(false, ` No rows were extracted from any table`, functionName);
  }

  return theResults(true, ` Collected ${allRowsFlat.length} row(s) from ${tables.length} table(s)`, functionName, allRowsFlat)
};



/**
 * Extracts tables from the tab that are marked with an unchecked check box,
 * returns the tables as an array of Table objects.
 *
 * @param {DocumentApp.Tab} docTab              - The Google Doc tab to extract tables from
 * @param {string}          unCheckedCheckboxChar - The character used to mark tables for transfer
 *
 * @returns {{ ok: boolean, message: string, data: Array<DocumentApp.Table>|null }}
 *   data = array of marked Table objects
 */
const getMarkedTables = (docTab, unCheckedCheckboxChar) => {
  const functionName = 'getMarkedTables';
  Logger.log(`${functionName}. Started.`);
  let returnResult;


  // --- Guard: validate inputs ---
  if (!docTab) return theResults(false, 'getMarkedTables: docTab is null or undefined', functionName);

  if (!unCheckedCheckboxChar || unCheckedCheckboxChar.trim() === '') {
    return theResults(false, 'unCheckedCheckboxChar is null or blank', functionName);
  }

  // --- Get all tables from the tab body ---
  let tables;
  try {
    tables = docTab.asDocumentTab().getBody().getTables();
  } catch (err) {
    return theResults(false, `Could not read tables from tab - ${err.message}`, functionName);
  }

  if (!tables || tables.length === 0) return theResults(false, 'No tables found in tab', functionName);
  Logger.log(`getMarkedTables: found ${tables.length} total tables`);

  // --- Filter to only the marked tables ---
  let tablesToProcess;
  try {
    tablesToProcess = tablesSubset(tables, unCheckedCheckboxChar);
  } catch (err) {
    return theResults(false, `getMarkedTables: Could not filter tables - ${err.message}`, functionName);
  }

  if (!tablesToProcess || tablesToProcess.length === 0) {
    return theResults(false, `No tables marked with: ${unCheckedCheckboxChar}`, functionName);
  }
  Logger.log(`getMarkedTables: found ${tablesToProcess.length} marked tables`);

  return theResults(true, `Found ${tablesToProcess.length} marked table(s)`, functionName, tablesToProcess);
};

/**
 * Opens a Google Doc by ID, navigates to a named top level tab,
 * then finds and returns a named child tab within it.
 *
 * @param {string} docId        - The ID of the Google Doc
 * @param {string} topTabTitle  - The title of the top level tab
 * @param {string} subTabTitle  - The title of the child tab
 *
 * @returns {{ ok: boolean, message: string, data: DocumentTab|null }}
 *   data = the child tab object if found
 */
 
const getDocSubTab = (docId, topTabTitle, subTabTitle) => {
  const functionName = 'getDocSubTab';
  Logger.log(`${functionName} Started.`);
  let returnResult;
  // Logger.log(`docId=${docId} topTabTitle=${topTabTitle} subTabTitle=${subTabTitle}`);

  // --- Open the Google Doc ---
  let doc;
  try {
    doc = DocumentApp.openById(docId);
  } catch (err) {
    return theResults(false, `Could not open document with id: ${docId} - ${err.message}`, functionName);
  }
  Logger.log(`getDocSubTab: opened doc: ${doc.getName()}`);

  // --- Find the top level tab ---
  const topTabs   = doc.getTabs();
  const topResult = getTabByTitle(topTabs, topTabTitle);
  if (!topResult.ok) {
    return theResults(false, ` Could not find top tab: ${topTabTitle} in doc: ${doc.getName()}`, functionName);
  }
  Logger.log(`getDocSubTab: found top tab: ${topResult.tab.getTitle()}`);

  // --- Find the child tab ---
  const subTabs   = topResult.tab.getChildTabs();
  const subResult = getTabByTitle(subTabs, subTabTitle);
  if (!subResult.ok) {
    return theResults(false, ` Could not find sub tab: ${subTabTitle} under top tab: ${topTabTitle}`, functionName);
  }
  Logger.log(`getDocSubTab: found sub tab: ${subResult.tab.getTitle()}`);

  // --- Return the sub tab ---
  return theResults(true, `found sub tab: ${subTabTitle}`, functionName, subResult.tab);
};

 /***************************************************************************************************************************
 */
const replaceCheckInTopLeftCell = table => {
  const checkedCheckboxChar = '✔';
  const checkedEmoji = '✅';

  // Pure: Don't mutate input; just report if invalid table
  const isInvalid =
    !table ||
    table.getNumRows() === 0 ||
    table.getRow(0).getNumCells() === 0;

  if (isInvalid) {
    // Explicit result object: not ok, return original table
    return { ok: false, table };
  }

  const topLeftCellText = table.getCell(0, 0).getText();

  if (!topLeftCellText.includes(checkedCheckboxChar)) {
    // Explicit result: nothing replaced, original returned
    return { ok: false, table };
  }

  // Immutability: create a copy before making changes
  const copy = table.copy();
  const topLeftCellCopy = copy.getCell(0, 0);
  topLeftCellCopy.setText(
    topLeftCellCopy.getText().replaceAll(checkedCheckboxChar, checkedEmoji)
  );

  // Explicit result: replacement made, return new table
  return { ok: true, table: copy };
};


/**
 * Extracts tables from the tab that are marked with an unchecked check box,
 * returns an array of arrays, i.e. arry[0][array of data representing a single table row]
 *
 * @param {string} docId
 * @param {string} topTabTitle
 * @param {string} subTabTitle
 * @param {string} unCheckedCheckboxChar
 *
 *
 * @returns {Result}
 *   data = array of Line Entries
 *   e.g.
 *   data = [
 *   ["col1", "col2", "col3"],
 *   ["col1", "col2", "col3"],
 *   ]
 *
 */
const getMarkedTablesFromDoc = (docId, topTabTitle, subTabTitle,  unCheckedCheckboxChar) => {
  const functionName = 'getMarkedTablesFromDoc';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  // --- Step: Get the correct Google Doc tab ---
  // TODO: make sure all upperlevel functions are returning the lower layer errors upwards.
  returnResult = getDocSubTab(docId, topTabTitle, subTabTitle,);
  if (!returnResult.ok)  return theResults(false, returnResult.message, functionName);
  const subTab = returnResult.data;
  Logger.log(`${functionName}: subTab found: ${subTab.getTitle()}`);

  // --- Step: Extract tables marked for transfer ---
  returnResult = getMarkedTables(subTab, unCheckedCheckboxChar);
  if (!returnResult.ok)  return theResults(false, returnResult.message, functionName);
  const markedTables = returnResult.data;
  Logger.log(`${functionName}: found ${markedTables.length} marked tables`);

  returnResult = collectRowsFromTables(markedTables);
  if (!returnResult.ok)  return theResults(false, returnResult.message, functionName);
  flatAllRowsFromTables = returnResult.data;

  // debug:
  // print out the resulting tabular data
  for (let i = 0; i < flatAllRowsFromTables.length; i++) {        // outer loop: each line
    let line = flatAllRowsFromTables[i];
    let lineStr = "";
   for (let j = 0; j < line.length; j++) {      // inner loop: each cell in line
      lineStr += (j > 0 ? "\t" : "") + line[j];  // separate cells with a tab
    }
    Logger.log(`${lineStr}`);
  }

  return theResults(true, 'Success', functionName, flatAllRowsFromTables);
};

const markTablesAsComplete = (docId, topTabTitle, subTabTitle,  unCheckedCheckboxChar, checkedCheckboxChar) => {
  const functionName = 'markTablesAsComplete';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  // --- Step: Get the correct Google Doc tab ---
  // TODO: make sure all upperlevel functions are returning the lower layer errors upwards.
  returnResult = getDocSubTab(docId, topTabTitle, subTabTitle,);
  if (!returnResult.ok) return theResults(false, returnResult.message, functionName);
  const subTab = returnResult.data;

  returnResult = replaceCharInTablesInPlace( subTab, unCheckedCheckboxChar, checkedCheckboxChar);
  if (!returnResult.ok) return theResults(false, returnResult.message, functionName);

  Logger.log(`markTablesAsComplete. Exiting`);
  return theResults(true, 'Success.', functionName);


};

/* TBH. I have never taken this function apart to really understand how it works.
 *
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
const replaceCharInTablesInPlace = ( docTab, findChar, replaceWithChar) => {
  const functionName = 'replaceCharInTablesInPlace';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  let tables;
  try {
    tables = docTab.asDocumentTab().getBody().getTables();
  } catch (err) {
    return theResults(false, `Could not read tables from tab - ${err.message}`, functionName);
  }

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

  let returnMessage = `No checkmarks found in any table(s).`;
  if (changedIndices.length > 0) returnMessage = `Changed ${changedIndices.length} table(s).`;
  return theResults(true, 'Complete. ' + returnMessage, functionName);

};