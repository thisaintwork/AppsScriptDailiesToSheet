/* 1️⃣
 *************************************************************************************************************************** 
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
const tablesSubset = (tables, checkedCheckboxChar = '✔') => 
  tables.filter(table => {
    if (table.getType() !== DocumentApp.ElementType.TABLE) return false;
    if (table.getNumRows() === 0) return false;
    if (table.getRow(0).getNumCells() === 0) return false;
    const topLeft = table.getCell(0, 0).getText().trim();
    return topLeft.includes(checkedCheckboxChar);
  });


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





const extractTablesRows = (tables,checkedCheckboxChar) => {

  // 
  const tablesToProcess = tablesSubset(tables,checkedCheckboxChar);
  Logger.log(`extractTablesRows; num tablesToProcess: ${tablesToProcess.length}`);

  // the result of this step is 
  // An array of tables
  // Each entry in the array of tables contains
  //   An array of rows.
  //   each entry in the array of rows continas
  //      An array of cells from that row.
  const allTablesRowsNested = tablesToProcess.map(extractTableRows);

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



  return allTablesRowsFlat;
}



/**
 * Extracts cell text from all rows (skipping the header) of a Google Doc table.
 * Fully functional implementation using map and slice.
 * * @param {DocumentApp.Table} table The Google Doc Table element to process.
 * @returns {Array<Array<string>>} An array of rows, where each row is an array of cell text strings.
 */
const extractTableRows = table => {
  const numRows = table.getNumRows();
  
  // 1. Create an array of row indices starting from 1 (to skip the header at index 0).
  //    The 'slice(1)' pattern is the most functional way to skip the first element.
  // 2. Filter: Skip the header row (index 0).
  // 3. Map: Transform each row index into an array of cell strings.
  const rowIndices = Array.from({ length: numRows }, (_, i) => i); // [0, 1, 2, 3, ...]
  const dataRowIndices = rowIndices.slice(1); // [1, 2, 3, ...]
  Logger.log(`extractTableRows; rowIndices: ${rowIndices.length}, dataRowIndices: ${dataRowIndices.length}`);
  
  const rowsData = dataRowIndices.map(r => {
    const row = table.getRow(r);
    const numCells = row.getNumCells();

    // Create an array of cell indices for the current row
    const cellIndices = Array.from({ length: numCells }, (_, c) => c);
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

  // Logger.log(`extractTableRows; rowsData: ${rowsData.length}`);
  //Logger.log(`extractTableRows; Extracted Rows: ${JSON.stringify(rowsData, null, 2)}`);
  // Logger.log('extractTableRows;  ===================================================================');
  // Logger.log(`extractTableRows; rowsData[0,0]: ${rowsData[0,0]}`);
  // Logger.log(`extractTableRows; rowsData[0,1]: ${rowsData[0,1]}`);
  // Logger.log(`extractTableRows; rowsData[0,1]: ${rowsData[0,2]}`);
  // Logger.log('extractTableRows;  ===================================================================');


  return rowsData;
}

