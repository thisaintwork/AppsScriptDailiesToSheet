// workflow.js
// =====================================================================
// Rule: Orchestration only.
// No low level detail. No Google API calls.
// Reads like a plain English description of the process.
// =====================================================================

const getRequiredInitInfo = () => ({
  sheetID: '1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ',
  sheetInpuTabTitle: 'DailiesXferMetaData',
  maxInputRows: 20,
});

const getConfigHash = () => ({
  sheetTabTitle:         undefined,
  unCheckedCheckboxChar: undefined,
  checkedCheckboxChar:   undefined,
  docId:                 undefined,
  topTabTitle:           undefined,
  subTabTitle:           undefined,
});

const getValidators = () => [
  checkIsAttributeUnique,
  checkIsAttributeKnownKey,
  checkIsAttributeValueDefined,
  assignValuesToHash,
];

// =====================================================================

/**
 * Main entry point for the Import Completed Dailies workflow.
 * Called from menu.js when the user selects 'Import Completed Dailies'.
 *
 * Steps:
 *   0. Read and validate configuration from DailiesXferMetaData tab
 *   1. Get the correct Google Doc tab for the current period
 *   2. Extract tables marked for transfer
 *   3. Write table data to the Google Sheet
 *   4. Mark transferred tables as complete
 */
function transferDailiesWorkflow() {
  Logger.log(`Entered: transferDailiesWorkflow`);
  const ui = SpreadsheetApp.getUi();

  // --- Step 0: Read raw config data from sheet ---
  const rawData = readTabAsTuples(getRequiredInitInfo().sheetID,getRequiredInitInfo().sheetInpuTabTitle,getRequiredInitInfo().maxInputRows);
  if (!rawData.ok) {
    ui.alert('Configuration Error', rawData.message, ui.ButtonSet.OK);
    return;
  }

  // Strip leading and trailing spaces from all tuple values
  const trimmedTuples = rawData.data.map(tuple => [
    tuple[0].trim(),
    tuple[1].trim(),
  ]);

  // --- Step 0: Validate and load config ---
  const configResult = processTuplesThroughValidators(
    trimmedTuples,
    getConfigHash(),
    getValidators(),
  );


  if (!configResult.ok) {
    ui.alert('Configuration Error', configResult.message, ui.ButtonSet.OK);
    return;
  }

  const config = configResult.data;
  Logger.log(`transferDailiesWorkflow: config loaded successfully`);

  // --- Step 1: Get the correct Google Doc tab ---
  // TODO: implement getDocSubTab()

  // --- Step 2: Extract tables marked for transfer ---
  // TODO: implement extractMarkedTables()

  // --- Step 3: Write table data to the Google Sheet ---
  // TODO: implement writeTableDataToSheet()

  // --- Step 4: Mark transferred tables as complete ---
  // TODO: implement markTablesAsComplete()

  Logger.log(`transferDailiesWorkflow: completed successfully`);
}