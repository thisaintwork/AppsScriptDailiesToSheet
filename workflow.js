// workflow.js
// =====================================================================
// Rule: Orchestration only.
// No low level detail. No Google API calls.
// Reads like a plain English description of the process.
// =====================================================================




// =====================================================================

/**
 * Main entry point for the Import Completed Dailies workflow.
 * Called from menu.js when the user selects 'Import Completed Dailies'.
 *
 * Steps:
 *   0. Read and validate input configuration
 *   2. Extract tables marked for transfer
 *   3. Write the extracted data to the Google Sheet
 *   4. Mark transferred tables as complete
 */
function transferDailiesWorkflow() {
  Logger.log(`Entered: transferDailiesWorkflow`);
  const errLogString = `transferDailiesWorkflow Error.`;
  const ui = SpreadsheetApp.getUi();
  let returnResult;


  // --- Step 0: Read raw config data from sheet ---
  const { sheetID, sheetInputTabTitle, inputRows } = getRequiredInitInfo();
  returnResult = readInput(sheetID,sheetInputTabTitle,inputRows,getConfigHash(),getValidators());
  if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }
  const config = returnResult.data;

  // debug:
  // print out the required data keys along with what
  // is in the actual config data
  Object.keys(getConfigHash()).forEach(function(key) {
    const value = config[key];
    Logger.log(`transferDailiesWorkflow. Input Data: config[getConfigHash()_key] = ${config[key]}`);
  });

  returnResult = getMarkedTablesFromDoc(config.docId, config.topTabTitle, config.subTabTitle, config.unCheckedCheckboxChar);
  if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }
  const arry = returnResult.data;

 // --- Step: Write the Marked Tables to the sheet
 // Logger.log(`transferDailiesWorkflow: getRequiredInitInfo().sheetID = ${getRequiredInitInfo().sheetID}, config.sheetTabTitle = ${config.sheetTabTitle}.`);
 returnResult = appendTableRowsToSheet(arry, getRequiredInitInfo().sheetID, config.sheetTabTitle);
 if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }

  // --- Step 4: Mark transferred tables as complete ---
  // TODO: implement markTablesAsComplete()
  returnResult = markTablesAsComplete(config.docId, config.topTabTitle, config.subTabTitle, config.unCheckedCheckboxChar, config.checkedCheckboxChar);
  if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }

  Logger.log(`transferDailiesWorkflow: Exiting`);
}