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
  returnResult = readInput( getRequiredInitInfo(),getConfigHash(),getValidators());
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
    Logger.log(`transferDailiesWorkflow. debug. Input Data: config[getConfigHash()_key] = ${value}`);
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
 returnResult = appendTableRowsToSheet(arry, config.sheetID, config.sheetTabTitle);
 if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }

  // --- Step 4: Mark transferred tables as complete ---
  returnResult = markTablesAsComplete(config.docId, config.topTabTitle, config.subTabTitle, config.unCheckedCheckboxChar, config.checkedCheckboxChar);
  if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }

  Logger.log(`transferDailiesWorkflow: Exiting`);
}

/**
 * Creates a timestamped snapshot of the specified sheet tab by copying its data and dimensions to a new sheet.
 * The new sheet is named using a combination of a prefix and a timestamp.
 * Alerts the user with a success or error message based on the operation outcome.
 *
 * @return {void} Does not return a value. Displays alerts to inform the user of success or failure. Handles errors gracefully within the method.
 */
function createTimestampedSnapshot() {
  const functionName = 'createTimestampedSnapshot';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  const ui = SpreadsheetApp.getUi(); // Get the UI object for alerts

  returnResult = readInput(getRequiredInitInfo(),getConfigHash(),getValidators());
  if (!returnResult.ok) {
    Logger.log(`${errLogString}` +  returnResult.message);
    ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
    return;
  }
  const config = returnResult.data;
  const newSheetTitle = config.copiedSheetPrefix + "-" + getTimestampString();


  returnResult = createCurrentSheetTabSnapshot(config.sheetTabTitle,newSheetTitle,config.dateHeader,config.topicHeader);
  if (!returnResult.ok) {
    ui.alert("Error", returnResult.message, ui.ButtonSet.OK);
    return;
  }
  ui.alert("Success", `Values and dimensions from '${config.sheetTabTitle}' have been copied to a new sheet: '${newSheetTitle}'!`, ui.ButtonSet.OK);
}


/**
 * Deletes all snapshot tabs in a Google Spreadsheet that match a specific prefix.
 *
 * The function retrieves initialization and configuration information required for the operation.
 * It validates the inputs and deletes all snapshot tabs matching the provided prefix.
 * Alerts are shown to the user in case of errors or to provide informational messages indicating the operation's results.
 * Internal logging is performed for debugging purposes.
 *
 * Process:
 * - Retrieves required initialization and configuration data.
 * - Validates the input parameters.
 * - Deletes all tabs that match the specified snapshot prefix.
 * - Alerts the user with success or error messages.
 *
 * Function relies on:
 * - `getRequiredInitInfo()` for initialization data.
 * - `getConfigHash()` for configuration data.
 * - `getValidators()` for validating input parameters.
 * - `deleteSnapshotTabs()` for the tab deletion process based on prefix.
 *
 * Alerts:
 * - Displays an alert if initialization or validation fails.
 * - Displays an alert after completing the delete operation or if an error occurs during the deletion process.
 */
const deleteAllSnapshotTabs = () => {
    const functionName = 'deleteAllSnapshotTabs';
    Logger.log(`${functionName}. Started.`);
    let returnResult;

    const ui = SpreadsheetApp.getUi(); // Get the UI object for alerts

    returnResult = readInput(getRequiredInitInfo(),getConfigHash(),getValidators());
    if (!returnResult.ok) {
      Logger.log(`${errLogString}` +  returnResult.message);
      ui.alert(`${errLogString}`, returnResult.message, ui.ButtonSet.OK);
      return;
    }
    const config = returnResult.data;

    returnResult = deleteSnapshotTabs(config.copiedSheetPrefix);
    if (!returnResult.ok) {
      ui.alert("Error", returnResult.message, ui.ButtonSet.OK);
      return;
    }
    ui.alert("Information", returnResult.message, ui.ButtonSet.OK);
}

