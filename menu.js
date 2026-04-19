/* menu.js
Only:

custom menu creation
user-triggered entry points
Rule: no real business logic here.
*/

/***********************************************************************************************************************************************************************************************
 * OPTIONAL: You can add a custom menu item to easily run this script.
 * This function runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Delete Snapshot Tabs', 'deleteSnapshotTabs')
      .addItem('create Timestamped Snapshot', 'createTimestampedSnapshot')
      .addItem('Import Completed Dailies', 'transferDailiesWorkflow')
      .addToUi();
}




/***********************************************************************************************************************************************************************************************
 * Deletes all sheets in the active spreadsheet that have
 * a name starting with the prefix defined below
 * This function now skips hidden and protected sheets.
 */
function deleteSnapshotTabs() {
  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi(); // Get the UI object for alerts

  // Get all sheets in the spreadsheet
  const sheets = spreadsheet.getSheets(); // Fetch sheets once at the beginning

  // Define the prefix to look for
  // Note: This prefix is case-sensitive ("Snapshot-" will not match "snapshot-")
  const prefix = "Snapshot-";

  let sheetsDeletedCount = 0; // Counter for successfully deleted sheets

  // Iterate through sheets in reverse order.
  // This avoids issues with index shifting when sheets are deleted.
  for (let i = sheets.length - 1; i >= 0; i--) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    const sheetId = sheet.getSheetId(); // Get the unique ID of the sheet

    // Check if the sheet is hidden
    if (sheet.isSheetHidden()) {
      Logger.log(`Skipping hidden sheet: ${sheetName} (hidden)`);
      continue; // Skip to the next sheet
    }

    // Check if the sheet is protected
    if (sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0) {
      Logger.log(`Skipping protected sheet: ${sheetName} (protected)`);
      continue; // Skip to the next sheet
    }

    // Check if the sheet name starts with the defined prefix
    if (sheetName.startsWith(prefix)) {
      // Before attempting deletion, check if it's the last visible sheet.
      // A spreadsheet must always have at least one visible sheet.
      const visibleSheets = spreadsheet.getSheets().filter(s => !s.isSheetHidden());
      if (visibleSheets.length === 1 && visibleSheets[0].getSheetId() === sheetId) {
        ui.alert("Warning", `Cannot delete the last visible sheet: '${sheetName}'. A spreadsheet must always have at least one visible sheet.`, ui.ButtonSet.OK);
        Logger.log(`Cannot delete last visible sheet: ${sheetName}`);
        continue; // Skip to the next sheet
      }

      // Attempt to delete the sheet with error handling
      try {
        ui.alert("Information", `Attempting to delete sheet: '${sheetName}', id: '${sheetId}', index: '${i}'.`, ui.ButtonSet.OK);
        spreadsheet.deleteSheet(sheet);
        sheetsDeletedCount++;
        Logger.log(`Successfully deleted sheet: ${sheetName}`);
        // No need to re-fetch sheets or adjust 'i' when iterating backwards
      } catch (e) {
        ui.alert("Error", `Failed to delete sheet: '${sheetName}'. Error: ${e.message}`, ui.ButtonSet.OK);
        Logger.log(`Error deleting sheet ${sheetName}: ${e.message}`);
        // If deletion fails, we still move to the next sheet in the backward iteration
      }
    }
  }

  Logger.log("Finished checking for snapshot tabs.");
  ui.alert("Information", `Deletions are complete. ${sheetsDeletedCount} sheet(s) were deleted.`, ui.ButtonSet.OK);
}



