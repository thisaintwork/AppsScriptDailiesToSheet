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
      .addItem('Delete Snapshot Tabs', 'deleteAllSnapshotTabs')
      .addItem('create Timestamped Snapshot', 'createTimestampedSnapshot')
      .addItem('Import Completed Dailies', 'transferDailiesWorkflow')
      .addToUi();
}



