/*
const transferDailiesWorkflow = () => {

  // Step 1 - Get the right tab
  const sub = getDailySubTab(docId, topTabTitle, subTabTitle);
  if (!sub.ok) { log and return; }

  // Step 2 - Extract marked table data
  const rows = extractTablesRows(sub.body, unCheckedCheckboxChar);

  // Step 3 - Write to sheet and mark as done
  if (appendTableRowsToSheet(rows, sheetId, sheetTabTitle)) {
    replaceCharInTablesInPlace(sub.body, unCheckedCheckboxChar, checkedCheckboxChar);
  }

};

*/




