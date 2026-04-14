

/* 1️⃣ Test
 ******************************************************************************************************************
 * 
 *
 */
function test_transferTablesBasedOnCheckboxFromAllTabs() {
  
  const sheetID ='1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ';
  const sheetTabTitle = 'Journal Input Queue';
  
  const docId = '15u2U-RSoiOVfaCPCwi0bf1Re_WsrO_QQG6aDw1oF3EM'; 
  const topTabTitle = 'Dailies';
  const subTabTitle = 'FY 2026 Q1 09';
  const unCheckedCheckboxChar = '🟪';
  const checkedCheckboxChar = '✔️';
  // Access the Google Doc and the Google Sheet
  const doc = DocumentApp.openById(docId);
  const topTabs = doc.getTabs();

  Logger.log(`Starting Google doc: ${doc.getName()}`);

  const top = getTabByTitle(topTabs, topTabTitle);
  if (top.ok) {
    Logger.log(`Found tab: ${top.tab.getTitle()}`);
  } else {
    Logger.log('Top tab ${topTabTitle} not found!');
    return;
  }

  const subTabs = top.tab.getChildTabs();
  const sub =  getTabByTitle(subTabs, subTabTitle);
  if (sub.ok) {
    Logger.log(`Found tab: ${sub.tab.getTitle()}`);
  } else {
    Logger.log('sub tab ${subTabTitle} not found!');
    return;
  }

  const tables = sub.tab.asDocumentTab().getBody().getTables();
  
  
  // const firstTable = tablesSubset(tables, checkedCheckboxChar )[0];
  //const rows =   extractTableRows(firstTable);
  if ( appendTableRowsToSheet(extractTablesRows(tables,unCheckedCheckboxChar),sheetID,sheetTabTitle)) {
    const replacedChar = replaceCharInTablesInPlace(sub.tab.asDocumentTab().getBody(),unCheckedCheckboxChar,checkedCheckboxChar)
    Logger.log(`${replacedChar.message}`);
  } else {
    Logger.log(`No rows were appended and no tables were marked as complete}`);
  }  


}

 
 
