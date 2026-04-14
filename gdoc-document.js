/* 1️⃣
 *************************************************************************************************************************** 
 * Searches an array of DocumentApp.Tab objects for a tab with a specific title.
 * @param {DocumentApp.Tab[]} tabs The array of tabs from the Google Doc.
 * @param {string} targetTitle The title of the tab to find.
 * @returns {DocumentApp.Tab|null} The DocumentApp.Tab object if found, otherwise null.
 */
const getTabByTitle = (tabs, targetTitle) => {
  // Use Array.prototype.find: declarative, no mutation
  const tab = tabs.find(tab => tab.getTitle() === targetTitle);

  // Return a result object for clarity (functional idiom)
  return tab
    ? { ok: true, tab }
    : { ok: false, tab: null };
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
  };
}


