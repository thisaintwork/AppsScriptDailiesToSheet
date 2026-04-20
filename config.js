//config.js

// These constants are hardcoded.
// They are rarely going to change and are needed for the workflow.
const getRequiredInitInfo = () => ({
  sheetID: '1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ',
  sheetInputTabTitle: 'DailiesXferMetaData',
  inputRows: 20,
});

// This is the canonical hash of all the config keys.
const getConfigHash = () => ({
  sheetTabTitle:         undefined,
  unCheckedCheckboxChar: undefined,
  checkedCheckboxChar:   undefined,
  docId:                 undefined,
  topTabTitle:           undefined,
  subTabTitle:           undefined,
});
