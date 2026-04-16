/* Test pattern:

1. Arrange  — set up inputs
2. Act      — call the function
3. Assert   — log expected vs actual

*/

 function test_processTuplesThroughValidators_happyPath() {

  // --- Test Input: the hash with predefined keys, all undefined ---
  const hash = {
    sheetID:               undefined,
    sheetTabTitle:         undefined,
    unCheckedCheckboxChar: undefined,
    checkedCheckboxChar:   undefined,
    docId:                 undefined,
    topTabTitle:           undefined,
    subTabTitle:           undefined
  };

  // --- Test Input: tuples representing what would be read from the sheet --
  // tuples[7][0]  // 'comment'
  // tuples[7][1]  // 'this is a note and should be skipped'

  const tuples = [
    ['sheetID',               '1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ'],
    ['sheetTabTitle',         'Journal Input Queue'],
    ['unCheckedCheckboxChar', '🟪'],
    ['checkedCheckboxChar',   '✔️'],
    ['docId',                 '15u2U-RSoiOVfaCPCwi0bf1Re_WsrO_QQG6aDw1oF3EM'],
    ['topTabTitle',           'Dailies'],
    ['subTabTitle',           'FY 2026 Q2 10'],
    ['comment',               'this is a note and should be skipped'],
    ['',                      'this empty row should be skipped']
  ];

  // --- Validators ---
  // The order is important
  const validators = [
    checkIsAttributeDuplicate,
    checkIsAttributeKnownKey,
    checkIsAttributeValueDefined,
    assignValueToHash,
  ];

  // --- Run ---
  const result = processTuplesThroughValidators(tuples, hash, validators);

  // --- Assertions ---
  Logger.log('=== test_processTuplesThroughValidators ===');
  Logger.log(`ok:      expected=true,  actual=${result.ok}`);
  Logger.log(`message: ${result.message}`);

  if (result.ok) {
    Logger.log(`sheetID:               actual=${result.data.sheetID}  expected=1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ`);
    Logger.log(`sheetTabTitle:         actual=${result.data.sheetTabTitle}                            expected=Journal Input Queue`);
    Logger.log(`unCheckedCheckboxChar: actual=${result.data.unCheckedCheckboxChar}                                             expected=🟪`);
    Logger.log(`checkedCheckboxChar:   actual=${result.data.checkedCheckboxChar}                                             expected=✔️`);
    Logger.log(`docId:                 actual=${result.data.docId}  expected=15u2U-RSoiOVfaCPCwi0bf1Re_WsrO_QQG6aDw1oF3EM`);
    Logger.log(`topTabTitle:           actual=${result.data.topTabTitle}                                        expected=Dailies`);
    Logger.log(`subTabTitle:           actual=${result.data.subTabTitle}                                  expected=FY 2026 Q2 10`);

  }

  Logger.log('=== end test ===');
}



 function test_processTuplesThroughValidators_duplicate() {

  const hash = {
    sheetID:   undefined,
    topTabTitle: undefined
  };

  // sheetID appears twice - should fail
  const tuples = [
    ['sheetID',     '1y2Frx8OJoK...'],
    ['sheetID',     'duplicate value'],
    ['topTabTitle', 'Dailies']
  ];

  const validators = [
    checkIsAttributeDuplicate,
    checkIsAttributeKnownKey,

  ];

  const result = processTuplesThroughValidators(tuples, hash, validators);

  Logger.log('=== test_duplicate ===');
  Logger.log(`ok:      expected=false  actual=${result.ok}`);
  Logger.log(`message: expected=Duplicate attribute found: sheetID  actual=${result.message}`);
  Logger.log('=== end test ===');
}


function test_processTuplesThroughValidators_unknownKey() {

  const hash = {
    sheetID: undefined
  };

  // unknownAttribute is not in the hash - should fail
  const tuples = [
    ['unknownAttribute', 'someValue']
  ];

  const validators = [
    checkIsAttributeDuplicate,
    checkIsAttributeKnownKey,

  ];

  const result = processTuplesThroughValidators(tuples, hash, validators);

  Logger.log('=== test_unknownKey ===');
  Logger.log(`ok:      expected=false  actual=${result.ok}`);
  Logger.log(`message: expected=Unknown attribute: unknownAttribute  actual=${result.message}`);
  Logger.log('=== end test ===');
}

// tests.gs
// =====================================================================
// Run all tests by calling runAllTests() from the Apps Script editor
// =====================================================================

function runAllTests() {

  Logger.log('=== STARTING TESTS ===');

  test_processTuplesThroughValidators_happyPath();
/*
  test_processTuplesThroughValidators_duplicate();
  test_processTuplesThroughValidators_unknownKey();
  test_processTuplesThroughValidators_emptyTuples();
  test_okToSkip();
*/
  // add more here as you build more functions
  Logger.log('=== ALL TESTS COMPLETE ===');
}

