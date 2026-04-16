/*
 * tests.gs
 * =====================================================================
 * Test pattern:
 *   1. Arrange  — set up inputs
 *   2. Act      — call the function
 *   3. Assert   — log expected vs actual
 *
 * Run all tests by calling runAllTests() from the Apps Script editor
 * =====================================================================
 */


/**
 * Returns all test cases for processTuplesThroughValidators.
 * Defined as a function so that validator references are resolved
 * at call time, not at load time.
 *
 * @returns {Object} testCases hash
 */
const getProcessTuplesTestCases = () => {
  return {

    happyPath: {
      hash: {
        sheetID:               undefined,
        sheetTabTitle:         undefined,
        unCheckedCheckboxChar: undefined,
        checkedCheckboxChar:   undefined,
        docId:                 undefined,
        topTabTitle:           undefined,
        subTabTitle:           undefined,
      },
      tuples: [
        ['sheetID',               '1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ'],
        ['sheetTabTitle',         'Journal Input Queue'],
        ['unCheckedCheckboxChar', '🟪'],
        ['checkedCheckboxChar',   '✔️'],
        ['docId',                 '15u2U-RSoiOVfaCPCwi0bf1Re_WsrO_QQG6aDw1oF3EM'],
        ['topTabTitle',           'Dailies'],
        ['subTabTitle',           'FY 2026 Q2 10'],
        ['comment',               'this is a note and should be skipped'],
        ['',                      'this empty row should be skipped'],
      ],
      validators: [
        checkIsAttributeDuplicate,
        checkIsAttributeKnownKey,
        checkIsAttributeValueDefined,
        assignValueToHash,
      ],
      expectedOk:      true,
      expectedMessage: 'All tuples processed successfully',
      expectedData: {
        sheetID:               '1y2Frx8OJoKtVdTtfGdngpabnipCA96DylhODX82HZwQ',
        sheetTabTitle:         'Journal Input Queue',
        unCheckedCheckboxChar: '🟪',
        checkedCheckboxChar:   '✔️',
        docId:                 '15u2U-RSoiOVfaCPCwi0bf1Re_WsrO_QQG6aDw1oF3EM',
        topTabTitle:           'Dailies',
        subTabTitle:           'FY 2026 Q2 10',
      },
    },

    duplicate: {
      hash: {
        sheetID:     undefined,
        topTabTitle: undefined,
      },
      tuples: [
        ['sheetID',     '1y2Frx8OJoK...'],
        ['sheetID',     'duplicate value'],
        ['topTabTitle', 'Dailies'],
      ],
      validators: [
        checkIsAttributeDuplicate,
        checkIsAttributeKnownKey,
      ],
      expectedOk:      false,
      expectedMessage: 'Duplicate attribute found: sheetID',
      expectedData:    null,
    },

    unknownKey: {
      hash: {
        sheetID: undefined,
      },
      tuples: [
        ['unknownAttribute', 'someValue'],
      ],
      validators: [
        checkIsAttributeDuplicate,
        checkIsAttributeKnownKey,
      ],
      expectedOk:      false,
      expectedMessage: 'Unknown attribute: unknownAttribute',
      expectedData:    null,
    },

  };
};


// =====================================================================
// TEST RUNNER
// =====================================================================

/**
 * Runs a single test case and logs results.
 * @param {string} testName  - The key from testCases
 * @param {Object} testCase  - The test case object
 */
const runOneTest = (testName, testCase) => {

  Logger.log(`=== START: ${testName} ===`);

  // --- Act ---
  const result = processTuplesThroughValidators(
    testCase.tuples,
    testCase.hash,
    testCase.validators,
  );

  // --- Assert: ok ---
  Logger.log(`ok:      expected=${testCase.expectedOk}  actual=${result.ok}  ${result.ok === testCase.expectedOk ? '✅' : '❌'}`);

  // --- Assert: message ---
  Logger.log(`message: expected="${testCase.expectedMessage}"`);
  Logger.log(`         actual="${result.message}"  ${result.message === testCase.expectedMessage ? '✅' : '❌'}`);

  // --- Assert: data (only if expectedData is defined) ---
  if (testCase.expectedData !== null && testCase.expectedData !== undefined) {
    Logger.log('--- data assertions ---');
    for (const key of Object.keys(testCase.expectedData)) {
      const expected = testCase.expectedData[key];
      const actual   = result.data ? result.data[key] : undefined;
      Logger.log(`  ${key}: expected="${expected}"  actual="${actual}"  ${actual === expected ? '✅' : '❌'}`);
    }
  }

  Logger.log(`=== END: ${testName} ===`);
};


// =====================================================================
// ENTRY POINT
// =====================================================================

function runAllTests() {
  Logger.log('=== STARTING ALL TESTS ===');

  const allTestCases = {
    ...getProcessTuplesTestCases(),
  };

  for (const testName of Object.keys(allTestCases)) {
    runOneTest(testName, allTestCases[testName]);
  }

  Logger.log('=== ALL TESTS COMPLETE ===');
}

