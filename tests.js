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
        sheetTabTitle:         undefined,
        unCheckedCheckboxChar: undefined,
        checkedCheckboxChar:   undefined,
        docId:                 undefined,
        topTabTitle:           undefined,
        subTabTitle:           undefined,
      },
      tuples: [
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
        checkIsAttributeKnownKey,
        checkIsAttributeUnique,
        checkIsAttributeValueDefined,
        assignValuesToHash,
      ],
      expectedOk:      true,
      expectedMessage: 'All tuples processed successfully',
      expectedData: {
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
        topTabTitle: undefined,
      },
      tuples: [
        ['topTabTitle', 'duplicate value'],
        ['topTabTitle', 'Dailies'],
      ],
      validators: [
        checkIsAttributeUnique,
,
      ],
      expectedOk:      false,
      expectedMessage: 'Duplicate attribute found: topTabTitle',
      expectedData:    null,
    },

    unknownKey: {
      hash: {
        topTabTitle: undefined,
      },
      tuples: [
        ['unknownAttribute', 'someValue'],
      ],
      validators: [
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

  Logger.log(`== runOneTest:${testName} START: ===`);

  // todo: I removed the validations on the test data. add back later.

  // --- Act ---
  const result = processTuplesThroughValidators(
    testCase.tuples,
    testCase.hash,
    testCase.validators,
  );
  Logger.log(`runOneTest: ${testName} Assertions.`);
  Logger.log(`runOneTest: ${testName} testCase Expected Ok: ${testCase.expectedOk}`);
  Logger.log(`runOneTest: ${testName} testCase Expected Message: ${testCase.expectedMessage}`);
  Logger.log(`runOneTest: ${testName} Actual.`);
  Logger.log(`runOneTest: ${testName} testCase Actual : ${result.ok}`);
  Logger.log(`runOneTest: ${testName} testCase Actual Message: ${result.message}`);
  Logger.log(`runOneTest: ${testName} Pass?  ${testCase.expectedOk === result.ok ? '✅' : '❌'}`);
  Logger.log(`== runOneTest:${testName} END: ===`);
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

