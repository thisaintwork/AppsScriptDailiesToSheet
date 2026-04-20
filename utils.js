// utils.js
/**
 * Standard result object for this project.
 *
 * @typedef {Object} Result
 * @property {boolean} ok       - True on success, false on failure.
 * @property {string}  message  - Description of the outcome.
 * @property {*}       data     - Payload data (or null on failure).
 */

// /**
//  * @param {string} message
//  * @param {*} [data=null]
//  * @param {string} [fn=""]
//  * @returns {Result}
//  */
// const okResult = (message, data = null, fn = "") => ({
//   ok:      true,
//   message: `${fn}. ${message}`,
//   data:    data
// });
//
// /**
//  * @param {string} message
//  * @param {string} [fn=""]
//  * @returns {Result}
//  */
// const failResult = (message) => {
//   Logger.log(message);
//   return ({
//     ok:      false,
//     message: `${message}`,
//     data:    null
//   });
// }

/**
 * @param {string} message
 * @param {string} functionName
 * @returns {Result}
 */
const theResults = (pass = false, message, functionName, data = null) => {
  const passString = pass ? "Success" : "Error";
  Logger.log(`${functionName}. ${passString}. ${message}. Is data null? = ${data === null}`);
  return ({
    ok:      pass,
    message: message,
    data:    data
  });
}

/**
 * This is the list of validators to run on the input data.
 * The order of the validators matters!
 * See the comments in each validator function for more details.
 * Don't change the order of the validators or make edits without understanding
 * the implications.
  */
const getValidators = () => [
  checkIsAttributeUnique,
  checkIsAttributeKnownKey,
  checkIsAttributeValueDefined,
  assignValuesToHash,
];

/*
 *************************************************************************************************************************** 
 * Generates a timestamp string for the current date and time in 
 * YYYYMMDDHHMMSS format (e.g., 20250930172629).
 * * @returns {string} The formatted timestamp string.
 */
const getTimestampString = () => {
  const pad = num => (num < 10 ? '0' : '') + num;
  const now = new Date();
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds())
  ].join('');
};


  /**
 * Processes an array of tuples through a pipeline of validator functions.
 * Each validator runs once and receives the full array of tuples and the
 * current state of the hash.
 *
 * Pre-conditions:
 *   - tuples is an array of [attributeString, valueString] pairs
 *   - hash has predefined keys with undefined values
 *   - validators is an array of functions with signature:
 *       (tuples: Array<Array<string>>, hash: Object) => { ok, message, data: updatedHash }
 *
 * @param {Array<Array<string>>} tuples      - Array of [attributeString, valueString] pairs
 * @param {Object}               hash        - Predefined keys with undefined values
 * @param {Array<Function>}      validators  - Array of validator functions
 *
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = final state of hash after all validators processed
 */
const populateInputValues = (tuples, hash, validators) => {
  const functionName = `populateInputValues`;
  Logger.log(`${functionName}. Started.`);
  let returnResult;


  // --- Guard: validate inputs ---
  if (!tuples || !Array.isArray(tuples)) {
    return theResults(false, 'tuples must be an array', functionName);

  }
  if (!hash || typeof hash !== 'object') {
    return theResults(false, 'hash must be an object', functionName);
  }
  if (!validators || !Array.isArray(validators)) {
    return theResults(false, 'validators must be an array', functionName);
  }

  // --- Start with a clean copy of the incoming hash ---
  let currentHash = { ...hash };

  // --- Loop: each validator runs once against the full tuples array ---
  Logger.log(`>>populateInputValues. Start Validator Loop`);
  for (const validator of validators) {

    Logger.log(`>> Running validator: ${validator.name}`);
    const result = validator(tuples, currentHash);
    if (!result.ok) {
      return theResults(false, result.message, functionName);
    }
    Logger.log(`>> ${validator.name} result: ${result.ok} - ${result.message}`);

    // Carry the updated hash forward to the next validator
    currentHash = { ...result.data };
  }
  Logger.log(`>>populateInputValues. Completed Validator Loop`);

  // print out the required data keys along with what
  // is in the actual config data
  Object.keys(hash).forEach(function(key) {
    const value = currentHash[key];
    Logger.log(`populateInputValues. Final. >${key}< = >${value}<`);
  });

  // --- All validators passed ---
  return theResults(true, 'Completed.', functionName, currentHash);
};

/**
 * Determines whether a tuple should be skipped entirely.
 * Called before the validator pipeline — not part of the validators array.
 *
 * Skip conditions:
 *   1. attributeName is empty or blank
 *   2. attributeName is the word 'comment' (case insensitive)
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   ok: true  = yes, skip this tuple
 *   ok: false = no, continue to validator pipeline
 */
const okToSkip = (tuple, hash) => {
  const functionName = 'okToSkip';
  // Logger.log(`${functionName}. Started.`);

  const trimmed = tuple[0].trim().toLowerCase();

  if (trimmed === '') {
    return theResults(true, 'empty attribute - skip', functionName);
  }

  if (trimmed === 'comment') {
    return theResults(true, 'comment - skip', functionName);
  }

  return theResults(false, 'not a skip condition - continue processing', functionName);
};


/**
 * Fails if any attribute name appears more than once in the tuples array.
 * Ignores comment and empty rows.
 * This must run after checkIsAttributeKnownKey
 *
 * @param {Array<Array<string>>} tuples
 * @param {Object}               hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 */
const checkIsAttributeUnique = (tuples, hash) => {
  const functionName = 'checkIsAttributeUnique';
  Logger.log(`${functionName}. Started.`);

  const seen       = {};
  const duplicates = [];

  for (const tuple of tuples) {
    const attributeName = tuple[0].trim().toLowerCase();

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    if (seen[attributeName]) {
      if (!duplicates.includes(attributeName)) {
        duplicates.push(attributeName);
      }
    } else {
      seen[attributeName] = true;
    }
  }

  if (duplicates.length > 0) {
    return theResults(false, `Duplicate attribute(s) found: ${duplicates.join(', ')}`, functionName);
  }

  return theResults(true, 'No duplicates found', functionName, {...hash});
};


/**
 * Fails if any attribute name in the tuples array is not a predefined key in the hash.
 * Unknown keys are not valid input.
 * Ignores comment and empty rows.
 * This must be the first validator that is run
 *
 * @param {Array<Array<string>>} tuples - Array of [attributeName, value] pairs
 * @param {Object}               hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeKnownKey = (tuples, hash) => {
  const functionName = 'checkIsAttributeKnownKey';
  Logger.log(`${functionName}. Started.`);
  let returnResult;

  const unknownKeys = [];
  const hashByHashKey = getConfigHash();

  for (const key of Object.keys(hashByHashKey)) {
    hashByHashKey[key] = key;
    //Logger.log(`checkIsAttributeKnownKey. key = ${key}, hashByHashKey[key] = ${hashByHashKey[key]}`);
  }
  for (const key of Object.keys(hashByHashKey)) {
      Logger.log(`checkIsAttributeKnownKey. key = ${key}, hashByHashKey[key] = ${hashByHashKey[key]}`);
  }

  for (const tuple of tuples) {

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    const attributeName = tuple[0];
    Logger.log(`checkIsAttributeKnownKey. attributeName = ${attributeName}, hashByHashKey[attributeName] = ${hashByHashKey[attributeName]}, ${attributeName === hashByHashKey[attributeName] ? '✅' : '❌'}`);
    if (hashByHashKey[attributeName] !== attributeName) {
      unknownKeys.push(attributeName);
    }
  }

  if (unknownKeys.length > 0) return theResults(false, `Unknown attribute(s) found: ${unknownKeys.join(', ')}`, functionName);
  return theResults(true, 'All attribute names are known keys', functionName, {...hash});
};




/**
 * Fails if any tuple value in the tuples array is missing, not a string, or blank.
 * Ignores comment and empty rows.
 * Must run after validator: checkIsAttributeKnownKey
 *
 * @param {Array<Array<string>>} tuples - Array of [attributeName, value] pairs
 * @param {Object}               hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeValueDefined = (tuples, hash) => {
  const functionName = 'checkIsAttributeValueDefined';
  Logger.log(`${functionName}. Started.`);
  const invalidValues = [];

  for (const tuple of tuples) {

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    const attributeName = tuple[0].trim() ;
    const value         = tuple[1].trim();

    if (value.trim() === '') {
      invalidValues.push(`[${attributeName}] value is blank`);
      continue;
    }

    if (value === undefined || value === null) {
      invalidValues.push(`[${attributeName}] value is undefined or null`);
      continue;
    }

    if (typeof value !== 'string') {
      invalidValues.push(`[${attributeName}] value is not a string: ${value}`);
      continue;
    }

  }

  if (invalidValues.length > 0) return theResults(false, `Invalid value(s) found: ${invalidValues.join(', ')}`, functionName);

  return theResults(true, 'All attribute values are valid', functionName, {...hash});
};

/**
 * Assigns values from the tuples array to the correct keys in the hash.
 * This should always be the last validator in the pipeline.
 * Ignores comment and empty rows.
 *
 * @param {Array<Array<string>>} tuples - Array of [attributeName, value] pairs
 * @param {Object}               hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = updated copy of hash with all values assigned
 */
const assignValuesToHash = (tuples, hash) => {
  const functionName = 'assignValuesToHash';
  Logger.log(`${functionName}. Started.`);
  const updatedHash = { ...hash };

  for (const tuple of tuples) {

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    const attributeName = tuple[0];
    const value         = tuple[1];

    updatedHash[attributeName] = value;
    Logger.log(`assignValuesToHash updatedHash[${attributeName}]=${value}`);
  }

  return theResults(true, 'All values assigned to hash', functionName, updatedHash);
};