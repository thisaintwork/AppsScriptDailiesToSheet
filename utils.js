// utils.js
/* 1️⃣
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

const okResult = (message, data = null) => ({
  ok:      true,
  message: message,
  data:    data
});

const failResult = (message) => ({
  ok:      false,
  message: message,
  data:    null
});


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
const processTuplesThroughValidators = (tuples, hash, validators) => {
  Logger.log(`Entered: processTuplesThroughValidators`);

  // --- Guard: validate inputs ---
  if (!tuples || !Array.isArray(tuples)) {
    return failResult('tuples must be an array');
  }
  if (!hash || typeof hash !== 'object') {
    return failResult('hash must be an object');
  }
  if (!validators || !Array.isArray(validators)) {
    return failResult('validators must be an array');
  }

  // --- Start with a clean copy of the incoming hash ---
  let currentHash = { ...hash };

  // --- Loop: each validator runs once against the full tuples array ---
  Logger.log(`>> Start Validator Loop`);
  for (const validator of validators) {

    Logger.log(`>> Running validator: ${validator.name}`);
    const result = validator(tuples, currentHash);
    Logger.log(`>> ${validator.name} result: ${result.ok} - ${result.message}`);

    // Any validator failure stops everything
    if (!result.ok) {
      return failResult(result.message);
    }

    // Carry the updated hash forward to the next validator
    currentHash = { ...result.data };
  }

  // --- All validators passed ---
  return okResult('All validators processed successfully', currentHash);
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
  const trimmed = tuple[0].trim().toLowerCase();

  if (trimmed === '') {
    return okResult('empty attribute - skip', { ...hash });
  }

  if (trimmed === 'comment') {
    return okResult('comment - skip', { ...hash });
  }

  return failResult('not a skip condition - continue processing');
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
    return failResult(`Duplicate attribute(s) found: ${duplicates.join(', ')}`);
  }

  return okResult('No duplicates found', { ...hash });
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

  if (unknownKeys.length > 0) {
    return failResult(`Unknown attribute(s) found: ${unknownKeys.join(', ')}`);
  }

  return okResult('All attribute names are known keys', { ...hash });
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
  const invalidValues = [];

  for (const tuple of tuples) {

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    const attributeName = tuple[0];
    const value         = tuple[1];

    if (value === undefined || value === null) {
      invalidValues.push(`[${attributeName}] is undefined or null`);
      continue;
    }

    if (typeof value !== 'string') {
      invalidValues.push(`[${attributeName}] is not a string: ${value}`);
      continue;
    }

    if (value.trim() === '') {
      invalidValues.push(`[${attributeName}] is blank`);
      continue;
    }
  }

  if (invalidValues.length > 0) {
    return failResult(`Invalid value(s) found: ${invalidValues.join(', ')}`);
  }

  return okResult('All attribute values are valid', { ...hash });
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
  const updatedHash = { ...hash };

  for (const tuple of tuples) {

    // Skip empty and comment rows
    if (okToSkip(tuple, hash).ok) continue;

    const attributeName = tuple[0];
    const value         = tuple[1];

    updatedHash[attributeName] = value;
    Logger.log(`assignValuesToHash updatedHash[${attributeName}]=${value}`);
  }

  return okResult('All values assigned to hash', updatedHash);
};