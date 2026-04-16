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
 * Each validator function receives the current attribute string and the
 * current state of the hash, and returns a result object with an updated
 * copy of the hash as data.
 *
 * Pre-conditions:
 *   - tuples is an array of [attributeString, valueString] pairs
 *   - hash has predefined keys with undefined values
 *   - validators is an array of functions with signature:
 *       (attributeString: string, hash: Object) => { ok, message, data: updatedHash }
 *
 * Skip conditions (handled before validator pipeline):
 *   - attributeString is empty or blank
 *   - attributeString is the word 'comment' (case insensitive)
 *
 * @param {Array<Array<string>>} tuples      - Array of [attributeString, valueString] pairs
 * @param {Object}               hash        - Predefined keys with undefined values
 * @param {Array<Function>}      validators  - Array of validator functions
 *
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = final state of hash after all tuples processed
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

  // --- Outer loop: each tuple ---
  Logger.log(`>> Start Outer Loop`);
  for (const tuple of tuples) {

    Logger.log(`>> OuterLoop: tuple[0]=${tuple[0]}, tuple[1]=${tuple[1]}`);
    // Guard: make sure this tuple is usable
    if (!Array.isArray(tuple) || tuple.length < 2) {
      return failResult(`Invalid tuple encountered: ${JSON.stringify(tuple)}`);
    }

    // --- Skip check: runs before the validator pipeline ---
    const skipCheck = okToSkip(tuple, currentHash);
    Logger.log(`SkipOK? ${skipCheck.ok}`);
    if (skipCheck.ok) {
      continue;  // move to next tuple cleanly
    }

    // --- Inner loop: each validator ---
    for (const validator of validators) {

       Logger.log(`>> >> Inner Loop, Running validator: ${validator.name} for tuple: ${JSON.stringify(tuple)}`);
       const result = validator(tuple, currentHash);
       Logger.log(`>> >> Inner Loop, ${validator.name} result:${result.ok} - ${result.message}`);

      // Any validator failure stops everything
      if (!result.ok) {
        return failResult(result.message);
      }

      // Carry the updated hash forward to the next validator
      currentHash = { ...result.data };
    }
  }

  // --- All tuples passed all validators ---
  return okResult('All tuples processed successfully', currentHash);
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
 * Fails if the attribute already has a defined value in the hash.
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeDuplicate = (tuple, hash) => {
  const attributeName = tuple[0];
  if (hash[attributeName] !== undefined) {
    return failResult(`Duplicate attribute found: ${attributeName}`);
  }
  return okResult('no duplicate found', { ...hash });
};


/**
 * Fails if the attribute is not a predefined key in the hash.
 * Unknown keys are not valid input.
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeKnownKey = (tuple, hash) => {
  const attributeName = tuple[0];
  if (!Object.prototype.hasOwnProperty.call(hash, attributeName)) {
    return failResult(`Unknown attribute: ${attributeName}`);
  }
  return okResult('known key', { ...hash });
};


/**
 * Assigns the value from the tuple to the correct key in the hash.
 * This should always be the last validator in the pipeline.
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = updated copy of hash with value assigned
 */
const assignValueToHash = (tuple, hash) => {
  const attributeName = tuple[0];
  const value         = tuple[1];
  const updatedHash   = { ...hash };

  updatedHash[attributeName] = value;
  Logger.log(`assignValueToHash updatedHash[${attributeName}]=${value}`);
  return okResult(`Assigned [${attributeName}]=[${value}]`, updatedHash);
};

/**
 * Fails if the tuple value is missing, not a string, or blank.
 *
 * @param {Array<string>} tuple - [attributeName, value]
 * @param {Object}        hash
 * @returns {{ ok: boolean, message: string, data: Object|null }}
 *   data = current hash unchanged
 */
const checkIsAttributeValueDefined = (tuple, hash) => {
  const attributeName = tuple[0];
  const value         = tuple[1];

  //Logger.log(`checkIsAttributeValueDefined attributeName=${attributeName} and value="${value}"`);

  if (value === undefined || value === null) {
    return failResult(`Value for attribute [${attributeName}] is undefined or null`);
  }

  if (typeof value !== 'string') {
    return failResult(`Value for attribute [${attributeName}] is not a string: ${value}`);
  }

  if (value.trim() === '') {
    return failResult(`Value for attribute [${attributeName}] is blank`);
  }

  return okResult(`Value for [${attributeName}] is valid`, { ...hash });
};