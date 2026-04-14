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