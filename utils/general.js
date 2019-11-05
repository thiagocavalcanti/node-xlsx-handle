/**
 * Verify progress.
 * @param   {number} count The total of interactions so far.
 * @param   {number} total The total of interactions.
 * @returns {number} The new % of the progress (if not, return undefinied)
 */
const verifyProgress = (count, total) => {
  if (Math.floor((100 * count) / total) > Math.floor((100 * (count - 1)) / total)) return Math.round((100 * count) / total);
};

/**
 * Splicing the array of documents into sub arrays.
 * @param   {Array}  BigArray The array with all the documents.
 * @param   {Number} size     The size of the sub arrays.
 * @returns {Array}  The array with the sub arrays.
 */
const creatingSmallerArrays = (BigArray, size = 100) => {
  const arrayOfArrays = [];
  for (let i = 0; i < BigArray.length; i += size) {
    arrayOfArrays.push(BigArray.slice(i, i + size));
  }
  return arrayOfArrays;
};

/**
 * Get a value within an object
 * @param {Object} obj  The object
 * @param {String} path The path
 *
 * @returns the value
 */
const getValue = (obj, path) => {
  return path
    .replace(/\[(\w+)\]/g, '.$1')
    .replace(/^\./, '')
    .split('.')
    .reduce((acc, part) => acc && acc[part], obj);
};

/**
 * Set a value within an object
 * @param {Object} obj   The object
 * @param {String} path  The path
 * @param {Any}    value The value
 *
 */
const setValue = (obj = {}, path, value) => {
  let i,
    array = path.replace(/^\./, '').split('.');
  for (i = 0; i < array.length - 1; i++) {
    if (!obj[array[i]]) obj[array[i]] = {};
    obj = obj[array[i]];
  }
  obj[array[i]] = value;
};

/**
 * Convert miliseconds to time format.
 * @param {Number} millisec The amount of miliseconds.
 *
 * @return {String} The time formatted.
 *
 */
const timeConversion = millisec => {
  const seconds = (millisec / 1000).toFixed(1);

  const minutes = (millisec / (1000 * 60)).toFixed(1);

  const hours = (millisec / (1000 * 60 * 60)).toFixed(1);

  const days = (millisec / (1000 * 60 * 60 * 24)).toFixed(1);

  if (seconds < 60) {
    return seconds + ' Sec';
  } else if (minutes < 60) {
    return minutes + ' Min';
  } else if (hours < 24) {
    return hours + ' Hrs';
  } else {
    return days + ' Days';
  }
};

module.exports = {
  verifyProgress,
  creatingSmallerArrays,
  getValue,
  setValue,
  timeConversion
};
