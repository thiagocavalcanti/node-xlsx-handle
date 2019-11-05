const config = require('xlsx-handle/config');

/**
 * Handle a row from xlsx
 * @param {string} value content of the cell
 * @param {string} type type of content
 * @return {any} the value formated
 */
const handleXlsxTypes = (value, type) => {
  switch (type) {
    case 'String':
      return value;
    case 'Boolean':
      return value === 'True';
    case 'Object':
      return JSON.parse(value);
    default:
      return value;
  }
};

/**
 * Remove special chars, used in dictionary.
 * @param {string} value content of the cell.
 * @return {string} the value formated.
 */
const removeSpecialChars = value => {
  const { dictionary } = config;
  let clearString = value;
  for (key in dictionary) {
    clearString = clearString.replace(dictionary.key, '');
  }
  return clearString;
};

/**
 * Validate primary key. Throw new error if any duplicate
 * @param {Array} primaryKeys All the primaryKeys
 */
const validatePrimaryKeys = primaryKeys => {
  const unique = [...new Set(primaryKeys)];
  unique.forEach(key => {
    const occorencies = primaryKeys.filter(primaryKey => primaryKey === key);
    if (occorencies.length > 1) {
      throw new Error(`Primary key is not unique in the sheet. Value: ${key}`);
    }
  });
};

/**
 * Handle xlsx head
 * @param {Array} head The head.
 *
 * @returns {Array} An array containaing: type -> type of the head; headerObjectPaths -> the path of the head; headerKey -> The header key
 */
const handleXlsxHead = head => {
  const [headerKey, type] = head.split(':').map(h => h.trim());
  const headerObjectPaths = headerKey.split('.').map(h => h.trim());
  return { type, headerObjectPaths, headerKey };
};

/**
 * Read the header
 * @param {Array} header The header.
 *
 * @returns {Object.Array}  The new header
 * @returns {Object.Number} The index of the id
 * @returns {Object.Array}  The array of the index of required cols
 */

const readHeader = header => {
  const { dictionary } = config;
  let id = -1;
  let required = [];
  const newHeader = header.map((h, index) => {
    switch (h[0]) {
      case dictionary.primaryKey:
        if (id !== -1) {
          throw new Error(`Two or more columns as ID. Require to be only one. Col: ${index + 1}`);
        } else {
          id = index;
          return handleXlsxHead(h.split(dictionary.primaryKey)[1]);
        }
      case dictionary.required:
        required.push(index);
        return handleXlsxHead(h.split(dictionary.required)[1]);
      default:
        return handleXlsxHead(h);
    }
  });

  return { header: newHeader, id, required };
};

module.exports = {
  handleXlsxTypes,
  removeSpecialChars,
  validatePrimaryKeys,
  handleXlsxHead,
  readHeader
};
