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

const handleXlsxHeader = header => {
  return header.map(head => {
    const [headerKey, type] = head.trim().split(':');
    const headerObjectPaths = headerKey.trim().split('.');
    return { type, headerObjectPaths, headerKey };
  });
};

module.exports = {
  handleXlsxTypes,
  removeSpecialChars,
  validatePrimaryKeys,
  handleXlsxHeader
};
