const config = require('xlsx-handle/config');
const utils = require('./general');

/**
 * Handle xlsx head
 * @type Private
 * @param {Array} head The head.
 *
 * @returns {Array} An array containaing: type -> type of the head; headerKey -> The header key
 */
const _handleXlsxHead = head => {
  const [headerKey, type] = head.split(':').map(h => h.trim());
  return { type, headerKey };
};

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
          return _handleXlsxHead(h.split(dictionary.primaryKey)[1]);
        }
      case dictionary.required:
        required.push(index);
        return _handleXlsxHead(h.split(dictionary.required)[1]);
      default:
        return _handleXlsxHead(h);
    }
  });

  return { header: newHeader, id, required };
};

const sendLog = (text, startTime) => {
  let sendText = `[XLSX-HANDLE] ${text} `;
  if (startTime) {
    sendText += utils.timeConversion(Date.now() - startTime);
  }
  console.log(sendText);
};

module.exports = {
  handleXlsxTypes,
  validatePrimaryKeys,
  readHeader,
  sendLog
};
