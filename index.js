const XLSX = require('xlsx');
const utils = require('./utils/general');
const xlsxUtils = require('./utils/xlsx');
const { dictionary } = require('./config');

/**
 * Controller of the xlsx
 * @param  {Array}  data file xlsx
 * @param  {Object} params params for configuration {type,debug}
 * @return {Object.Array}  the matrix xlsx
 * @return {Object.Array}  the header
 * @return {Object.Number}  the id index
 */
const convertXlsxToArray = (data, params) => {
  const { type = 'base64', debug = false } = params;
  try {
    data = data.split(`${type},`)[1];
    const benchmark = Date.now();
    const workbook = XLSX.read(data, { type });
    if (debug) xlsxUtils.sendLog('Time to read xlsx file:', benchmark);
    const [colsCount, rowsCount] = [workbook.Sheets[workbook.SheetNames[1]]['B1'].v, workbook.Sheets[workbook.SheetNames[1]]['B2'].v];
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const cells_names = Object.keys(worksheet).filter(c => !c.split('!')[1]);
    const header_cells_names = cells_names.splice(0, colsCount);
    const cells_values = cells_names.map(cell => worksheet[cell].v);
    const header_values = header_cells_names.map(cell => worksheet[cell].v);

    let xlsx = [];
    let { id, required, header } = xlsxUtils.readHeader(header_values);

    let blankValues = 0;
    for (let i = 0; i < rowsCount - 1; i++) {
      xlsx.push([]);
      for (let j = 0; j < colsCount; j++) {
        const cell_name = cells_names[i * colsCount + j - blankValues];
        if (cell_name && cell_name.charCodeAt(0) - 65 === j) {
          xlsx[i].push(cells_values[i * colsCount + j - blankValues]);
        } else {
          if (cell_name && (required.includes(j) || id === j)) {
            const colunm = String.fromCharCode(65 + j);
            throw new Error(`Missing required value in cell ${colunm}${i + 2}`);
          } else {
            xlsx[i].push('');
            blankValues++;
          }
        }
      }
    }

    xlsxUtils.validatePrimaryKeys(xlsx.map(row => row[id]));
    return {
      xlsx,
      header,
      id
    };
  } catch (error) {
    return { error };
  }
};

/**
 * Controller of the xlsx
 * @param {array} data array rowsXcols of the xlsx
 * @param {string} type type of content
 * @return {array} all the docs created
 */
const handleXlsx = (data, params) => {
  const { debug = false, subDocuments = 0 } = params;
  try {
    let benchmark = Date.now();
    const { xlsx, header, id, error } = convertXlsxToArray(data, params);

    if (debug) xlsxUtils.sendLog('Time to convert xlsx to matrix:', benchmark);
    if (error) {
      throw new Error(error);
    } else {
      let documents = [];
      benchmark = Date.now();
      for (let index = 0; index < xlsx.length; index++) {
        const row = xlsx[index];
        let doc = documents.find(d => {
          if (!d) return false;
          const { headerKey } = header[id];
          if (row[id] !== d[headerKey]) return false;
          return true;
        });

        // [x] Handle row
        if (doc) {
          _handleXlsxRow(doc, row, header, id);
        } else {
          doc = _handleXlsxRow({}, row, header, id);
          documents.push(doc);
        }
        const newProgress = utils.verifyProgress(index, xlsx.length);
        if (debug && newProgress) xlsxUtils.sendLog(`Handling file... ${newProgress}%`);
      }
      if (debug) xlsxUtils.sendLog('Time to handle file:', benchmark);
      return subDocuments ? { documents: utils.creatingSmallerArrays(documents, subDocuments) } : { documents };
    }
  } catch (error) {
    return { error };
  }
};

/**
 * Handle a row from xlsx
 * @param {Object} doc
 * @param {Array} row
 * @param {Array} header
 * @param {Array} id
 * @return {Object} a new/updated doc
 */
const _handleXlsxRow = (doc, row, header, id) => {
  header.forEach((h, index) => {
    const { type, headerKey } = h;

    // [x] Get Value
    const value = xlsxUtils.handleXlsxTypes(row[index], type ? type.trim() : undefined);

    // [x] Creating doc if necessary (only id)
    if (!Object.keys(doc).length) {
      doc[headerKey] = row[id];
    }

    // [x] Return if id
    if (id === index) return;

    utils.setValue(doc, headerKey, value);
  });
  return doc;
};

module.exports = {
  handleXlsx,
  convertXlsxToArray
};
