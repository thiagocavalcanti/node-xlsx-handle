const XLSX = require('xlsx');
const utils = require('./utils/general');
const xlsxUtils = require('./utils/xlsx');
const { dictionary } = require('./config');

/**
 * Controller of the xlsx
 * @param  {array}  data file xlsx
 * @param  {object} params params for configuration {type,debug}
 * @return {array}  the matrix xlsx
 * @return {array}  the header
 * @return {array}  the ids index
 */
const convertXlsxToArray = (data, params) => {
  const { type = 'base64', debug = false } = params;
  try {
    data = data.split(`${type},`)[1];
    const benchmark = Date.now();
    const workbook = XLSX.read(data, { type });
    if (debug) console.log('Time to read xlsx file: ', utils.timeConversion(Date.now() - benchmark));
    delete benchmark;
    const [cols, rows] = [workbook.Sheets[workbook.SheetNames[1]]['B1'].v, workbook.Sheets[workbook.SheetNames[1]]['B2'].v];
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const cells_names = Object.keys(worksheet).filter(c => !c.split('!')[1]);
    const cells = cells_names.map(cell => worksheet[cell].v);
    let xlsx = [];
    let header = cells.filter((c, index) => index < cols);
    let ids = [];
    let required = [];
    header = header.map((h, index) => {
      switch (h[0]) {
        case dictionary.primaryKey:
          ids.push(index);
          return h.split(dictionary.primaryKey)[1];
        case dictionary.required:
          required.push(index);
          return h.split(dictionary.required)[1];
        default:
          return h;
      }
    });

    let blankValues = 0;
    for (let i = 1; i < rows; i++) {
      xlsx.push([]);
      for (let j = 0; j < cols; j++) {
        const cell_name = cells_names[i * cols + j - blankValues];
        if (cell_name && cell_name.charCodeAt(0) - 65 === j) {
          xlsx[i - 1].push(cells[i * cols + j - blankValues]);
        } else {
          if (cell_name && (required.includes(j) || ids.includes(j))) {
            const colunm = String.fromCharCode(65 + j);
            throw new Error(`Missing required value in cell ${colunm}${i + 1}`);
          } else {
            xlsx[i - 1].push('');
            blankValues++;
          }
        }
      }
    }

    xlsxUtils.validatePrimaryKeys(xlsx.map(row => row[ids[0]]));

    return {
      xlsx,
      header,
      ids
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
    const { xlsx, header, ids, error } = convertXlsxToArray(data, params);
    header = xlsxUtils.handleXlsxHeader(header);
    if (debug) console.log('Time to convert xlsx to matrix: ', utils.timeConversion(Date.now() - benchmark));
    if (error) {
      throw new Error(error);
    } else {
      let documents = [];
      benchmark = Date.now();
      for (let index = 0; index < xlsx.length; index++) {
        const row = xlsx[index];
        let doc = documents.find(d => {
          if (!d) return false;
          for (let i = 0; i < ids.length; i++) {
            const header_key = header[ids[i]].trim().split(':')[0];
            if (row[ids[i]] !== d[header_key]) return false;
          }
          return true;
        });

        // [x] Handle row
        if (doc) {
          handleXlsxRow(doc, row, header, ids);
        } else {
          doc = handleXlsxRow({}, row, header, ids);
          documents.push(doc);
        }
        const newProgress = utils.verifyProgress(index, xlsx.length);
        if (debug && newProgress) console.log(`Handling file... ${newProgress}%`);
      }
      if (debug) console.log('Time to handle file: ', utils.timeConversion(Date.now() - benchmark));
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
 * @param {Array} ids
 * @return {Object} a new/updated doc
 */
const handleXlsxRow = (doc, row, header, ids) => {
  // [x] Model: { array: '', name: '', value: '', index: -1 }
  let key_array = [];

  header.forEach((h, index) => {
    const { type, headerObjectPaths } = h;

    // [x] Get Value
    const value = xlsxUtils.handleXlsxTypes(row[index], type ? type.trim() : undefined);

    // [x] Check nested objects. Ex: A.B.[C].D
    let object_ref = {};

    // [x] Creating doc if necessary (only ids)
    if (!Object.keys(doc).length) {
      for (let i = 0; i < ids.length; i++) {
        const header_key = header[i].trim().split(':')[0];
        doc[header_key] = row[ids[i] - 1];
      }
    }

    // [x] Return if ids
    if (ids.find(i => i === (index + 1).toString())) return;

    // [x] Interects over nested objects
    headerObjectPaths.forEach((obj, obj_index) => {
      // [x] Get reference
      if (!obj_index) object_ref = doc;
      let key_array_item = key_array.find(a => a.index === obj_index);
      if (key_array_item) {
        if (key_array_item.array === obj.replace(/[\[\]]/g, '')) {
          object_ref = object_ref[key_array_item.array].find(a => a[key_array_item.name] === key_array_item.value);
          return;
        } else delete key_array_item;
      }
      // [x] Check for array
      const array = obj.replace(/[\[\]]/g, '');

      // [x] Not Array
      if (array.length === obj.length) {
        // [x] Not last object
        if (obj_index !== h_objects.length - 1) {
          if (!object_ref[obj]) object_ref[obj] = {};
          object_ref = object_ref[obj];
        }
        // [x] Last object
        else {
          // [x] Object of an array
          if (Array.isArray(object_ref))
            object_ref.push({
              [obj]: value
            });
          // [x] Regular value of a key
          else object_ref[obj] = value;
        }
      }

      // [x] Array
      else {
        // [x] Check if array already exists
        if (!object_ref[xlsxUtils.removeSpecialChars(array)]) object_ref[xlsxUtils.removeSpecialChars(array)] = [];

        // [x] Not last object
        if (obj_index !== h_objects.length - 1) {
          // [x] Check if key array
          if (array.includes(dictionary.primaryKey)) {
            key_array.push({
              array: xlsxUtils.removeSpecialChars(array),
              name: h_objects[obj_index + 1],
              value,
              index: obj_index
            });
          }
          object_ref = object_ref[xlsxUtils.removeSpecialChars(array)];

          // [x] Check if array already exists to get reference
          let array_item = object_ref.find(a => a[h_objects[obj_index + 1]] === value);
          if (array_item) object_ref = array_item;
        }
        // [x] Last object
        else {
          // [x] Check if the value is an array
          if (Array.isArray(value)) {
            value.forEach(v => {
              if (!object_ref[array].includes(v)) object_ref[array].push(v);
            });
          } else {
            value.split(';').forEach(v => !object_ref[array].includes(v) && object_ref[array].push(v));
          }
        }
      }
    });
  });
  return doc;
};

module.exports = {
  handleXlsx,
  convertXlsxToArray
};
