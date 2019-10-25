const XLSX = require('xlsx');

/**
 * Dictionary for special chars.
 */
const dictionary = {
  primaryKey: '$',
  required: '*'
};

/**
 * Controller of the xlsx
 * @param {array} data file xlsx
 * @param {string} type formatation of file (default = base64)
 * @return {array} the matrix xlsx
 * @return {array} the header
 * @return {array} the ids index
 */
const convertXlsxToArray = (data, type = 'base64') => {
  try {
    data = data.split(`${type},`)[1];
    const workbook = XLSX.read(data, { type });
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
        if (cell_name.charCodeAt(0) - 65 === j) {
          xlsx[i - 1].push(cells[i * cols + j - blankValues]);
        } else {
          if (required.includes(j) || ids.includes(j)) {
            const colunm = String.fromCharCode(65 + j);
            throw new Error(`Missing required value in cell ${colunm}${i + 1}`);
          } else {
            xlsx[i - 1].push('');
            blankValues++;
          }
        }
      }
    }
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
const handleXlsx = (data, type = 'base64') => {
  try {
    const { xlsx, header, ids, error } = convertXlsxToArray(data, type);
    if (error) {
      throw new Error(error);
    } else {
      let documents = [];
      xlsx.forEach(row => {
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
      });
      return { documents };
    }
  } catch (error) {
    return { error };
  }
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
 * Handle a row from xlsx
 * @param {string} value content of the cell
 * @param {string} type type of content
 * @return {any} the value formated
 */
const removeSpecialChars = value => {
  let clearString = value;
  for (key in dictionary) {
    clearString = clearString.replace(dictionary.key, '');
  }
  return clearString;
};

/**
 * Handle a row from xlsx
 * @param {obj} doc
 * @param {array} row
 * @param {array} header
 * @param {array} ids
 * @return {obj} a new/updated doc
 */
const handleXlsxRow = (doc, row, header, ids) => {
  // [x] Model: { array: '', name: '', value: '', index: -1 }
  let key_array = [];

  header.forEach((h, index) => {
    // [x]Check type
    const [h_rest, type] = h.trim().split(':');

    // [x] Get Value
    const value = handleXlsxTypes(row[index], type ? type.trim() : undefined);

    // [x] Check nested objects. Ex: A.B.[C].D
    let object_ref = {};
    const h_objects = h_rest.trim().split('.');

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
    h_objects.forEach((obj, obj_index) => {
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
        if (!object_ref[removeSpecialChars(array)]) object_ref[removeSpecialChars(array)] = [];

        // [x] Not last object
        if (obj_index !== h_objects.length - 1) {
          // [x] Check if key array
          if (array.includes(dictionary.primaryKey)) {
            key_array.push({
              array: removeSpecialChars(array),
              name: h_objects[obj_index + 1],
              value,
              index: obj_index
            });
          }
          object_ref = object_ref[removeSpecialChars(array)];

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
