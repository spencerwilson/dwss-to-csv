const fs = require('fs');
const path = require('path');

const XLSX = require('xlsx');

const { logger } = require('./logging');
const Utils = require('./utils');

// Object<String, Dataset>
const Datasets = Utils.symbolMirror([
  'AllEmployees',
  'DuesDeducted',
  'InitiationFees',
]);
exports.datasets = Datasets;

// Object<Dataset, Array<ColumnDescriptor>>
const SCHEMA = JSON.parse(fs.readFileSync(path.join(__dirname, 'schema.json')));
Object.keys(Datasets).forEach((dsDescription) => {
  SCHEMA[Datasets[dsDescription]] = SCHEMA[dsDescription];
  delete SCHEMA[dsDescription];
});
exports.SCHEMA = SCHEMA;

/**
 * @returns {Boolean}
 *        True iff the given map contains at least the entries keyed by
 *        those in the given dataset.
 */
function isDataset(dataset, map) {
  const columnNames = SCHEMA[dataset].map(({name}) => name);
  return columnNames.every(map.has.bind(map));
}

/**
 * @param {Array<Sheet>} workbook
 * @returns Map<Symbol, Maybe<{sheet, headerInfo}>>
 */
exports.correspondence = function correspondence(wb) {
  const result = new Map();
  const remainingDatasets = new Set(Object.values(Datasets));

  nextSheet: for (const sheet of Object.values(wb.Sheets)) {
    // TODO: Handle when there are no headers
    const headersResult = findHeaders(sheet);

    for (const dataset of remainingDatasets) {
      if (isDataset(dataset, headersResult.columns)) {
        result.set(dataset, { sheet, headersResult });
        remainingDatasets.delete(dataset);
        continue nextSheet;
      }
    }
  }

  // TODO make placeholder for non-found datasets?

  return result;
}

// sheet -> { headerRow: Number, columns: (name => column Number) }
function findHeaders(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const cellsPerRow = range.e.c - range.s.c + 1;
  const result = {};

  // The first row for whom a majority of cells have non-empty string values...
  let headerRow;
  for (headerRow of Utils.range(range.s.r, range.e.r + 1)) {
    const stringCellCount = Array.from(cellsInRow(sheet, headerRow)).reduce((count, cell) => {
      if (cell && cell.t === 's' && cell.v.length !== 0) {
        count += 1;
      }
      return count;
    }, 0);

    const majorityOfCellsAreStrings = (stringCellCount / cellsPerRow) > 0.5;
    if (majorityOfCellsAreStrings) {
      break;
    }
  }

  const columns = Array.from(cellsInRow(sheet, headerRow)).reduce((map, cell, i) => {
    if (!cell) {
      // e.g., in the case of merged cells.
      return map;
    }

    if (!cell.v) {
      // e.g., in the case of unmerged but simply empty cells.
      return map;
    }

    const columnName = String(cell.v).trim();
    map.set(columnName, i);
    return map;
  }, new Map());

  return {
    headerRow,
    columns,
  }
}

function* cellsInRow(sheet, row) {
  const sheetRange = XLSX.utils.decode_range(sheet['!ref']);

  for (const i of Utils.range(sheetRange.s.c, sheetRange.e.c + 1)) {
    const cellAddress = XLSX.utils.encode_cell({r: row, c: i});
    yield sheet[cellAddress];
  }
}

const MONTHS = [
  ['January', 'Jan'],
  ['February', 'Feb'],
  ['March', 'Mar'],
  ['April', 'Apr'],
  ['May'],
  ['June', 'Jun'],
  ['July', 'Jul'],
  ['August', 'Aug'],
  ['September', 'Sept'],
  ['October', 'Oct'],
  ['November', 'Nov'],
  ['December', 'Dec'],
].map((aliases) => aliases.map((a) => a.toLowerCase()));

/**
 * @params {Array<String>} sheetNames
 * @returns {Result<(Year, Month)>}
 */
exports.inferYearMonth = function inferYearMonth(sheetNames) {
  //find a sheet name that contains a \d{4}, and use it
  // Scan the sheet names and infer that the first sequence of 4 digits
  // that's found is the year.
  let matchedYear;
  for (const name of sheetNames) {
    const result = name.match(/\d{4}/);

    if (result) {
      matchedYear = Number.parseInt(result[0]);
      break;
    }
  }

  if (!matchedYear) {
    return { kind: 'failure', message: 'No year found' };
  }

  let matchedMonth;
  let lowerSheetNames = sheetNames.map((n) => n.toLowerCase());

  nextSheetName: for (const name of lowerSheetNames) {
    for (const mIndex in MONTHS) {
      const aliases = MONTHS[mIndex];
      const aliasMatch = aliases.some((a) => name.includes(a));

      if (aliasMatch) {
        matchedMonth = Number.parseInt(mIndex) + 1;
        break nextSheetName;
      }
    }
  }

  if (!matchedMonth) {
    return { kind: 'failure', message: 'No month found' };
  }

  return {
    kind: 'success',
    value: {
      year: matchedYear,
      month: matchedMonth,
    },
  };
}

exports.formatCsvName = function formatCsvName(inferenceResult, wbPath, dataset) {
  const description = dataset.toString().match(/\((.*)\)/)[1];
  const namePrefix = `DWSS_${description}`;
  const partialResult = {
    dir: path.dirname(wbPath),
    // name: filled in below
    ext: '.csv',
  };

  if (inferenceResult.kind === 'failure') {
    return path.format(Object.assign(partialResult, {
      name: namePrefix + '_' + path.basename(wbPath),
    }));
  }

  const { year, month } = inferenceResult.value;

  const name = [
    namePrefix,
    year.toString(),
    month.toString().padStart(2, '0'),
  ].join('_');

  return path.format(Object.assign(partialResult, { name }));
};

