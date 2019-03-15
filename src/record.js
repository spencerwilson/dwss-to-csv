const { FailureKind } = require('./constants');
const { logger } = require('./logging');

const CURRENCY_FORMATTER = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
});

/**
 * @param {Function} predicate
 * @param {Array<*>} arr
 * @returns {Array<*>}
 *        A shallow copy of the given array, with the first element
 *        satisfying the given predicate moved to the front of the
 *        array.
 */
function moveToFront(predicate, arr) {
  arr = [].concat(arr);

  const i = arr.findIndex(predicate);
  const [el] = arr.splice(i, 1);
  arr.unshift(el);

  return arr;
}

/**
 * @returns {Boolean}
 *        True iff the given candidate record conforms to the given schema.
 */
exports.matches = (schema) => (candidate) => {
  // Ensure perner check is the first check performed,
  // since invalid perners are treated specially upstream.
  schema = moveToFront((s) => s.type === 'perner', schema);

  for (const columnSchema of schema) {
    const result = checkColumn(columnSchema, candidate);
    if (result.failure) {
      return result;
    }
  }

  return { failure: false };
};

function checkColumn(column, candidate) {
  const candidateValue = candidate[column.name];

  if (column.optional && !Boolean(candidateValue)) {
    return true;
  }

  function validationError(failureKind, expected) {
    const error = {
      failure: true,
      kind: failureKind,
      column: column.name,
    };

    switch (failureKind) {
      case FailureKind.TYPE:
        error.message = `Column "${column.name}" type mismatch: expected ${expected}, got ${typeof candidateValue}`;
        break;
      case FailureKind.STRUCTURE:
        error.message = `Column "${column.name}" structure mismatch: expected ${expected}, got "${candidate[column.name]}"`;
        break;
    }

    return error;
  }

  switch (column.type || 'string') {
    case 'string':
      if (typeof candidateValue !== 'string') return validationError(FailureKind.TYPE, 'string');
      break;
    case 'perner':
      if (typeof candidateValue !== 'string') return validationError(FailureKind.TYPE, 'string');

      if (!candidateValue.match(/^\d{8}$/)) {
        return validationError(FailureKind.STRUCTURE, '8 digits');
      }
      break;
    case 'ssn':
      if (typeof candidateValue !== 'string') return validationError(FailureKind.TYPE, 'string');

      if (!candidateValue.match(/^\d{3}-\d{2}-\d{4}$/)) {
        return validationError(FailureKind.STRUCTURE, 'nnn-nn-nnnn');
      }
      break;
    case 'currency':
      if (typeof candidateValue !== 'number') return validationError(FailureKind.TYPE, 'number');
      break;
    case 'percentage':
      if (typeof candidateValue !== 'number') return validationError(FailureKind.TYPE, 'number');
      break;
    default:
      throw new Error('Unknown column type: ' + column.type);
  }

  return { failure: false };
}

/**
 * @returns {Array}
 *        The record, formatted as an array of strings.
 */
exports.format = (schema) => (record) => {
  return schema.map((column) => {
    let value = record[column.name];

    if (column.optional && !Boolean(value)) {
      return '';
    }

    switch (column.type || 'string') {
      case 'currency':
        let isNegative = false;
        if (value < 0) {
          isNegative = true;
          value *= -1;
        }

        // 1234.655 -> '$1,234.66'
        const formatted = CURRENCY_FORMATTER.format(value);

        return isNegative ? `(${formatted})` : formatted;
      case 'percentage':
        value *= 100;
        return value.toFixed(2) + '%';
      default:
        return value;
    }
  });
};
