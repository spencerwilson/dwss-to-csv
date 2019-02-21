const CURRENCY_FORMATTER = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
});

/**
 * @returns {Boolean}
 *        True iff the given candidate record conforms to the given schema.
 */
exports.matches = (schema) => (candidate) => {
  return schema.every((column) => {
    const candidateValue = candidate[column.name];

    if (column.optional && !Boolean(candidateValue)) {
      return true;
    }

    switch (column.type || 'string') {
      case 'string':
        if (typeof candidateValue !== 'string') {
          return false;
        }
        break;
      case 'perner':
        if (typeof candidateValue !== 'string') {
          return false;
        }

        if (!candidateValue.match(/^\d{8}$/)) {
          return false;
        }
        break;
      case 'ssn':
        if (typeof candidateValue !== 'string') {
          return false;
        }

        if (!candidateValue.match(/^\d{3}-\d{2}-\d{4}$/)) {
          return false;
        }
        break;
      case 'currency':
        if (typeof candidateValue !== 'number') {
          return false;
        }
        break;
      case 'percentage':
        if (typeof candidateValue !== 'number') {
          return false;
        }
        break;
      default:
        throw new Error('Unknown column type: ' + column.type);
    }

    return true;
  });
};

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
