const fs = require('fs');
const path = require('path');
const util = require('util');

const winston = require('winston');
let currentLoggerId;

// An object that acts like a winston.Logger, but the particular logger
// instance that it is is decided dynamically.
exports.logger = new Proxy({}, {
  get: function (target, property, receiver) {
    return getCurrentLogger()[property];
  },
});

function getLogName(wbPath) {
  return path.format({
    dir: process.cwd(),
    name: path.parse(wbPath).name,
    ext: '.log',
  });
}

function getCurrentLogger() {
  return winston.loggers.get(currentLoggerId);
}

/**
 * Add a logger to the container for the given workbook.
 */
exports.createWorkbookLogger = function createWorkbookLogger(wbPath) {
  winston.loggers.add(wbPath, {
    level: 'debug',
    format: winston.format.combine(
      winston.format.simple(),
      winston.format.colorize(),
    ),
    transports: [
      new winston.transports.File({
        filename: getLogName(wbPath),
        options: {
          // Overwrite any existing log file
          flags: 'w',
        },
      }),
    ],
  });
};

exports.changeLogger = function changeLogger(wbPath) {
  currentLoggerId = wbPath;
};

exports.printWorkbookLogs = async function printWorkbookLogs(wbPath) {
  const logs = await util.promisify(fs.readFile)(getLogName(wbPath), 'utf-8');
  console.log(logs);
};
