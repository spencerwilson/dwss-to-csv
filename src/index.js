const fs = require('fs');
const path = require('path');
const util = require('util');

const CFB = require('cfb');
const csvStringify = require('csv-stringify/lib/sync');
const inquirer = require('inquirer');
const ProgressBar = require('progress');
const XLSX = require('xlsx');
const yargs = require('yargs');

const decrypt = require('./decrypt');
const { FailureKind } = require('./constants');
const Logging = require('./logging');
const logger = Logging.logger;
const Record = require('./record');
const Utils = require('./utils');
const Workbook = require('./workbook');

const MANIFEST = JSON.parse(fs.readFileSync(path.join(__dirname, '../package.json'), 'utf-8'));

// Write-once, the first time a password-protected workbook is encountered.
let PASSWORD;

let bar;

main();

async function main() {
  const argv = yargs
    .usage('Usage: $0 <file> [...]')
    .demandCommand(1, 'You need to provide at least one Excel workbook to process.')
    .option('p', {
      alias: 'password',
      describe: 'Password used to read password-protected workbooks',
      nargs: 1,
      type: 'string',
    })
    .version(MANIFEST.version)
    .example('$0 -p hunter2 DWSS_Reference.xlsx', 'Process the given workbook using password "hunter2"')
    .example('$0 a.xlsx b.xlsx c.xlsx', 'Extract CSVs from many workbooks all at once')
    .argv;

  if (argv.password) {
    PASSWORD = argv.password;
  }

  await passwordPromptIfNecessary();

  const workbookPaths = argv._.map((p) => path.resolve(p));
  bar = new ProgressBar('[:bar] :percent || :completed/:all || :eta seconds remaining', {
    total: workbookPaths.length * 3,
    width: 40,
    complete: '*',
    incomplete: '.',
  });

  let completed = 0;
  bar.tick(0, {
    completed,
    all: workbookPaths.length,
  });

  let errorOccurred = false;
  for (const wbPath of workbookPaths) {
    Logging.createWorkbookLogger(wbPath);
    Logging.changeLogger(wbPath);

    try {
      await processWorkbook(wbPath);
    } catch (err) {
      errorOccurred = true;
      setImmediate(() => logger.error(err.stack));
    }

    completed++;
    bar.tick({
      completed,
      all: workbookPaths.length,
    });
  }

  // Allow the logs to [almost surely] be flushed to disk.
  await new Promise((resolve) => setTimeout(resolve, 300));
}

/**
 * Side effects: emits CSVs, writes to stdout and log file.
 *
 * @param {String} wbPath
 *        Absolute path to a workbook.
 * @returns {Promise}
 *        Rejects if
 */
async function processWorkbook(wbPath) {
  logger.info(`WORKBOOK: ${wbPath}`);

  const wb = await deserializeWorkbook(wbPath);
  bar.tick();

  const sheets = Object.values(wb.Sheets);
  const yearMonthInference = Workbook.inferYearMonth(wb.SheetNames);
  if (yearMonthInference.kind === 'failure') {
    logger.warning(`Year/month inference failed: ${yearMonthInference.message}`);
  } else {
    const { year, month } = yearMonthInference.value;
    logger.debug(`Year/month inference result: ${year}, ${month}`);
  }

  const correspondence = Workbook.correspondence(wb);
  logger.info(`Dataset-to-sheet correspondence: ${Utils.serialize(correspondence)}`);

  Array.from(correspondence).forEach((entry) => {
    const dataset = entry[0];
    const csvFileName = Workbook.formatCsvName(yearMonthInference, wbPath, dataset);

    entry[1].sheet = wb.Sheets[entry[1].sheetName];
    const formattedRecords = processDataset(entry, csvFileName);

    writeCsv(csvFileName, formattedRecords);
  });
}

/**
 * @param {String} wbPath
 *        Absolute path to a workbook.
 * @returns {Promise.<Workbook, Error>}
 *        Rejects if
 *          - no file exists at the given path
 *          - file is not a CFB nor ZIP archive
 *          - any necessary decryption throws an error
 *          - deserialization (possibly of a post-decryption result) fails
 *        Fulfills otherwise with the deserialized workbook.
 */
async function deserializeWorkbook(wbPath) {
  let wbBuffer = await util.promisify(fs.readFile)(wbPath);

  // Assumption: A workbook file is either
  //   - an encrypted ECMA-376 document, packaged in a CFB in accord
  //     with MS-OFFCRYPTO.
  //   - an unencrypted, ZIP-packaged ECMA-376 document.
  let wbCfb;
  try {
    wbCfb = CFB.parse(wbBuffer, {});
  } catch (err) {
    logger.error(`Failed to parse file as CFB or ZIP: ${err.stack}`);
    throw err;
  }

  const isEncrypted = Boolean(CFB.find(wbCfb, '/EncryptedPackage'));

  if (isEncrypted) {
    try {
      wbBuffer = await decrypt(wbCfb, PASSWORD);
    } catch (err) {
      logger.error(`Decryption failed: ${err.stack}`);
    }
  }

  bar.tick();

  let result;
  try {
    result = XLSX.read(wbBuffer, { type: 'buffer' });
  } catch (err) {
    logger.error(`Workbook parse failed. Wrong password? ${err.stack}`);
  }

  return result;
}

/**
 * @returns {Promise}
 *        If PASSWORD is already defined (either by CLI option or by
 *        a previous prompt), fulfills immediately.
 *        Otherwise, fulfills after setting PASSWORD to the prompt's answer.
 */
async function passwordPromptIfNecessary() {
  if (typeof PASSWORD !== 'undefined') {
    return;
  }

  const answers = await inquirer.prompt([{
    name: 'password',
    message: 'Password to use for protected workbooks:',
  }]);

  PASSWORD = answers.password;
}

/**
 * @param {Array}
 *        An entry from Workbook.correspondence.
 * @param {String}
 *        The absolute path of the CSV to be saved.
 */
function processDataset([dataset, { sheet, headersResult }], csvFileName) {
  // TODO: Handle when the workbook lacks a dataset
  const schema = Workbook.SCHEMA[dataset];
  const allRows = XLSX.utils.sheet_to_json(sheet, {range: headersResult.headerRow});

  let invalidCount = 0;
  let invalidPernerCount = 0;

  const formattedRecords = allRows
    .filter((r) => {
      const result = Record.matches(schema)(r);

      if (result.failure) {
        if (result.column === 'Perner') {
          invalidPernerCount++;
        }
        invalidCount++;
        return false;
      }

      return true;
    })
    .map(Record.format(schema));

  // Log async so as not to come in during the middle of the progress bar!
  if (invalidCount > invalidPernerCount) {
    setImmediate(() => {
      logger.warning(`${Utils.description(dataset)}: ${invalidCount} record${invalidCount > 1 ? 's' : ''} omitted (${invalidPernerCount} due to invalid Perner)`);
      const logFile = path.join(logger.transports[1].dirname, logger.transports[1].filename);
      logger.warning(`Check the log file: ${logFile}`);
    });
  } else {
    setImmediate(() => {
      logger.info(`${Utils.description(dataset)}: ${invalidCount} record${invalidCount > 1 ? 's' : ''} omitted (${invalidPernerCount} due to invalid Perner)`);
    });
  }

  return formattedRecords;
}

/**
 * @param {String}
 *        Absolute path to a CSV file to write.
 * @param {Array.<Array.<String>>}
 *        Records to be written to the CSV.
 */
function writeCsv(filename, formattedRecords) {
  const output = csvStringify(formattedRecords, {
    header: false,
    recordDelimiter: '\r\n',
  });

  logger.info(`Writing CSV: ${filename}, records: ${formattedRecords.length}`);
  fs.writeFileSync(filename, output, 'utf-8');
}
