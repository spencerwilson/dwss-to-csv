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
const Record = require('./record');
const Workbook = require('./workbook');

const MANIFEST = JSON.parse(fs.readFileSync(path.join(__dirname, '../package.json'), 'utf8'));

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

  for (const wbPath of workbookPaths) {
    await processWorkbook(wbPath);

    completed++;
    bar.tick({
      completed,
      all: workbookPaths.length,
    });
  }
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
  const wb = await deserializeWorkbook(wbPath);
  bar.tick();
  const sheets = Object.values(wb.Sheets);

  Array.from(Workbook.correspondence(wb)).forEach((entry) => {
    const dataset = entry[0];
    const csvFileName = Workbook.formatCsvName(wb, wbPath, dataset);

    const formattedRecords = processDataset(entry);

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
  const wbCfb = CFB.parse(wbBuffer, {});
  const isEncrypted = Boolean(CFB.find(wbCfb, '/EncryptedPackage'));

  if (isEncrypted) {
    wbBuffer = await decrypt(wbCfb, PASSWORD);
  }

  bar.tick();

  return XLSX.read(wbBuffer, { type: 'buffer' });
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
 * @param {String} 
 *        The absolute path of the CSV to be saved.
 * @param {Array}
 *        An entry from Workbook.correspondence.
 */
function processDataset([dataset, { sheet, headersResult }]) {
  // TODO: Handle when the workbook lacks a dataset
  const schema = Workbook.SCHEMA[dataset];
  const allRows = XLSX.utils.sheet_to_json(sheet, {range: headersResult.headerRow});

  const formattedRecords = allRows
    .filter(Record.matches(schema))
    .map(Record.format(schema));

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

  fs.writeFileSync(filename, output, 'utf-8');
}
