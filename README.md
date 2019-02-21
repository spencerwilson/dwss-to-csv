# DWSS to CSV

DWSS to CSV is a command line program that reads Excel workbooks (AKA .xlsx files) provided by Disney Worldwide Shared Services (DWSS) and emits CSVs with clean data.

Though it was created for that specific purpose, the source code itself is unaware of the workbook format being considered. At runtime a JSON document that describes the datasets within the workbook is read from the filesystem. By altering this declarative schema file, the program can be adapted (even by less-technical users) to keep pace with evolution in a workbook's schema—or repurposed entirely, for totally new datasets in other workbooks.

Developers: see [DEVELOPING](DEVELOPING.md).

## Installing

Go to Releases and find a build for your platform.

## Using

```sh
# See usage options
dwss_to_csv.exe --help
```

## License

This project is licensed under the MIT License; see the [LICENSE](LICENSE) file for details.

## Acknowledgments

* SheetJS (NPM: [xlsx](https://www.npmjs.com/package/xlsx)) maintainers, for producing a great library for working with xlsx files (and for friendly email correspondence!)
* Dave T. Johnson ([dtjohnson](https://github.com/dtjohnson)), whose decryption code is used by this project, enabling support for reading password-protected xlsx files. See [src/encryption.js](src/encryption.js) for details.
