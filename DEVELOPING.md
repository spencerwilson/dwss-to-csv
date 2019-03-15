# Developing

For those wanting to develop the program.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. All steps assume a Unix-like environment.

### Prerequisites

```sh
# Get a good version of `node` and `npm` on your $PATH
nvm install
```

### Installing

```sh
# Install dependencies
npm install

# Basic test: print the program version
npm start -- --version
```

## Running the tests

A precondition to the acceptance tests is having certain files in a res/ directory. Check the script to see what it expects.

```sh
# Run acceptance tests
npm run acceptance-tests -- "<workbook password to be used for all test workbooks>"
```

## Releasing

Run `npm run build`, publish, and tag the resulting ZIP archive.
