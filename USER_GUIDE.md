# User Guide

This guide walks through the purpose of the "DWSS to CSV" program and how to use it.

## What is DWSS to CSV?

### The problem

Periodically, you receive Excel spreadsheets from DWSS containing data describing employees and their relationship to the union. In particular, the data describes peoples' membership status, initiation fees, and routine wage deductions for union dues. This data is eventually imported into Voltar, where it underpins Voltar's query and reporting faculties. It's therefore imperative that the data is high quality.

What does "high quality" mean? In order for the import into Voltar to be successful, there are strict requirements that must be met on the data. For example, SSNs and currency amounts must be formatted in particular ways, and many fields must have *some* value (i.e., they may not be blank). Believe it or not, the data from DWSS does not always comply with these requirements.

What's more, even if it did, there are other steps that must be taken in order to transform the spreadsheet data into the CSV (comma-separated value) format that Voltar knows how to interpret. The spreadsheets come with things like header rows at the top of sheets, describing the column contents. These can be helpful to a human reviewing the spreadsheet, but they're superfluous to Voltar and in fact they must **not** be present; Voltar wants only those rows which have a meaningful record of data in them, and needs them to comply with the requirements described above.

### The solution

DWSS to CSV is a simple computer program that automates

1. Ensuring data records are valid.
2. Converting the data from an Excel file into CSV files that Voltar can import.

Often, no work is required from the user beyond feeding the DWSS-provided Excel file through the program. Effortlessly, three CSVs—one for each of the All Employees, Dues Deducted, Initiation Fees datasets—are produced which are immediately importable into Voltar.

Occasionally, there will be anomalies in the data for which there is no mechanical "recipe" to resolve. Say, some crucial number is missing in a particular record. What should you do? In these cases, DWSS to CSV surfaces the anomaly to the user so that they may use their domain expertise to resolve the issue.

The rest of this guide walks through how to use the program and what to do in cases where human intervention is needed.

## Basic use

When you install the program, you will find its two components:

1. An .exe file that is the "core" program. One can run this while using a Command Prompt, but this an advanced usage and is not the primary way to run the program.
2. A .bat file, or "batch file", which is comparatively very easy to use. Just drag and drop Excel files (.xlsx files) received from DWSS onto this batch file and the program will process the files.

### Installation

1. Download the latest release of the program from the [Releases](https://github.com/spencerwilson/dwss-to-csv/releases) page.

2. Unzip it to some location on your computer where it's safe (e.g., won't inadvertently get deleted).

3. Locate the batch file. It's called "dwss_to_csv_drag_and_drop.bat". This is the file onto which you may simply drag and drop Excel files.

   **The following steps are optional,** but might make the program easier to locate and run in the future.

4. Right-click the batch file and select "Create shortcut". This will create a shortcut file in the same folder as the batch file.

5. Drag and drop the created shortcut onto your Desktop. This shortcut behaves just like the file which it's a shortcut to: just drag and drop onto it and you're off to the races.

### Running the program

Say you have a file with the name "DWSS_April_2019.xlsx". If you drag and drop that onto the batch file (or a shortcut to it, as described above), a window will open with the following prompt:

```
? Password to use for protected workbooks: 
```

As the Excel files (also called ***workbooks***) that DWSS provides are often password-protected, this prompt is asking you "In case any of the workbooks you dragged onto me seem password-protected, what shall I use for the password?".

Type in a password and press Enter. Suppose this results in the following output:

```
? Password to use for protected workbooks: **********
[****************************************] 100% || 1/1 || 0.0 seconds remaining
Press any key to continue . . .
```

What we see here is a progress bar indicating how close the program is to being done. After a few seconds, `1/1` is shown, indicating that 1 out of 1 workbooks have been completed.

This relatively sparse output indicates that there were no anomalies in the input workbook's data. Hooray! 🎉 No further work is required from you. You can find the resulting CSV files in the same folder as the input Excel file. If the program detected that the data is from April 2019, then the CSV files will be named

* DWSS_2019_04_AllEmployees.csv
* DWSS_2019_04_DuesDeducted.csv
* DWSS_2019_04_InitiationFees.csv

These can be imported straight into Voltar.

***Tips:***

* You can drag and drop many Excel files onto the program all at once. You'll only be prompted for a password once, and the password will be used to access all of the password-protected workbooks that are given.
* When locating the output CSVs, it's helpful to sort the files by file name (like a dictionary would sort things). This way, the three CSVs from a given month will all be next to one another rather than scattered throughout the larger list.

## Handling anomalies

Occasionally, things won't go as smoothly as above and there will be anomalies in the data from DWSS. What does that look like?

```
? Password to use for protected workbooks: **********
[****************************************] 100% || 1/1 || 0.0 seconds remaining
warning: InitiationFees: 3 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
warning: AllEmployees: 4 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
```

In the above example the program was run with a workbook that contains anomalies. Let's go line by line:

```
warning: InitiationFees: 3 records omitted (2 of which for interesting reasons)
```

The InitiationFees dataset had 3 records (non-header rows) that **did not** comply with the expected format. A row will be omitted if it's *clearly* not a legitimate record (defined as: the row is totally blank, or has an invalid Perner), but also will be omitted if it has a valid Perner but is invalid in some other way. The latter kind—a row with valid Perner but otherwise invalid data—is said to be invalid for "interesting reasons". In the above example, 1 omitted record was clearly illegitimate and the other 2 were invalid for more nuanced, interesting reasons. Examples of interesting reasons include missing values for "EE %" or "Initiation Fee", or data that doesn't look quite right like a social security number with 8 digits instead of 9.

Where can we go to see the omitted records? The next line clues us in:

```
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
```

For each Excel file that's processed, a ***log file*** gets created that contains details about the processing. Following every message that records were omitted is the information about where to find all the information needed to understand what happened. This message tells you where to find the log file, but know that it's the same pattern as the CSVs: the log file gets output to the same folder as the Excel file that it corresponds with.

The final two lines are like the two we've just discussed: the first summarizes the number of omitted records in the dataset (this time "AllEmployees"; in the previous example, the dataset was "InitiationFees"), and the second conveys the relevant log file that contains the details.

```
warning: AllEmployees: 4 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
```

### The log file

Here's a sample log file, with some details elided:

```
info: WORKBOOK: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1.xlsx
info: Year/month inference result: 2018, 12
info: Dataset-to-sheet correspondence: [
  "InitiationFees => December  2018 IF",
  "AllEmployees => December  2018 Dues",
  "DuesDeducted => Dues"
]
info: 
info: ========== Starting dataset: InitiationFees ==========
...
warning: InitiationFees: 3 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
info: Writing CSV: /Users/ssw/code/dwss-to-csv/res/DWSS_2018_12_InitiationFees.csv, records: 101
info: ========== Completed: InitiationFees ==========
info: 
info: ========== Starting dataset: AllEmployees ==========
...
warning: AllEmployees: 4 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
info: Writing CSV: /Users/ssw/code/dwss-to-csv/res/DWSS_2018_12_AllEmployees.csv, records: 2338
info: ========== Completed: AllEmployees ==========
info: 
info: ========== Starting dataset: DuesDeducted ==========
...
warning: DuesDeducted: 14 records omitted (3 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
info: Writing CSV: /Users/ssw/code/dwss-to-csv/res/DWSS_2018_12_DuesDeducted.csv, records: 1549
info: ========== Completed: DuesDeducted ==========
info: 
info: Workbook processing complete
```

Piece by piece:

```
info: WORKBOOK: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1.xlsx
info: Year/month inference result: 2018, 12
info: Dataset-to-sheet correspondence: [
  "InitiationFees => December  2018 IF",
  "AllEmployees => December  2018 Dues",
  "DuesDeducted => Dues"
]
```

This preamble, which appears at the beginning of every log file, states the Excel file that correponds to the log file, the year and month that were inferred from the sheet names, and the answer to the question "Which sheet contains which dataset?"—at least as far as the program thought.

After the preamble there are three principal sections, one per dataset. Here's one of them:

```
info: ========== Starting dataset: InitiationFees ==========
...
warning: InitiationFees: 3 records omitted (2 of which for interesting reasons)
warning: Check the log file: /Users/ssw/code/dwss-to-csv/res/DWSS_TestCase_1-LOG.log
info: Writing CSV: /Users/ssw/code/dwss-to-csv/res/DWSS_2018_12_InitiationFees.csv, records: 101
info: ========== Completed: InitiationFees ==========
```

Here we see some of the same summary sections as the main program output, but in the `...` lie the details of records omitted from this dataset. The way to interpret that information is to examine take consecutive pairs of lines the following:

```
info: Omitted record: {"PSA Text":"NABET Eng Daily","Perner":"01234567","New SSN":"123-45-678","Last Name":"Dart","First Name":"Joe","CC":"1264","BA":"265","PA Text":"ABC TV Network - CA","PSA":"M200","ESG":"07","PA":"64","Cost Center":"5486099","EEWT":"8350","EE Amt":306.6,"EE %":0.022500440321728408}
info: ...reason: Column "Contrib EE" type mismatch: expected number, got undefined
```

This would have been categorized as an "interesting" reason for omission. It's saying that this record seemed to entirely lack a value in the "Contrib EE" column. Interesting indeed!

At this point you have to exercise your judgment about how to proceed. How important is this record? Can it be "repaired"? Maybe that'd involve a conversation with someone else, maybe even someone at DWSS. If you figure something out, you can hand-edit the data in the Excel file and rerun the program until you see that the record is no longer omitted. In contrast, if you're willing to have the row omitted (and as a result not imported into Voltar), no further action is needed; use the CSV that the program has output.

### A playbook

Here's a playbook for processing data:

1. Drag and drop DWSS Excel files onto the program's batch file.
2. For each `warning` message, take note of which log files and datasets therein had "interesting" omissions.
3. Open the log files and inspect the omissions and their reasons. For those you want to pursue further, use something decently unique in the record (e.g., the Perner, or maybe even the last name) to look up that record in the original spreadsheet.
4. On a case by case basis, decide whether you want to A) hand-alter the data in the spreadsheet and Save the file, and rerun the program so that the record is now a valid one, or B) don't change it, and thereby accept its omission.