#!/usr/bin/env bash -x

rm -f res/DWSS_2018_12*.csv
node src/index.js -p "$1" res/DWSS_Reference.xlsx

diff -u res/DWSS_2018_12_AllEmployees.csv res/DWSS_Reference_AllEmployees.csv
diff -u res/DWSS_2018_12_DuesDeducted.csv res/DWSS_Reference_DuesDeducted.csv
diff -u res/DWSS_2018_12_InitiationFees.csv res/DWSS_Reference_InitiationFees.csv

rm -f res/DWSS_2018_12*.csv
node src/index.js -p "$1" res/DWSS_TestCase_1.xlsx

diff -u res/DWSS_2018_12_AllEmployees.csv res/DWSS_TestCase_1_AllEmployees.csv
diff -u res/DWSS_2018_12_DuesDeducted.csv res/DWSS_TestCase_1_DuesDeducted.csv
diff -u res/DWSS_2018_12_InitiationFees.csv res/DWSS_TestCase_1_InitiationFees.csv
