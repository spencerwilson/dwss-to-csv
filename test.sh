#!/usr/bin/env bash -x

rm -f res/DWSS*2018_12.csv
node src/index.js -p "$1" res/DWSS_Reference.xlsx

diff -u res/DWSS_AllEmployees_2018_12.csv res/DWSS_Reference_AllEmployees.csv
diff -u res/DWSS_DuesDeducted_2018_12.csv res/DWSS_Reference_DuesDeducted.csv
diff -u res/DWSS_InitiationFees_2018_12.csv res/DWSS_Reference_InitiationFees.csv

rm -f res/DWSS*2018_12.csv
node src/index.js -p "$1" res/DWSS_TestCase_1.xlsx

diff -u res/DWSS_AllEmployees_2018_12.csv res/DWSS_TestCase_1_AllEmployees.csv
diff -u res/DWSS_DuesDeducted_2018_12.csv res/DWSS_TestCase_1_DuesDeducted.csv
diff -u res/DWSS_InitiationFees_2018_12.csv res/DWSS_TestCase_1_InitiationFees.csv
