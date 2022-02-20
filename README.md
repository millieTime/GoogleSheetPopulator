# Overview

This project helped me learn more about Google's Apps Script and how it integrates with Sheets.

A teacher had a spreadsheet they would update with values from a csv every week to keep track of how all their students were doing. But it was tedious work, copying and pasting values, and would lend itself to errors when the csv's rows were not ordered the same each time. So I built this script to help them!

### The Files

The script is contained in the `Apps script code.docx` file.

`Sample Record.xlsx` is a spreadsheet similar to the one my teacher used. It must be uploaded to Google Drive and converted to a Google Sheet for this to work.

`grades.csv` is a csv file similar to the one my teacher would get her student's grades and attendance from. It must be uploaded to Google Drive for this to work.

`W08.csv` is a csv file similar to the one my teacher would use to get her student's self-reported engagement and time spent on the course. It must be uploaded to Google Drive for this to work.

### How to Use

1. Upload `grades.csv`, `W08.csv`, and `Sample Record.xlsx` to your Google Drive and convert the `.xlsx` file to a Google Sheet.
2. In `Sample Record`, click the Extensions menu item, and then Apps Script. This should open a new tab.
3. Copy all text from `Apps script code.docx` and paste it into the Apps Script tab that opened. Click Save and click Run.
4. When it asks for permissions, choose your Google account and give permission to see and download your files in Drive, and edit the Sheet it is attached to.
5. Return to the `Sample Record` tab and notice there is a new menu item! Click on it, and then either select Import Grades or Import Weekly Report and enter the name of the corresponding file. After a couple seconds, the spreadsheet will update with the values it finds in the csv file.

### Driving Goals

I wanted this software to be as safe and easily usable as possible.

If my script ever finds that a student name is missing, or one of the column names is not what it expected, it will abort and display an error pop up for the user. I decided to do this to minimize the chances that the script would be given erroneous inputs and ruin the sheet because of that. Instead of a garbage-in-garbage-out mindset, I wanted to make sure this script would be as free from damaging results as possible.

From semester to semester, the layout of the sheet may change a little bit. Because of this, I made all important cell references constants and put them at the top of the Script. Next semester, if the teacher needs to update anything, it will be very easy for them to change a single variable and no algorithms will need edited.

In case there does happen to be a bug, or something needs updated, I used clear variable names and plenty of comments. It should be easy for me or someone else with a little background it programming to make any edits they need.

### Full process

The following video shows the full process of setting up the script, running it, and gives a walkthrough of the code.

[Software Demo Video](https://youtu.be/0DOpxd6LG-g)

# Development Environment

The development of this code was extremely simple, thanks to Google. All my code was written and run in an Apps Script tab, with no special libraries.

Google Apps Script is built on Javascript. All basic JS functions work, but more functionality is provided through Google giving access to libraries that allow you to interact with Sheets, Docs, and other files in Google Drive.

# Useful Websites

* [Stack Overflow](https://stackoverflow.com)
* [How to deal with user input in csv files](https://stackoverflow.com/questions/1293147/example-javascript-code-to-parse-csv-data/14991797#14991797)
* [Apps Script Documentation](https://developers.google.com/apps-script/guides/sheets)
* [MDN Web Docs](https://developer.mozilla.org/en-US/)

# Future Work

* If a student has two last names, or two first names (separated by a space) this script won't work. I included a comment about this in the top of the `Apps script code.docx` file, as well as how to fix it. Because it requires changing my teacher's spreadsheet, I did not implement it. Instead, I made them aware of the problem so that if it were to pop up it could be quickly remedied.
* Make the code even more robust. In the case of a missing student or column, I'd rather have the script not write anything to the Sheet than let it write until it reached the missing value.