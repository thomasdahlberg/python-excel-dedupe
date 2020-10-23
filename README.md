# Python Excel Dedupe

This program is the start of a suite of simple python programs to help everyday users of spreadsheets to automate some of their tasks.

I'm starting with a deduplication program because often I know I need to dedupe a data set before even opening the file and often it's the only work I have to do with that file.

Booting up Excel, as powerful as it is, often wastes a fair amount of time. In the pursuit of efficiency, here goes!

## Prerequisites

This is a CLI program meant to be called in a Python runtime environment in the terminal. It depends on the functionality of openpyxl, so installing that is a prerequesite.

## Getting Started

- Download the program folder from github and place in a working directory such as 'Desktop.'
- Add the Excel document that you want to manipulate to the 'Workdocs' folder.
- Open terminal and navigate to the program's directory.

## Calling the Program

In the terminal call 'dedupe.py' in python like so:

>$python3 dedupe.py

- dedupe.py will display a numbered list of excel (.xlsx) documents that are in the /workdocs directory and ask you to pick a document by number.
- After you have selected your spreadsheet document you will be given a numbered list of the sheets in the document. Select the sheet you want to work with.
- Next, a list of the column headers for that sheet will be displayed in the CLI and you will first be asked if you want to sort rows then if you want to deduplicate rows. You may perform either operation or both.
- If you are sorting you will be asked what column you want to sort by. Pick the number of the column by which you want to sort.
- If you are deduplicating rows you will them be asked what column or columns by which you want to deduplicate rows. For a single column deduplcation criteria, simply enter the column number. For multiple column criteria, enter each column number separated by a comma, no spaces.
- After entering this criteria, dedupe.py will sort/dedupe your sheet displaying the sorting and dedupe criteria you entered while it runs the process. It will also display the original number of rows in your sheet and the number of rows after the deduplication process if you chose to dedupe.
- The process completes by creating a new sheet on your document with the naming convention of <original_sheet_name-deduped> or <original_sheet_name-sorted> depending on which processes you ran. On that sheet the processed data is written and then saved. Your original sheet on the document is never rewritten or mutated, only read.


