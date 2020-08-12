# Python Excel Dedupe

This program is the start of what may become a suite of simple python programs to help everyday users of spreadsheets automate some of their tasks.

I'm starting with a dedupe program because often I know I need to dedupe a data set before even opening the file and often it's the only work I have to do with that file.

Booting up Excel, as powerful as it is, often wastes a fair amount of time. In the pursuit of efficiency, here goes!

## Prerequisites

This is a CLI program meant to be called in a Python runtime environment in the terminal. It depends on the functionality of openpyxl, so installing that is a prerequesite.

## Getting Started

- Download the program folder from github and place in a working directory such as 'Desktop.'
- Add the Excel document that you want to manipulate to the 'Workdocs' folder.
- Open terminal and navigate to the program's directory.

## Calling the Program

In the terminal call 'dedupe.py' in python like so:

>$python3 dedupe.py <file_name> <1st_arg> <2nd_arg>

Please note:
- <file_name> is the name of the spreadsheet file without its file extension identifier ('.xlsx').
- <1st_arg> is the place holder for the dedupe column(s) criteria. In other words the program will only keep rows with unique values in the specified column(s). You may list one integer to specify the dedupe column or any number of integers separated by a comma, no spaces.
- <2nd arg> is a place holder for the sort column criteria. It accepts one integer value.
- Both arguments are zero-indexed.

## Example Call of 'dedupe.py'

In the example below the user is calling 'dedupe.py' on spreadsheet 'example' to deduplicate rows based on unique values in column 0, 1 and 4 and sort the resulting rows by column 2.

>$python3 dedupe.py example 0,1,4 2

