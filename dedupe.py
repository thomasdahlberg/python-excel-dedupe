# Sort and Dedupe Program

import openpyxl
import pprint
import os

print('Opening workbook')

os.chdir('../')
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['sheet1']
header_values = []
row_keys = range(2,sheet.max_row)
document = []

for i in range(1, sheet.max_column):
    header_values.append(sheet.cell(row=2, column=i).value)

print('Headers: ', header_values)
print('Row Keys: ', row_keys)

for key in row_keys:
    row_dict = {}
    for i in range(1, sheet.max_column):
        row_dict[header_values[i-1]] = sheet.cell(row=key, column=i).value
    document.append(row_dict)

print('Document: ', document)