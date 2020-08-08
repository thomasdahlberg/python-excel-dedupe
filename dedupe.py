# Sort and Dedupe Program

import openpyxl
import pprint
import os

print('Opening workbook')

os.chdir('../')
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['sheet1']
header_values = []
row_keys = range(3,sheet.max_row)
document = []

for i in range(1, sheet.max_column + 1):
    header_values.append(sheet.cell(row=2, column=i).value)

print('Headers: ', header_values)
print('Row Keys: ', row_keys)

for key in row_keys:
    row_dict = {}
    for i in range(1, sheet.max_column +1):
        row_dict[header_values[i-1]] = sheet.cell(row=key, column=i).value
    document.append(row_dict)

def sort_by(e):
    return e['Admit Term']

document.sort(key=sort_by)

print('Document: ', document)

# wb.remove(wb['Sorted'])

wb.create_sheet(index=1, title='Sorted')
sorted_sheet = wb['Sorted']
sorted_sheet.append(header_values)

for i in range(1, len(document)):
    sorted_sheet.append(list(document[i].values()))

wb.save(filename='example.xlsx')

print('Sheetnames: ', wb.sheetnames)