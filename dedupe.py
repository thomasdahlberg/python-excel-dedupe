# Sort and Dedupe Program

import openpyxl
import pprint
import os

print('Opening workbook')

os.chdir('../')
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['sheet1']
header_values = []

for i in range(1, sheet.max_column):
    header_values.append(sheet.cell(row=2, column=i).value)

print('Headers: ', header_values)