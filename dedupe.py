# Sort and Dedupe Program

import openpyxl
import pprint
import os
import sys

print('Opening workbook')

file_arg = sys.argv[1]
dedupe_arg = sys.argv[2]


os.chdir('./workdocs')
wb = openpyxl.load_workbook(f"{file_arg}.xlsx")
sheet = wb['sheet1']
header_values = []
row_keys = range(3,sheet.max_row)
document = []

for i in range(1, sheet.max_column + 1):
    header_values.append(sheet.cell(row=2, column=i).value)

print('Headers: ', header_values)
print('Row Keys: ', row_keys)

filter_col = header_values[int(dedupe_arg)]

for key in row_keys:
    row_dict = {}
    for i in range(1, sheet.max_column +1):
        row_dict[header_values[i-1]] = sheet.cell(row=key, column=i).value
    document.append(row_dict)


deduped_document = []
unique_values = []

# for sub in document:
#     unique_values.append(sub[unique_key])
# unique_values = list(set(unique_values))

# print('Unique IDs: ', str(unique_values))
# print('# of IDs: ', len(unique_values))

# deduped_document = [d for d in document if d['ID'] in unique_values]

for d in document:
    if d[filter_col] not in unique_values:
        unique_values.append(d[filter_col])
        deduped_document.append(d)


# print('Original length: ', len(document))
print('Deduped length: ', len(deduped_document))


def sort_by(e):
    return e['Admit Term']

# document.sort(key=sort_by)

# print('Document: ', document)


# wb.create_sheet(index=1, title='Sorted')
# sorted_sheet = wb['Sorted']
# sorted_sheet.append(header_values)

# for i in range(1, len(document)):
#     sorted_sheet.append(list(document[i].values()))
sheet_exists = wb['Deduped']

if sheet_exists:
    wb.remove(wb['Deduped'])

wb.create_sheet(index=1, title='Deduped')
deduped_sheet = wb['Deduped']
deduped_sheet.append(header_values)

for i in range(1, len(deduped_document)):
    deduped_sheet.append(list(deduped_document[i].values()))

wb.save(filename=f"{file_arg}.xlsx")

print('Sheetnames: ', wb.sheetnames)