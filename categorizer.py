import openpyxl
from openpyxl.styles import Alignment
import string
import re

workbook = openpyxl.load_workbook('C:/Users/cipri/OneDrive/Desktop/Financial/04_23.xlsx')
source_sheet = workbook['Sheet1']

if 'Sheet2' not in workbook.sheetnames:
    workbook.create_sheet('Sheet2')
sheet2 = workbook['Sheet2']

keyword_categories = {
    'Shopping': ['SRL', 'SCO', 'SHOPPING'],
    'Salary': ["Alim conventie"]
}

letters_uppercase = string.ascii_uppercase

# Iterate over the keys of the keyword_categories dictionary
for i, key in enumerate(keyword_categories.keys()):
    first_let = letters_uppercase[i * 2]
    last_let = letters_uppercase[i * 2 + 1] if i * 2 + 1 < len(letters_uppercase) else first_let
    sheet2.merge_cells(f"{first_let}1:{last_let}1")
    sheet2[f'{first_let}1'] = f'{key}'
    sheet2[f'{first_let}2'] = 'nume'
    sheet2[f'{last_let}2'] = 'suma'
    varA = sheet2[f'{first_let}1']
    varA.alignment = Alignment(horizontal='center', vertical='center')
    varB1 = sheet2[f'{first_let}2']
    varB2 = sheet2[f'{last_let}2']
    varB1.alignment = Alignment(horizontal='center', vertical='center')
    varB2.alignment = Alignment(horizontal='center', vertical='center')

# Shopping
pattern = r'Card/Terminale OP \d+(?:/\d+)? Card nr\d+|Utilizare POS comerciant .*'
sheet2_row = 3
first_vals = list(keyword_categories.values())[0]
for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
    for cell in row:
        if isinstance(cell, float):
            continue
        else:
            cell = re.sub(pattern, '', str(cell).strip())
        for i in first_vals:
            if i in str(cell):

                sheet2.cell(row=sheet2_row, column=1).value = cell.strip()
                sheet2.cell(row=sheet2_row, column=2).value = float(row[0])
                sheet2_row += 1


# Salary
sheet2_row = 3
second_vals = list(keyword_categories.values())[1]
for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
    for cell in row:
        if isinstance(cell, float):
            continue
        else:
            cell = cell.split('Automat OPH00004837 Dl CAPATA CIPRIAN GHEORGHE ')[-1].strip()
        for i in second_vals:
            if i in str(cell):
                sheet2.cell(row=sheet2_row, column=3).value = cell.strip()
                sheet2.cell(row=sheet2_row, column=4).value = float(row[0])
                sheet2_row += 1

# Sheet 3
if 'Sheet3' not in workbook.sheetnames:
    workbook.create_sheet('Sheet3')
sheet3 = workbook['Sheet3']

sheet3['A1'] = 'Centralizator'
letters = []
for i in range(len(keyword_categories.keys())):
    letter = letters_uppercase[i]
    letters.append(letter)
sheet3.merge_cells(f'A1:{letters[-1]}1')
sheet3['A1'].alignment = Alignment(horizontal='center', vertical='center')
for i, key in enumerate(keyword_categories.keys()):
    first_let = letters_uppercase[i]
    second_let = letters_uppercase[i * 2 + 1]
    sheet3[f'{first_let}2'] = f'{key}'
    sheet3[f'{first_let}3'] = f'=SUM(Sheet2!{second_let}3:{second_let}100)'

workbook.save('example.xlsx')
