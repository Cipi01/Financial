import re

pattern = r'Card/Terminale OP \d+(?:/\d+)? Card nr\d+|Utilizare POS comerciant .*'


# Supermarkets
def supermarkets(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    first_vals = list(keyword_categories.values())[0]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            if isinstance(cell, float):
                continue
            else:
                cell = re.sub(pattern, '', str(cell).strip())
            for i in first_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=1).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=2).value = float(row[0])
                    written_sheet_row += 1


# Salary
def salary(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[1]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            if isinstance(cell, float) or isinstance(cell, int):
                continue
            else:
                cell = cell.split('Automat OPH00004837 Dl CAPATA CIPRIAN GHEORGHE ')[-1].strip()
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=3).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=4).value = float(row[0])
                    written_sheet_row += 1


# Deliveries
def deliveries(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[2]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            if isinstance(cell, float):
                continue
            else:
                cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=5).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=6).value = float(row[0])
                    written_sheet_row += 1


# Datorii IN
def datorii_in(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[3]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            # if isinstance(cell, float):
            #  continue
            # else:
            # cell = re.sub(r'^\w+\s+\w+\s+\w+\s+(Dl\s+\S+\s+)RO\d+\w+\s(.*)', r'\1\2', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=7).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=8).value = float(row[0])
                    written_sheet_row += 1


# Datorii OUT
def datorii_out(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[4]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=9).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=10).value = float(row[0])
                    written_sheet_row += 1


# Fuel
def fuel(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[5]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=11).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=12).value = float(row[0])
                    written_sheet_row += 1


# ATM Extractions
def extractions(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[6]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=13).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=14).value = float(row[0])
                    written_sheet_row += 1


# Taxes
def taxes(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[7]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=15).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=16).value = float(row[0])
                    written_sheet_row += 1


# Fastfood
def fastfood(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[8]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=17).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=18).value = float(row[0])
                    written_sheet_row += 1


# Diverse
def diverse(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[9]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=19).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=20).value = float(row[0])
                    written_sheet_row += 1

# Clothes
def clothes(keyword_categories, source_sheet, written_sheet):
    written_sheet_row = 3
    second_vals = list(keyword_categories.values())[10]
    for row in source_sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        for cell in row:
            #  if isinstance(cell, float):
            #   continue
            # else:
            # cell = re.sub(pattern, '', str(cell).strip())
            for i in second_vals:
                if i in str(cell):
                    written_sheet.cell(row=written_sheet_row, column=21).value = cell.strip()
                    written_sheet.cell(row=written_sheet_row, column=22).value = float(row[0])
                    written_sheet_row += 1
