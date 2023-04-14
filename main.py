import PyPDF2
import re
import pandas as pd
pdf_file = open('SV08870913300_2023_03.pdf', 'rb')

pdf_reader = PyPDF2.PdfReader(pdf_file)
num_pages = len(pdf_reader.pages)

table_data = ''
for i in range(num_pages):
    page = pdf_reader.pages[i]
    table_data += page.extract_text()

pdf_file.close()

unwanted_pg_nr = r'Pag. \d+ /'
table_data = re.sub(unwanted_pg_nr, '', table_data)

pattern = r"Start balance\s+(.*?)Nr. Zile"
table = re.findall(pattern, table_data, re.MULTILINE | re.DOTALL)

# Remove unwanted expression from the table
unwanted_expr_forlist = 'BRD-Groupe Societe Generale S.A.Turn BRD - Bdul. Ion Mihalache nr. 1-'
unwanted_expr = ['7, 011171 Bucuresti, Romania Tel:+4021.301.61.00 Fax:+4021.301.66.36',
                 'http://www.brd.roCAPITAL SOCIAL IN RON: 696.901.518 lei; R.C. J40/608/19.02.1991; RB -',
                 'PJR - 40 - 007 /18.02.1999; C.U.I./C.I.F. RO361579 ; EUID:',
                 'ROONRC.J40/608/1991 Atestat CNVM nr. 255/06.08.2008, inregistrata in',
                 'Registrul Public al CNVM cu nr. PJR01INCR/400008 7']
unwanted_columns = ['Value date Transaction descriptionCredit Debit Data val. Descriere operatiune Data oper.',
                    'Trans.Date']
table = table[0].replace(unwanted_expr_forlist, '')
for expr in unwanted_expr:
    table = table.replace(expr, '')
for expr in unwanted_columns:
    table = table.replace(expr, '')

table = table.replace('\n\n', '\n')
table = table.replace('\n\n', '\n')
table = table.replace('\n\n', '\n')
table = table.replace('\n\n', '\n')

print("Table extracted using regular expression:\n")
print(table)

