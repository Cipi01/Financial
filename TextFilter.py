import re
import PyPDF2


def pdf_extractor(file_name):
    pdf_file = open(file_name, 'rb')

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
    #table = re.sub(r'OP \d{6,}/\d+([0-9])', r'OP \g<0> \1', table)

    return table


def table_modifier(table):
    # Split the table into lines
    lines = table.split('\n')

    # Define the pattern for matching the date format
    date_pattern = r"\d{2}/\d{2}/\d{4}"

    # Iterate over the lines and check if there is a date in the line
    # If there is a date, remove it from the line and print it followed by the date
    modified_lines = []
    for i, line in enumerate(lines):
        if re.search(date_pattern, line):
            date = re.search(date_pattern, line)[0]
            modified_line = re.sub(date_pattern, '', line)
            modified_lines.append(modified_line)
            modified_lines.append(date)
        else:
            modified_lines.append(line)

    # Join the modified lines into a single string
    modified_table = '\n'.join(modified_lines)

    modified_table = modified_table.replace('.', '')
    return modified_table


def dict_maker(modified_table):
    lines = modified_table.split("\n")
    transactions = []
    # Initialize a dictionary with empty values
    transaction = {"Sum": "", "Details": "", "Date": "", "Status": ""}

    # Loop through the lines and fill in the dictionary
    for line in lines:
        if line.count("/") == 2 and line.replace("/", "").isdigit():
            # This line contains the date
            transaction["Date"] = line.strip()
            if transaction["Sum"] != "":
                # If a previous transaction has been started, add it to the list of transactions
                if "credit" in transaction["Details"].lower():
                    transaction["Status"] = "Credit"
                else:
                    transaction["Status"] = "Debit"
                transactions.append(transaction)
                # Start a new transaction
                transaction = {"Sum": "", "Details": "", "Date": "", "Status": ""}
        elif line.replace(',', '').replace('.', '').isdigit():
            # This line contains the sum
            line = line.replace(',', '.')
            transaction["Sum"] = float(line.strip())
        else:
            # This line contains the transaction details
            transaction["Details"] += line.strip() + " "

    # Strip whitespace from the Details column
    for transaction in transactions:
        transaction["Details"] = transaction["Details"].strip()

    # Add the last transaction to the list of transactions
    if "credit" in transaction["Details"].lower():
        transaction["Status"] = "Credit"
    else:
        transaction["Status"] = "Debit"
    transactions.append(transaction)
    transactions.pop()
    return transactions


if __name__ == '__main__':
    a = pdf_extractor('SV08870913300_2023_03.pdf')
    print(a)