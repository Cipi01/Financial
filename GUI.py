import sys
import string
from categories import taxes, extractions, fuel, deliveries, supermarkets, salary, datorii_in, datorii_out, diverse, \
    fastfood, clothes
from classes import ExcelFile, align, color_red, color_green
from dict_categories import get_dict
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton

WINDOW_SIZE = 235


class ExcelFileSelector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Excel File Selector')
        self.setGeometry(100, 100, 500, 500)

        self.init_ui()

    def init_ui(self):
        self.button = QPushButton('Select Excel File', self)
        self.button.setGeometry(200, 200, 100, 50)
        self.button.clicked.connect(self.select_file)

    def select_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter('Excel Files (*.xlsx)')
        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]
            self.process_file(selected_file)
            self.close_window()

    def close_window(self):
        self.close()

    def process_file(self, file_path):
        with ExcelFile(file_path) as wb:
            sheet1 = wb['Sheet1']

            if 'Sheet2' not in wb.sheetnames:
                wb.create_sheet('Sheet2')
            sheet2 = wb['Sheet2']

            keyword_categories = get_dict()
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
                align(varA)
                varB1 = sheet2[f'{first_let}2']
                varB2 = sheet2[f'{last_let}2']
                align(varB1)
                align(varB2)

                if key == "Salary" or key == "Plata datorii":
                    color_green(varA)
                else:
                    color_red(varA)

            supermarkets(keyword_categories, sheet1, sheet2)
            salary(keyword_categories, sheet1, sheet2)
            deliveries(keyword_categories, sheet1, sheet2)
            datorii_in(keyword_categories, sheet1, sheet2)
            datorii_out(keyword_categories, sheet1, sheet2)
            fuel(keyword_categories, sheet1, sheet2)
            extractions(keyword_categories, sheet1, sheet2)
            taxes(keyword_categories, sheet1, sheet2)
            fastfood(keyword_categories, sheet1, sheet2)
            diverse(keyword_categories, sheet1, sheet2)
            clothes(keyword_categories, sheet1, sheet2)

            # Sheet 3
            if 'Sheet3' not in wb.sheetnames:
                wb.create_sheet('Sheet3')
            sheet3 = wb['Sheet3']

            sheet3['A1'] = 'Centralizator'
            letters = []
            for i in range(len(keyword_categories.keys())):
                letter = letters_uppercase[i]
                letters.append(letter)
            sheet3.merge_cells(f'A1:{letters[-1]}1')
            align(sheet3['A1'])
            for i, key in enumerate(keyword_categories.keys()):
                first_let = letters_uppercase[i]
                second_let = letters_uppercase[i * 2 + 1]
                sheet3[f'{first_let}2'] = f'{key}'
                sheet3[f'{first_let}3'] = f'=SUM(Sheet2!{second_let}3:{second_let}100)'
                varA = sheet3[f'{first_let}2']

                if key == "Salary" or key == "Plata datorii":
                    color_green(varA)
                else:
                    color_red(varA)

            sheet3['A6'] = 'TOTAL'
            sheet3.merge_cells('A6:B6')
            align(sheet3['A6'])

            sheet3['A7'] = 'CREDIT'
            sheet3['B7'] = 'DEBIT'
            align(sheet3['A7'])
            align(sheet3['B7'])

            sheet3['A8'] = '=SUMIF(Sheet1!D2:D1000,"Credit",Sheet1!A2:A1000)'
            sheet3['B8'] = '=SUMIF(Sheet1!D2:D1000,"Debit",Sheet1!A2:A1000)'
            align(sheet3['A8'])
            align(sheet3['B8'])
            color_green(sheet3['A8'])
            color_red(sheet3['B8'])

            wb.save('test.xlsx')


if __name__ == "__main__":
    app = QApplication([])
    window = ExcelFileSelector()
    window.show()
    sys.exit(app.exec())
