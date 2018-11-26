import csv
import openpyxl
import os.path
import re


# Folder with xlsx files
XSLX_DIR = 'xslx'

# Folders with csv files
CSV_DIR = 'csv'


# Generator of all xslx files in the folder
def get_xslx_files():
    xslx_files = (f for f in os.listdir(XSLX_DIR) if os.path.isfile(os.path.join(XSLX_DIR, f)))
    return xslx_files


# Generator of all rows in xslx file
def all_xlsx_rows(ws):
    for row in ws.iter_rows():
        yield [cell.value for cell in row]


for xlsx_file in get_xslx_files():
    # Path to xlsx file
    xlsx_files_path = os.path.join(XSLX_DIR, xlsx_file)

    # Open xlsx file
    wb = openpyxl.load_workbook(xlsx_files_path)

    # Select active sheet
    ws = wb.active

    # csv file in which will be converted xlsx file
    csv_file = re.sub(r'xlsx?', 'csv', xlsx_file)

    # Path to csv file
    csv_save_path = os.path.join(CSV_DIR, csv_file)

    print('Converting', xlsx_file, 'to', csv_file, '...')

    # Create csv file in which will be converted xlsx file
    with open(csv_save_path, 'w', encoding='utf-8', newline='') as f:
        csv_writer = csv.writer(f)

        # Line by line writing of row contents of xslx into csv
        for i in all_xlsx_rows(ws):
            csv_writer.writerow(i)
