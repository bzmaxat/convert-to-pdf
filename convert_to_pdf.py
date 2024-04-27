import logging
import os
import tkinter as tk
from tkinter import filedialog

from docx2pdf import convert
from openpyxl import load_workbook
from win32com import client


def docx2pdf(file_path, save_path):
    try:
        # Конвертация .docx в .pdf
        convert(file_path, save_path)
        logging.info(f'Output file: {save_path}')
    except Exception as error:
        print('Something went wrong: ', error)
        logging.error(error)


def xlsx2pdf(file_path, save_path):
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0
    output_path = f'{save_path}/output.xlsx'
    output_pdf = f'{save_path}/output.pdf'

    try:
        # Обработка .xlsx файла с помощью библиотеки openpyxl
        workbook = load_workbook(file_path)
        sheet = workbook.active
        treeData = [["Type", "Leaf Color", "Height"], ["Maple", "Red", 549], ["Oak", "Green", 783], ["Pine", "Green", 1204]]
        for row in treeData:
            sheet.append(row)
        workbook.save(output_path)

        # Конвертация .xlsx в .pdf
        sheets = excel.Workbooks.Open(output_path)
        work_sheets = sheets.Worksheets[0]
        work_sheets.ExportAsFixedFormat(0, output_pdf)
        logging.info(f'Output file: {output_pdf}')
    except Exception as error:
        print('Something went wrong: ', error)
        logging.error(error)


def input_file():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()
    filename, file_extension = os.path.splitext(file_path)
    save_path = os.path.dirname(file_path)
    logging.info(f'Input file: {file_path}')
    if file_extension == '.docx':
        docx2pdf(file_path, save_path)
    elif file_extension == '.xlsx':
        xlsx2pdf(file_path, save_path)


def main():
    logging.basicConfig(level=logging.INFO, filename="logs.txt", filemode="w", format="%(asctime)s %(levelname)s %(message)s")

    input_file()


if __name__ == "__main__":
    main()
