import json
import xlsxwriter
import tkinter as tk
from tkinter import filedialog


def read_lev(lev_up, bc_list):
    if type(lev_up) == list:
        for lev_key in lev_up:
            read_lev(lev_key, bc_list)
        return bc_list
    elif type(lev_up) == dict:
        if 'ChildBarcodes' in lev_up:
            read_lev(lev_up['ChildBarcodes'], bc_list)
        elif 'Barcodes' in lev_up:
            read_lev(lev_up['Barcodes'], bc_list)
        elif 'Barcode' in lev_up:
            bc_list.append(lev_up['Barcode'])
        return bc_list


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    barcode_list = list()
    with open(file_path) as read_file:
        templates = json.load(read_file)
        lev1 = templates["TaskMarks"]
        read_lev(lev1, barcode_list)

    counter = 1
    workbook = xlsxwriter.Workbook('gematogen.xlsx')
    worksheet = workbook.add_worksheet()
    for barcode in barcode_list:
        worksheet.write('A'+str(counter), barcode)
        counter += 1
    workbook.close()
