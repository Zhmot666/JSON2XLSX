import json
import copy
import xlsxwriter
import tkinter as tk
from tkinter import filedialog


def write_bc(lev_dic, level, wr_barcode):
    lev_dic['lv'+str(level)] = wr_barcode
    return lev_dic


def read_lev(lev_up, bc_list, lev_dic):
    if isinstance(lev_up, list):
        for lev_key in lev_up:
            bc_list = read_lev(lev_key, bc_list, lev_dic)
        return bc_list
    elif isinstance(lev_up, dict):
        if 'ChildBarcodes' in lev_up:
            if 'level' in lev_up:
                lev_dic = write_bc(lev_dic, lev_up['level'], lev_up['Barcode'])
            bc_list = read_lev(lev_up['ChildBarcodes'], bc_list, lev_dic)
        elif 'Barcodes' in lev_up:
            bc_list = read_lev(lev_up['Barcodes'], bc_list, lev_dic)
        elif 'Barcode' in lev_up:
            if 'level' in lev_up:
                lev_dic = write_bc(lev_dic, lev_up['level'], lev_up['Barcode'])
            copy_lev_dic = copy.copy(lev_dic)
            bc_list.append(copy_lev_dic)
        return bc_list


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title='Выберите JSON-файл для конвертирования', defaultextension='*.json',
                                           filetypes=[('JSON файлы', '*.json'), ('All files', '*.*')])

    barcode_list = list()
    level_dic = dict()
    with open(file_path, encoding='utf-8') as read_file:
        for line in read_file:
            if 'level' in line:
                num_level = (line.replace('"level": ', '')).strip()
                if 'lv' + num_level not in level_dic:
                    level_dic['lv' + num_level] = ''
    with open(file_path, encoding='utf-8') as read_file:
        templates = json.load(read_file)
        lev1 = templates["TaskMarks"]
        read_lev(lev1, barcode_list, level_dic)

    workbook = xlsxwriter.Workbook(file_path.replace('json', 'xlsx'))
    counter = 1
    worksheet = workbook.add_worksheet()
    for barcode in barcode_list:
        for i in range(int(num_level) + 1):
            worksheet.write(chr(65+i) + str(counter), barcode['lv'+str(i)])
        counter += 1
    workbook.close()
