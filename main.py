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
            if variant == 1:
                bc_list.append(lev_up['Barcode'])
            else:
                copy_lev_dic = copy.copy(lev_dic)
                bc_list.append(copy_lev_dic)
        return bc_list


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    variant = 2  # Вариант 1 - список, вариант 2 - словарь

    barcode_list = list()
    with open(file_path, encoding='utf-8') as read_file:
        templates = json.load(read_file)
        lev1 = templates["TaskMarks"]
        level_dic = dict()
        level_dic['lv2'] = ''
        level_dic['lv1'] = ''
        level_dic['lv0'] = ''
        read_lev(lev1, barcode_list, level_dic)

    counter = 1
    workbook = xlsxwriter.Workbook(file_path.replace('json', 'xlsx'))
    worksheet = workbook.add_worksheet()
    for barcode in barcode_list:
        if variant == 1:
            worksheet.write('A' + str(counter), barcode)
        else:
            worksheet.write('A' + str(counter), barcode['lv0'])
            worksheet.write('B' + str(counter), barcode['lv1'])
            worksheet.write('C' + str(counter), barcode['lv2'])
        counter += 1
    workbook.close()
