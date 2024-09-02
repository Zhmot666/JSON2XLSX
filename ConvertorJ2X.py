import sys
import json
import copy
import xlsxwriter
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QFileDialog, QLabel, QPushButton, QMessageBox

import design


class MainApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.level_dic = dict()
        self.barcode_list = list()
        self.num_level = ''
        self.setupUi(self)
        self.setWindowTitle('Convertor JSON to XLSX')
        select_button = self.groupBox.findChild(QPushButton, 'SelectFile')
        select_button.clicked.connect(self.press_select)
        convert_button = self.findChild(QPushButton, 'ConvertFile')
        convert_button.clicked.connect(self.convert_file)

    def press_select(self):
        dlg = QFileDialog.getOpenFileName(self, 'Выберите файл', 'd:/', 'JSON-файлы (*.json)')
        file_name = self.groupBox.findChild(QLabel, 'SelectedFile')
        file_name.setText(dlg[0])
        count_lv0, count_lv1, count_lv2 = 0, 0, 0
        if dlg[0] != '':
            with open(dlg[0], encoding='utf-8') as read_file:
                for line in read_file:
                    if 'level' in line:
                        self.num_level = (line.replace('"level": ', '')).strip()
                        match int(self.num_level):
                            case 0:
                                count_lv0 += 1
                            case 1:
                                count_lv1 += 1
                            case 2:
                                count_lv2 += 1
                        if 'lv' + self.num_level not in self.level_dic:
                            self.level_dic['lv' + self.num_level] = ''
                self.groupBox.findChild(QLabel, 'items_lv_0').setText(str(count_lv0))
                self.groupBox.findChild(QLabel, 'items_lv_1').setText(str(count_lv1) + ' по ' + str(count_lv0/count_lv1) +
                                                                      ' шт в коробке')
                self.groupBox.findChild(QLabel, 'items_lv_2').setText(str(count_lv2))

    def write_bc(self, lev_dic, level, wr_barcode):
        lev_dic['lv' + str(level)] = wr_barcode
        return lev_dic

    def read_lev(self, lev_up, bc_list, lev_dic):
        if isinstance(lev_up, list):
            for lev_key in lev_up:
                bc_list = self.read_lev(lev_key, bc_list, lev_dic)
            return bc_list
        elif isinstance(lev_up, dict):
            if 'ChildBarcodes' in lev_up:
                if 'level' in lev_up:
                    lev_dic = self.write_bc(lev_dic, lev_up['level'], lev_up['Barcode'])
                bc_list = self.read_lev(lev_up['ChildBarcodes'], bc_list, lev_dic)
            elif 'Barcodes' in lev_up:
                bc_list = self.read_lev(lev_up['Barcodes'], bc_list, lev_dic)
            elif 'Barcode' in lev_up:
                if 'level' in lev_up:
                    lev_dic = self.write_bc(lev_dic, lev_up['level'], lev_up['Barcode'])
                copy_lev_dic = copy.copy(lev_dic)
                bc_list.append(copy_lev_dic)
            return bc_list

    def convert_file(self):
        workbook, worksheet = '', ''
        self.barcode_list = list()
        file_path = self.groupBox.findChild(QLabel, 'SelectedFile').text()
        separator_num = 0 if self.max_items_lv0.text().strip() == '' else int(self.max_items_lv0.text())
        with open(file_path, encoding='utf-8') as read_file:
            templates = json.load(read_file)
            lev1 = templates["TaskMarks"]
            self.read_lev(lev1, self.barcode_list, self.level_dic)

        if separator_num != 0:
            file_name = file_path.replace('.json', '')
            counter_files = 1
            counter_lines = 1
            for barcode in self.barcode_list:
                if counter_lines == 1:
                    workbook = xlsxwriter.Workbook(file_name + '_' + str(counter_files) + '.xlsx')
                    worksheet = workbook.add_worksheet()
                for i in range(int(self.num_level) + 1):
                    worksheet.write(chr(65 + i) + str(counter_lines), barcode['lv' + str(i)])
                counter_lines += 1
                if counter_lines > separator_num:
                    counter_files += 1
                    counter_lines = 1
                    workbook.close()
            workbook.close()

        workbook = xlsxwriter.Workbook(file_path.replace('.json', '_big.xlsx'))
        worksheet = workbook.add_worksheet()
        counter_lines = 1
        for barcode in self.barcode_list:
            for i in range(int(self.num_level) + 1):
                worksheet.write(chr(65 + i) + str(counter_lines), barcode['lv' + str(i)])
            counter_lines += 1
        workbook.close()
        dialog = QMessageBox(parent=self, text="Конвертация окончена")
        dialog.setWindowTitle("   ")
        dialog.exec()


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
