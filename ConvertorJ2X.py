import sys
import json
import copy
from pathlib import Path

import xlsxwriter
from PyQt6 import QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMessageBox

import design
import taskmarks_aggregation

_APP_STYLESHEET = """
QMainWindow {
    background: #eef1f6;
}
QWidget#centralwidget {
    background: #eef1f6;
}
QGroupBox {
    font-weight: 600;
    color: #1e293b;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    margin-top: 14px;
    padding: 18px;
    padding-top: 12px;
    background: #ffffff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 8px;
    background-color: transparent;
}
QLabel#SelectedFile {
    color: #0f172a;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    min-height: 20px;
}
QLabel {
    color: #334155;
}
QFrame#line {
    color: #e2e8f0;
    max-height: 2px;
}
QLineEdit {
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px 10px;
    background: #ffffff;
    selection-background-color: #2563eb;
}
QLineEdit:focus {
    border-color: #2563eb;
}
QPushButton#SelectFile {
    background: #ffffff;
    color: #1d4ed8;
    border: 1px solid #93c5fd;
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: 600;
}
QPushButton#SelectFile:hover {
    background: #eff6ff;
    border-color: #3b82f6;
}
QPushButton#SelectFile:pressed {
    background: #dbeafe;
}
QPushButton#ConvertFile {
    background: #2563eb;
    color: #ffffff;
    border: none;
    border-radius: 8px;
    padding: 10px 20px;
    font-weight: 600;
    font-size: 14px;
}
QPushButton#ConvertFile:hover {
    background: #1d4ed8;
}
QPushButton#ConvertFile:pressed {
    background: #1e40af;
}
QPushButton#ExportAggregation {
    background: #ffffff;
    color: #1d4ed8;
    border: 1px solid #93c5fd;
    border-radius: 8px;
    padding: 10px 16px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton#ExportAggregation:hover {
    background: #eff6ff;
    border-color: #3b82f6;
}
QPushButton#ExportAggregation:pressed {
    background: #dbeafe;
}
QStatusBar {
    background: #f1f5f9;
    border-top: 1px solid #e2e8f0;
    color: #64748b;
}
"""


class MainApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.level_dic = dict()
        self.barcode_list = list()
        self.num_level = ''
        self.setupUi(self)
        self.setStyleSheet(_APP_STYLESHEET)
        self.setWindowTitle('JSON → XLSX')
        self.SelectFile.clicked.connect(self.press_select)
        self.ConvertFile.clicked.connect(self.convert_file)
        self.ExportAggregation.clicked.connect(self.export_aggregation_report)

    def press_select(self):
        dlg = QFileDialog.getOpenFileName(self, 'Выберите файл', '', 'JSON-файлы (*.json)')
        self.SelectedFile.setText(dlg[0] if dlg[0] else 'Файл не выбран')
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
                self.items_lv_0.setText(str(count_lv0))
                if count_lv1:
                    per_box = count_lv0 / count_lv1
                    self.items_lv_1.setText(f'{count_lv1} (~{per_box:.1f} изд. на коробку)')
                else:
                    self.items_lv_1.setText(str(count_lv1))
                self.items_lv_2.setText(str(count_lv2))

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
        file_path = self.SelectedFile.text()
        if not file_path or file_path == 'Файл не выбран':
            QMessageBox.warning(self, 'Нет файла', 'Сначала выберите JSON-файл.')
            return
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
        QMessageBox.information(self, 'Готово', 'Конвертация завершена.')

    def export_aggregation_report(self):
        file_path = self.SelectedFile.text()
        if not file_path or file_path == 'Файл не выбран':
            QMessageBox.warning(self, 'Нет файла', 'Сначала выберите JSON-файл.')
            return
        p = Path(file_path)
        pid = self.participantIdInput.text().strip()
        if not pid:
            QMessageBox.warning(
                self,
                'Нет participantId',
                'Заполните поле «Участник (participantId)» для отчёта агрегации.',
            )
            return
        try:
            pg = self.productGroupInput.text().strip() or None
            out_json, out_csv = taskmarks_aggregation.process_file(
                p,
                taskmarks_aggregation.SCHEMA_PATH,
                validate=True,
                product_group=pg,
                participant_id=pid,
            )
        except Exception as exc:
            QMessageBox.critical(self, 'Ошибка экспорта', str(exc))
            return
        QMessageBox.information(
            self,
            'Готово',
            f'Созданы файлы:\n{out_json}\n{out_csv}',
        )

    @staticmethod
    def _default_font():
        f = QtGui.QFont()
        if sys.platform == 'win32':
            f.setFamily('Segoe UI')
        f.setPointSize(10)
        return f


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setFont(MainApp._default_font())
    window = MainApp()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
