import sys
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QDialog, QDialogButtonBox, QVBoxLayout, QLabel, QLineEdit, QPushButton

import design


class MainApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Convertor JSON to XLSX')
        select_button = self.groupBox.findChild(QPushButton, 'SelectFile')
        select_button.clicked.connect(self.press_select)
        convert_button = self.findChild(QPushButton, 'ConvertFile')
        convert_button.clicked.connect(self.convert_file)

    def press_select(self):
        pass
        # level_dic = dict()
        # with open(file_path, encoding='utf-8') as read_file:
        #     for line in read_file:
        #         if 'level' in line:
        #             num_level = (line.replace('"level": ', '')).strip()
        #             if 'lv' + num_level not in level_dic:
        #                 level_dic['lv' + num_level] = ''

    def convert_file(self):
        pass


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
