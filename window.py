""" -*- coding: utf-8 -*-
 Form implementation generated from reading ui file 'window.ui'
 Created by: PyQt5 UI code generator 5.13.2
 WARNING! All changes made in this file will be lost!"""

import sys
import os
from time import time
from itertools import chain
from random import choice
import colorama
from PyQt5.QtWidgets import QWidget, QApplication
from PyQt5 import QtCore, QtWidgets
from file_manager import FileManager, warning_box
from record_checker import RecordChecker

DIR = os.getcwd()

colorama.init()


class UiMainWindow(QWidget):
    def setupUi(self, main_window):
        main_window.setObjectName("MainWindow")
        main_window.resize(307, 165)

        self.centralwidget = QtWidgets.QWidget(main_window)
        self.centralwidget.setObjectName("centralwidget")

        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 100, 290, 25))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")

        self.welcome_label = QtWidgets.QLabel(self.centralwidget)
        self.welcome_label.setEnabled(True)
        self.welcome_label.setGeometry(QtCore.QRect(100, 0, 120, 20))
        self.welcome_label.setObjectName("welcome_label")

        self.compile_btn = QtWidgets.QPushButton(self.centralwidget)
        self.compile_btn.setEnabled(False)
        self.compile_btn.setGeometry(QtCore.QRect(200, 70, 89, 23))
        self.compile_btn.setObjectName("compile_btn")

        self.loadfile_btn = QtWidgets.QPushButton(self.centralwidget)
        self.loadfile_btn.setGeometry(QtCore.QRect(10, 30, 89, 23))
        self.loadfile_btn.setObjectName("loadfile_btn")

        self.combo = QtWidgets.QComboBox(self.centralwidget)
        self.combo.setEnabled(False)
        self.combo.setGeometry(QtCore.QRect(10, 70, 113, 23))

        self.percent_box = QtWidgets.QLineEdit(self.centralwidget)
        self.percent_box.setEnabled(False)
        self.percent_box.setGeometry(QtCore.QRect(130, 70, 40, 23))
        self.percent_box.setObjectName("percent_box")

        main_window.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(main_window)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 299, 20))
        self.menubar.setObjectName("menubar")
        main_window.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(main_window)
        self.statusbar.setObjectName("statusbar")
        main_window.setStatusBar(self.statusbar)

        self.file_name = None
        self.compile_btn.clicked.connect(self.compile_file)
        self.loadfile_btn.clicked.connect(self.open_file_browser)

        self.retranslateUi(main_window)
        QtCore.QMetaObject.connectSlotsByName(main_window)

    def open_file_browser(self):
        """
        Open file browser
        :return: object
        """
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Open XLSX File',
                                                            DIR, 'Excel Files (*.xls *.xml *.xlsx *.xlsm)')

        def test_file(file_path):
            """
            Tests if file can be opened
            :param file_path:
            :return:
            """
            try:
                file = FileManager(file_path)
                del file
                return True
            except FileNotFoundError:
                warning_box('File not found!', 'File not found.')
                sys.exit(1)
            except PermissionError:
                warning_box('Permision Error!',
                            'File is running in different process. '
                            'Please close other processes to work with this file.')
                sys.exit(1)
            except IOError:
                warning_box('Opening file failed!', 'Cannot open this file.')
                sys.exit(1)

        if test_file(filename):
            workbook = FileManager(filename)
            wb = workbook.load_file()

            available_person = []

            for row in range(2, workbook.get_max_row() + 1):
                record = RecordChecker(wb, row)
                if record.col_h not in available_person:
                    available_person.append(record.col_h)

            self.combo.addItems(sorted(set(available_person)))

            self.compile_btn.setEnabled(True)
            self.percent_box.setEnabled(True)
            self.combo.setEnabled(True)
            self.file_name = filename

    def compile_file(self):
        """
        Compiles records in file due to criteria
        :return: None
        """

        start_time = time()

        records_length = 0

        workbook_file = FileManager(self.file_name)

        if workbook_file.check_col_jk():
            print(colorama.Fore.RED + 'Some of records are not REVIEWED/PENDING')
            print(colorama.Style.RESET_ALL, end='')

        workbook = workbook_file.load_file()

        for row in range(2, workbook_file.get_max_row() + 1):
            record = RecordChecker(workbook, row)
            if record.col_h == self.combo.currentText():
                records_length += 1
        self.progressBar.setProperty("value", 20)

        print('Length of all records of a person :', records_length)

        # do math.ceil(number) to round up the float to whole integer
        records_length_percent = round(records_length * float(self.percent_box.text()))
        category_ed_percent = round(records_length_percent * 0.4)
        category_fg_percent = round(records_length_percent * 0.3)
        # category_J_percent = round(records_length_percent * 0.2)
        category_m_percent = round(records_length_percent * 0.1)

        print('All records person should check: ', records_length_percent)
        print('ED category records to check: ', category_ed_percent)
        print('FG category records to check: ', category_fg_percent)
        # print(category_J_percent)
        print('M category records to check: ', category_m_percent)

        category_ed = []
        category_fg = []
        # category_J = []
        category_m = []

        to_check_category_ed = []
        to_check_category_fg = []
        # to_check_category_J = []
        to_check_category_m = []
        # all_categories = category_ed, category_fg, category_m

        for row in range(2, workbook_file.get_max_row() + 1):
            record = RecordChecker(workbook, row)
            if record.col_h == self.combo.currentText():
                if record.check_category_ed():
                    category_ed.append(row)
                if record.check_category_fg():
                    category_fg.append(row)
                # if record.check_category_J() and row not in chain(*all_categories):
                #     category_J.append(row)
                if record.check_category_m():
                    category_m.append(row)
        self.progressBar.setProperty("value", 50)

        print('List of records of a given category:')
        print('ED list :', category_ed)
        print('FG list :', category_fg)
        # category_J
        print('M list :', category_m)

        if category_ed_percent > len(category_ed):
            category_fg_percent += category_ed_percent - len(category_ed)
            category_ed_percent = len(category_ed)
        if category_fg_percent > len(category_fg):
            category_m_percent += category_fg_percent - len(category_fg)
            category_fg_percent = len(category_fg)

        while category_ed_percent != 0:
            if len(category_ed) == 0:
                break
            record_to_add = choice(category_ed)
            if record_to_add not in to_check_category_ed:
                to_check_category_ed.append(record_to_add)
                category_ed_percent -= 1
        self.progressBar.setProperty("value", 55)

        while category_fg_percent != 0:
            if len(category_fg) == 0:
                break
            record_to_add = choice(category_fg)
            if record_to_add not in to_check_category_fg:
                to_check_category_fg.append(record_to_add)
                category_fg_percent -= 1
        self.progressBar.setProperty("value", 62)

        while category_m_percent != 0:
            if len(category_m) == 0:
                break
            record_to_add = choice(category_m)
            if record_to_add not in to_check_category_m:
                to_check_category_m.append(record_to_add)
                category_m_percent -= 1
        self.progressBar.setProperty("value", 68)

        # bar_record_randomize.finish()
        # if bar_record_randomize.max > bar_record_randomize.index:
        #     print(colorama.Fore.RED + 'Cannot find enough records to fill criteria.')
        #     print(colorama.Style.RESET_ALL, end='')

        print(colorama.Fore.YELLOW + 'Rows with chosen records to check: ')
        print('ED rows: ', to_check_category_ed)
        print('FG rows: ', to_check_category_fg)
        print('M rows :', to_check_category_m, colorama.Style.RESET_ALL)

        self.progressBar.setProperty("value", 82)

        # chained_categories = chain(to_check_category_ed, to_check_category_fg, to_check_category_m)
        records_to_check_list = chain(to_check_category_ed, to_check_category_fg, to_check_category_m)
        workbook_file.save_workbook(records_to_check_list)

        self.progressBar.setProperty("value", 100)

        print(f'\n Executing time: {(time() - start_time):.3f}s')

    def retranslateUi(self, main_window):
        """
        PyQt5 method
        :param main_window:
        :return: None
        """
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.welcome_label.setText(_translate("MainWindow", "Audit Excel Compiler"))
        self.compile_btn.setText(_translate("MainWindow", "Compile"))
        self.loadfile_btn.setText(_translate("MainWindow", "Load File"))
        # self.login_box.setText(_translate("MainWindow", "Person login"))
        self.combo.setItemText(0, _translate("MainWindow", "ComboBox"))
        self.percent_box.setText(_translate("MainWindow", "0.2"))


class ExampleApp(QtWidgets.QMainWindow, UiMainWindow):
    """
    Executing code
    """
    def __init__(self, parent=None):
        super(ExampleApp, self).__init__(parent)
        self.setupUi(self)


def main():
    """
    Runs all together
    :return: QtPy5 object
    """
    app = QApplication(sys.argv)
    form = ExampleApp()
    form.show()
    app.exec_()


if __name__ == '__main__':
    main()
