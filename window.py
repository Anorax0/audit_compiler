# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'window.ui'
# Created by: PyQt5 UI code generator 5.13.2
# WARNING! All changes made in this file will be lost!

from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog, QComboBox
from PyQt5 import QtCore, QtGui, QtWidgets
from file_manager import FileManager
from record_checker import RecordChecker
import sys
import os

DIR = os.getcwd()


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

        # self.combo = QComboBox(self)
        self.combo = QtWidgets.QComboBox(self.centralwidget)
        self.combo.setEnabled(False)
        self.combo.setGeometry(QtCore.QRect(10, 70, 113, 23))
        # self.login_box.setObjectName("login_box")

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
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Open XLSX File', DIR,
                                                            'Excel Files (*.xls *.xml *.xlsx *.xlsm)')

        def test_file(file_path):
            try:
                file = FileManager(filename)
                return True
            except FileNotFoundError:
                return False

        if test_file(filename):
            workbook = FileManager(filename)
            wb = workbook.load_file()

            available_person = []

            for row in range(2, workbook.get_max_row() + 1):
                record = RecordChecker(wb, row)
                if record.col_H not in available_person:
                    available_person.append(record.col_H)

            self.combo.addItems(available_person)


            self.compile_btn.setEnabled(True)
            self.percent_box.setEnabled(True)
            self.combo.setEnabled(True)
            self.file_name = filename

    def compile_file(self):
        print(self.file_name)

    def retranslateUi(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.welcome_label.setText(_translate("MainWindow", "Audit Excel Compiler"))
        self.compile_btn.setText(_translate("MainWindow", "Compile"))
        self.loadfile_btn.setText(_translate("MainWindow", "Load File"))
        # self.login_box.setText(_translate("MainWindow", "Person login"))
        self.combo.setItemText(0, _translate("MainWindow", "ComboBox"))
        self.percent_box.setText(_translate("MainWindow", "0.2"))


class ExampleApp(QtWidgets.QMainWindow, UiMainWindow):
    def __init__(self, parent=None):
        super(ExampleApp, self).__init__(parent)
        self.setupUi(self)


def main():
    app = QApplication(sys.argv)
    form = ExampleApp()
    form.show()
    app.exec_()


if __name__ == '__main__':
    main()
