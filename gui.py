# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MyWindow(object):
    def setupUi(self, MyWindow):
        MyWindow.setObjectName("MyWindow")
        MyWindow.resize(222, 187)
        self.centralwidget = QtWidgets.QWidget(MyWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.printer_combo = QtWidgets.QComboBox(self.centralwidget)
        self.printer_combo.setGeometry(QtCore.QRect(10, 30, 201, 22))
        self.printer_combo.setObjectName("printer_combo")
        self.printer_label = QtWidgets.QLabel(self.centralwidget)
        self.printer_label.setGeometry(QtCore.QRect(50, 10, 141, 16))
        self.printer_label.setObjectName("printer_label")
        self.new_label = QtWidgets.QLabel(self.centralwidget)
        self.new_label.setGeometry(QtCore.QRect(10, 60, 91, 16))
        self.new_label.setObjectName("new_label")
        self.book_num_text_label = QtWidgets.QLineEdit(self.centralwidget)
        self.book_num_text_label.setGeometry(QtCore.QRect(10, 80, 81, 20))
        self.book_num_text_label.setObjectName("book_num_text_label")
        self.room_text_label = QtWidgets.QLineEdit(self.centralwidget)
        self.room_text_label.setGeometry(QtCore.QRect(100, 80, 111, 20))
        self.room_text_label.setObjectName("room_text_label")
        self.room_label = QtWidgets.QLabel(self.centralwidget)
        self.room_label.setGeometry(QtCore.QRect(120, 60, 71, 16))
        self.room_label.setObjectName("room_label")
        self.select_file_button = QtWidgets.QPushButton(self.centralwidget)
        self.select_file_button.setGeometry(QtCore.QRect(10, 110, 201, 23))
        self.select_file_button.setObjectName("select_file_button")
        self.print_button = QtWidgets.QPushButton(self.centralwidget)
        self.print_button.setGeometry(QtCore.QRect(10, 140, 201, 23))
        self.print_button.setObjectName("print_button")
        MyWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MyWindow)
        self.statusbar.setObjectName("statusbar")
        MyWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MyWindow)
        QtCore.QMetaObject.connectSlotsByName(MyWindow)

    def retranslateUi(self, MyWindow):
        _translate = QtCore.QCoreApplication.translate
        MyWindow.setWindowTitle(_translate("MyWindow", "MainWindow"))
        self.printer_label.setText(_translate("MyWindow", "프린터를 선택해주세요"))
        self.new_label.setText(_translate("MyWindow", "신간 기준 번호"))
        self.book_num_text_label.setText(_translate("MyWindow", ""))
        self.room_text_label.setText(_translate("MyWindow", ""))
        self.room_label.setText(_translate("MyWindow", "자료실 이름"))
        self.select_file_button.setText(_translate("MyWindow", "제공자료(xslx) 선택"))
        self.print_button.setText(_translate("MyWindow", "프린트"))
