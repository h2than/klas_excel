import sys
import os
import win32com.client
import win32print
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QPushButton
from PyQt5.QtCore import QSettings
from gui import Ui_MyWindow
import glob
import pandas as pd
import time
import re
import chromedriver_autoinstaller

class MyWindow(QMainWindow, Ui_MyWindow):
    xlsx_path = ""
    driver = None
    rooms = ""
    new = None
    printer_name = ""
    download_folder = os.path.expanduser("~/Downloads")

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.print_button.clicked.connect(self.print_excel_file)

        self.populate_printer_list()
        self.populate_book_number()
        self.room_select()
        try :
            chromedriver_autoinstaller.install()
        except:
            pass

    def get_xls(self):
        pattern = "제공승인목록*.xls"
        files = glob.glob(os.path.join(self.download_folder, pattern ))
        files.sort(key=os.path.getctime, reverse=True)
        self.xlsx_path = files[0]

    def populate_printer_list(self):
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        self.printer_combo.addItems(printers)
        
        settings = QSettings('MyApp', 'MyWindow')
        last_printer = settings.value('last_printer', type=str)
        if last_printer in printers:
            self.printer_combo.setCurrentText(last_printer)
        self.printer_combo.activated.connect(self.save_selected_printer)

    def save_selected_printer(self):
        selected_printer = self.printer_combo.currentText()
        settings = QSettings('MyApp', 'MyWindow')
        settings.setValue('last_printer', selected_printer)

    def populate_book_number(self):
        settings = QSettings('MyApp', 'MyWindow')
        last_book_number = settings.value('last_book_number', type=str)
        if last_book_number:
            self.book_num_text_label.setText(last_book_number)

        self.book_num_text_label.textChanged.connect(self.save_book_number)

    def save_book_number(self):
        book_number = self.book_num_text_label.text()
        settings = QSettings('MyApp', 'MyWindow')
        settings.setValue('last_book_number', book_number)

    def room_select(self):
        settings = QSettings('MyApp', 'MyWindow')
        last_room_text = settings.value('last_room_text', type=str)
        if last_room_text:
            self.room_text_label.setText(last_room_text)
        
        self.room_text_label.textChanged.connect(self.save_room_text)

    def save_room_text(self):
        room_text = self.room_text_label.text()
        settings = QSettings('MyApp', 'MyWindow')
        settings.setValue('last_room_text', room_text)

    def show_message(self, text):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(text)
        msg_box.setWindowTitle("Info")
        msg_box.exec_()

    def print_excel_file(self):
        self.print_button.setDisabled(True)
        self.get_xls()

        if self.xlsx_path == "":
            self.show_message("다운로드된 .xls 파일을 찾을 수 없습니다.")
            return

        self.rooms = (self.room_text_label.text().split())
        self.new = int(self.book_num_text_label.text())
        self.printer_name = self.printer_combo.currentText()

        if self.rooms == None or self.new == None :
            self.show_message("옵션을 다시 확인해 주십시오.")
            return
        if self.printer_name == "":
            self.show_message("프린터를 찾을 수 없습니다.")
            return

        data = pd.read_excel(self.xlsx_path, sheet_name='전체', skiprows=2)
        pattern = '|'.join(self.rooms)
        subset = data.iloc[:, 3:7]
        subset = subset[subset.iloc[:, 3].astype(str).str.contains(pattern, na=False)]
        subset = subset.iloc[:, :-1]
        subset = subset.sort_values(by=subset.columns[2])

        self.xlsx_path = self.download_folder + "\\print_.xlsx"
        if os.path.exists(self.xlsx_path):
            os.remove(self.xlsx_path)
        subset.to_excel(self.xlsx_path, index=False)
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(self.xlsx_path)
        worksheet = workbook.Worksheets[0]

        current_row = 2  # 현재 출력 중인 행
        last_row = 2
        total_rows = worksheet.UsedRange.Rows.Count+1

        for row in range(2, total_rows):
            rowA = worksheet.Cells(row, 1).Value
            cell_new_value = int(rowA[2:])
            
            for column_index in range(1, 3):
                cell = worksheet.Cells(row, column_index)
                if cell_new_value >= self.new :
                    cell.Font.Color = 255

        worksheet.Cells.Font.Size = 20
        worksheet.Rows.AutoFit()
        worksheet.Columns.AutoFit()
        worksheet.PageSetup.Orientation = 2
        worksheet.PageSetup.PaperSize = 9
        worksheet.PageSetup.FitToPagesWide = 1
        worksheet.PageSetup.FitToPagesTall = False
        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.LeftMargin = 0.2
        worksheet.PageSetup.RightMargin = 0.2
        worksheet.PageSetup.TopMargin = 0.2
        worksheet.PageSetup.BottomMargin = 0.2

        while True :
            if current_row == total_rows-1 :
                print_range = worksheet.Range(worksheet.Cells(last_row, 1), worksheet.Cells(current_row, 3))
                print_range.PrintOut(ActivePrinter=self.printer_name)
                break

            if current_row % 20 == 0:
                print_range = worksheet.Range(worksheet.Cells(last_row, 1), worksheet.Cells(current_row, 3))
                last_row = current_row
                print_range.PrintOut(ActivePrinter=self.printer_name)

            current_row += 1

        workbook.Close(SaveChanges=False)
        excel.Quit()
        self.show_message("작업이 완료되었습니다.")
        os.remove(self.xlsx_path)
        self.print_button.setDisabled(False)

if __name__ == '__main__':
    MyApp = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(MyApp.exec_())