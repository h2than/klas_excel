import sys
import os
import win32com.client
import win32print
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QPushButton
from PyQt5.QtCore import QSettings
from gui import Ui_MyWindow
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import glob
import time

import chromedriver_autoinstaller

class MyWindow(QMainWindow, Ui_MyWindow):
    xlsx_path = ""
    driver = None
    download_folder = os.path.expanduser("~/Downloads")

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.print_button.clicked.connect(self.print_excel_file)

        self.populate_printer_list()
        self.populate_book_number()
        self.room_select()
        try:
            self.get_xls()
        except:
            return
        try :
            chromedriver_autoinstaller.install()
        except:
            pass

    def get_xls(self):
        wait_time = 5
        while wait_time > 0 :
            files = glob.glob(os.path.join(self.download_folder, '*'))
            files.sort(key=os.path.getctime, reverse=True)

            if files:
                most_recent_file = files[0]
                file_extension = os.path.splitext(most_recent_file)[-1]
                
                if file_extension == '.xls' and 'download' not in os.path.basename(most_recent_file):
                    self.xlsx_path = most_recent_file
                    break
            else:
                time.sleep(1)
                wait_time -= 1

        if self.xlsx_path == "":
            self.show_message("다운로드된 .xls 파일을 찾을 수 없거나 대기 시간을 초과했습니다.")

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
        try:        
            rooms = self.room_text_label.text()
            room = rooms.split()
            new = int(self.book_num_text_label.text())

            printer_name = self.printer_combo.currentText()
            if not hasattr(self, 'xlsx_path'):
                self.show_message('No Excel file selected')
                return

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(self.xlsx_path)

            for sheet in workbook.Sheets:
                if not sheet == workbook.Worksheets[0]:
                    sheet.Delete()
            worksheet = workbook.Worksheets[0]
            worksheet.Cells.ClearFormats()

            worksheet.Rows("1:3").EntireRow.Delete()
            columns_to_delete = [1, 2, 3, 8, 9, 10, 11,12]
            for column_index in sorted(columns_to_delete, reverse=True):
                worksheet.Columns(column_index).Delete()

            worksheet.UsedRange.Sort(worksheet.UsedRange.Columns("C"), Header=1)

            start = None
            total_rows = worksheet.UsedRange.Rows.Count + 1
            end = total_rows
            
            for row in range(1, total_rows):
                cell_room_value = worksheet.Cells(row, 4).Value                
                condition1 = any(name in cell_room_value for name in room)
                
                if condition1 and start is None:
                    start = row
                elif start is not None and not condition1:
                    end = row
                    break

            worksheet.Columns(4).Delete()

            if start is None :
                self.show_message("해당 자료실의 제공자료가 존재하지 않습니다.")
                workbook.Close(SaveChanges=False)
                excel.Quit()
                return
            else :
                if start > 2 :
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(start - 1, 3)).EntireRow.Delete()
                if end < total_rows :
                    worksheet.Range(worksheet.Cells(end, 1), worksheet.Cells(total_rows, 3)).EntireRow.Delete()
            
            for row in range(1, total_rows):
                rowA = worksheet.Cells(row, 1).Value
                cell_new_value = int(rowA[2:])
                for column_index in range(1, 4):
                    cell = worksheet.Cells(row, column_index)
                    if cell_new_value >= new :
                        cell.Font.Color = 255
                    cell.Font.Size = 20

            worksheet.Rows.AutoFit()
            worksheet.Columns.AutoFit()
            
            rows_per_page = 20  # A4 용지에 출력할 행 수
            current_row = 1  # 현재 출력 중인 행
            last_row = 1
            
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
                    print_range.PrintOut(ActivePrinter=printer_name)
                    break

                if current_row % rows_per_page == 0:
                    print_range = worksheet.Range(worksheet.Cells(last_row, 1), worksheet.Cells(current_row, 3))
                    last_row = current_row
                    print_range.PrintOut(ActivePrinter=printer_name)

                current_row += 1

            # workbook.SaveAs(self.download_folder + "\\test.xls")
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.show_message("작업이 완료되었습니다.")
        except Exception as e:
            self.show_message(str(e))


if __name__ == '__main__':
    MyApp = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(MyApp.exec_())