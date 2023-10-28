import sys
import os
import win32com.client
import win32print
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QPushButton
from PyQt5.QtCore import QSettings
from gui import Ui_MyWindow
from selenium import webdriver

class MyWindow(QMainWindow, Ui_MyWindow):
    xlsx_path = ""
    driver = None

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.select_file_button.clicked.connect(self.select_excel_file)
        self.print_button.clicked.connect(self.print_excel_file)
        self.populate_printer_list()
        self.populate_book_number()
        self.room_select()

    def find_chrome_profile():
        user_data_path = os.path.expanduser("~")  # 현재 사용자 홈 디렉토리
        chrome_user_data_path = os.path.join(user_data_path, 'AppData', 'Local', 'Google', 'Chrome', 'User Data')
        
        profile_name = os.listdir(chrome_user_data_path)[0] if os.path.exists(chrome_user_data_path) else "Default"
        
        profile_path = os.path.join(chrome_user_data_path, profile_name)
        return profile_path
    
    def connect_chrome_session(self):
        try :
            chrome_profile_path = self.find_chrome_profile()
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument(f'--user-data-dir={chrome_profile_path}')
            self.driver = webdriver.Chrome(options=chrome_options)
        except:
            self.show_message("크롬 Klas에 접속에 실패 했습니다.")
            return False
    
    def download_excel(self):
        self.driver.get("https://klas.jeonju.go.kr/klas3/Admin/")

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

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_dialog = QFileDialog()
        file_dialog.setOptions(options)
        download_folder = os.path.expanduser("~/Downloads")
        file_dialog.setDirectory(download_folder)
        file_path, _ = file_dialog.getOpenFileName(self, 'Select Excel File', '', 'Excel Files (*.xlsx *.xls)')

        if file_path:
            self.xlsx_path = file_path

    def show_message(self, text):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(text)
        msg_box.setWindowTitle("Info")
        msg_box.exec_()

    def print_excel_file(self):
        try:
            if self.connect_chrome_session(self):
                return
            
            rooms = self.room_text_label.text()
            room = rooms.split()
            new = int(self.book_num_text_label.text())

            printer_name = self.printer_combo.currentText()
            if not hasattr(self, 'xlsx_path'):
                raise Exception('No Excel file selected')
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(self.xlsx_path)
            output_file_path = os.path.join(os.path.expanduser("~/Desktop"), "제공자료.xls")

            if os.path.exists(output_file_path):
                os.remove(output_file_path)

            for sheet in workbook.Sheets:
                if not sheet == workbook.Worksheets[0]:
                    sheet.Delete()
            worksheet = workbook.Worksheets[0]
            worksheet.Cells.ClearFormats()
            worksheet.Rows("1:3").EntireRow.Delete()

            columns_to_delete = [1, 2, 3, 8, 9, 10, 11]
            for column_index in sorted(columns_to_delete, reverse=True):
                worksheet.Columns(column_index).Delete()

            worksheet.UsedRange.Sort(worksheet.UsedRange.Columns("D"), Header=1)

            start = None
            end = None
            
            for row in range(1, worksheet.UsedRange.Rows.Count + 1):
                cell_room_value = worksheet.Cells(row, 4).Value                
                condition1 = any(name in cell_room_value for name in room)
                
                if condition1 and start is None:
                    start = row
                elif start is not None and not condition1:
                    end = row - 1
                    break
                
                if row is worksheet.UsedRange.Rows.Count :
                    end = row -1

            worksheet.Columns(5).ClearFormats()
            worksheet.Columns(5).Delete()
            worksheet.Columns(4).ClearFormats()
            worksheet.Columns(4).Delete()

            if start is not None and end is not None:
                worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(start - 1, 3)).EntireRow.Delete()
                worksheet.Range(worksheet.Cells(end + 1, 1), worksheet.Cells(worksheet.UsedRange.Rows.Count + 1, 3)).EntireRow.Delete()
            
            if end is None and start is None :
                self.show_message("해당 자료실의 제공자료가 존재하지 않습니다.")
                workbook.Close(SaveChanges=False)
                excel.Quit()
                return

            worksheet.UsedRange.Sort(worksheet.UsedRange.Columns("C"), Header=1)
            
            for row in range(1, worksheet.UsedRange.Rows.Count + 1):
                rowA = worksheet.Cells(row, 1).Value
                cell_new_value = int(rowA[2:])
                for column_index in range(1, 4):
                    cell = worksheet.Cells(row, column_index)
                    if cell_new_value >= new :
                        cell.Font.Color = 255
                    cell.Font.Size = 20

            worksheet.Rows.AutoFit()
            worksheet.Columns.AutoFit()

            worksheet.PageSetup.PaperSize = 9  # A4
            worksheet.PageSetup.Orientation = 2  # Landscape
            worksheet.PageSetup.LeftMargin = 0.2
            worksheet.PageSetup.RightMargin = 0.2
            worksheet.PageSetup.TopMargin = 0.2
            worksheet.PageSetup.BottomMargin = 0.2

            worksheet.PageSetup.Zoom = False
            worksheet.PageSetup.FitToPagesWide = 1
            worksheet.PageSetup.FitToPagesTall = 1

            # workbook.SaveAs(output_file_path)
            worksheet.PrintOut(ActivePrinter=printer_name)
            workbook.Close(SaveChanges=False)
            excel.Quit()
            self.show_message("작업이 완료되었습니다.")
        except Exception as e:
            self.show_message(str(e))
            # workbook.SaveAs(output_file_path)
            workbook.Close(SaveChanges=False)
            excel.Quit()


if __name__ == '__main__':
    MyApp = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(MyApp.exec_())
