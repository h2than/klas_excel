import sys
import os
import win32com.client
import win32print
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QPushButton
from PyQt5.QtCore import QSettings
from gui import Ui_MyWindow

class MyWindow(QMainWindow, Ui_MyWindow):
    xlsx_path = ""

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.select_file_button.clicked.connect(self.select_excel_file)
        self.print_button.clicked.connect(self.print_excel_file)
        self.populate_printer_list()
        self.populate_book_number()
        self.room_select()

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
            rooms = self.room_text_label.text()
            room = rooms.split()
            if ',' in rooms :
                room = rooms.split(',')
            new = self.book_num_text_label.text()

            printer_name = self.printer_combo.currentText()
            if not hasattr(self, 'xlsx_path'):
                raise Exception('No Excel file selected')
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(self.xlsx_path)
            output_file_path = os.path.join(os.path.expanduser("~/Desktop"), "제공자료.xls")
            self.xlsx_path = output_file_path

            if os.path.exists(self.xlsx_path):
                os.remove(self.xlsx_path)

            workbook.SaveAs(self.xlsx_path)
            workbook.Close(SaveChanges=False)
            workbook = excel.Workbooks.Open(self.xlsx_path)

            for sheet in workbook.Sheets:
                if not sheet == workbook.Worksheets[0]:
                    sheet.Delete()
            worksheet = workbook.Worksheets[0]
            worksheet.Rows("1:3").EntireRow.Delete()

            columns_to_delete = [1, 2, 3, 8, 9, 10, 11]
            for column_index in sorted(columns_to_delete, reverse=True):
                worksheet.Columns(column_index).Delete()

            total_rows = worksheet.UsedRange.Rows.Count + 1
            total_columns = worksheet.UsedRange.Columns.Count + 1
            rows_to_delete = []
                
            for row in range(1, total_rows):
                cell_room_value = worksheet.Cells(row, 4).Value
                cell_new_value = worksheet.Cells(row, 1).Value

                # Check if any of the room values in the list do not exist in cell_room_value
                if any(name not in cell_room_value for name in room):
                    rows_to_delete.append(row)

                # Check if the numeric part of cell_new_value is greater than new
                elif int(cell_new_value[2:]) > int(new):
                    for column_index in range(1, total_columns):
                        cell = worksheet.Cells(row, column_index)
                        cell.Font.Color = 255  # Red
                        cell.Font.Size = 20
                else :
                    for column_index in range(1, total_columns):
                        cell = worksheet.Cells(row, column_index)
                        cell.Font.Size = 20  # Set the font size for all cells in the row                
                
            for row_index in reversed(rows_to_delete):
                worksheet.Rows(row_index).Delete()

            worksheet.Columns(4).Delete()

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

            workbook.Save()
            worksheet.PrintOut(ActivePrinter=printer_name)
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
