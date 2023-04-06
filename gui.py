# create a pyqt6 application which shows 2 buttons at the top, one to add a new excel file and one to generate the pdf and saves it to a table below. If you click on one of the entries of the table, the pdf is displayed on the bottom right of the gui.

import sys
import csv
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QFileDialog, QHeaderView
from PyQt6.QtWidgets import QMainWindow, QMenu, QTableWidget
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6 import QtCore
from PyQt6.QtGui import QDesktopServices
from fpdf import FPDF
import unicodedata
from datetime import date
from excelparser import parse_excel_file

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "Portfolio Research Manager"
        self.top = 100
        self.left = 100
        self.width = 1280
        self.height = 720
        
        # set context menu policy for the table
        self.table = QTableWidget(self)
        self.table.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)

        self.init_ui()


    # Define the init_ui method
    def init_ui(self):
        # Set the window title and dimensions
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Create a QTableWidget object
        self.table = QTableWidget()
        
        # Columns
        table_columns = ["Date", "Company Name", "Ticker", "Close", "Target Value", "Invested?",  "Rating", "Performance since Report Date", "% Difference from Target Value"]
        
        # Set the number of columns to length of table_columns
        self.table.setColumnCount(len(table_columns))
        # Set the horizontal header labels
        self.table.setHorizontalHeaderLabels(table_columns)

        # Set the resize mode for the horizontal header sections
        for i in range(len(table_columns)):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)

        # Set the column widths
        for i in range(len(table_columns)):
            self.table.setColumnWidth(i, 100)
        
        # Setup the inital data for the table from the records.csv file and update the table with it
        with open("records.csv", "r") as f:
            reader = csv.reader(f)
            for row_number, row_data in enumerate(reader):
                self.table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.table.setItem(row_number, column_number, QTableWidgetItem(data))

        # Set the edit triggers, selection behavior, and selection mode for the table
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        
        # Double click on a row to open the pdf
        self.table.doubleClicked.connect(self.open_pdf)

        # Create a QHBoxLayout object for the buttons
        button_layout = QHBoxLayout()

        # Create a QPushButton object for adding Excel files and connect it to the add_excel_file slot
        button_add = QPushButton("Add Excel File")
        button_add.clicked.connect(self.add_excel_file)

        # Add the add_excel_file button to the button layout
        button_layout.addWidget(button_add)

        # Create a QVBoxLayout object for the overall layout and add the button layout and table to it
        layout = QVBoxLayout()
        layout.addLayout(button_layout)
        layout.addWidget(self.table)
        
        # Set the layout of the window to the QVBoxLayout object
        self.setLayout(layout)

        # Show the window
        self.show()
        
        
    def open_pdf(self):
            # Get the current row
            current_row = self.table.currentRow()
            # Get the current item in the first column of the current row
            current_ticker = self.table.item(current_row, 2)
            # Get the text of the current item
            current_item_text = current_ticker.text()
            # Open the pdf with the current item text
            QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(f"{current_item_text.lower()}.pdf"))


    def add_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            ### Read excel file and add it to the records.csv file and update the table
            company_name, company_ticker, company_stock_close, company_targetvalue, stock_rating, company_description = parse_excel_file(file_name)
            company_targetvalue = round(float(company_targetvalue), 2)
            company_stock_close = round(float(company_stock_close), 2)
            change_till_targetvalue = round((company_targetvalue/company_stock_close-1) * 100, 1)
            
            # update the table with the new data which is in csv format and then put the data to the csv file
            update_csv = f"{date.today().strftime('%d/%m/%Y')},{company_name},{company_ticker},{company_stock_close},{company_targetvalue},No,{stock_rating},Empty,{change_till_targetvalue}\n"
            
            # update the table with "update_csv"
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            for i in range(len(update_csv.split(","))):
                self.table.setItem(row_position, i, QTableWidgetItem(update_csv.split(",")[i]))
            
            # write the data to the CSV file
            with open("records.csv", "a") as f:
                f.write(update_csv)
            
            
            ### PDF generator
            pdf = FPDF(orientation="P", unit="pt", format="A4")

            # add page to the pdf
            pdf.add_page()

            # set a font
            pdf.set_font(family="Arial", size=24, style="B")

            # create a title (w is the width of the cell, if 0 then it will be the width of the page)
            # align = C means center, border = 1 adds a border, ln = 1 means go to the next line
            pdf.cell(w=0, h=50, txt=f"{company_name} Report", align="C", border=1, ln=1)

            # set a font
            pdf.set_font(family="Arial", size=12)

            pdf.cell(w=0, h=16, txt="", ln=1)
            pdf.cell(w=0, h=16, txt=f"Report date: {date.today().strftime('%d/%m/%Y')}", ln=1)
            pdf.cell(w=0, h=16, txt=f"Ticker: {company_ticker}", ln=1)
            pdf.cell(w=0, h=16, txt=f"Close: {company_stock_close}$", ln=1)
            pdf.cell(w=0, h=16, txt="", ln=1)

            # about the company
            pdf.set_font(family="Arial", size=18, style="B")
            pdf.cell(w=0, h=16, txt=f"About", ln=1)

            pdf.set_font(family="Arial", size=12)
            # multi_cell is used for multiple lines of text
            pdf.multi_cell(w=0, h=16, txt=f"{unicodedata.normalize('NFKD', company_description).encode('ASCII', 'ignore').decode('ASCII')}")

            # output the pdf
            pdf.output(f"{company_ticker.lower()}.pdf")


# open the app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())