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
from excelparser import company_name, company_ticker, company_stock_close, company_marketcap, company_n_shares_outstanding, df_dcf, company_description, company_targetvalue

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "Portfolio Research Manager"
        self.top = 100
        self.left = 100
        self.width = 800
        self.height = 600
        
        # set context menu policy for the table
        self.table = QTableWidget(self)
        self.table.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.create_context_menu)

        self.init_ui()

    # Define the init_ui method
    def init_ui(self):
        # Set the window title and dimensions
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Create a QTableWidget object
        self.table = QTableWidget()
        # Set the number of columns to 2
        self.table.setColumnCount(5)
        # Set the horizontal header labels
        self.table.setHorizontalHeaderLabels(["Date", "Company Name", "Ticker", "Close", "Target Value"])

        # Set the resize mode for the horizontal header sections
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        # Set the column widths
        self.table.setColumnWidth(0, 300)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(1, 100)
        
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
        
        # Connect the customContextMenuRequested signal of the table to the create_context_menu method
        self.table.customContextMenuRequested.connect(self.create_context_menu)

        # Show the window
        self.show()


    def add_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_name:            
            # add the sample to the records.csv file
            with open("records.csv", "a") as f:
                f.write(f"{date.today().strftime('%d/%m/%Y')},{company_name},{company_ticker},{company_stock_close},{company_targetvalue}\n")
                
            self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.table.customContextMenuRequested.connect(self.right_click_menu)

            # read the data from the CSV file and add it to the table
            with open("records.csv", "r") as f:
                reader = csv.reader(f)
                for row in reader:
                    row_position = self.table.rowCount()
                    self.table.insertRow(row_position)
                    for i in range(len(row)):
                        self.table.setItem(row_position, i, QTableWidgetItem(row[i]))

            # row_position = self.table.rowCount()
            # self.table.insertRow(row_position)
            # self.table.setItem(row_position, 1, QTableWidgetItem(company_name))
            # self.table.setItem(row_position, 0, QTableWidgetItem(date.today().strftime("%d/%m/%Y")))
            # self.table.setItem(row_position, 2, QTableWidgetItem(company_ticker))
            # self.table.setItem(row_position, 3, QTableWidgetItem(str(company_stock_close)))
            # self.table.setItem(row_position, 4, QTableWidgetItem(str(company_targetvalue)))


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
            pdf.output(f"{company_name}.pdf")
            
            # save the pdf to the table
            self.table.setItem(row_position, 2, QTableWidgetItem(f"{company_name}.pdf"))
            
    
    def right_click_menu(self, position):
        menu = QMenu()
        delete_action = menu.addAction("Delete")
        action = menu.exec_(self.table.mapToGlobal(position))
        if action == delete_action:
            current_row = self.table.currentRow()
            self.table.removeRow(current_row)
            with open("records.csv", "r") as f:
                rows = list(csv.reader(f))
                rows.pop(current_row)
            with open("records.csv", "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
                
                
    def show_pdf(self, row, column):
        current_item = self.table.item(row, column)
        if current_item:
            company = current_item.text()
            pdf_file = f"{company}.pdf"
            QDesktopServices.openUrl(f"file:///{pdf_file}")
            
    def create_context_menu(self, pos):
        # create context menu
        menu = QMenu(self)
        edit_action = QAction("Edit", self)
        delete_action = QAction("Delete", self)
        menu.addAction(edit_action)
        menu.addAction(delete_action)
        
        # display context menu at cursor position
        action = menu.exec_(self.table.viewport().mapToGlobal(pos))
        
        # handle selected action
        if action == edit_action:
            # edit the selected item
            pass
        elif action == delete_action:
            # delete the selected item
            pass

# open the app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())