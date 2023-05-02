import sys
import csv
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6 import QtCore
from datetime import date
from excelparser import parse_excel_file
from report_generator import create_report
from decorators import timeit
import shutil
import pandas as pd
import webbrowser
from table_formatting import DollarSignFormat, TitleCaseFormat, PercentFormat, DateFormat


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "Portfolio Research Manager"
        self.top = 100
        self.left = 100
        self.width = 1280
        self.height = 720

        # set context menu policy for the table
        self.init_ui()
        self.init_contextmenu()

    def init_contextmenu(self):
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.entry_contextmenu)

    def entry_contextmenu(self, position):
        self.menu = QMenu()

        # delete action
        self.delete_action = QAction("Delete")
        self.menu.addAction(self.delete_action)
        self.delete_action.triggered.connect(self.delete_entry)

        # open report
        self.show_report_action = QAction("Open Report")
        self.menu.addAction(self.show_report_action)
        self.show_report_action.triggered.connect(self.open_pdf)

        # open model
        self.show_model_action = QAction("Open Model")
        self.menu.addAction(self.show_model_action)
        self.show_model_action.triggered.connect(self.open_model)

        # open online resources
        self.search_online = QAction("Search Online")
        self.menu.addAction(self.search_online)
        self.search_online.triggered.connect(self.search_online_resources)

        self.menu.popup(QCursor.pos())

    @timeit
    def init_ui(self):
        # Set the window title and dimensions
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Create a QTableWidget object
        self.table = QTableWidget()

        # Columns
        table_columns = ["Date", "Company Name", "Ticker", "Close", "Target Value", "Invested?",
                         "Rating", "Performance since Report Date", "% Difference from Target Value"]

        self.table.setColumnCount(len(table_columns))
        self.table.setHorizontalHeaderLabels(table_columns)

        # Set the resize mode for the horizontal header sections
        for i in range(len(table_columns)):
            self.table.horizontalHeader().setSectionResizeMode(
                i, QHeaderView.ResizeMode.Stretch)

        # Set the column widths
        for i in range(len(table_columns)):
            self.table.setColumnWidth(i, 100)

        # Setup the inital data for the table from the records.csv file and update the table with it
        with open("records.csv", "r") as f:
            reader = csv.reader(f)
            for row_number, row_data in enumerate(reader):
                self.table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.table.setItem(
                        row_number, column_number, QTableWidgetItem(data))

        # setup column formatting
        self.table.setItemDelegateForColumn(3, DollarSignFormat(self.table))
        self.table.setItemDelegateForColumn(4, DollarSignFormat(self.table))
        self.table.setItemDelegateForColumn(1, TitleCaseFormat(self.table))
        self.table.setItemDelegateForColumn(8, PercentFormat(self.table))
        self.table.setItemDelegateForColumn(0, DateFormat(self.table))

        # Set the edit triggers, selection behavior, and selection mode for the table
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows)
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
        current_row = self.table.currentRow()
        current_ticker = self.table.item(current_row, 2).text()
        current_date = self.table.item(current_row, 0).text()

        QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(
            f"reports/{current_date}_{current_ticker.lower()}.pdf"))

    def open_model(self):
        current_row = self.table.currentRow()
        current_ticker = self.table.item(current_row, 2).text()
        current_date = self.table.item(current_row, 0).text()

        QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(
            f"models/{current_date}_{current_ticker.lower()}.xlsx"))

    def delete_entry(self):
        current_row = self.table.currentRow()
        self.table.removeRow(current_row)
        print(f"Deleted row {current_row}.")

        try:
            df = pd.read_csv('records.csv', header=None)
            print(f"Number of rows: {df.shape[0]}")

            if df.shape[0] == 1:
                df = pd.DataFrame()
            else:
                df.drop(current_row, inplace=True)
                df.dropna(how='all', inplace=True)
            df.to_csv('records.csv', index=False, header=False)
        except:
            print("Error cleaning csv file.")

    def add_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            # Read excel file and add it to the records.csv file and update the table
            company_name, company_ticker, company_stock_close, company_targetvalue, stock_rating, company_description = parse_excel_file(
                file_name)
            company_targetvalue = round(float(company_targetvalue), 2)
            company_stock_close = round(float(company_stock_close), 2)
            change_till_targetvalue = round(
                (company_targetvalue/company_stock_close-1) * 100, 1)
            todays_date = date.today().strftime('%Y%m%d')

            # update the table with the new data which is in csv format and then put the data to the csv file
            update_csv = f"{todays_date},{company_name},{company_ticker},{company_stock_close},{company_targetvalue},No,{stock_rating},Empty,{change_till_targetvalue}\n"

            # update the table
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            for i in range(len(update_csv.split(','))):
                self.table.setItem(row_position, i, QTableWidgetItem(
                    update_csv.split(",")[i]))

            # write the data to the CSV file
            with open("records.csv", "a") as f:
                f.write(update_csv)

            # generate the pdf report
            create_report(company_ticker, company_name, company_stock_close,
                          todays_date, company_targetvalue, stock_rating)

            # copy the xlsx file and save it to the models folder
            shutil.copyfile(
                file_name, f"models/{todays_date}_{company_ticker.lower()}.xlsx")

    def search_online_resources(self):
        current_row = self.table.currentRow()
        current_ticker = self.table.item(current_row, 2).text()

        websites = [f"https://seekingalpha.com/symbol/{current_ticker}",
                    f"http://openinsider.com/search?q={current_ticker}",
                    f"https://www.stratosphere.io/company/{current_ticker}/"]
        for site in websites:
            webbrowser.open(site)


# open the app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
