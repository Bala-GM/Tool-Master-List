import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem, QTabWidget, QAbstractItemView
from openpyxl import load_workbook
from PyQt5.QtCore import Qt

class ExcelViewerWithTabs(QWidget):
    def __init__(self, file_path):
        super().__init__()
        self.setWindowTitle("Excel Viewer with Tabs and Search - Read Only")
        self.setGeometry(100, 100, 800, 600)

        # Load the Excel file with openpyxl
        self.workbook = load_workbook(file_path, data_only=True)

        # Set up the search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search...")
        self.search_bar.textChanged.connect(self.search)

        # Set up the tab widget to hold each sheet
        self.tab_widget = QTabWidget()

        # Initialize a dictionary to keep track of table widgets per sheet
        self.tables = {}
        self.load_sheets()

        # Set up layout
        layout = QVBoxLayout()
        layout.addWidget(self.search_bar)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def load_sheets(self):
        # Iterate through each sheet in the workbook
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]

            # Create a QTableWidget for the sheet
            table_widget = QTableWidget()
            table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only mode
            table_widget.setRowCount(sheet.max_row)
            table_widget.setColumnCount(sheet.max_column)

            # Load data and formatting
            self.load_data(sheet, table_widget)
            self.tables[sheet_name] = table_widget

            # Add table widget to a new tab
            self.tab_widget.addTab(table_widget, sheet_name)

    def load_data(self, sheet, table_widget):
        # Populate table widget with sheet data
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                item = QTableWidgetItem(str(cell.value))
                item.setTextAlignment(Qt.AlignCenter)
                table_widget.setItem(cell.row - 1, cell.column - 1, item)

        # Handle merged cells
        for merged_range in sheet.merged_cells.ranges:
            start_row = merged_range.min_row - 1
            end_row = merged_range.max_row - 1
            start_col = merged_range.min_col - 1
            end_col = merged_range.max_col - 1
            table_widget.setSpan(start_row, start_col, end_row - start_row + 1, end_col - start_col + 1)

    def search(self, text):
        # Perform case-insensitive search across all tabs
        for sheet_name, table_widget in self.tables.items():
            for row in range(table_widget.rowCount()):
                for col in range(table_widget.columnCount()):
                    item = table_widget.item(row, col)
                    if item:
                        # Check if the item text contains the search term
                        item.setBackground(Qt.white)  # Reset background
                        if text.lower() in item.text().lower():
                            item.setBackground(Qt.yellow)  # Highlight matching cells

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelViewerWithTabs("C:\\Users\\Bala Ganesh\\Desktop\\ToolMaster.xlsx")  # Replace with your Excel file path
    viewer.show()
    sys.exit(app.exec_())


    
    
    
    # Replace with your Excel file path ("C:\\Users\\Bala Ganesh\\Desktop\\ToolMaster.xlsx")