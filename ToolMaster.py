import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem, QTabWidget, QAbstractItemView
from PyQt5.QtCore import Qt
from openpyxl import load_workbook

class ExcelViewerWithHomePage(QWidget):
    def __init__(self, file_path):
        super().__init__()
        self.setWindowTitle("Excel Viewer with Home and Search Navigation")
        self.setGeometry(100, 100, 800, 600)

        # Load the Excel file
        self.workbook = load_workbook(file_path, data_only=True)

        # Set up search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search...")
        self.search_bar.textChanged.connect(self.search)

        # Set up the tab widget for all sheets
        self.tab_widget = QTabWidget()

        # Create a "Home" tab for search results
        self.home_table = QTableWidget()
        self.home_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.home_table.cellClicked.connect(self.navigate_to_sheet)
        self.tab_widget.addTab(self.home_table, "Home")

        # Initialize dictionary to track table widgets by sheet
        self.tables = {}
        self.load_sheets()

        # Set up layout
        layout = QVBoxLayout()
        layout.addWidget(self.search_bar)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def load_sheets(self):
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]

            # Create a QTableWidget for each sheet
            table_widget = QTableWidget()
            table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table_widget.setRowCount(sheet.max_row)
            table_widget.setColumnCount(sheet.max_column)

            # Load data into the table widget
            self.load_data(sheet, table_widget)
            self.tables[sheet_name] = table_widget

            # Add table widget as a new tab
            self.tab_widget.addTab(table_widget, sheet_name)

    def load_data(self, sheet, table_widget):
        # Populate table with data from the Excel sheet
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
        # Clear previous search results on the Home tab
        self.home_table.clear()
        self.home_table.setRowCount(0)
        self.home_table.setColumnCount(3)
        self.home_table.setHorizontalHeaderLabels(["Sheet", "Cell", "Content"])

        # Search across all sheets and highlight matching cells
        for sheet_name, table_widget in self.tables.items():
            for row in range(table_widget.rowCount()):
                for col in range(table_widget.columnCount()):
                    item = table_widget.item(row, col)
                    if item and text.lower() in item.text().lower():
                        # Add matching cells to the home table as a search result
                        result_row = self.home_table.rowCount()
                        self.home_table.insertRow(result_row)
                        self.home_table.setItem(result_row, 0, QTableWidgetItem(sheet_name))
                        self.home_table.setItem(result_row, 1, QTableWidgetItem(f"{chr(65 + col)}{row + 1}"))
                        self.home_table.setItem(result_row, 2, QTableWidgetItem(item.text()))
                        item.setBackground(Qt.yellow)  # Highlight match in original sheet

    def navigate_to_sheet(self, row, column):
        # Retrieve sheet name and cell location from the clicked search result
        sheet_name = self.home_table.item(row, 0).text()
        cell_location = self.home_table.item(row, 1).text()
        target_row = int(cell_location[1:]) - 1
        target_col = ord(cell_location[0].upper()) - 65

        # Switch to the sheet tab and select the cell
        self.tab_widget.setCurrentWidget(self.tables[sheet_name])
        self.tables[sheet_name].setCurrentCell(target_row, target_col)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelViewerWithHomePage("C:\\Users\\Bala Ganesh\\Desktop\\ToolMaster.xlsx")  # Replace with your Excel file path
    viewer.show()
    sys.exit(app.exec_())



    
    
    
    # Replace with your Excel file path ("C:\\Users\\Bala Ganesh\\Desktop\\ToolMaster.xlsx")