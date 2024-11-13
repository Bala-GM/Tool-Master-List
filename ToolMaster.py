import sys
import msoffcrypto
import io
import json
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem,
                             QTabWidget, QAbstractItemView, QPushButton, QLabel, QMessageBox, QStackedWidget)
from PyQt5.QtCore import Qt
from openpyxl import load_workbook


# Dummy data to store credentials
USERS = {
    "admin": {"password": "admin#123", "role": "admin"},
    "Process": {"password": "Pro123", "role": "process"},
    "OP": {"password": "OP", "role": "operator"}
}

# Global variable to store file paths and passwords
config_data = {
    "file_path_process": None,
    "file_path_operator": None,
    "workbook_password": None,
    "sheet_password": None,
    "history_path": None
}


class LoginPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.setWindowTitle("Login")
        layout = QVBoxLayout()

        # Username input
        self.username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)

        # Password input
        self.password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)

        # Login button
        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.authenticate)
        layout.addWidget(self.login_button)

        self.setLayout(layout)

    def authenticate(self):
        username = self.username_input.text()
        password = self.password_input.text()
        user_info = USERS.get(username)

        if user_info and user_info["password"] == password:
            self.parent.user_role = user_info["role"]
            self.parent.show_role_specific_page()
        else:
            QMessageBox.warning(self, "Login Failed", "Invalid username or password")


class AdminConfigurationPage(QWidget):
    def __init__(self, layout):
        super().__init__()
        self.setWindowTitle("Admin Configuration")
        self.layout = layout
        self.setup_widgets()

    def setup_widgets(self):
        # Create widgets for file path and passwords input
        self.file_path_process_label = QLabel("Set Process Excel File Path:")
        self.file_path_process_input = QLineEdit()
        self.file_path_operator_label = QLabel("Set Operator Excel File Path:")
        self.file_path_operator_input = QLineEdit()
        self.workbook_password_label = QLabel("Set Workbook Password:")
        self.workbook_password_input = QLineEdit()
        self.workbook_password_input.setEchoMode(QLineEdit.Password)
        self.sheet_password_label = QLabel("Set Sheet Password:")
        self.sheet_password_input = QLineEdit()
        self.sheet_password_input.setEchoMode(QLineEdit.Password)
        self.history_path_label = QLabel("Set Process Log History Path:")
        self.history_path_input = QLineEdit()

        # Save button to store all configurations
        self.save_button = QPushButton("Save Configuration")
        self.save_button.clicked.connect(self.save_configuration)

        # Add widgets to the layout
        self.layout.addWidget(self.file_path_process_label)
        self.layout.addWidget(self.file_path_process_input)
        self.layout.addWidget(self.file_path_operator_label)
        self.layout.addWidget(self.file_path_operator_input)
        self.layout.addWidget(self.workbook_password_label)
        self.layout.addWidget(self.workbook_password_input)
        self.layout.addWidget(self.sheet_password_label)
        self.layout.addWidget(self.sheet_password_input)
        self.layout.addWidget(self.history_path_label)
        self.layout.addWidget(self.history_path_input)
        self.layout.addWidget(self.save_button)

    def save_configuration(self):
        global config_data

        # Get values from the input fields
        file_path_process = self.file_path_process_input.text()
        file_path_operator = self.file_path_operator_input.text()
        workbook_password = self.workbook_password_input.text()
        sheet_password = self.sheet_password_input.text()
        history_path = self.history_path_input.text()

        # Store the values in the global configuration data
        config_data["file_path_process"] = file_path_process
        config_data["file_path_operator"] = file_path_operator
        config_data["workbook_password"] = workbook_password
        config_data["sheet_password"] = sheet_password
        config_data["history_path"] = history_path

        # Save configuration data to a JSON file
        try:
            with open("config_data.json", "w") as f:
                json.dump(config_data, f)
            QMessageBox.information(self, "Success", "Configuration has been successfully saved.")
        except IOError as e:
            QMessageBox.warning(self, "Error", f"Failed to save configuration: {str(e)}")


class ExcelViewerWithHomePage(QWidget):
    def __init__(self, file_path=None, role="operator"):
        super().__init__()
        self.user_role = role
        self.setWindowTitle("Excel Viewer with Home and Search Navigation")
        self.setGeometry(100, 100, 800, 600)
        self.history_path = config_data["history_path"]  # Path to store history for Process role

        # Initialize layout before adding any widgets
        layout = QVBoxLayout()  # Initialize the layout
        self.setLayout(layout)  # Set this layout to the widget

        # Widgets for admin file path and sheet name setup
        if self.user_role == "admin":
            self.setup_admin_widgets(layout)

        # If a file path is provided and exists, attempt to load the workbook
        if file_path:
            self.file_path = file_path
            self.load_workbook()

        # Setup for Process role to log actions
        if self.user_role == "process":
            self.log_action("Process login initiated")

        # Set up search and home page for all roles
        self.setup_search_and_home(layout)

    def setup_admin_widgets(self, layout):
        # File path setup input and button for Admin
        self.file_path_label = QLabel("Set Excel File Path:")
        self.file_path_input = QLineEdit()
        self.file_path_button = QPushButton("Set File Path")

        # Connect the button to the method to set file path
        self.file_path_button.clicked.connect(self.set_file_path)

        # Add widgets to layout
        layout.addWidget(self.file_path_label)
        layout.addWidget(self.file_path_input)
        layout.addWidget(self.file_path_button)

    def set_file_path(self):
        global config_data  # Use global variable to store the file path
        self.file_path = self.file_path_input.text()
        config_data["file_path_process"] = self.file_path  # Set global file path
        if self.file_path:
            QMessageBox.information(self, "File Path Set", "File path has been successfully set.")
        else:
            QMessageBox.warning(self, "Error", "Please provide a valid file path")

    def load_workbook(self):
        # Decrypt and load the password-protected Excel file for viewing
        try:
            decrypted_file = io.BytesIO()
            with open(self.file_path, "rb") as f:
                encrypted = msoffcrypto.OfficeFile(f)
                encrypted.load_key(password=config_data["workbook_password"])  # Use actual password or prompt admin to enter
                encrypted.decrypt(decrypted_file)

            # Load workbook from decrypted file content
            decrypted_file.seek(0)
            self.workbook = load_workbook(decrypted_file, data_only=True, keep_vba=True)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load workbook: {str(e)}")

    def setup_search_and_home(self, layout):
        # Set up search bar and tab widget
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search...")
        self.search_bar.textChanged.connect(self.search)

        self.tab_widget = QTabWidget()
        self.home_table = QTableWidget()
        self.home_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.home_table.cellClicked.connect(self.navigate_to_sheet)
        self.tab_widget.addTab(self.home_table, "Home")

        layout.addWidget(self.search_bar)
        layout.addWidget(self.tab_widget)

    def search(self, text):
        # Implement search logic here for all roles
        pass

    def navigate_to_sheet(self, row, column):
        # This method will be called when a cell in home_table is clicked.
        sheet_name = self.home_table.item(row, 0).text()  # Assuming first column has sheet names
        QMessageBox.information(self, "Sheet Navigation", f"Navigating to {sheet_name} (row: {row}, column: {column})")
        # Add logic here to switch to the corresponding sheet in the workbook or display its content

    def log_action(self, action):
        if self.user_role == "process":
            log_message = f"{datetime.now()}: {action}"
            try:
                with open(self.history_path, "a") as f:
                    f.write(log_message + "\n")
            except Exception as e:
                QMessageBox.warning(self, "Log Error", f"Failed to log action: {str(e)}")


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ToolMaster")
        self.setGeometry(100, 100, 800, 600)
        self.layout = QVBoxLayout()
        self.user_role = None
        self.login_page = LoginPage(self)
        self.layout.addWidget(self.login_page)
        self.setLayout(self.layout)

    def show_role_specific_page(self):
        # Show the page based on the user role
        if self.user_role == "admin":
            admin_view = ExcelViewerWithHomePage(role="admin")
            self.layout.addWidget(admin_view)
        else:
            operator_view = ExcelViewerWithHomePage(role=self.user_role)
            self.layout.addWidget(operator_view)

        self.login_page.hide()


# Initialize application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
