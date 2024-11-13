import sys
import json
import os
import msoffcrypto
import io
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem,
                             QTabWidget, QAbstractItemView, QPushButton, QLabel, QMessageBox, QFileDialog, QStackedWidget)
from PyQt5.QtCore import Qt
from openpyxl import load_workbook

# Function to load config.json
def load_config(config_path="config.json"):
    if os.path.exists(config_path):
        with open(config_path, "r") as config_file:
            return json.load(config_file)
    else:
        QMessageBox.warning(None, "Error", "Config file not found!")
        return {}

# Dummy data to store credentials
USERS = {
    "admin": {"password": "admin#123", "role": "admin"},
    "Process": {"password": "Pro123", "role": "process"},
    "OP": {"password": "OP", "role": "operator"}
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

class ExcelViewerWithHomePage(QWidget):
    def __init__(self, role="operator"):
        super().__init__()
        self.user_role = role
        self.setWindowTitle("Excel Viewer")
        self.setGeometry(100, 100, 800, 600)

        self.config_path = "config.json"  # Default config path
        self.history_path = "history_log.json"
        self.config_data = load_config(self.config_path)

        self.file_path = self.config_data.get(f"{self.user_role}_file_path", "")
        self.workbook_password = self.config_data.get("workbook_password", "")
        self.sheet_passwords = self.config_data.get("sheet_passwords", {})

        if self.file_path:
            self.load_workbook()

        self.setup_widgets()

    def load_workbook(self):
        # Decrypt and load the password-protected Excel file for viewing
        decrypted_file = io.BytesIO()
        with open(self.file_path, "rb") as f:
            encrypted = msoffcrypto.OfficeFile(f)
            encrypted.load_key(password=self.workbook_password)  # Use the password from config
            encrypted.decrypt(decrypted_file)

        decrypted_file.seek(0)
        self.workbook = load_workbook(decrypted_file, data_only=True, keep_vba=True)

    def setup_widgets(self):
        # Set up the UI (search bar, tabs, etc.)
        self.tab_widget = QTabWidget()

        # Home tab
        self.home_table = QTableWidget()
        self.home_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tab_widget.addTab(self.home_table, "Home")

        # Load config.json content into the home page if the role is admin
        if self.user_role == "admin":
            self.load_config_data_to_home()

            # Admin-specific settings (like file path configuration)
            self.setup_admin_widgets()

        layout = QVBoxLayout()
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def load_config_data_to_home(self):
        # Load config data into home tab as a table
        self.home_table.clear()
        self.home_table.setColumnCount(2)
        self.home_table.setRowCount(len(self.config_data))
        self.home_table.setHorizontalHeaderLabels(["Key", "Value"])

        for row, (key, value) in enumerate(self.config_data.items()):
            self.home_table.setItem(row, 0, QTableWidgetItem(key))
            self.home_table.setItem(row, 1, QTableWidgetItem(str(value)))

    def setup_admin_widgets(self):
        # Admin can set file paths and passwords
        settings_tab = QWidget()
        settings_layout = QVBoxLayout()

        # Config File Path Display
        self.config_path_label = QLabel("Config File Path:")
        self.config_path_display = QLabel(self.config_path)
        self.config_browse_button = QPushButton("Change Config File Path")
        self.config_browse_button.clicked.connect(self.change_config_path)

        # Process File Path
        self.process_file_path_label = QLabel("Process File Path:")
        self.process_file_path_input = QLineEdit(self.config_data.get("process_file_path", ""))
        self.process_file_path_button = QPushButton("Browse Process File Path")
        self.process_file_path_button.clicked.connect(self.browse_process_file_path)

        # OP File Path
        self.op_file_path_label = QLabel("OP File Path:")
        self.op_file_path_input = QLineEdit(self.config_data.get("op_file_path", ""))
        self.op_file_path_button = QPushButton("Browse OP File Path")
        self.op_file_path_button.clicked.connect(self.browse_op_file_path)

        # Workbook Password
        self.workbook_password_label = QLabel("Workbook Password:")
        self.workbook_password_input = QLineEdit(self.workbook_password)

        # Save Button
        self.save_config_button = QPushButton("Save Settings")
        self.save_config_button.clicked.connect(self.save_config)

        # Add widgets to layout
        settings_layout.addWidget(self.config_path_label)
        settings_layout.addWidget(self.config_path_display)
        settings_layout.addWidget(self.config_browse_button)

        settings_layout.addWidget(self.process_file_path_label)
        settings_layout.addWidget(self.process_file_path_input)
        settings_layout.addWidget(self.process_file_path_button)

        settings_layout.addWidget(self.op_file_path_label)
        settings_layout.addWidget(self.op_file_path_input)
        settings_layout.addWidget(self.op_file_path_button)

        settings_layout.addWidget(self.workbook_password_label)
        settings_layout.addWidget(self.workbook_password_input)
        settings_layout.addWidget(self.save_config_button)

        settings_tab.setLayout(settings_layout)
        self.tab_widget.addTab(settings_tab, "Settings")

    def change_config_path(self):
        # Allow the admin to change the config.json path
        new_path, _ = QFileDialog.getOpenFileName(self, "Select Config File", "", "JSON Files (*.json)")
        if new_path:
            self.config_path = new_path
            self.config_path_display.setText(self.config_path)
            self.config_data = load_config(self.config_path)  # Reload config with new path
            self.load_config_data_to_home()  # Update the home table with new config data
            self.update_fields()

    def update_fields(self):
        # Update fields with loaded config values
        self.process_file_path_input.setText(self.config_data.get("process_file_path", ""))
        self.op_file_path_input.setText(self.config_data.get("op_file_path", ""))
        self.workbook_password_input.setText(self.config_data.get("workbook_password", ""))

    def browse_process_file_path(self):
        # Select file path for Process file
        new_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx;*.xlsm)")
        if new_path:
            self.process_file_path_input.setText(new_path)

    def browse_op_file_path(self):
        # Select file path for OP file
        new_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx;*.xlsm)")
        if new_path:
            self.op_file_path_input.setText(new_path)

    def save_config(self):
        # Save updated config data to config.json
        self.config_data["process_file_path"] = self.process_file_path_input.text()
        self.config_data["op_file_path"] = self.op_file_path_input.text()
        self.config_data["workbook_password"] = self.workbook_password_input.text()

        with open(self.config_path, "w") as config_file:
            json.dump(self.config_data, config_file, indent=4)

        QMessageBox.information(self, "Settings Saved", "Configuration has been saved successfully.")

class MainApplication(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Application")
        self.user_role = None

        layout = QVBoxLayout()

        # Create stacked widget for login and role-based views
        self.stacked_widget = QStackedWidget()
        self.login_page = LoginPage(self)
        self.stacked_widget.addWidget(self.login_page)

        layout.addWidget(self.stacked_widget)
        self.setLayout(layout)

    def show_role_specific_page(self):
        # Show specific view for each role after login
        if self.user_role == "admin":
            admin_view = ExcelViewerWithHomePage(role="admin")
            self.stacked_widget.addWidget(admin_view)
            self.stacked_widget.setCurrentWidget(admin_view)
        elif self.user_role == "process":
            process_view = ExcelViewerWithHomePage(role="process")
            self.stacked_widget.addWidget(process_view)
            self.stacked_widget.setCurrentWidget(process_view)
        elif self.user_role == "operator":
            operator_view = ExcelViewerWithHomePage(role="operator")
            self.stacked_widget.addWidget(operator_view)
            self.stacked_widget.setCurrentWidget(operator_view)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_app = MainApplication()
    main_app.show()
    sys.exit(app.exec_())
