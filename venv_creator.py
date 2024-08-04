import sys
import os
import subprocess
import re
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QComboBox,
    QFileDialog, QMessageBox, QTextEdit, QHBoxLayout
)
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import QSize


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class VenvCreator(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Python VENV Creator')
        self.setFixedSize(400, 400)
        self.setWindowIcon(QIcon(resource_path('asheshicon.ico')))

        layout = QVBoxLayout()

        # Banner
        banner = QLabel(self)
        pixmap = QPixmap(resource_path('asheshdevkitbanner.png'))
        banner.setPixmap(pixmap.scaled(self.width(), 100))
        layout.addWidget(banner)

        # Path selection
        path_layout = QHBoxLayout()
        self.path_label = QLabel('Select Path:')
        path_layout.addWidget(self.path_label)

        self.path_input = QLineEdit(self)
        path_layout.addWidget(self.path_input)

        self.browse_button = QPushButton('...', self)
        self.browse_button.setFixedWidth(30)
        self.browse_button.clicked.connect(self.browse_path)
        path_layout.addWidget(self.browse_button)

        layout.addLayout(path_layout)

        # Interpreter selection
        self.interpreter_label = QLabel('Select Python Interpreter:')
        layout.addWidget(self.interpreter_label)

        self.interpreter_combo = QComboBox(self)
        self.interpreter_combo.addItems(self.get_python_interpreters())
        layout.addWidget(self.interpreter_combo)

        # Create button
        self.create_button = QPushButton('Create VENV', self)
        self.create_button.clicked.connect(self.create_venv)
        layout.addWidget(self.create_button)

        # Console for debug logs
        self.console = QTextEdit(self)
        self.console.setReadOnly(True)
        layout.addWidget(self.console)

        self.setLayout(layout)

        # Author credit
        self.author_label = QLabel('Ashesh Development © 2024', self)
        layout.addWidget(self.author_label)

    def browse_path(self):
        path = QFileDialog.getExistingDirectory(self, 'Select Directory')
        if path:
            self.path_input.setText(path)

    def get_python_interpreters(self):
        interpreters = []
        try:
            result = subprocess.run(["py", "--list"], capture_output=True, text=True)
            matches = re.findall(r' - (\d+\.\d+\.\d+)', result.stdout)
            for match in matches:
                interpreters.append(f"python{match}")
        except Exception as e:
            self.log(f"Error fetching Python interpreters: {e}", 'error')

        # Ensure we have a minimum set of versions if fetching fails or no interpreters found
        if not interpreters:
            interpreters = ['python3', 'python3.8', 'python3.9', 'python3.10', 'python3.11', 'python3.12.4']

        return interpreters

    def create_venv(self):
        path = self.path_input.text()
        interpreter = self.interpreter_combo.currentText()

        if not path:
            QMessageBox.warning(self, 'Error', 'Please select a path.')
            return

        venv_path = os.path.join(path, 'venv')

        try:
            self.log(f'Creating virtual environment at {venv_path} using {interpreter}...')
            subprocess.check_call([interpreter, '-m', 'venv', venv_path])
            self.log(f'Success: Virtual environment created at {venv_path}', 'success')
            QMessageBox.information(self, 'Success', f'Virtual environment created at {venv_path}')
        except subprocess.CalledProcessError as e:
            self.log(f'Error: Failed to create virtual environment: {e}', 'error')
            QMessageBox.critical(self, 'Error', f'Failed to create virtual environment: {e}')

    def log(self, message, level='info'):
        color = {
            'info': 'black',
            'success': 'green',
            'error': 'red'
        }.get(level, 'black')
        self.console.append(f'<span style="color:{color}">{message}</span>')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = VenvCreator()
    ex.show()
    sys.exit(app.exec())
