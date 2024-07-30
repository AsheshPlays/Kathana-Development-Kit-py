import sys
import os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, \
    QFileDialog, QSizePolicy, QTextEdit, QSpacerItem
from PyQt6.QtGui import QPixmap, QIcon, QPalette, QColor, QFont, QPainter, QPolygon
from PyQt6.QtCore import Qt, QProcess, QThread, pyqtSignal, QPoint
from kathana_clx_pyqt6 import copy_and_sort_files, generate_fbx_files, generate_combined_fbx_batch_file, clean_up

KATHANA_DISPLAY_NAMES = ["Kathana Global", "Kathana 2", "Kathana 3", "Kathana 3.2", "Kathana 4", "Kathana 5.2",
                         "Kathana 6"]
KATHANA_VERSIONS = [r"B:\\Kathana\\Kathana-Global", r"B:\\Kathana\\Kathana2", r"B\\Kathana\\Kathana3",
                    r"B\\Kathana\\Kathana3.2", r"B\\Kathana\\Kathana4", r"B\\Kathana\\Kathana5.2",
                    r"B\\Kathana\\Kathana6"]

class Worker(QThread):
    progress = pyqtSignal(int)
    output = pyqtSignal(str)
    error = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, task, *args, **kwargs):
        super().__init__()
        self.task = task
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            self.task(*self.args, **self.kwargs)
        except Exception as e:
            self.error.emit(str(e))
        self.finished.emit()

class ArrowButton(QPushButton):
    def __init__(self, direction, parent=None):
        super().__init__(parent)
        self.direction = direction
        self.setFixedSize(30, 30)
        self.setStyleSheet("background-color: transparent; border: none;")

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor("red"))
        painter.setPen(Qt.PenStyle.NoPen)

        if self.direction == "left":
            points = [QPoint(self.width(), 0), QPoint(0, self.height() // 2), QPoint(self.width(), self.height())]
        elif self.direction == "right":
            points = [QPoint(0, 0), QPoint(self.width(), self.height() // 2), QPoint(0, self.height())]

        triangle = QPolygon(points)
        painter.drawPolygon(triangle)

class KathanaVersionTool(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_index = 0
        self.selected_version = KATHANA_VERSIONS[self.selected_index]
        self.process = QProcess(self)
        self.version_set = False
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Kathana Version Selector')
        self.setFixedSize(800, 600)
        self.setWindowIcon(QIcon('asheshicon.png'))

        QApplication.setFont(QFont("Dotum", 8))

        layout = QVBoxLayout()

        self.banner_label = QLabel(self)
        pixmap = QPixmap('asheshdevkitbanner.png')
        self.banner_label.setPixmap(pixmap)
        self.banner_label.setScaledContents(True)
        self.banner_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.banner_label)

        version_selector_layout = QHBoxLayout()
        version_selector_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))

        left_arrow_btn = ArrowButton("left", self)
        left_arrow_btn.clicked.connect(self.select_previous_version)
        version_selector_layout.addWidget(left_arrow_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        self.selected_version_label = QLabel(KATHANA_DISPLAY_NAMES[self.selected_index])
        self.selected_version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.selected_version_label.setFont(QFont("Dotum", 16, QFont.Weight.Bold))
        palette = self.selected_version_label.palette()
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.red if self.version_set else QColor("orange"))
        self.selected_version_label.setPalette(palette)
        version_selector_layout.addWidget(self.selected_version_label, alignment=Qt.AlignmentFlag.AlignCenter)

        right_arrow_btn = ArrowButton("right", self)
        right_arrow_btn.clicked.connect(self.select_next_version)
        version_selector_layout.addWidget(right_arrow_btn, alignment=Qt.AlignmentFlag.AlignRight)

        version_selector_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        layout.addLayout(version_selector_layout)

        set_version_btn = QPushButton('Set Version', self)
        set_version_btn.clicked.connect(self.set_version)
        layout.addWidget(set_version_btn)

        browse_btn = QPushButton('Browse for Kathana Version Directory', self)
        browse_btn.clicked.connect(self.browse_version)
        layout.addWidget(browse_btn)

        buttons_layout = QHBoxLayout()

        col1_layout = QVBoxLayout()
        self.add_button_row(
            col1_layout, [('Copy and Sort PC Files', lambda: self.run_task('PC')),
                          ('Copy and Sort NPC Files', lambda: self.run_task('NPC')),
                          ('Copy and Sort Monster Files', lambda: self.run_task('Monster')),
                          ('Copy and Sort All Entity Files', lambda: self.run_task('All'))]
        )
        buttons_layout.addLayout(col1_layout)

        col2_layout = QVBoxLayout()
        self.add_button_row(
            col2_layout, [('Generate PC FBX Files', lambda: self.run_fbx_task('PC')),
                          ('Generate NPC FBX Files', lambda: self.run_fbx_task('NPC')),
                          ('Generate Monster FBX Files', lambda: self.run_fbx_task('Monster')),
                          ('Generate All Entity FBX Files', lambda: self.run_fbx_task('All'))]
        )
        buttons_layout.addLayout(col2_layout)

        col3_layout = QVBoxLayout()
        self.add_button_row(
            col3_layout, [('Generate PC FBX Batch File Only', lambda: self.run_fbx_task('PC', True)),
                          ('Generate NPC FBX Batch File Only', lambda: self.run_fbx_task('NPC', True)),
                          ('Generate Monster FBX Batch File Only', lambda: self.run_fbx_task('Monster', True)),
                          ('Generate All FBX Batch File Only', self.run_combined_fbx_batch_file)]
        )
        buttons_layout.addLayout(col3_layout)

        layout.addLayout(buttons_layout)

        control_buttons_layout = QHBoxLayout()
        self.add_button_row(
            control_buttons_layout,
            [('Clean Up', self.clean_up), ('Stop', self.stop_processes), ('Restart', self.restart_application)]
        )
        layout.addLayout(control_buttons_layout)

        console_section_layout = QVBoxLayout()

        console_label = QLabel('Console')
        console_label.setFont(QFont("Dotum", 10, QFont.Weight.Bold))
        console_section_layout.addWidget(console_label)

        console_layout = QHBoxLayout()

        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        console_palette = self.console_output.palette()
        console_palette.setColor(QPalette.ColorRole.Base, QColor(200, 200, 200))
        self.console_output.setPalette(console_palette)
        console_layout.addWidget(self.console_output)

        clear_console_btn = QPushButton('Clear \nConsole')
        clear_console_btn.setStyleSheet("background-color: red; color: white;")
        clear_console_btn.setFont(QFont("Dotum", 10))
        clear_console_btn.clicked.connect(self.clear_console)
        clear_console_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        clear_console_btn.setFixedWidth(150)
        console_layout.addWidget(clear_console_btn)

        console_section_layout.addLayout(console_layout)

        layout.addLayout(console_section_layout)

        author_version_layout = QHBoxLayout()
        author_label = QLabel('Ashesh Development © 2024')
        author_label.setFont(QFont("Dotum", 8))
        version_label = QLabel('Version: 1.0.4')
        version_label.setFont(QFont("Dotum", 8))
        author_version_layout.addWidget(author_label)
        author_version_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        author_version_layout.addWidget(version_label)
        layout.addLayout(author_version_layout)

        self.setLayout(layout)

        self.process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self.process.readyReadStandardOutput.connect(self.on_readyReadStandardOutput)
        self.process.readyReadStandardError.connect(self.on_readyReadStandardError)
        self.process.finished.connect(self.on_task_finished)

    def add_button_row(self, parent_layout, buttons):
        for label, callback in buttons:
            button = QPushButton(label)
            button.clicked.connect(callback)
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            parent_layout.addWidget(button)

    def update_selected_version(self):
        self.selected_version = KATHANA_VERSIONS[self.selected_index]
        self.selected_version_label.setText(KATHANA_DISPLAY_NAMES[self.selected_index])
        self.update_version_label_color()

    def select_previous_version(self):
        self.selected_index = (self.selected_index - 1) % len(KATHANA_VERSIONS)
        self.update_selected_version()

    def select_next_version(self):
        self.selected_index = (self.selected_index + 1) % len(KATHANA_VERSIONS)
        self.update_selected_version()

    def set_version(self):
        self.version_set = True
        self.update_version_label_color()
        QMessageBox.information(
            self, 'Version Set', f'Selected version set to: {KATHANA_DISPLAY_NAMES[self.selected_index]}'
        )

    def update_version_label_color(self):
        palette = self.selected_version_label.palette()
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.red if self.version_set else QColor("orange"))
        self.selected_version_label.setPalette(palette)

    def browse_version(self):
        options = QFileDialog.Options()
        options |= QFileDialog.Option.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self, "Select Kathana Version Directory", options=options)
        if directory:
            self.selected_version = directory
            self.selected_version_label.setText(os.path.basename(directory))
            self.version_set = True
            self.update_version_label_color()

    def run_task(self, entity_type):
        if self.selected_version:
            self.worker = Worker(copy_and_sort_files, self.selected_version, entity_type)
            self.worker.output.connect(self.append_output)
            self.worker.error.connect(self.append_error)
            self.worker.finished.connect(self.on_task_finished)
            self.worker.start()
        else:
            QMessageBox.warning(self, 'Error', 'Please choose a Kathana version first.')

    def run_fbx_task(self, entity_type, generate_batch_only=False):
        if self.selected_version:
            self.worker = Worker(generate_fbx_files, self.selected_version, entity_type, generate_batch_only)
            self.worker.output.connect(self.append_output)
            self.worker.error.connect(self.append_error)
            self.worker.finished.connect(self.on_task_finished)
            self.worker.start()
        else:
            QMessageBox.warning(self, 'Error', 'Please choose a Kathana version first.')

    def run_combined_fbx_batch_file(self):
        if self.selected_version:
            self.worker = Worker(generate_combined_fbx_batch_file, self.selected_version)
            self.worker.output.connect(self.append_output)
            self.worker.error.connect(self.append_error)
            self.worker.finished.connect(self.on_task_finished)
            self.worker.start()
        else:
            QMessageBox.warning(self, 'Error', 'Please choose a Kathana version first.')

    def clean_up(self):
        self.worker = Worker(clean_up)
        self.worker.output.connect(self.append_output)
        self.worker.error.connect(self.append_error)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def stop_processes(self):
        self.process.kill()
        QMessageBox.information(self, 'Process Stopped', 'All running processes have been stopped.')

    def restart_application(self):
        QApplication.quit()
        os.execl(sys.executable, sys.executable, *sys.argv)

    def clear_console(self):
        self.console_output.clear()

    def append_output(self, text):
        self.console_output.setTextColor(QColor("blue"))
        self.console_output.append(text)
        self.console_output.setTextColor(QColor("black"))

    def append_error(self, text):
        self.console_output.setTextColor(QColor("red"))
        self.console_output.append(f"ERROR: {text}")
        self.console_output.setTextColor(QColor("black"))

    def on_task_finished(self):
        self.console_output.setTextColor(QColor("green"))
        self.console_output.append("Task Finished: The task has been completed successfully.")
        self.console_output.setTextColor(QColor("black"))

    def on_readyReadStandardOutput(self):
        text = self.process.readAllStandardOutput().data().decode()
        self.append_output(text)

    def on_readyReadStandardError(self):
        text = self.process.readAllStandardError().data().decode()
        self.append_error(text)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    tool = KathanaVersionTool()
    tool.show()
    sys.exit(app.exec())
