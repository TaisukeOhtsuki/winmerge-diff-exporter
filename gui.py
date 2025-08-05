# -*- coding: UTF-8 -*-
from typing import List, Optional
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QTextEdit, QPushButton, QFileDialog, QMessageBox
)
from PyQt6.QtGui import QFont, QPainter, QLinearGradient, QColor
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QTimer, QRect

from winmergexlsx import WinMergeXlsx

class DropLineEdit(QLineEdit):
    paths_dropped = pyqtSignal(list)

    def __init__(self, path_list: Optional[List[str]] = None) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.path_list = path_list or []

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        new_paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and path not in self.path_list:
                self.path_list.append(path)
                new_paths.append(path)
        self.setText("; ".join(self.path_list))
        self.paths_dropped.emit(new_paths)


class FadingProgressBar(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._value = 0
        self._minimum = 0
        self._maximum = 100
        self._is_complete = False
        self.setMinimumHeight(20)

    def setRange(self, minimum: int, maximum: int) -> None:
        self._minimum = minimum
        self._maximum = maximum
        self.update()

    def setValue(self, value: int) -> None:
        self._value = value
        self._is_complete = False
        self.update()

    def setComplete(self) -> None:
        self._is_complete = True
        self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        rect = self.rect()
        painter.fillRect(rect, QColor("#eeeeee"))

        if self._is_complete:
            painter.fillRect(rect, QColor(0, 255, 0))
        else:
            range_width = rect.width()
            percent = (self._value - self._minimum) / (self._maximum - self._minimum)
            center_x = int(range_width * percent)

            fade_width = int(range_width * 5 / (self._maximum - self._minimum))
            left_x = max(0, center_x - fade_width)
            right_x = min(range_width, center_x + fade_width)

            gradient = QLinearGradient(left_x, 0, right_x, 0)
            gradient.setColorAt(0.0, QColor(0, 255, 0, 0))
            gradient.setColorAt(0.5, QColor(0, 255, 0, 255))
            gradient.setColorAt(1.0, QColor(0, 255, 0, 0))

            painter.fillRect(QRect(left_x, 0, right_x - left_x, rect.height()), gradient)

        painter.setPen(Qt.GlobalColor.black)
        painter.drawRect(rect)


class Worker(QObject):
    finished = pyqtSignal()
    log_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, base: str, latest: str, output: str) -> None:
        super().__init__()
        self.base = base
        self.latest = latest
        self.output = output

    def emit_log(self, message: str, progress: Optional[int] = None) -> None:
        self.log_signal.emit(message)
        if progress is not None:
            self.progress_signal.emit(progress)

    def run(self) -> None:
        try:
            def log_callback(msg: str, value: int = None) -> None:
                self.emit_log(msg, value)

            diff = WinMergeXlsx(
                self.base, self.latest, self.output,
                log_callback=log_callback
            )
            diff.generate()
            self.emit_log(f"Completed! Output: {self.output}", 100)
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            self.finished.emit()


class DiffApp(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("WinMerge Diff to Excel")
        self.setGeometry(100, 100, 800, 360)

        self.base_paths: List[str] = []
        self.latest_paths: List[str] = []

        self.animation_timer = QTimer()
        self.animation_value = 0

        self.setup_widgets()
        self.setup_layout()
        self.setup_connections()
        self.apply_style()

    def setup_widgets(self) -> None:
        self.base_input = DropLineEdit(self.base_paths)
        self.latest_input = DropLineEdit(self.latest_paths)
        self.output_input = QLineEdit("output.xlsx")

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        self.progress_bar = FadingProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setVisible(False)

        self.run_button = QPushButton("Run (Compare and Export to Excel)")

    def setup_layout(self) -> None:
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)

        title = QLabel("WinMerge Diff to Excel")
        title.setFont(QFont("Helvetica", 20))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        layout.addLayout(self.create_path_row("1 Base Folder", self.base_input, self.browse_base))
        layout.addLayout(self.create_path_row("2 Comparison Target Folder", self.latest_input, self.browse_latest))
        layout.addLayout(self.create_path_row("3 Output File", self.output_input, self.browse_output))

        layout.addWidget(self.run_button)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_text)
        self.setLayout(layout)

    def setup_connections(self) -> None:
        self.run_button.clicked.connect(self.run_process)
        self.base_input.paths_dropped.connect(self.on_base_dropped)
        self.latest_input.paths_dropped.connect(self.on_latest_dropped)

    def apply_style(self) -> None:
        self.setStyleSheet("""
            QWidget {
                background-color: #fdfdfd;
                font-family: Helvetica, Arial, sans-serif;
                font-size: 14px;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #ccc;
                padding: 6px;
                border-radius: 4px;
            }
            QPushButton {
                background-color: #000;
                color: #fff;
                padding: 10px 20px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #333;
            }
        """)

    def create_path_row(self, label_text: str, line_edit: QLineEdit, browse_func) -> QHBoxLayout:
        layout = QHBoxLayout()
        label = QLabel(label_text)
        label.setFixedWidth(180)
        button = QPushButton("Browse")
        button.setFixedWidth(85)
        button.clicked.connect(browse_func)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button)
        return layout

    def browse_base(self) -> None:
        self.select_folder(self.base_paths, self.base_input)

    def browse_latest(self) -> None:
        self.select_folder(self.latest_paths, self.latest_input)

    def browse_output(self) -> None:
        file, _ = QFileDialog.getSaveFileName(self, "Save Output File", "output.xlsx", "Excel Files (*.xlsx)")
        if file:
            self.output_input.setText(file)

    def select_folder(self, path_list: List[str], input_widget: DropLineEdit) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if path and path not in path_list:
            path_list.append(path)
            input_widget.setText("; ".join(path_list))

    def on_base_dropped(self, paths: List[str]) -> None:
        self.base_paths.extend([p for p in paths if p not in self.base_paths])

    def on_latest_dropped(self, paths: List[str]) -> None:
        self.latest_paths.extend([p for p in paths if p not in self.latest_paths])

    def log(self, message: str) -> None:
        self.log_text.append(message)

    def update_progress(self, value: int) -> None:
        self.stop_progress_animation()
        self.progress_bar.setValue(value)
        if value >= 100:
            self.progress_bar.setComplete()

    def start_progress_animation(self) -> None:
        self.animation_value = 0
        self.animation_timer.timeout.connect(self.animate_progress)
        self.animation_timer.start(50)

    def animate_progress(self) -> None:
        self.animation_value = (self.animation_value + 2) % 100
        self.progress_bar.setValue(self.animation_value)

    def stop_progress_animation(self) -> None:
        self.animation_timer.stop()
        try:
            self.animation_timer.timeout.disconnect(self.animate_progress)
        except TypeError:
            pass

    def run_process(self) -> None:
        if not self.base_paths or not self.latest_paths:
            QMessageBox.critical(self, "Error", "Please specify both folders.")
            return

        self.run_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.start_progress_animation()
        self.log("Starting process...")

        self.thread = QThread()
        self.worker = Worker(
            self.base_paths[0],
            self.latest_paths[0],
            self.output_input.text()
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.error_signal.connect(lambda msg: QMessageBox.critical(self, "Error", msg))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(lambda: self.run_button.setEnabled(True))

        self.thread.start()


