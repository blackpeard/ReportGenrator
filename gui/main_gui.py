"""
gui_app.py
==========
PyQt6 distributable GUI for the Report Engine.

Features
--------
- Modern dark-themed interface
- Live log output & progress bar during generation
- Save / load client profiles (JSON)
- File/folder pickers for Excel & POC
- Runs report generation in a background thread (non-blocking UI)

Run:
    python gui_app.py

Build exe (Windows):
    pyinstaller --onefile --windowed --name ReportGenerator gui_app.py
"""

from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import threading
from pathlib import Path

from PyQt6.QtCore import (
    QObject, QThread, Qt, QTimer, pyqtSignal, pyqtSlot,
)
from PyQt6.QtGui import QColor, QFont, QIcon, QPalette, QTextCharFormat
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
    QSizePolicy, QSpacerItem, QSplitter, QStatusBar,
    QVBoxLayout, QWidget, QScrollArea, QFrame,
)

# Add local paths
sys.path.insert(0, os.path.dirname(__file__))

from report_engine import ReportConfig, ReportEngine, ReportResult

# ── Constants ──────────────────────────────────────────────────────────────

APP_NAME    = "Security Report Generator"
APP_VERSION = "2.0.0"
PROFILE_DIR = Path.home() / ".report_generator" / "profiles"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)

DARK_STYLE = """
QMainWindow, QWidget {
    background-color: #1a1d23;
    color: #e2e8f0;
    font-family: 'Segoe UI', 'SF Pro Display', sans-serif;
    font-size: 13px;
}

QGroupBox {
    border: 1px solid #2d3748;
    border-radius: 8px;
    margin-top: 12px;
    padding: 12px 8px 8px 8px;
    background-color: #1e2330;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #63b3ed;
    font-weight: 600;
    font-size: 12px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

QLineEdit, QComboBox {
    background-color: #2d3748;
    border: 1px solid #4a5568;
    border-radius: 6px;
    padding: 7px 10px;
    color: #e2e8f0;
    selection-background-color: #3182ce;
}
QLineEdit:focus, QComboBox:focus {
    border-color: #63b3ed;
    background-color: #2d3a4f;
}
QLineEdit:disabled {
    color: #718096;
    background-color: #252d3d;
}

QComboBox::drop-down {
    border: none;
    padding-right: 8px;
}
QComboBox::down-arrow {
    width: 10px;
    height: 10px;
}
QComboBox QAbstractItemView {
    background-color: #2d3748;
    border: 1px solid #4a5568;
    selection-background-color: #3182ce;
    outline: none;
}

QPushButton {
    background-color: #2d3748;
    border: 1px solid #4a5568;
    border-radius: 6px;
    padding: 7px 16px;
    color: #e2e8f0;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #3a4a5c;
    border-color: #63b3ed;
}
QPushButton:pressed {
    background-color: #2a3a4a;
}
QPushButton:disabled {
    color: #4a5568;
    background-color: #1e2330;
    border-color: #2d3748;
}

QPushButton#btn_generate {
    background-color: #2b6cb0;
    border-color: #3182ce;
    color: #ffffff;
    font-size: 14px;
    font-weight: 700;
    padding: 12px 32px;
    border-radius: 8px;
}
QPushButton#btn_generate:hover {
    background-color: #3182ce;
    border-color: #63b3ed;
}
QPushButton#btn_generate:disabled {
    background-color: #1e3a5f;
    border-color: #2c5282;
    color: #718096;
}

QPushButton#btn_save_profile {
    background-color: #276749;
    border-color: #38a169;
    color: #ffffff;
}
QPushButton#btn_save_profile:hover {
    background-color: #38a169;
}

QPushButton#btn_load_profile {
    background-color: #553c9a;
    border-color: #6b46c1;
    color: #ffffff;
}
QPushButton#btn_load_profile:hover {
    background-color: #6b46c1;
}

QProgressBar {
    background-color: #2d3748;
    border: 1px solid #4a5568;
    border-radius: 6px;
    text-align: center;
    color: #e2e8f0;
    font-weight: 600;
    height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #2b6cb0, stop:1 #38a169);
    border-radius: 5px;
}

QPlainTextEdit {
    background-color: #0d1117;
    border: 1px solid #2d3748;
    border-radius: 6px;
    color: #a8d8a8;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    padding: 8px;
}

QScrollBar:vertical {
    background: #1a1d23;
    width: 10px;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background: #4a5568;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QLabel#lbl_title {
    color: #63b3ed;
    font-size: 20px;
    font-weight: 700;
}
QLabel#lbl_sub {
    color: #718096;
    font-size: 11px;
}

QStatusBar {
    background-color: #141720;
    color: #718096;
    border-top: 1px solid #2d3748;
}

QFrame#divider {
    background-color: #2d3748;
    max-height: 1px;
}

QSplitter::handle {
    background-color: #2d3748;
    width: 2px;
}
"""


# ── Worker thread ──────────────────────────────────────────────────────────

class GeneratorWorker(QObject):
    """Runs ReportEngine in a background thread."""

    log     = pyqtSignal(str)
    progress= pyqtSignal(int)
    finished= pyqtSignal(object)   # ReportResult

    def __init__(self, config: ReportConfig):
        super().__init__()
        self._config = config

    @pyqtSlot()
    def run(self):
        engine = ReportEngine(
            self._config,
            progress_callback=self.progress.emit,
            log_callback=self.log.emit,
        )
        result = engine.run()
        self.finished.emit(result)


# ── File picker row ────────────────────────────────────────────────────────

class FilePickerRow(QWidget):
    def __init__(self, placeholder: str, folder: bool = False, parent=None):
        super().__init__(parent)
        self._folder = folder
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.line = QLineEdit()
        self.line.setPlaceholderText(placeholder)
        layout.addWidget(self.line)

        btn = QPushButton("Browse")
        btn.setFixedWidth(72)
        btn.clicked.connect(self._browse)
        layout.addWidget(btn)

    def _browse(self):
        if self._folder:
            path = QFileDialog.getExistingDirectory(self, "Select Folder")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select File", "", "Excel Files (*.xlsx *.xls)"
            )
        if path:
            self.line.setText(path)

    def text(self) -> str:
        return self.line.text().strip()

    def setText(self, v: str):
        self.line.setText(v)


# ── Main window ────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(960, 700)
        self.resize(1100, 760)

        self._thread: QThread | None = None
        self._worker: GeneratorWorker | None = None

        self._build_ui()
        self._apply_style()
        self._status("Ready")

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 16, 20, 12)
        root.setSpacing(12)

        # Header
        root.addWidget(self._header())

        divider = QFrame()
        divider.setObjectName("divider")
        divider.setFrameShape(QFrame.Shape.HLine)
        root.addWidget(divider)

        # Splitter: left form | right log
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._left_panel())
        splitter.addWidget(self._right_panel())
        splitter.setSizes([560, 480])
        root.addWidget(splitter, stretch=1)

        # Bottom bar
        root.addWidget(self._bottom_bar())

        # Status bar
        self.setStatusBar(QStatusBar())

    def _header(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)

        title = QLabel(APP_NAME)
        title.setObjectName("lbl_title")

        sub = QLabel(f"Security Audit Report Generator  ·  v{APP_VERSION}")
        sub.setObjectName("lbl_sub")

        h.addWidget(title)
        h.addSpacing(12)
        h.addWidget(sub)
        h.addStretch()

        # Profile buttons
        btn_load = QPushButton("⬆  Load Profile")
        btn_load.setObjectName("btn_load_profile")
        btn_load.clicked.connect(self._load_profile)

        btn_save = QPushButton("💾  Save Profile")
        btn_save.setObjectName("btn_save_profile")
        btn_save.clicked.connect(self._save_profile)

        h.addWidget(btn_load)
        h.addSpacing(6)
        h.addWidget(btn_save)
        return w

    def _left_panel(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        vbox = QVBoxLayout(container)
        vbox.setSpacing(12)
        vbox.setContentsMargins(4, 4, 8, 4)

        vbox.addWidget(self._group_report_settings())
        vbox.addWidget(self._group_client_info())
        vbox.addWidget(self._group_files())
        vbox.addStretch()

        scroll.setWidget(container)
        return scroll

    def _group_report_settings(self) -> QGroupBox:
        grp = QGroupBox("Report Settings")
        form = QFormLayout(grp)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.cmb_type = QComboBox()
        self.cmb_type.addItems(["Web", "Api", "Mobile"])
        form.addRow("Report Type:", self.cmb_type)

        self.cmb_env = QComboBox()
        self.cmb_env.addItems(["Production", "Uat"])
        form.addRow("Environment:", self.cmb_env)

        return grp

    def _group_client_info(self) -> QGroupBox:
        grp = QGroupBox("Client Information")
        form = QFormLayout(grp)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.in_client     = QLineEdit(); self.in_client.setPlaceholderText("e.g. Acme Bank Pvt. Ltd.")
        self.in_app        = QLineEdit(); self.in_app.setPlaceholderText("e.g. Internet Banking Portal")
        self.in_app_type   = QLineEdit(); self.in_app_type.setPlaceholderText("e.g. External Web Application")
        self.in_period     = QLineEdit(); self.in_period.setPlaceholderText("e.g. 01-01-2026 - 15-01-2026")
        self.in_url        = QLineEdit(); self.in_url.setPlaceholderText("https://example.com")
        self.in_method     = QLineEdit(); self.in_method.setPlaceholderText("Grey Box / Black Box")

        form.addRow("Client Name:",    self.in_client)
        form.addRow("App Name:",       self.in_app)
        form.addRow("App Type:",       self.in_app_type)
        form.addRow("Audit Period:",   self.in_period)
        form.addRow("Target URL:",     self.in_url)
        form.addRow("Test Method:",    self.in_method)

        return grp

    def _group_files(self) -> QGroupBox:
        grp = QGroupBox("Files & Output")
        form = QFormLayout(grp)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.pick_excel  = FilePickerRow("Select findings Excel file (.xlsx)")
        self.pick_poc    = FilePickerRow("Optional: POC screenshots folder", folder=True)
        self.pick_output = QLineEdit(); self.pick_output.setPlaceholderText("Leave blank for auto-named output")

        btn_out = QPushButton("Browse")
        btn_out.setFixedWidth(72)
        btn_out.clicked.connect(self._browse_output)

        out_row = QWidget()
        out_h = QHBoxLayout(out_row)
        out_h.setContentsMargins(0, 0, 0, 0)
        out_h.setSpacing(6)
        out_h.addWidget(self.pick_output)
        out_h.addWidget(btn_out)

        form.addRow("Excel File:", self.pick_excel)
        form.addRow("POC Folder:", self.pick_poc)
        form.addRow("Output File:", out_row)

        return grp

    def _right_panel(self) -> QWidget:
        w = QWidget()
        vbox = QVBoxLayout(w)
        vbox.setContentsMargins(8, 0, 0, 0)
        vbox.setSpacing(8)

        lbl = QLabel("Live Output")
        lbl.setObjectName("lbl_sub")
        lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #718096; text-transform: uppercase; letter-spacing: 0.5px;")
        vbox.addWidget(lbl)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("Logs will appear here once generation starts…")
        vbox.addWidget(self.log_view, stretch=1)

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        vbox.addWidget(self.progress_bar)

        btn_clear = QPushButton("Clear Log")
        btn_clear.setFixedHeight(28)
        btn_clear.clicked.connect(self.log_view.clear)
        vbox.addWidget(btn_clear)

        return w

    def _bottom_bar(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)

        h.addStretch()

        self.btn_generate = QPushButton("⚡  Generate Report")
        self.btn_generate.setObjectName("btn_generate")
        self.btn_generate.setFixedHeight(46)
        self.btn_generate.clicked.connect(self._on_generate)
        h.addWidget(self.btn_generate)

        return w

    def _apply_style(self):
        self.setStyleSheet(DARK_STYLE)

    # ── Actions ────────────────────────────────────────────────────────

    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Report As", "", "Word Documents (*.docx)"
        )
        if path:
            self.pick_output.setText(path)

    def _on_generate(self):
        config = self._build_config()
        if not config:
            return

        self.btn_generate.setEnabled(False)
        self.log_view.clear()
        self.progress_bar.setValue(0)
        self._status("Generating report…")

        self._worker = GeneratorWorker(config)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._on_log)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()

    def _build_config(self) -> ReportConfig | None:
        errors = []
        excel = self.pick_excel.text()
        if not excel:
            errors.append("Excel file is required.")
        elif not os.path.exists(excel):
            errors.append(f"Excel file not found:\n{excel}")

        if errors:
            QMessageBox.warning(self, "Validation Error", "\n".join(errors))
            return None

        return ReportConfig(
            client_name  = self.in_client.text().strip(),
            app_name     = self.in_app.text().strip(),
            app_type     = self.in_app_type.text().strip(),
            audit_period = self.in_period.text().strip(),
            url          = self.in_url.text().strip(),
            method       = self.in_method.text().strip(),
            report_type  = self.cmb_type.currentText(),
            environment  = self.cmb_env.currentText(),
            excel_file   = excel,
            poc_folder   = self.pick_poc.text(),
            output_file  = self.pick_output.text().strip(),
        )

    @pyqtSlot(str)
    def _on_log(self, msg: str):
        self.log_view.appendPlainText(msg)
        self._status(msg.replace("✅", "").replace("❌", "").strip())

    @pyqtSlot(int)
    def _on_progress(self, val: int):
        self.progress_bar.setValue(val)

    @pyqtSlot(object)
    def _on_finished(self, result: ReportResult):
        self.btn_generate.setEnabled(True)
        if result.success:
            self._status(f"Done  ✓  →  {result.output_path}")
            QMessageBox.information(
                self,
                "Report Generated",
                f"✅ Report created successfully!\n\n"
                f"📄 Output:  {result.output_path}\n"
                f"📊 Observations:  {result.observations_count}\n\n"
                "⚠️  Please review before sharing.",
            )
        else:
            self._status(f"Error: {result.error}")
            QMessageBox.critical(
                self,
                "Generation Failed",
                f"❌ Report generation failed:\n\n{result.error}",
            )

    def _status(self, msg: str):
        self.statusBar().showMessage(msg, 8000)

    # ── Profile save / load ────────────────────────────────────────────

    def _profile_data(self) -> dict:
        return {
            "client_name":  self.in_client.text(),
            "app_name":     self.in_app.text(),
            "app_type":     self.in_app_type.text(),
            "audit_period": self.in_period.text(),
            "url":          self.in_url.text(),
            "method":       self.in_method.text(),
            "report_type":  self.cmb_type.currentText(),
            "environment":  self.cmb_env.currentText(),
            "excel_file":   self.pick_excel.text(),
            "poc_folder":   self.pick_poc.text(),
            "output_file":  self.pick_output.text(),
        }

    def _apply_profile_data(self, data: dict):
        self.in_client.setText(data.get("client_name", ""))
        self.in_app.setText(data.get("app_name", ""))
        self.in_app_type.setText(data.get("app_type", ""))
        self.in_period.setText(data.get("audit_period", ""))
        self.in_url.setText(data.get("url", ""))
        self.in_method.setText(data.get("method", ""))
        idx = self.cmb_type.findText(data.get("report_type", "Web"))
        if idx >= 0: self.cmb_type.setCurrentIndex(idx)
        idx = self.cmb_env.findText(data.get("environment", "Production"))
        if idx >= 0: self.cmb_env.setCurrentIndex(idx)
        self.pick_excel.setText(data.get("excel_file", ""))
        self.pick_poc.setText(data.get("poc_folder", ""))
        self.pick_output.setText(data.get("output_file", ""))

    def _save_profile(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Profile", str(PROFILE_DIR), "JSON Profiles (*.json)"
        )
        if not path:
            return
        if not path.endswith(".json"):
            path += ".json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self._profile_data(), f, indent=2)
        self._status(f"Profile saved: {path}")
        QMessageBox.information(self, "Profile Saved", f"✅ Profile saved:\n{path}")

    def _load_profile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Profile", str(PROFILE_DIR), "JSON Profiles (*.json)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._apply_profile_data(data)
            self._status(f"Profile loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Load Failed", f"Could not load profile:\n{e}")


# ── Entry point ────────────────────────────────────────────────────────────

def launch_gui():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)

    # High-DPI
    try:
        app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)
    except AttributeError:
        pass

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    launch_gui()