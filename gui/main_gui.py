"""
gui_app.py
==========
PyQt6 distributable GUI for the Report Engine.
3-page wizard:
  Page 1 — General Info    (Prepared By, Reviewed By, Doc Version, Client History, Limitation)
  Page 2 — Report Details  (Client info, dates, Excel OR Manual observation mode)
  Page 3 — Generate        (Summary + live logs + progress bar)

Run:
    python gui/gui_app.py
    python gui_app.py      (from gui/ folder)

Build exe:
    pyinstaller --onefile --windowed --name "ReportGenerator"
      --add-data "templates;templates" --add-data "src;src" --add-data "data;data"
      --paths "." gui/gui_app.py
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

# ── Project root on sys.path so report_engine + db_manager are importable ────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PyQt6.QtCore  import QDate, QObject, QThread, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui   import QColor
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QListWidget, QListWidgetItem, QComboBox, QDateEdit,
    QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QCheckBox, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QScrollArea, QSizePolicy, QSplitter,
    QStackedWidget, QStatusBar, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget, QFrame,
)

from report_engine import ReportConfig, ReportEngine, ReportResult
from db_manager    import DBManager

# ── Constants ─────────────────────────────────────────────────────────────────
APP_NAME    = "Advance Report_Generator Machine"
APP_VERSION = "4.0.0"
PROFILE_DIR = Path.home() / ".report_generator" / "profiles"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)

# ── Dark stylesheet ───────────────────────────────────────────────────────────
THEMES = {
    "Dark": """
QMainWindow,QWidget{background:#1a1d23;color:#e2e8f0;font-family:'Segoe UI',sans-serif;font-size:13px}
QGroupBox{border:1px solid #2d3748;border-radius:8px;margin-top:12px;padding:12px 8px 8px;background:#1e2330}
QGroupBox::title{subcontrol-origin:margin;left:12px;padding:0 6px;color:#63b3ed;font-weight:600;font-size:11px;text-transform:uppercase}
QLineEdit,QComboBox,QTextEdit,QPlainTextEdit,QDateEdit{background:#2d3748;border:1px solid #4a5568;border-radius:6px;padding:6px 10px;color:#e2e8f0}
QLineEdit:focus,QComboBox:focus,QTextEdit:focus,QDateEdit:focus{border-color:#63b3ed;background:#2d3a4f}
QComboBox::drop-down,QDateEdit::drop-down{border:none;padding-right:8px}
QComboBox QAbstractItemView{background:#2d3748;border:1px solid #4a5568;selection-background-color:#3182ce;outline:none}
QPushButton{background:#2d3748;border:1px solid #4a5568;border-radius:6px;padding:7px 16px;color:#e2e8f0;font-weight:500}
QPushButton:hover{background:#3a4a5c;border-color:#63b3ed}
QPushButton:checked{background:#2b4c7e;border-color:#63b3ed;color:#90cdf4}
QPushButton:disabled{color:#4a5568;background:#1e2330;border-color:#2d3748}
QPushButton#btn_next{background:#2b6cb0;border-color:#3182ce;color:#fff;font-size:13px;font-weight:700;padding:10px 28px;border-radius:8px}
QPushButton#btn_next:hover{background:#3182ce}
QPushButton#btn_generate{background:#276749;border-color:#38a169;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px}
QPushButton#btn_generate:hover{background:#38a169}
QPushButton#btn_generate:disabled{background:#1a3a2a;color:#4a5568}
QPushButton#btn_save_profile{background:#276749;border-color:#38a169;color:#fff}
QPushButton#btn_load_profile{background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_add_obs{background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_lib{background:#1a365d;border-color:#2b6cb0;color:#63b3ed}
QPushButton#btn_lib:hover{background:#2b6cb0;color:#fff}
QPushButton#btn_del{background:#742a2a;border-color:#c53030;color:#fff;padding:3px 10px}
QPushButton#btn_del:hover{background:#c53030}
QProgressBar{background:#2d3748;border:1px solid #4a5568;border-radius:6px;text-align:center;color:#e2e8f0;font-weight:600;height:22px}
QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #2b6cb0,stop:1 #38a169);border-radius:5px}
QTableWidget{background:#1e2330;border:1px solid #2d3748;gridline-color:#2d3748;border-radius:6px}
QTableWidget::item{padding:5px}
QTableWidget::item:selected{background:#2b6cb0;color:#fff}
QHeaderView::section{background:#2d3748;color:#a0aec0;border:none;padding:6px;font-size:12px;font-weight:600}
QListWidget{background:#1e2330;border:1px solid #2d3748;border-radius:6px}
QListWidget::item{padding:6px 10px}
QListWidget::item:selected{background:#2b6cb0;color:#fff}
QScrollBar:vertical{background:#1a1d23;width:8px;border-radius:4px}
QScrollBar::handle:vertical{background:#4a5568;border-radius:4px;min-height:24px}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0}
QStatusBar{background:#141720;color:#718096;border-top:1px solid #2d3748}
QSplitter::handle{background:#2d3748;width:2px}
QFrame#divider{background:#2d3748;max-height:1px}
QLabel#lbl_title{color:#63b3ed;font-size:20px;font-weight:700}
QLabel#lbl_sub{color:#718096;font-size:11px}
QLabel#step_label{color:#63b3ed;font-size:11px;font-weight:600;text-transform:uppercase}
QLabel#page_title{color:#e2e8f0;font-size:16px;font-weight:700}
QLabel#page_sub{color:#718096;font-size:12px}
""",
    "Light": """
QMainWindow,QWidget{background:#f7f8fa;color:#1a202c;font-family:'Segoe UI',sans-serif;font-size:13px}
QGroupBox{border:1px solid #cbd5e0;border-radius:8px;margin-top:12px;padding:12px 8px 8px;background:#ffffff}
QGroupBox::title{subcontrol-origin:margin;left:12px;padding:0 6px;color:#2b6cb0;font-weight:600;font-size:11px;text-transform:uppercase}
QLineEdit,QComboBox,QTextEdit,QPlainTextEdit,QDateEdit{background:#ffffff;border:1px solid #cbd5e0;border-radius:6px;padding:6px 10px;color:#1a202c}
QLineEdit:focus,QComboBox:focus,QTextEdit:focus,QDateEdit:focus{border-color:#3182ce;background:#ebf8ff}
QComboBox::drop-down,QDateEdit::drop-down{border:none;padding-right:8px}
QComboBox QAbstractItemView{background:#ffffff;border:1px solid #cbd5e0;selection-background-color:#bee3f8;color:#1a202c}
QPushButton{background:#edf2f7;border:1px solid #cbd5e0;border-radius:6px;padding:7px 16px;color:#1a202c;font-weight:500}
QPushButton:hover{background:#e2e8f0;border-color:#3182ce}
QPushButton:checked{background:#bee3f8;border-color:#3182ce;color:#1a365d}
QPushButton:disabled{color:#a0aec0;background:#f7fafc}
QPushButton#btn_next{background:#2b6cb0;border-color:#2c5282;color:#fff;font-size:13px;font-weight:700;padding:10px 28px;border-radius:8px}
QPushButton#btn_next:hover{background:#3182ce}
QPushButton#btn_generate{background:#276749;border-color:#22543d;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px}
QPushButton#btn_generate:hover{background:#38a169}
QPushButton#btn_generate:disabled{background:#c6f6d5;color:#a0aec0}
QPushButton#btn_save_profile{background:#276749;border-color:#22543d;color:#fff}
QPushButton#btn_load_profile{background:#553c9a;border-color:#44337a;color:#fff}
QPushButton#btn_add_obs{background:#553c9a;border-color:#44337a;color:#fff}
QPushButton#btn_lib{background:#ebf8ff;border-color:#90cdf4;color:#2b6cb0}
QPushButton#btn_del{background:#fff5f5;border-color:#fc8181;color:#c53030;padding:3px 10px}
QProgressBar{background:#e2e8f0;border:1px solid #cbd5e0;border-radius:6px;text-align:center;color:#1a202c;font-weight:600;height:22px}
QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #3182ce,stop:1 #38a169);border-radius:5px}
QTableWidget{background:#ffffff;border:1px solid #e2e8f0;gridline-color:#e2e8f0;border-radius:6px}
QTableWidget::item{padding:5px;color:#1a202c}
QTableWidget::item:selected{background:#bee3f8;color:#1a365d}
QHeaderView::section{background:#edf2f7;color:#4a5568;border:none;padding:6px;font-size:12px;font-weight:600}
QListWidget{background:#ffffff;border:1px solid #e2e8f0;border-radius:6px}
QListWidget::item{padding:6px 10px;color:#1a202c}
QListWidget::item:selected{background:#bee3f8;color:#1a365d}
QScrollBar:vertical{background:#f7f8fa;width:8px;border-radius:4px}
QScrollBar::handle:vertical{background:#cbd5e0;border-radius:4px;min-height:24px}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0}
QStatusBar{background:#edf2f7;color:#718096;border-top:1px solid #cbd5e0}
QSplitter::handle{background:#e2e8f0;width:2px}
QFrame#divider{background:#e2e8f0;max-height:1px}
QLabel#lbl_title{color:#2b6cb0;font-size:20px;font-weight:700}
QLabel#lbl_sub{color:#718096;font-size:11px}
QLabel#step_label{color:#2b6cb0;font-size:11px;font-weight:600;text-transform:uppercase}
QLabel#page_title{color:#1a202c;font-size:16px;font-weight:700}
QLabel#page_sub{color:#718096;font-size:12px}
""",
    "Midnight Blue": """
QMainWindow,QWidget{background:#0d1b2a;color:#cdd9e5;font-family:'Segoe UI',sans-serif;font-size:13px}
QGroupBox{border:1px solid #1e3a5f;border-radius:8px;margin-top:12px;padding:12px 8px 8px;background:#112240}
QGroupBox::title{subcontrol-origin:margin;left:12px;padding:0 6px;color:#64b5f6;font-weight:600;font-size:11px;text-transform:uppercase}
QLineEdit,QComboBox,QTextEdit,QPlainTextEdit,QDateEdit{background:#1e3a5f;border:1px solid #2d5a8e;border-radius:6px;padding:6px 10px;color:#cdd9e5}
QLineEdit:focus,QComboBox:focus,QTextEdit:focus,QDateEdit:focus{border-color:#64b5f6;background:#1a3a6e}
QComboBox::drop-down,QDateEdit::drop-down{border:none;padding-right:8px}
QComboBox QAbstractItemView{background:#1e3a5f;border:1px solid #2d5a8e;selection-background-color:#1565c0;color:#cdd9e5}
QPushButton{background:#1e3a5f;border:1px solid #2d5a8e;border-radius:6px;padding:7px 16px;color:#cdd9e5;font-weight:500}
QPushButton:hover{background:#2d5a8e;border-color:#64b5f6}
QPushButton:checked{background:#1565c0;border-color:#64b5f6;color:#e3f2fd}
QPushButton:disabled{color:#4a6a8a;background:#0d1b2a}
QPushButton#btn_next{background:#1565c0;border-color:#1976d2;color:#fff;font-size:13px;font-weight:700;padding:10px 28px;border-radius:8px}
QPushButton#btn_generate{background:#1b5e20;border-color:#2e7d32;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px}
QPushButton#btn_generate:hover{background:#2e7d32}
QPushButton#btn_save_profile{background:#1b5e20;border-color:#2e7d32;color:#fff}
QPushButton#btn_load_profile{background:#4a148c;border-color:#6a1b9a;color:#fff}
QPushButton#btn_add_obs{background:#4a148c;border-color:#6a1b9a;color:#fff}
QPushButton#btn_lib{background:#0d2137;border-color:#1565c0;color:#64b5f6}
QPushButton#btn_del{background:#7f0000;border-color:#b71c1c;color:#fff;padding:3px 10px}
QProgressBar{background:#1e3a5f;border:1px solid #2d5a8e;border-radius:6px;text-align:center;color:#cdd9e5;font-weight:600;height:22px}
QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1565c0,stop:1 #2e7d32);border-radius:5px}
QTableWidget{background:#112240;border:1px solid #1e3a5f;gridline-color:#1e3a5f;border-radius:6px}
QTableWidget::item{padding:5px;color:#cdd9e5}
QTableWidget::item:selected{background:#1565c0;color:#fff}
QHeaderView::section{background:#1e3a5f;color:#90b4d4;border:none;padding:6px;font-size:12px;font-weight:600}
QListWidget{background:#112240;border:1px solid #1e3a5f;border-radius:6px}
QListWidget::item{padding:6px 10px;color:#cdd9e5}
QListWidget::item:selected{background:#1565c0;color:#fff}
QScrollBar:vertical{background:#0d1b2a;width:8px;border-radius:4px}
QScrollBar::handle:vertical{background:#2d5a8e;border-radius:4px;min-height:24px}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0}
QStatusBar{background:#0a1628;color:#64b5f6;border-top:1px solid #1e3a5f}
QSplitter::handle{background:#1e3a5f;width:2px}
QFrame#divider{background:#1e3a5f;max-height:1px}
QLabel#lbl_title{color:#64b5f6;font-size:20px;font-weight:700}
QLabel#lbl_sub{color:#4a6a8a;font-size:11px}
QLabel#step_label{color:#64b5f6;font-size:11px;font-weight:600;text-transform:uppercase}
QLabel#page_title{color:#cdd9e5;font-size:16px;font-weight:700}
QLabel#page_sub{color:#4a6a8a;font-size:12px}
""",
}
DARK = THEMES["Dark"]   # default
DARK = """
QMainWindow,QWidget{
    background:#1a1d23;color:#e2e8f0;
    font-family:'Segoe UI','SF Pro Display',sans-serif;font-size:13px}
QGroupBox{
    border:1px solid #2d3748;border-radius:8px;
    margin-top:12px;padding:12px 8px 8px;background:#1e2330}
QGroupBox::title{
    subcontrol-origin:margin;left:12px;padding:0 6px;
    color:#63b3ed;font-weight:600;font-size:11px;
    letter-spacing:0.5px;text-transform:uppercase}
QLineEdit,QComboBox,QTextEdit,QPlainTextEdit,QDateEdit{
    background:#2d3748;border:1px solid #4a5568;
    border-radius:6px;padding:6px 10px;color:#e2e8f0}
QLineEdit:focus,QComboBox:focus,QTextEdit:focus,QDateEdit:focus{
    border-color:#63b3ed;background:#2d3a4f}
QLineEdit:disabled{color:#718096;background:#252d3d}
QComboBox::drop-down{border:none;padding-right:8px}
QComboBox QAbstractItemView{
    background:#2d3748;border:1px solid #4a5568;
    selection-background-color:#3182ce;outline:none}
QDateEdit::drop-down{border:none;padding-right:8px}
QCalendarWidget{background:#1e2330;color:#e2e8f0}
QPushButton{
    background:#2d3748;border:1px solid #4a5568;
    border-radius:6px;padding:7px 16px;color:#e2e8f0;font-weight:500}
QPushButton:hover{background:#3a4a5c;border-color:#63b3ed}
QPushButton:pressed{background:#2a3a4a}
QPushButton:disabled{color:#4a5568;background:#1e2330;border-color:#2d3748}
QPushButton:checked{background:#2b4c7e;border-color:#63b3ed;color:#90cdf4}
QPushButton#btn_next{
    background:#2b6cb0;border-color:#3182ce;color:#fff;
    font-size:13px;font-weight:700;padding:10px 28px;border-radius:8px}
QPushButton#btn_next:hover{background:#3182ce}
QPushButton#btn_back{padding:10px 20px;border-radius:8px}
QPushButton#btn_generate{
    background:#276749;border-color:#38a169;color:#fff;
    font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px}
QPushButton#btn_generate:hover{background:#38a169}
QPushButton#btn_generate:disabled{
    background:#1a3a2a;border-color:#276749;color:#4a5568}
QPushButton#btn_save_profile{
    background:#276749;border-color:#38a169;color:#fff}
QPushButton#btn_save_profile:hover{background:#38a169}
QPushButton#btn_load_profile{
    background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_load_profile:hover{background:#6b46c1}
QPushButton#btn_add_obs{
    background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_add_obs:hover{background:#6b46c1}
QPushButton#btn_lib{
    background:#1a365d;border-color:#2b6cb0;color:#63b3ed}
QPushButton#btn_lib:hover{background:#2b6cb0;color:#fff}
QPushButton#btn_del{
    background:#742a2a;border-color:#c53030;color:#fff;padding:3px 10px}
QPushButton#btn_del:hover{background:#c53030}
QProgressBar{
    background:#2d3748;border:1px solid #4a5568;border-radius:6px;
    text-align:center;color:#e2e8f0;font-weight:600;height:22px}
QProgressBar::chunk{
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #2b6cb0,stop:1 #38a169);border-radius:5px}
QTableWidget{
    background:#1e2330;border:1px solid #2d3748;
    gridline-color:#2d3748;border-radius:6px}
QTableWidget::item{padding:5px;border:none}
QTableWidget::item:selected{background:#2b6cb0;color:#fff}
QHeaderView::section{
    background:#2d3748;color:#a0aec0;border:none;
    padding:6px;font-size:12px;font-weight:600}
QListWidget{background:#1e2330;border:1px solid #2d3748;border-radius:6px}
QListWidget::item{padding:6px 10px}
QListWidget::item:hover{background:#2d3748}
QListWidget::item:selected{background:#2b6cb0;color:#fff}
QScrollBar:vertical{background:#1a1d23;width:8px;border-radius:4px}
QScrollBar::handle:vertical{background:#4a5568;border-radius:4px;min-height:24px}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0}
QStatusBar{background:#141720;color:#718096;border-top:1px solid #2d3748}
QSplitter::handle{background:#2d3748;width:2px}
QFrame#divider{background:#2d3748;max-height:1px}
QLabel#lbl_title{color:#63b3ed;font-size:20px;font-weight:700}
QLabel#lbl_sub{color:#718096;font-size:11px}
QLabel#step_label{color:#63b3ed;font-size:11px;font-weight:600;
    letter-spacing:0.5px;text-transform:uppercase}
QLabel#page_title{color:#e2e8f0;font-size:16px;font-weight:700}
QLabel#page_sub{color:#718096;font-size:12px}
"""

SEV_COLORS = {
    "Critical": "#EE0000", "High": "#EE0000",
    "Medium": "#FFC000", "Low": "#00B050", "Info": "#0070C0"
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _divider() -> QFrame:
    f = QFrame(); f.setObjectName("divider")
    f.setFrameShape(QFrame.Shape.HLine); return f


class FilePicker(QWidget):
    def __init__(self, placeholder="", folder=False, save=False, parent=None):
        super().__init__(parent)
        self._folder = folder; self._save = save
        h = QHBoxLayout(self); h.setContentsMargins(0,0,0,0); h.setSpacing(6)
        self.line = QLineEdit(); self.line.setPlaceholderText(placeholder)
        btn = QPushButton("Browse"); btn.setFixedWidth(70)
        btn.clicked.connect(self._browse)
        h.addWidget(self.line); h.addWidget(btn)

    def _browse(self):
        if self._folder:
            p = QFileDialog.getExistingDirectory(self, "Select Folder")
        elif self._save:
            p, _ = QFileDialog.getSaveFileName(self,"Save As","","Word Documents (*.docx)")
        else:
            p, _ = QFileDialog.getOpenFileName(self,"Select File","","Excel Files (*.xlsx *.xls)")
        if p: self.line.setText(p)

    def text(self): return self.line.text().strip()
    def setText(self, v): self.line.setText(v)


# ── Worker ────────────────────────────────────────────────────────────────────

class GeneratorWorker(QObject):
    log      = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(object)

    def __init__(self, config: ReportConfig):
        super().__init__(); self._config = config

    @pyqtSlot()
    def run(self):
        engine = ReportEngine(self._config,
                              progress_callback=self.progress.emit,
                              log_callback=self.log.emit)
        self.finished.emit(engine.run())


# ── Observation Library Dialog ────────────────────────────────────────────────

class ObsLibraryDialog(QDialog):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent)
        self.db = db; self.selected = None
        self.setWindowTitle("Observation Library")
        self.setMinimumSize(760, 500); self.setStyleSheet(DARK)
        self._build(); self._load("")

    def _build(self):
        v = QVBoxLayout(self); v.setSpacing(10); v.setContentsMargins(16,16,16,16)

        h = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search title, category, description…")
        self.search.textChanged.connect(self._load)
        self.cmb_cat = QComboBox(); self.cmb_cat.addItem("All Categories")
        for c in self.db.get_categories(): self.cmb_cat.addItem(c)
        self.cmb_cat.currentTextChanged.connect(lambda _: self._load(self.search.text()))
        h.addWidget(self.search, stretch=1); h.addWidget(self.cmb_cat)
        v.addLayout(h)

        self.lst = QListWidget()
        self.lst.itemDoubleClicked.connect(self._accept)
        self.lst.currentItemChanged.connect(self._on_select)
        v.addWidget(self.lst, stretch=1)

        grp = QGroupBox("Preview"); pv = QVBoxLayout(grp)
        self.preview = QPlainTextEdit(); self.preview.setReadOnly(True)
        self.preview.setFixedHeight(110)
        pv.addWidget(self.preview); v.addWidget(grp)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self._accept); btns.rejected.connect(self.reject)
        v.addWidget(btns)

    def _load(self, query: str):
        cat  = self.cmb_cat.currentText()
        rows = self.db.search_observations(query)
        if cat != "All Categories":
            rows = [r for r in rows if r["category"] == cat]
        self.lst.clear(); self._rows = rows
        for r in rows:
            item = QListWidgetItem(f"[{r['severity']:8s}]  {r['title']}  —  {r['category']}")
            item.setForeground(QColor(SEV_COLORS.get(r["severity"], "#718096")))
            item.setData(Qt.ItemDataRole.UserRole, r["id"])
            self.lst.addItem(item)

    def _on_select(self, cur, _):
        if not cur: return
        row = self.db.get_observation_by_id(cur.data(Qt.ItemDataRole.UserRole))
        if row:
            self.preview.setPlainText(
                f"Description:\n{row['description'][:250]}\n\n"
                f"Impact:\n{row['impact'][:200]}\n\nCVE: {row['cve']}"
            )

    def _accept(self):
        cur = self.lst.currentItem()
        if not cur:
            QMessageBox.warning(self, "Select", "Please select an observation."); return
        self.selected = self.db.get_observation_by_id(cur.data(Qt.ItemDataRole.UserRole))
        self.accept()


# ── Manage Employees Dialog ───────────────────────────────────────────────────

class EmployeeDialog(QDialog):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent)
        self.db = db; self.setWindowTitle("Manage Employees")
        self.setMinimumSize(620, 420); self.setStyleSheet(DARK)
        self._build(); self._load()

    def _build(self):
        v = QVBoxLayout(self); v.setSpacing(8); v.setContentsMargins(16,16,16,16)
        self.tbl = QTableWidget(0, 6)
        self.tbl.setHorizontalHeaderLabels(["Name","Designation","Email","Department","Qualifications/Certifications","CERT-In Listed"])
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        v.addWidget(self.tbl)

        grp = QGroupBox("Add New Employee"); f = QFormLayout(grp); f.setSpacing(8)
        self.e_name  = QLineEdit(); self.e_desig = QLineEdit()
        self.e_email = QLineEdit(); self.e_dept  = QLineEdit()
        f.addRow("Name:",        self.e_name)
        f.addRow("Designation:", self.e_desig)
        f.addRow("Email:",       self.e_email)
        self.e_qual    = QLineEdit()
        self.e_cert_in = QComboBox(); self.e_cert_in.addItems(["No", "Yes"])
        f.addRow("Qualifications:", self.e_qual)
        f.addRow("CERT-In Listed:", self.e_cert_in)
        btn = QPushButton("Add Employee"); btn.clicked.connect(self._add)
        f.addRow("", btn); v.addWidget(grp)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(self.accept); v.addWidget(btns)

    def _load(self):
        self.tbl.setRowCount(0)
        for emp in self.db.get_employees():
            r = self.tbl.rowCount(); self.tbl.insertRow(r)
            for c, k in enumerate(["name","designation","email","department", "qualifications","cert_in_listed"]):
                self.tbl.setItem(r, c, QTableWidgetItem(emp.get(k,"")))

    def _add(self):
        name  = self.e_name.text().strip()
        desig = self.e_desig.text().strip()
        if not name or not desig:
            QMessageBox.warning(self,"Required","Name and Designation are required."); return
        self.db.add_employee(name, desig, self.e_email.text().strip(), self.e_dept.text().strip(), self.e_qual.text().strip(), self.e_cert_in.currentText())
        self.e_name.clear(); self.e_desig.clear()
        self.e_email.clear(); self.e_dept.clear()
        self.e_qual.clear()
        self._load()
      


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — General Info
# ═══════════════════════════════════════════════════════════════════════════════

class Page1General(QWidget):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(10)

        lbl = QLabel("General Information"); lbl.setObjectName("page_title")
        sub = QLabel("Document metadata and client background — these appear on the cover page.")
        sub.setObjectName("page_sub"); sub.setWordWrap(True)
        v.addWidget(lbl); v.addWidget(sub); v.addWidget(_divider())

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        inner = QWidget(); iv = QVBoxLayout(inner)
        iv.setSpacing(12); iv.setContentsMargins(4,4,8,4)
        inner.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.MinimumExpanding)

        # ── Personnel ────────────────────────────────────────────────────────
        grp1 = QGroupBox("Personnel"); f1 = QFormLayout(grp1); f1.setSpacing(8)
        f1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        emp_names = [""] + [e["name"] for e in self.db.get_employees()]

        self.cmb_prepared = QComboBox(); self.cmb_prepared.addItems(emp_names)
        self.cmb_reviewed = QComboBox(); self.cmb_reviewed.addItems(emp_names)
        self.lbl_prep_desig = QLabel(""); self.lbl_prep_desig.setStyleSheet("color:#718096")
        self.lbl_rev_desig  = QLabel(""); self.lbl_rev_desig.setStyleSheet("color:#718096")

        self.cmb_prepared.currentTextChanged.connect(self._upd_prep)
        self.cmb_reviewed.currentTextChanged.connect(self._upd_rev)

        h_prep = QHBoxLayout(); h_prep.setSpacing(8)
        h_prep.addWidget(self.cmb_prepared, stretch=1)
        h_prep.addWidget(self.lbl_prep_desig)

        h_rev = QHBoxLayout(); h_rev.setSpacing(8)
        h_rev.addWidget(self.cmb_reviewed, stretch=1)
        h_rev.addWidget(self.lbl_rev_desig)

        btn_emp = QPushButton("Manage Employees →")
        btn_emp.clicked.connect(self._manage_emp)

        f1.addRow("Prepared By:", h_prep)  #New Changes 
        f1.addRow("Reviewed By:", h_rev)
        f1.addRow("", btn_emp)

        # Approved By
        self.cmb_approved = QComboBox(); self.cmb_approved.addItems(emp_names)
        self.lbl_appr_desig = QLabel(""); self.lbl_appr_desig.setStyleSheet("color:#718096")
        self.cmb_approved.currentTextChanged.connect(self._upd_appr)
        h_appr = QHBoxLayout(); h_appr.setSpacing(8)
        h_appr.addWidget(self.cmb_approved, stretch=1)
        h_appr.addWidget(self.lbl_appr_desig)
        f1.addRow("Approved By:", h_appr)

        # Released By
        self.cmb_released = QComboBox(); self.cmb_released.addItems(emp_names)
        self.lbl_rel_desig = QLabel(""); self.lbl_rel_desig.setStyleSheet("color:#718096")
        self.cmb_released.currentTextChanged.connect(self._upd_rel)
        h_rel = QHBoxLayout(); h_rel.setSpacing(8)
        h_rel.addWidget(self.cmb_released, stretch=1)
        h_rel.addWidget(self.lbl_rel_desig)
        f1.addRow("Released By:", h_rel)

        # Release Date
        self.in_release_date = QDateEdit()
        self.in_release_date.setCalendarPopup(True)
        self.in_release_date.setDate(QDate.currentDate())
        self.in_release_date.setDisplayFormat("dd-MM-yyyy")
        f1.addRow("Release Date:", self.in_release_date) #to Here

        # ── Auditing Team ────────────────────────────────────────────────────
        grp_team = QGroupBox("Auditing Team")
        vt = QVBoxLayout(grp_team); vt.setSpacing(8)

        lbl_hint = QLabel("Select employees who performed this audit:")
        lbl_hint.setStyleSheet("color:#718096;font-size:12px")
        vt.addWidget(lbl_hint)

        self.team_list = QTableWidget(0, 5)
        self.team_list.setHorizontalHeaderLabels([
            "Name","Designation","Email","Qualifications","CERT-In"])
        th = self.team_list.horizontalHeader()
        th.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        th.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.team_list.setFixedHeight(150)
        vt.addWidget(self.team_list)

        # Quick-add from employees dropdown
        add_team_h = QHBoxLayout(); add_team_h.setSpacing(6)
        self.cmb_add_member = QComboBox()
        self.cmb_add_member.addItems(emp_names)
        btn_add_member = QPushButton("Add to Team"); btn_add_member.setObjectName("btn_add_obs")
        btn_add_member.clicked.connect(self._add_team_member)
        add_team_h.addWidget(self.cmb_add_member, stretch=1)
        add_team_h.addWidget(btn_add_member)
        vt.addLayout(add_team_h)

        
        iv.addWidget(grp1)
        iv.addWidget(grp_team)
    
        # ── Document details ─────────────────────────────────────────────────
        grp2 = QGroupBox("Document Details"); f2 = QFormLayout(grp2); f2.setSpacing(8)
        f2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.in_version = QLineEdit("1.0")
        self.in_version.setPlaceholderText("e.g. 1.0  /  2.0  /  1.1")
        f2.addRow("Document Version:", self.in_version)
        iv.addWidget(grp2)

        # ── Client history ───────────────────────────────────────────────────
        grp3 = QGroupBox("Client History / Background")
        v3 = QVBoxLayout(grp3)
        self.txt_history = QTextEdit()
        self.txt_history.setPlaceholderText(
            "Brief background about the client organisation, their industry, "
            "and purpose of this assessment…")
        self.txt_history.setFixedHeight(110)
        v3.addWidget(self.txt_history); iv.addWidget(grp3)

        # ── Limitation ───────────────────────────────────────────────────────
        grp4 = QGroupBox("Limitation / Constraints")
        v4 = QVBoxLayout(grp4)
        self.txt_limitation = QTextEdit()
        self.txt_limitation.setPlaceholderText(
            "Any limitations during the assessment, e.g. no source code access, "
            "restricted testing hours, out-of-scope systems…")
        self.txt_limitation.setFixedHeight(90)
        v4.addWidget(self.txt_limitation); iv.addWidget(grp4)

        # ── Tools / Software Used ────────────────────────────────────────────
        grp5 = QGroupBox("Tools / Software Used")
        v5 = QVBoxLayout(grp5); v5.setSpacing(8)

        # Filter bar
        filter_h = QHBoxLayout()
        self.cmb_tool_cat = QComboBox()
        self.cmb_tool_cat.addItem("All Categories")
        for c in self.db.get_tool_categories():
            self.cmb_tool_cat.addItem(c)
        self.cmb_tool_cat.currentTextChanged.connect(self._filter_tools)
        filter_h.addWidget(QLabel("Filter:")); filter_h.addWidget(self.cmb_tool_cat)
        filter_h.addStretch()
        v5.addLayout(filter_h)

        # Tools table
        self.tools_tbl = QTableWidget(0, 6)
        self.tools_tbl.setHorizontalHeaderLabels(
            ["✓", "Tool Name", "Version", "Type", "Category", ""])
        hdr = self.tools_tbl.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.tools_tbl.setMinimumHeight(180)
        self.tools_tbl.setMaximumHeight(280)
        v5.addWidget(self.tools_tbl)

        # Add custom tool row
        add_h = QHBoxLayout(); add_h.setSpacing(6)
        self.in_tool_name    = QLineEdit(); self.in_tool_name.setPlaceholderText("Tool name")
        self.in_tool_ver     = QLineEdit(); self.in_tool_ver.setPlaceholderText("Version")
        self.cmb_tool_type   = QComboBox(); self.cmb_tool_type.addItems(["Open Source","Licensed"])
        self.cmb_tool_newcat = QComboBox()
        self.cmb_tool_newcat.addItems(["Web","API","Mobile","Source Code","Red Team","Internal","General"])
        btn_add_tool = QPushButton("Add"); btn_add_tool.setObjectName("btn_add_obs")
        btn_add_tool.setFixedWidth(60)
        btn_add_tool.clicked.connect(self._add_custom_tool)
        add_h.addWidget(self.in_tool_name, stretch=2)
        add_h.addWidget(self.in_tool_ver,  stretch=1)
        add_h.addWidget(self.cmb_tool_type)
        add_h.addWidget(self.cmb_tool_newcat)
        add_h.addWidget(btn_add_tool)
        v5.addLayout(add_h)

        iv.addWidget(grp5)
        self._filter_tools("All Categories")   # initial load

        iv.addStretch()
        scroll.setWidget(inner); v.addWidget(scroll, stretch=1)



    def _emp_by_name(self, name: str) -> dict:
        return next((e for e in self.db.get_employees() if e["name"] == name), {})

    def _upd_prep(self, name: str):
        self.lbl_prep_desig.setText(self._emp_by_name(name).get("designation",""))

    def _upd_rev(self, name: str):
        self.lbl_rev_desig.setText(self._emp_by_name(name).get("designation",""))

    def _manage_emp(self):
        # Track existing employees before dialog opens
        existing_names = {e["name"] for e in self.db.get_employees()}

        dlg = EmployeeDialog(self.db, self)
        dlg.exec()

        # Refresh all dropdowns
        all_emps = self.db.get_employees()
        names    = [""] + [e["name"] for e in all_emps]

        for cmb, cur in [
            (self.cmb_prepared,   self.cmb_prepared.currentText()),
            (self.cmb_reviewed,   self.cmb_reviewed.currentText()),
            (self.cmb_approved,   self.cmb_approved.currentText()),
            (self.cmb_released,   self.cmb_released.currentText()),
            (self.cmb_add_member, self.cmb_add_member.currentText()),
        ]:
            cmb.clear(); cmb.addItems(names)
            if cur in names: cmb.setCurrentText(cur)

        # Auto-add NEW employees to audit team table
        new_emps = [e for e in all_emps if e["name"] not in existing_names]
        for emp in new_emps:
            self._add_team_member_from_emp(emp)

    def get_data(self) -> dict:
        prep = self._emp_by_name(self.cmb_prepared.currentText())
        rev  = self._emp_by_name(self.cmb_reviewed.currentText())
        return {
            "prepared_by":             self.cmb_prepared.currentText(),
            "prepared_by_designation": prep.get("designation",""),
            "reviewed_by":             self.cmb_reviewed.currentText(),
            "reviewed_by_designation": rev.get("designation",""),
            "doc_version":             self.in_version.text().strip() or "1.0",
            "client_history":          self.txt_history.toPlainText().strip(),
            "limitation":              self.txt_limitation.toPlainText().strip(),
            "approved_by":              self.cmb_approved.currentText(),
            "approved_by_designation":  self._emp_by_name(self.cmb_approved.currentText()).get("designation",""),
            "released_by":              self.cmb_released.currentText(),
            "released_by_designation":  self._emp_by_name(self.cmb_released.currentText()).get("designation",""),
            "release_date":             self.in_release_date.date().toString("dd-MM-yyyy"),
            "selected_tools":           self.get_selected_tools(),
            "team_members":            self.get_team_members(),
        }

    def set_data(self, d: dict):
        for name, cmb in [(d.get("prepared_by",""), self.cmb_prepared),
                           (d.get("reviewed_by",""),  self.cmb_reviewed)]:
            if name and cmb.findText(name) >= 0:
                cmb.setCurrentText(name)
        self.in_version.setText(d.get("doc_version","1.0"))
        self.txt_history.setPlainText(d.get("client_history",""))
        self.txt_limitation.setPlainText(d.get("limitation",""))
        for name, cmb in [(d.get("approved_by",""), self.cmb_approved),
                   (d.get("released_by",""),  self.cmb_released)]:
            if name and cmb.findText(name) >= 0:
                cmb.setCurrentText(name)
        if d.get("release_date"):
            self.in_release_date.setDate(QDate.fromString(d["release_date"], "dd-MM-yyyy"))

    def _filter_tools(self, cat: str):
        tools = self.db.get_tools("" if cat == "All Categories" else cat)
        self.tools_tbl.setRowCount(0)
        self._tool_ids = []
        for t in tools:
            r = self.tools_tbl.rowCount(); self.tools_tbl.insertRow(r)
            # Column 0 — checkbox for selection
            chk = QCheckBox()
            chk.setStyleSheet("margin-left:8px")
            self.tools_tbl.setCellWidget(r, 0, chk)
            # Columns 1-4 — data
            self.tools_tbl.setItem(r, 1, QTableWidgetItem(t["tool_name"]))
            self.tools_tbl.setItem(r, 2, QTableWidgetItem(t["tool_version"]))
            self.tools_tbl.setItem(r, 3, QTableWidgetItem(t["tool_type"]))
            self.tools_tbl.setItem(r, 4, QTableWidgetItem(t["category"]))
            # Column 5 — delete button
            btn = QPushButton("✕"); btn.setObjectName("btn_del")
            btn.clicked.connect(lambda _, tid=t["id"]: self._remove_tool(tid))
            self.tools_tbl.setCellWidget(r, 5, btn)
            self.tools_tbl.setRowHeight(r, 32)
            self._tool_ids.append(t["id"])

    def _remove_tool(self, tool_id: int):
            self.db.delete_tool(tool_id)
            self._filter_tools(self.cmb_tool_cat.currentText())

    def _add_custom_tool(self):
            name = self.in_tool_name.text().strip()
            if not name:
                return
            self.db.add_tool(
                name,
                self.in_tool_ver.text().strip(),
                self.cmb_tool_type.currentText(),
                self.cmb_tool_newcat.currentText(),
            )
            self.in_tool_name.clear(); self.in_tool_ver.clear()
            self._filter_tools(self.cmb_tool_cat.currentText())

    def get_selected_tools(self) -> list[dict]:
            """Returns only tools the user has selected (checked) in the table."""
            selected = []
            for row in range(self.tools_tbl.rowCount()):
                chk = self.tools_tbl.cellWidget(row, 0)
                if chk and chk.isChecked():
                    selected.append({
                        "tool_name":    self.tools_tbl.item(row, 1).text() if self.tools_tbl.item(row, 1) else "",
                        "tool_version": self.tools_tbl.item(row, 2).text() if self.tools_tbl.item(row, 2) else "",
                        "tool_type":    self.tools_tbl.item(row, 3).text() if self.tools_tbl.item(row, 3) else "",
                        "category":     self.tools_tbl.item(row, 4).text() if self.tools_tbl.item(row, 4) else "",
                        "tool_id":      self._tool_ids[row] if row < len(self._tool_ids) else 0,
                    })
            return selected
    def _add_team_member(self):
       name = self.cmb_add_member.currentText().strip()
       if not name:
            return
       emp = self._emp_by_name(name)
       self._add_team_member_from_emp(emp)

    def _add_team_member_from_emp(self, emp: dict):
        """Add an employee dict directly to the audit team table."""
        name = emp.get("name", "")
        # Prevent duplicates
        for r in range(self.team_list.rowCount()):
            if self.team_list.item(r, 0) and self.team_list.item(r, 0).text() == name:
                return
        r = self.team_list.rowCount(); self.team_list.insertRow(r)
        self.team_list.setItem(r, 0, QTableWidgetItem(emp.get("name", "")))
        self.team_list.setItem(r, 1, QTableWidgetItem(emp.get("designation", "")))
        self.team_list.setItem(r, 2, QTableWidgetItem(emp.get("email", "")))
        self.team_list.setItem(r, 3, QTableWidgetItem(emp.get("qualifications", "")))
        self.team_list.setItem(r, 4, QTableWidgetItem(emp.get("cert_in_listed", "No")))
        self.team_list.setRowHeight(r, 32)

    def get_team_members(self) -> list[dict]:
        members = []
        for r in range(self.team_list.rowCount()):
            members.append({
                "name":           self._tbl_cell(r, 0),
                "designation":    self._tbl_cell(r, 1),
                "email":          self._tbl_cell(r, 2),
                "qualifications": self._tbl_cell(r, 3),
                "cert_in_listed": self._tbl_cell(r, 4),
            })
        return members

    def _tbl_cell(self, row: int, col: int) -> str:
        item = self.team_list.item(row, col)
        return item.text().strip() if item else ""

    def _upd_appr(self, name: str): #helper method
        self.lbl_appr_desig.setText(self._emp_by_name(name).get("designation", ""))

    def _upd_rel(self, name: str):
        self.lbl_rel_desig.setText(self._emp_by_name(name).get("designation", "")) #to here



# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — Report Details
# ═══════════════════════════════════════════════════════════════════════════════

class ObsTable(QWidget):
    """Manual observation entry table with Add / Library-picker / Delete."""

    COLS = ["#","Title","Severity","Affected URL","CVE",
            "Description","Impact","Recommendation",""]

    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(6)

        h = QHBoxLayout()
        btn_add = QPushButton("＋  Add Row");   btn_add.setObjectName("btn_add_obs")
        btn_lib = QPushButton("⚡  From Library"); btn_lib.setObjectName("btn_lib")
        btn_add.clicked.connect(self._add_blank)
        btn_lib.clicked.connect(self._pick_lib)
        h.addWidget(btn_add); h.addWidget(btn_lib); h.addStretch()
        v.addLayout(h)

        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        hdr = self.tbl.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        for c in [5,6,7]: hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(8, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.setMinimumHeight(400)
        self.tbl.setWordWrap(True)
        self.tbl.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.tbl.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )
        v.addWidget(self.tbl)

    def _add_blank(self): self._add_row({})

    def _pick_lib(self):
        dlg = ObsLibraryDialog(self.db, self)
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.selected:
            self._add_row(dlg.selected)

    def _add_row(self, data: dict):
        r = self.tbl.rowCount(); self.tbl.insertRow(r)

        sev_cmb = QComboBox()
        sev_cmb.addItems(["Critical","High","Medium","Low","Info"])
        sev = data.get("severity","Medium")
        idx = sev_cmb.findText(sev, Qt.MatchFlag.MatchFixedString)
        if idx >= 0: sev_cmb.setCurrentIndex(idx)
        # Colour severity dropdown text
        sev_cmb.currentTextChanged.connect(
            lambda t, cb=sev_cmb: cb.setStyleSheet(
                f"color:{SEV_COLORS.get(t,'#e2e8f0')}"))
        sev_cmb.setStyleSheet(f"color:{SEV_COLORS.get(sev,'#e2e8f0')}")

        self.tbl.setItem(r, 0, QTableWidgetItem(str(r+1)))
        self.tbl.setItem(r, 1, QTableWidgetItem(data.get("title","")))
        self.tbl.setCellWidget(r, 2, sev_cmb)
        self.tbl.setItem(r, 3, QTableWidgetItem(data.get("affected_url","")))
        self.tbl.setItem(r, 4, QTableWidgetItem(data.get("cve","")))
        for col, key in [(5,"description"),(6,"impact"),(7,"recommendation")]:
            te = QTextEdit()
            te.setPlainText(data.get(key, ""))
            te.setMinimumHeight(75)
            te.setStyleSheet(
                "background:#2d3748;color:#e2e8f0;"
                "border:none;font-size:12px;padding:4px;"
            )
            self.tbl.setCellWidget(r, col, te)

        btn_del = QPushButton("✕"); btn_del.setObjectName("btn_del")
        btn_del.clicked.connect(lambda _, row=r: self._del(row))
        self.tbl.setCellWidget(r, 8, btn_del)
        self.tbl.setRowHeight(r, 85)

    def _del(self, row: int):
        self.tbl.removeRow(row)
        for i in range(self.tbl.rowCount()):
            self.tbl.setItem(i, 0, QTableWidgetItem(str(i+1)))

    def _cell(self, r: int, c: int) -> str:
        widget = self.tbl.cellWidget(r, c)
        if isinstance(widget, QTextEdit):
            return widget.toPlainText().strip()
        item = self.tbl.item(r, c)
        return item.text().strip() if item else ""

    def get_observations(self) -> list[dict]:
        obs = []
        for r in range(self.tbl.rowCount()):
            wgt = self.tbl.cellWidget(r, 2)
            obs.append({
                "sr_no":          self._cell(r,0),
                "title":          self._cell(r,1),
                "severity":       wgt.currentText() if wgt else "Medium",
                "affected_url":   self._cell(r,3),
                "cve":            self._cell(r,4),
                "description":    self._cell(r,5),
                "impact":         self._cell(r,6),
                "recommendation": self._cell(r,7),
            })
        return obs

    def set_observations(self, obs_list: list[dict]):
        self.tbl.setRowCount(0)
        for obs in obs_list: self._add_row(obs)


class Page2Report(QWidget):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(10)

        lbl = QLabel("Report Details"); lbl.setObjectName("page_title")
        sub = QLabel("Client info, audit period, report type, and observation data source.")
        sub.setObjectName("page_sub"); sub.setWordWrap(True)
        v.addWidget(lbl); v.addWidget(sub); v.addWidget(_divider())

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        inner = QWidget(); iv = QVBoxLayout(inner)
        iv.setSpacing(12); iv.setContentsMargins(4,4,8,4)
        inner.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.MinimumExpanding)

        # ── Client & Application ─────────────────────────────────────────────
        grp1 = QGroupBox("Client & Application"); f1 = QFormLayout(grp1); f1.setSpacing(8)
        f1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.in_client   = QLineEdit(); self.in_client.setPlaceholderText("e.g. Acme Bank Pvt. Ltd.")
        self.in_app      = QLineEdit(); self.in_app.setPlaceholderText("e.g. Internet Banking Portal")

        self.in_app_type = QComboBox()
        self.in_app_type.addItems(["External","Internal"])

        # Date pickers (kept from your original)
        self.in_start = QDateEdit(); self.in_start.setCalendarPopup(True)
        self.in_start.setDate(QDate.currentDate()); self.in_start.setDisplayFormat("dd-MM-yyyy")
        self.in_end   = QDateEdit(); self.in_end.setCalendarPopup(True)
        self.in_end.setDate(QDate.currentDate().addDays(14))
        self.in_end.setDisplayFormat("dd-MM-yyyy")
        date_h = QHBoxLayout()
        date_h.addWidget(self.in_start); date_h.addWidget(QLabel("to"))
        date_h.addWidget(self.in_end);   date_h.addStretch()

        self.in_url    = QLineEdit(); self.in_url.setPlaceholderText("https://example.com")
        self.in_method = QComboBox(); self.in_method.addItems(["Grey Box","Black Box","White Box"])



        f1.addRow("Client Name:",  self.in_client)
        f1.addRow("App Name:",     self.in_app)
        f1.addRow("App Type:",     self.in_app_type)
        f1.addRow("Audit Period:", date_h)
        f1.addRow("Target URL:",   self.in_url)
        f1.addRow("Test Method:",  self.in_method)
        iv.addWidget(grp1)

        # ── Report Settings ──────────────────────────────────────────────────
        grp2 = QGroupBox("Report Settings"); f2 = QFormLayout(grp2); f2.setSpacing(8)
        f2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.cmb_type = QComboBox(); self.cmb_type.addItems(["Web","Api","Mobile"])
        self.cmb_env  = QComboBox(); self.cmb_env.addItems(["Production","Uat"])
        f2.addRow("Report Type:", self.cmb_type)
        f2.addRow("Environment:", self.cmb_env)
        iv.addWidget(grp2)

        # ── Client Contact ───────────────────────────────────────────────────
        grp_contact = QGroupBox("Client Contact Person")
        fc = QFormLayout(grp_contact); fc.setSpacing(8)
        fc.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.in_contact_name  = QLineEdit(); self.in_contact_name.setPlaceholderText("e.g. John Smith")
        self.in_contact_desig = QLineEdit(); self.in_contact_desig.setPlaceholderText("e.g. IT Manager")
        self.in_contact_email = QLineEdit(); self.in_contact_email.setPlaceholderText("e.g. john@client.com")
        fc.addRow("Contact Person:", self.in_contact_name)
        fc.addRow("Designation:",    self.in_contact_desig)
        fc.addRow("Email:",          self.in_contact_email)
        iv.addWidget(grp_contact)

        # ── Data Source ──────────────────────────────────────────────────────
        grp3 = QGroupBox("Observation Data Source")
        v3 = QVBoxLayout(grp3); v3.setSpacing(8)

        mode_h = QHBoxLayout()
        self.btn_excel  = QPushButton("📊  Excel File")
        self.btn_manual = QPushButton("✏️  Manual Entry")
        for b in [self.btn_excel, self.btn_manual]:
            b.setCheckable(True); b.setFixedHeight(32)
        self.btn_excel.setChecked(True)
        self.btn_excel.clicked.connect(lambda: self._mode("excel"))
        self.btn_manual.clicked.connect(lambda: self._mode("manual"))
        mode_h.addWidget(self.btn_excel); mode_h.addWidget(self.btn_manual)
        mode_h.addStretch(); v3.addLayout(mode_h)

        # Excel panel
        self.pnl_excel = QWidget()
        pe = QFormLayout(self.pnl_excel); pe.setSpacing(8)
        pe.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.pick_excel = FilePicker("Select findings .xlsx file")
        pe.addRow("Excel File:", self.pick_excel)
        v3.addWidget(self.pnl_excel)

        # Manual panel
        self.pnl_manual = QWidget()
        pm = QVBoxLayout(self.pnl_manual)
        pm.setContentsMargins(0,0,0,0); pm.setSpacing(0)
        self.obs_table = ObsTable(self.db)
        self.obs_table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        pm.addWidget(self.obs_table)
        self.pnl_manual.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.pnl_manual.hide()

        v3.addWidget(self.pnl_manual)
        iv.addWidget(grp3)

        # ── Files & Output ───────────────────────────────────────────────────
        grp4 = QGroupBox("Files & Output"); f4 = QFormLayout(grp4); f4.setSpacing(8)
        f4.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.pick_poc    = FilePicker("Optional: POC screenshots folder", folder=True)
        self.pick_output = FilePicker("Leave blank for auto-named output", save=True)
        f4.addRow("POC Folder:",  self.pick_poc)
        f4.addRow("Output File:", self.pick_output)
        iv.addWidget(grp4)

        iv.addStretch()
        scroll.setWidget(inner); v.addWidget(scroll, stretch=1)

    def _mode(self, m: str):
        excel = (m == "excel")
        self.btn_excel.setChecked(excel); self.btn_manual.setChecked(not excel)
        self.pnl_excel.setVisible(excel); self.pnl_manual.setVisible(not excel)

    def is_manual(self) -> bool: return self.btn_manual.isChecked()

    def audit_period(self) -> str:
        return (f"{self.in_start.date().toString('dd-MM-yyyy')} - "
                f"{self.in_end.date().toString('dd-MM-yyyy')}")

    def get_data(self) -> dict:
        return {
            "client_name":  self.in_client.text().strip(),
            "app_name":     self.in_app.text().strip(),
            "app_type":     self.in_app_type.currentText(),
            "audit_period": self.audit_period(),
            "url":          self.in_url.text().strip(),
            "method":       self.in_method.currentText(),
            "report_type":  self.cmb_type.currentText(),
            "environment":  self.cmb_env.currentText(),
            "excel_file":   self.pick_excel.text(),
            "poc_folder":   self.pick_poc.text(),
            "output_file":  self.pick_output.text(),
            "manual_mode":  self.is_manual(),
            "manual_obs":   self.obs_table.get_observations() if self.is_manual() else [],
           "client_contact_person": self.in_contact_name.text().strip(),
            "client_designation":    self.in_contact_desig.text().strip(),
            "client_email":          self.in_contact_email.text().strip(),
        }

    def set_data(self, d: dict):
        self.in_client.setText(d.get("client_name",""))
        self.in_app.setText(d.get("app_name",""))
        idx = self.in_app_type.findText(d.get("app_type","External"))
        if idx >= 0: self.in_app_type.setCurrentIndex(idx)
        self.in_url.setText(d.get("url",""))
        idx = self.in_method.findText(d.get("method","Grey Box"))
        if idx >= 0: self.in_method.setCurrentIndex(idx)
        idx = self.cmb_type.findText(d.get("report_type","Web"))
        if idx >= 0: self.cmb_type.setCurrentIndex(idx)
        idx = self.cmb_env.findText(d.get("environment","Production"))
        if idx >= 0: self.cmb_env.setCurrentIndex(idx)
        self.pick_excel.setText(d.get("excel_file",""))
        self.pick_poc.setText(d.get("poc_folder",""))
        self.pick_output.setText(d.get("output_file",""))
        if d.get("manual_mode"):
            self._mode("manual")
            if d.get("manual_obs"):
                self.obs_table.set_observations(d["manual_obs"])
        # Restore dates if saved
        if d.get("start_date"):
            self.in_start.setDate(QDate.fromString(d["start_date"], "dd-MM-yyyy"))
        if d.get("end_date"):
            self.in_end.setDate(QDate.fromString(d["end_date"], "dd-MM-yyyy"))
        self.in_contact_name.setText(d.get("client_contact_person",""))
        self.in_contact_desig.setText(d.get("client_designation",""))
        self.in_contact_email.setText(d.get("client_email",""))

    def get_profile_dates(self) -> dict:
        return {
            "start_date": self.in_start.date().toString("dd-MM-yyyy"),
            "end_date":   self.in_end.date().toString("dd-MM-yyyy"),
        }


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — Generate
# ═══════════════════════════════════════════════════════════════════════════════

class Page3Generate(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(10)

        lbl = QLabel("Generate Report"); lbl.setObjectName("page_title")
        sub = QLabel("Review the summary, then click Generate. Watch live progress on the right.")
        sub.setObjectName("page_sub"); sub.setWordWrap(True)
        v.addWidget(lbl); v.addWidget(sub); v.addWidget(_divider())

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left: summary
        sw = QWidget(); sv = QVBoxLayout(sw); sv.setContentsMargins(0,0,8,0)
        sv.addWidget(QLabel("Report Summary"))
        self.summary = QPlainTextEdit(); self.summary.setReadOnly(True)
        sv.addWidget(self.summary, stretch=1)
        splitter.addWidget(sw)

        # Right: logs + progress
        lw = QWidget(); lv = QVBoxLayout(lw); lv.setContentsMargins(8,0,0,0)
        lv.addWidget(QLabel("Live Output"))
        self.log_view = QPlainTextEdit(); self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("Logs will appear here…")
        lv.addWidget(self.log_view, stretch=1)
        self.progress_bar = QProgressBar(); self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        lv.addWidget(self.progress_bar)
        btn_clr = QPushButton("Clear Log"); btn_clr.setFixedHeight(26)
        btn_clr.clicked.connect(self.log_view.clear); lv.addWidget(btn_clr)
        splitter.addWidget(lw)

        splitter.setSizes([420, 460])
        v.addWidget(splitter, stretch=1)

    def set_summary(self, text: str): self.summary.setPlainText(text)
    def append_log(self, msg: str):   self.log_view.appendPlainText(msg)
    def set_progress(self, v: int):   self.progress_bar.setValue(v)
    def reset(self): self.log_view.clear(); self.progress_bar.setValue(0)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1020, 700); self.resize(1200, 780)
        self.db = DBManager()
        self._thread = self._worker = None
        self._build_ui()
        self.setStyleSheet(DARK)
        self.setStatusBar(QStatusBar())
        self._status("Ready")

    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20,14,20,10); root.setSpacing(10)
        root.addWidget(self._header())
        root.addWidget(_divider())

        self.step_lbl = QLabel("Step 1 of 3 — General Information")
        self.step_lbl.setObjectName("step_label")
        root.addWidget(self.step_lbl)

        self.stack  = QStackedWidget()
        self.page1  = Page1General(self.db)
        self.page2  = Page2Report(self.db)
        self.page3  = Page3Generate()
        for p in [self.page1, self.page2, self.page3]:
            self.stack.addWidget(p)
        root.addWidget(self.stack, stretch=1)
        root.addWidget(_divider())
        root.addWidget(self._nav_bar())
    
    def _apply_theme(self, theme_name: str):
        self.setStyleSheet(THEMES.get(theme_name, THEMES["Dark"]))

    def _header(self) -> QWidget:
        w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0,0,0,0)
        title = QLabel(APP_NAME); title.setObjectName("lbl_title")
        sub   = QLabel(f"Security Audit Report Generator  ·  v{APP_VERSION}")
        sub.setObjectName("lbl_sub")
        h.addWidget(title); h.addSpacing(12); h.addWidget(sub); h.addStretch()
        btn_load = QPushButton("⬆  Load Profile"); btn_load.setObjectName("btn_load_profile")
        btn_save = QPushButton("💾  Save Profile"); btn_save.setObjectName("btn_save_profile")
        btn_load.clicked.connect(self._load_profile)
        btn_save.clicked.connect(self._save_profile)

        # Theme switcher
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        self.cmb_theme.setFixedWidth(130)
        self.cmb_theme.currentTextChanged.connect(self._apply_theme)

        h.addWidget(QLabel("Theme:"))
        h.addWidget(self.cmb_theme)
        h.addSpacing(10)
        h.addWidget(btn_load); h.addSpacing(6); h.addWidget(btn_save)
        return w

    def _nav_bar(self) -> QWidget:
        w = QWidget(); h = QHBoxLayout(w); h.setContentsMargins(0,0,0,0)
        self.btn_back = QPushButton("← Back"); self.btn_back.setObjectName("btn_back")
        self.btn_back.setFixedHeight(42); self.btn_back.setEnabled(False)
        self.btn_back.clicked.connect(self._back)

        self.btn_next = QPushButton("Next →"); self.btn_next.setObjectName("btn_next")
        self.btn_next.setFixedHeight(42); self.btn_next.clicked.connect(self._next)

        self.btn_generate = QPushButton("⚡  Generate Report")
        self.btn_generate.setObjectName("btn_generate")
        self.btn_generate.setFixedHeight(42); self.btn_generate.hide()
        self.btn_generate.clicked.connect(self._on_generate)

        h.addWidget(self.btn_back); h.addStretch()
        h.addWidget(self.btn_next); h.addWidget(self.btn_generate)
        return w

    # ── Navigation ────────────────────────────────────────────────────────────

    def _next(self):
        idx = self.stack.currentIndex()
        if idx == 1:
            d2 = self.page2.get_data()
            if not d2["manual_mode"] and not d2["excel_file"]:
                QMessageBox.warning(self, "Missing Input",
                    "Please select an Excel file or switch to Manual mode."); return
            if d2["manual_mode"] and not d2["manual_obs"]:
                QMessageBox.warning(self, "No Observations",
                    "Please add at least one observation."); return
            self._refresh_summary()
        if idx < 2:
            self.stack.setCurrentIndex(idx + 1)
            self._upd_nav()

    def _back(self):
        idx = self.stack.currentIndex()
        if idx > 0:
            self.stack.setCurrentIndex(idx - 1)
            self._upd_nav()

    def _upd_nav(self):
        idx   = self.stack.currentIndex()
        steps = ["Step 1 of 3 — General Information",
                 "Step 2 of 3 — Report Details",
                 "Step 3 of 3 — Generate Report"]
        self.step_lbl.setText(steps[idx])
        self.btn_back.setEnabled(idx > 0)
        last = (idx == 2)
        self.btn_next.setVisible(not last)
        self.btn_generate.setVisible(last)

    def _refresh_summary(self):
        d1 = self.page1.get_data(); d2 = self.page2.get_data()
        mode = "Manual" if d2["manual_mode"] else "Excel"
        obs  = len(d2["manual_obs"]) if d2["manual_mode"] else "from Excel"
        self.page3.set_summary(
            f"Client:        {d2['client_name']}\n"
            f"Application:   {d2['app_name']}\n"
            f"Type:          {d2['app_type']}\n"
            f"Period:        {d2['audit_period']}\n"
            f"URL:           {d2['url']}\n"
            f"Method:        {d2['method']}\n"
            f"Report Type:   {d2['report_type']}\n"
            f"Environment:   {d2['environment']}\n"
            f"\n"
            f"Prepared By:   {d1['prepared_by']} ({d1['prepared_by_designation']})\n"
            f"Reviewed By:   {d1['reviewed_by']} ({d1['reviewed_by_designation']})\n"
            f"Doc Version:   {d1['doc_version']}\n"
            f"\n"
            f"Data Source:   {mode}\n"
            f"Observations:  {obs}\n"
            f"POC Folder:    {d2['poc_folder'] or 'Not set'}\n"
            f"Output:        {d2['output_file'] or 'Auto-named'}"
        )
        self.page3.reset()

    # ── Generate ──────────────────────────────────────────────────────────────

    def _on_generate(self):
        config = self._build_config()
        if not config: return
        self.btn_generate.setEnabled(False)
        self._status("Generating report…")

        self._worker = GeneratorWorker(config)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self.page3.append_log)
        self._worker.progress.connect(self.page3.set_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _build_config(self) -> ReportConfig | None:
        d1 = self.page1.get_data(); d2 = self.page2.get_data()
        cfg = ReportConfig(
            prepared_by              = d1["prepared_by"],
            prepared_by_designation  = d1["prepared_by_designation"],
            reviewed_by              = d1["reviewed_by"],
            reviewed_by_designation  = d1["reviewed_by_designation"],
            doc_version              = d1["doc_version"],
            client_history           = d1["client_history"],
            limitation               = d1["limitation"],
            client_name              = d2["client_name"],
            app_name                 = d2["app_name"],
            app_type                 = d2["app_type"],
            audit_period             = d2["audit_period"],
            url                      = d2["url"],
            method                   = d2["method"],
            report_type              = d2["report_type"],
            environment              = d2["environment"],
            poc_folder               = d2["poc_folder"],
            output_file              = d2["output_file"],
            approved_by              = d1["approved_by"],
            approved_by_designation  = d1["approved_by_designation"],
            released_by              = d1["released_by"],
            released_by_designation  = d1["released_by_designation"],
            release_date             = d1["release_date"],
            client_contact_person    = d2["client_contact_person"],
            client_designation       = d2["client_designation"],
            client_email             = d2["client_email"],
            selected_tools           = d1["selected_tools"],
            team_members             = d1["team_members"],
        )
        if d2["manual_mode"]:
            cfg.manual_observations = d2["manual_obs"]
        else:
            cfg.excel_file = d2["excel_file"]
        return cfg

    @pyqtSlot(object)
    def _on_finished(self, result: ReportResult):
        self.btn_generate.setEnabled(True)
        if result.success:
            d1 = self.page1.get_data(); d2 = self.page2.get_data()
            self.db.save_report_history(d2["client_name"], d2["app_name"],
                                        d2["report_type"], result.output_path,
                                        d1["prepared_by"])
            self._status(f"Done ✓  →  {result.output_path}")
            QMessageBox.information(self, "Report Generated",
                f"✅ Report created successfully!\n\n"
                f"📄 Output:  {result.output_path}\n"
                f"📊 Observations:  {result.observations_count}\n\n"
                "⚠️  Please review before sharing.")
        else:
            self._status(f"Error: {result.error}")
            QMessageBox.critical(self, "Generation Failed",
                f"❌ Report generation failed:\n\n{result.error}")

    def _status(self, msg: str): self.statusBar().showMessage(msg, 8000)

    # ── Profiles ──────────────────────────────────────────────────────────────

    def _full_profile(self) -> dict:
        d2 = self.page2.get_data()
        d2.update(self.page2.get_profile_dates())   # save dates too
        return {"page1": self.page1.get_data(), "page2": d2}

    def _save_profile(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Profile", str(PROFILE_DIR), "JSON Profiles (*.json)")
        if not path: return
        if not path.endswith(".json"): path += ".json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self._full_profile(), f, indent=2)
        self._status(f"Profile saved: {path}")
        QMessageBox.information(self, "Saved", f"✅ Profile saved:\n{path}")

    def _load_profile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Profile", str(PROFILE_DIR), "JSON Profiles (*.json)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f: data = json.load(f)
            if "page1" in data: self.page1.set_data(data["page1"])
            if "page2" in data: self.page2.set_data(data["page2"])
            self._status(f"Profile loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Load Failed", f"Could not load profile:\n{e}")


# ── Entry point ────────────────────────────────────────────────────────────────

def launch_gui():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    try:
        app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)
    except AttributeError:
        pass
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    launch_gui()
