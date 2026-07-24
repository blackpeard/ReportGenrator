"""
gui_app.py
==========
PyQt6 GUI for the Report Engine — redesigned shell.

Layout
------
HOME (launcher): your background image fills the window; six glassy buttons
  (General Info, Generate Excel, Word from Excel, Manual Word, Misc, Credit).
Click one -> view fades to the MAIN layout: those buttons become a left
  sidebar, the chosen section fills the right. "Home" returns to the launcher.

Section map (where your old wizard pages went)
  General Info     -> Page1General (personnel/tools/team/doc) + ReportDetailsForm
                      (client/app/dates/contact/output) shown as two tabs.
  Generate Excel   -> ExcelTemplateGenerator (report-type picker + Generate).
  Word from Excel  -> single/batch Excel pickers + generate + live output.
  Manual Word      -> manual observation table + generate + live output.
  Misc             -> manage employees, observation library, theme, profiles.
  Credit           -> about screen (shows the background image).

Both generate sections read the shared General Info data, exactly like the old
_build_config did (page1 ∪ report-details), then add their own excel/manual data.

Background image: set BACKGROUND_IMAGE below, or call shell.set_background_image().
It shows on the launcher, the main-view margins, and the Credit screen; the
forms stay on the solid theme so they remain easy to read.

Run:  python gui/gui_app.py
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PyQt6.QtCore import (
    QDate, QObject, QThread, Qt, pyqtSignal, pyqtSlot,
    QPropertyAnimation, QEasingCurve, QRect,
)
from PyQt6.QtGui import QColor, QPixmap, QPainter, QLinearGradient
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QListWidget, QListWidgetItem, QComboBox, QDateEdit,
    QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QCheckBox, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QScrollArea, QSizePolicy, QSplitter,
    QStackedWidget, QStatusBar, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget, QFrame, QGridLayout, QTabWidget,
    QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
)

from report_engine import ReportConfig, ReportEngine, ReportResult, BatchRunner, BatchResult
from db_manager    import DBManager
from src.excel_template import ExcelTemplateGenerator

# ── Constants ─────────────────────────────────────────────────────────────────
APP_NAME    = "Advance Report_Generator Machine"
APP_VERSION = "5.0.0"
PROFILE_DIR = Path.home() / ".report_generator" / "profiles"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)

# Set to your image (PNG/JPG). Empty -> built-in gradient.
BACKGROUND_IMAGE = ""

# key, label, icon glyph, subtitle
SECTIONS = [
    ("general_info",    "General Info",    "\u270e", "Enter every report input here"),
    ("generate_excel",  "Generate Excel",  "\u25a6", "Build a findings workbook"),
    ("word_from_excel", "Word from Excel", "\u25a4", "Excel findings \u2192 .docx report"),
    ("manual_word",     "Manual Word",     "\u270d", "Type observations, then generate"),
    ("misc",            "Misc",            "\u2699", "Employees, tools, library, theme"),
    ("credit",          "Credit",          "\u2605", "About this tool"),
]

SEV_COLORS = {
    "Critical": "#C00000", "High": "#EE0000",
    "Medium": "#FFC000", "Low": "#00B050", "Info": "#0070C0"
}

# ── Dark stylesheet (unchanged from your build) ───────────────────────────────
DARK = """
QMainWindow,QWidget{background:#1a1d23;color:#e2e8f0;font-family:'Segoe UI','SF Pro Display',sans-serif;font-size:13px}
QGroupBox{border:1px solid #2d3748;border-radius:8px;margin-top:12px;padding:12px 8px 8px;background:#1e2330}
QGroupBox::title{subcontrol-origin:margin;left:12px;padding:0 6px;color:#63b3ed;font-weight:600;font-size:11px;letter-spacing:0.5px;text-transform:uppercase}
QLineEdit,QComboBox,QTextEdit,QPlainTextEdit,QDateEdit{background:#2d3748;border:1px solid #4a5568;border-radius:6px;padding:6px 10px;color:#e2e8f0}
QLineEdit:focus,QComboBox:focus,QTextEdit:focus,QDateEdit:focus{border-color:#63b3ed;background:#2d3a4f}
QLineEdit:disabled{color:#718096;background:#252d3d}
QComboBox::drop-down{border:none;padding-right:8px}
QComboBox QAbstractItemView{background:#2d3748;border:1px solid #4a5568;selection-background-color:#3182ce;outline:none}
QDateEdit::drop-down{border:none;padding-right:8px}
QCalendarWidget{background:#1e2330;color:#e2e8f0}
QPushButton{background:#2d3748;border:1px solid #4a5568;border-radius:6px;padding:7px 16px;color:#e2e8f0;font-weight:500}
QPushButton:hover{background:#3a4a5c;border-color:#63b3ed}
QPushButton:pressed{background:#2a3a4a}
QPushButton:disabled{color:#4a5568;background:#1e2330;border-color:#2d3748}
QPushButton:checked{background:#2b4c7e;border-color:#63b3ed;color:#90cdf4}
QPushButton#btn_generate{background:#276749;border-color:#38a169;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:8px}
QPushButton#btn_generate:hover{background:#38a169}
QPushButton#btn_generate:disabled{background:#1a3a2a;border-color:#276749;color:#4a5568}
QPushButton#btn_save_profile{background:#276749;border-color:#38a169;color:#fff}
QPushButton#btn_load_profile{background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_add_obs{background:#553c9a;border-color:#6b46c1;color:#fff}
QPushButton#btn_add_obs:hover{background:#6b46c1}
QPushButton#btn_lib{background:#1a365d;border-color:#2b6cb0;color:#63b3ed}
QPushButton#btn_lib:hover{background:#2b6cb0;color:#fff}
QPushButton#btn_del{background:#742a2a;border-color:#c53030;color:#fff;padding:3px 10px}
QPushButton#btn_del:hover{background:#c53030}
QProgressBar{background:#2d3748;border:1px solid #4a5568;border-radius:6px;text-align:center;color:#e2e8f0;font-weight:600;height:22px}
QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #2b6cb0,stop:1 #38a169);border-radius:5px}
QTableWidget{background:#1e2330;border:1px solid #2d3748;gridline-color:#2d3748;border-radius:6px}
QTableWidget::item{padding:5px;border:none}
QTableWidget::item:selected{background:#2b6cb0;color:#fff}
QHeaderView::section{background:#2d3748;color:#a0aec0;border:none;padding:6px;font-size:12px;font-weight:600}
QListWidget{background:#1e2330;border:1px solid #2d3748;border-radius:6px}
QListWidget::item{padding:6px 10px}
QListWidget::item:hover{background:#2d3748}
QListWidget::item:selected{background:#2b6cb0;color:#fff}
QTabWidget::pane{border:1px solid #2d3748;border-radius:8px;top:-1px}
QTabBar::tab{background:#1e2330;color:#a0aec0;padding:8px 18px;border:1px solid #2d3748;border-bottom:none;border-top-left-radius:8px;border-top-right-radius:8px;margin-right:3px}
QTabBar::tab:selected{background:#2b4c7e;color:#fff}
QScrollBar:vertical{background:#1a1d23;width:8px;border-radius:4px}
QScrollBar::handle:vertical{background:#4a5568;border-radius:4px;min-height:24px}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0}
QStatusBar{background:#141720;color:#718096;border-top:1px solid #2d3748}
QSplitter::handle{background:#2d3748;width:2px}
QFrame#divider{background:#2d3748;max-height:1px}
QLabel#lbl_title{color:#63b3ed;font-size:20px;font-weight:700}
QLabel#lbl_sub{color:#718096;font-size:11px}
QLabel#page_title{color:#e2e8f0;font-size:18px;font-weight:700}
QLabel#page_sub{color:#718096;font-size:12px}
"""

# Shell chrome (launcher, sidebar, cards, credit). Appended to the active theme.
SHELL_QSS = """
#rootStack,#mainView,#launcher,#launchWrap,#creditPage,#contentStack{background:transparent}
#launchTitle{font-size:34px;font-weight:800;color:#ffffff}
#launchSub{font-size:14px;color:rgba(255,255,255,0.72)}
#navCard{background:rgba(24,28,38,0.72);border:1px solid rgba(255,255,255,0.09);border-radius:18px;text-align:left}
#navCard:hover{background:rgba(43,71,116,0.66);border:1px solid rgba(99,179,237,0.60)}
#navCardIcon{font-size:26px;color:#8fc4f5}
#navCardTitle{font-size:18px;font-weight:700;color:#ffffff}
#navCardSub{font-size:12px;color:rgba(255,255,255,0.62)}
#sidebar{background:#141823;border:1px solid #2d3748;border-radius:18px}
#brand{font-size:15px;font-weight:800;color:#e2e8f0}
#sideVer{font-size:11px;color:#4a5568}
#sideHome,#sideItem{text-align:left;padding:10px 14px;border-radius:10px;background:transparent;border:1px solid transparent;color:#cbd5e0;font-weight:600;font-size:14px}
#sideHome{color:#63b3ed}
#sideHome:hover,#sideItem:hover{background:#1e2330;border:1px solid #2d3748}
#sideItem:checked{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 rgba(99,179,237,0.30),stop:1 rgba(99,179,237,0.08));border:1px solid rgba(99,179,237,0.45);color:#fff}
#creditTitle{font-size:30px;font-weight:800;color:#ffffff}
#creditText{font-size:14px;color:rgba(255,255,255,0.82)}
QPushButton#primary{background:#2b6cb0;border:1px solid #3182ce;color:#fff;font-weight:700;padding:10px 22px;border-radius:10px}
QPushButton#primary:hover{background:#3182ce}
"""


# ── Helpers ───────────────────────────────────────────────────────────────────
def _divider() -> QFrame:
    f = QFrame(); f.setObjectName("divider")
    f.setFrameShape(QFrame.Shape.HLine); return f


def _shadow(widget, blur=26, y=10, alpha=160):
    eff = QGraphicsDropShadowEffect(widget)
    eff.setBlurRadius(blur); eff.setXOffset(0); eff.setYOffset(y)
    eff.setColor(QColor(0, 0, 0, alpha)); widget.setGraphicsEffect(eff)
    return eff


def _click_through(label: QLabel):
    label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)


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


# ── Background painter ────────────────────────────────────────────────────────
class BackgroundHost(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self._pix = QPixmap()

    def set_pixmap(self, pix: QPixmap):
        self._pix = pix if pix and not pix.isNull() else QPixmap()
        self.update()

    def paintEvent(self, _evt):
        p = QPainter(self); p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        r = self.rect()
        if not self._pix.isNull():
            scaled = self._pix.scaled(
                r.size(), Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation)
            x = (scaled.width() - r.width()) // 2
            y = (scaled.height() - r.height()) // 2
            p.drawPixmap(r, scaled, QRect(x, y, r.width(), r.height()))
            ov = QLinearGradient(0, 0, 0, r.height())
            ov.setColorAt(0.0, QColor(8, 10, 16, 150)); ov.setColorAt(1.0, QColor(8, 10, 16, 205))
            p.fillRect(r, ov)
        else:
            g = QLinearGradient(0, 0, r.width(), r.height())
            g.setColorAt(0.0, QColor("#0e1320")); g.setColorAt(0.55, QColor("#14182a"))
            g.setColorAt(1.0, QColor("#1c1430"))
            p.fillRect(r, g)
        p.end()


# ── Workers (unchanged) ───────────────────────────────────────────────────────
class GeneratorWorker(QObject):
    log = pyqtSignal(str); progress = pyqtSignal(int); finished = pyqtSignal(object)
    def __init__(self, config: ReportConfig):
        super().__init__(); self._config = config
    @pyqtSlot()
    def run(self):
        engine = ReportEngine(self._config, progress_callback=self.progress.emit,
                              log_callback=self.log.emit)
        self.finished.emit(engine.run())


class BatchWorker(QObject):
    log = pyqtSignal(str); progress = pyqtSignal(int); finished = pyqtSignal(object)
    def __init__(self, base_config: ReportConfig, excel_folder: str, output_folder: str):
        super().__init__()
        self._base_config = base_config; self._excel_folder = excel_folder
        self._output_folder = output_folder
    @pyqtSlot()
    def run(self):
        runner = BatchRunner(base_config=self._base_config, excel_folder=self._excel_folder,
                             output_folder=self._output_folder,
                             progress_callback=self.progress.emit, log_callback=self.log.emit)
        self.finished.emit(runner.run())


# ── Dialogs (unchanged) ───────────────────────────────────────────────────────
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
        self.search = QLineEdit(); self.search.setPlaceholderText("Search title, category, description…")
        self.search.textChanged.connect(self._load)
        self.cmb_cat = QComboBox(); self.cmb_cat.addItem("All Categories")
        for c in self.db.get_categories(): self.cmb_cat.addItem(c)
        self.cmb_cat.currentTextChanged.connect(lambda _: self._load(self.search.text()))
        h.addWidget(self.search, stretch=1); h.addWidget(self.cmb_cat); v.addLayout(h)
        self.lst = QListWidget()
        self.lst.itemDoubleClicked.connect(self._accept)
        self.lst.currentItemChanged.connect(self._on_select)
        v.addWidget(self.lst, stretch=1)
        grp = QGroupBox("Preview"); pv = QVBoxLayout(grp)
        self.preview = QPlainTextEdit(); self.preview.setReadOnly(True); self.preview.setFixedHeight(110)
        pv.addWidget(self.preview); v.addWidget(grp)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self._accept); btns.rejected.connect(self.reject); v.addWidget(btns)

    def _load(self, query: str):
        cat  = self.cmb_cat.currentText()
        rows = self.db.search_observations(query)
        if cat != "All Categories":
            rows = [r for r in rows if r["category"] == cat]
        self.lst.clear(); self._rows = rows
        for r in rows:
            item = QListWidgetItem(f"[{r['severity']:8s}]  {r['title']}  —  {r['category']}")
            item.setForeground(QColor(SEV_COLORS.get(r["severity"], "#718096")))
            item.setData(Qt.ItemDataRole.UserRole, r["id"]); self.lst.addItem(item)

    def _on_select(self, cur, _):
        if not cur: return
        row = self.db.get_observation_by_id(cur.data(Qt.ItemDataRole.UserRole))
        if row:
            self.preview.setPlainText(
                f"Description:\n{row['description'][:250]}\n\n"
                f"Impact:\n{row['impact'][:200]}\n\nCVE: {row['cve']}")

    def _accept(self):
        cur = self.lst.currentItem()
        if not cur:
            QMessageBox.warning(self, "Select", "Please select an observation."); return
        self.selected = self.db.get_observation_by_id(cur.data(Qt.ItemDataRole.UserRole))
        self.accept()


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
        f.addRow("Name:", self.e_name); f.addRow("Designation:", self.e_desig); f.addRow("Email:", self.e_email)
        self.e_qual = QLineEdit(); self.e_cert_in = QComboBox(); self.e_cert_in.addItems(["No", "Yes"])
        f.addRow("Qualifications:", self.e_qual); f.addRow("CERT-In Listed:", self.e_cert_in)
        btn = QPushButton("Add Employee"); btn.clicked.connect(self._add); f.addRow("", btn); v.addWidget(grp)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(self.accept); v.addWidget(btns)

    def _load(self):
        self.tbl.setRowCount(0)
        for emp in self.db.get_employees():
            r = self.tbl.rowCount(); self.tbl.insertRow(r)
            for c, k in enumerate(["name","designation","email","department","qualifications","cert_in_listed"]):
                self.tbl.setItem(r, c, QTableWidgetItem(emp.get(k,"")))

    def _add(self):
        name  = self.e_name.text().strip(); desig = self.e_desig.text().strip()
        if not name or not desig:
            QMessageBox.warning(self,"Required","Name and Designation are required."); return
        self.db.add_employee(name, desig, self.e_email.text().strip(), self.e_dept.text().strip(),
                             self.e_qual.text().strip(), self.e_cert_in.currentText())
        self.e_name.clear(); self.e_desig.clear(); self.e_email.clear(); self.e_dept.clear(); self.e_qual.clear()
        self._load()


# ═══ Page1General (unchanged) ═════════════════════════════════════════════════
class Page1General(QWidget):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(10)
        lbl = QLabel("Personnel · Tools · Document"); lbl.setObjectName("page_title")
        sub = QLabel("Who ran the audit, what tools were used, and document metadata.")
        sub.setObjectName("page_sub"); sub.setWordWrap(True)
        v.addWidget(lbl); v.addWidget(sub); v.addWidget(_divider())

        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        inner = QWidget(); iv = QVBoxLayout(inner); iv.setSpacing(12); iv.setContentsMargins(4,4,8,4)
        inner.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.MinimumExpanding)

        grp1 = QGroupBox("Personnel"); f1 = QFormLayout(grp1); f1.setSpacing(8)
        f1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        emp_names = [""] + [e["name"] for e in self.db.get_employees()]
        self.cmb_prepared = QComboBox(); self.cmb_prepared.addItems(emp_names)
        self.cmb_reviewed = QComboBox(); self.cmb_reviewed.addItems(emp_names)
        self.lbl_prep_desig = QLabel(""); self.lbl_prep_desig.setStyleSheet("color:#718096")
        self.lbl_rev_desig  = QLabel(""); self.lbl_rev_desig.setStyleSheet("color:#718096")
        self.cmb_prepared.currentTextChanged.connect(self._upd_prep)
        self.cmb_reviewed.currentTextChanged.connect(self._upd_rev)
        h_prep = QHBoxLayout(); h_prep.setSpacing(8); h_prep.addWidget(self.cmb_prepared, stretch=1); h_prep.addWidget(self.lbl_prep_desig)
        h_rev = QHBoxLayout(); h_rev.setSpacing(8); h_rev.addWidget(self.cmb_reviewed, stretch=1); h_rev.addWidget(self.lbl_rev_desig)
        btn_emp = QPushButton("Manage Employees →"); btn_emp.clicked.connect(self._manage_emp)
        f1.addRow("Prepared By:", h_prep); f1.addRow("Reviewed By:", h_rev); f1.addRow("", btn_emp)
        self.cmb_approved = QComboBox(); self.cmb_approved.addItems(emp_names)
        self.lbl_appr_desig = QLabel(""); self.lbl_appr_desig.setStyleSheet("color:#718096")
        self.cmb_approved.currentTextChanged.connect(self._upd_appr)
        h_appr = QHBoxLayout(); h_appr.setSpacing(8); h_appr.addWidget(self.cmb_approved, stretch=1); h_appr.addWidget(self.lbl_appr_desig)
        f1.addRow("Approved By:", h_appr)
        self.cmb_released = QComboBox(); self.cmb_released.addItems(emp_names)
        self.lbl_rel_desig = QLabel(""); self.lbl_rel_desig.setStyleSheet("color:#718096")
        self.cmb_released.currentTextChanged.connect(self._upd_rel)
        h_rel = QHBoxLayout(); h_rel.setSpacing(8); h_rel.addWidget(self.cmb_released, stretch=1); h_rel.addWidget(self.lbl_rel_desig)
        f1.addRow("Released By:", h_rel)
        self.in_release_date = QDateEdit(); self.in_release_date.setCalendarPopup(True)
        self.in_release_date.setDate(QDate.currentDate()); self.in_release_date.setDisplayFormat("dd-MM-yyyy")
        f1.addRow("Release Date:", self.in_release_date)

        grp_team = QGroupBox("Auditing Team"); vt = QVBoxLayout(grp_team); vt.setSpacing(8)
        lbl_hint = QLabel("Select employees who performed this audit:"); lbl_hint.setStyleSheet("color:#718096;font-size:12px")
        vt.addWidget(lbl_hint)
        self.team_list = QTableWidget(0, 5)
        self.team_list.setHorizontalHeaderLabels(["Name","Designation","Email","Qualifications","CERT-In"])
        th = self.team_list.horizontalHeader()
        th.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        th.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        th.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.team_list.setFixedHeight(150); vt.addWidget(self.team_list)
        add_team_h = QHBoxLayout(); add_team_h.setSpacing(6)
        self.cmb_add_member = QComboBox(); self.cmb_add_member.addItems(emp_names)
        btn_add_member = QPushButton("Add to Team"); btn_add_member.setObjectName("btn_add_obs")
        btn_add_member.clicked.connect(self._add_team_member)
        add_team_h.addWidget(self.cmb_add_member, stretch=1); add_team_h.addWidget(btn_add_member); vt.addLayout(add_team_h)

        iv.addWidget(grp1); iv.addWidget(grp_team)

        grp2 = QGroupBox("Document Details"); f2 = QFormLayout(grp2); f2.setSpacing(8)
        f2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.in_version = QLineEdit("1.0"); self.in_version.setPlaceholderText("e.g. 1.0  /  2.0  /  1.1")
        f2.addRow("Document Version:", self.in_version)
        self.cmb_status = QComboBox(); self.cmb_status.addItems(["Draft", "Final"])
        f2.addRow("Document Status:", self.cmb_status); iv.addWidget(grp2)

        grp3 = QGroupBox("Client History / Background"); v3 = QVBoxLayout(grp3)
        self.txt_history = QTextEdit()
        self.txt_history.setPlaceholderText("Brief background about the client organisation, their industry, and purpose of this assessment…")
        self.txt_history.setFixedHeight(110); v3.addWidget(self.txt_history); iv.addWidget(grp3)

        grp4 = QGroupBox("Limitation / Constraints"); v4 = QVBoxLayout(grp4)
        self.txt_limitation = QTextEdit()
        self.txt_limitation.setPlaceholderText("Any limitations during the assessment, e.g. no source code access, restricted testing hours, out-of-scope systems…")
        self.txt_limitation.setFixedHeight(90); v4.addWidget(self.txt_limitation); iv.addWidget(grp4)

        grp5 = QGroupBox("Tools / Software Used"); v5 = QVBoxLayout(grp5); v5.setSpacing(8)
        filter_h = QHBoxLayout()
        self.cmb_tool_cat = QComboBox(); self.cmb_tool_cat.addItem("All Categories")
        for c in self.db.get_tool_categories(): self.cmb_tool_cat.addItem(c)
        self.cmb_tool_cat.currentTextChanged.connect(self._filter_tools)
        filter_h.addWidget(QLabel("Filter:")); filter_h.addWidget(self.cmb_tool_cat); filter_h.addStretch(); v5.addLayout(filter_h)
        self.tools_tbl = QTableWidget(0, 6)
        self.tools_tbl.setHorizontalHeaderLabels(["✓", "Tool Name", "Version", "Type", "Category", ""])
        hdr = self.tools_tbl.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in (2,3,4,5): hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self.tools_tbl.setMinimumHeight(180); self.tools_tbl.setMaximumHeight(280); v5.addWidget(self.tools_tbl)
        add_h = QHBoxLayout(); add_h.setSpacing(6)
        self.in_tool_name = QLineEdit(); self.in_tool_name.setPlaceholderText("Tool name")
        self.in_tool_ver  = QLineEdit(); self.in_tool_ver.setPlaceholderText("Version")
        self.cmb_tool_type = QComboBox(); self.cmb_tool_type.addItems(["Open Source","Licensed"])
        self.cmb_tool_newcat = QComboBox(); self.cmb_tool_newcat.addItems(["Web","API","Mobile","Source Code","Red Team","Internal","General"])
        btn_add_tool = QPushButton("Add"); btn_add_tool.setObjectName("btn_add_obs"); btn_add_tool.setFixedWidth(60)
        btn_add_tool.clicked.connect(self._add_custom_tool)
        add_h.addWidget(self.in_tool_name, stretch=2); add_h.addWidget(self.in_tool_ver, stretch=1)
        add_h.addWidget(self.cmb_tool_type); add_h.addWidget(self.cmb_tool_newcat); add_h.addWidget(btn_add_tool)
        v5.addLayout(add_h); iv.addWidget(grp5); self._filter_tools("All Categories")
        iv.addStretch(); scroll.setWidget(inner); v.addWidget(scroll, stretch=1)

    def _emp_by_name(self, name): return next((e for e in self.db.get_employees() if e["name"] == name), {})
    def _upd_prep(self, name): self.lbl_prep_desig.setText(self._emp_by_name(name).get("designation",""))
    def _upd_rev(self, name): self.lbl_rev_desig.setText(self._emp_by_name(name).get("designation",""))
    def _upd_appr(self, name): self.lbl_appr_desig.setText(self._emp_by_name(name).get("designation",""))
    def _upd_rel(self, name): self.lbl_rel_desig.setText(self._emp_by_name(name).get("designation",""))

    def _manage_emp(self):
        existing_names = {e["name"] for e in self.db.get_employees()}
        dlg = EmployeeDialog(self.db, self); dlg.exec()
        all_emps = self.db.get_employees(); names = [""] + [e["name"] for e in all_emps]
        for cmb, cur in [(self.cmb_prepared, self.cmb_prepared.currentText()),
                         (self.cmb_reviewed, self.cmb_reviewed.currentText()),
                         (self.cmb_approved, self.cmb_approved.currentText()),
                         (self.cmb_released, self.cmb_released.currentText()),
                         (self.cmb_add_member, self.cmb_add_member.currentText())]:
            cmb.clear(); cmb.addItems(names)
            if cur in names: cmb.setCurrentText(cur)
        for emp in [e for e in all_emps if e["name"] not in existing_names]:
            self._add_team_member_from_emp(emp)

    def get_data(self):
        prep = self._emp_by_name(self.cmb_prepared.currentText())
        rev  = self._emp_by_name(self.cmb_reviewed.currentText())
        return {
        "prepared_by": self.cmb_prepared.currentText(),
        "prepared_by_designation": prep.get("designation",""),
        "reviewed_by": self.cmb_reviewed.currentText(),
        "reviewed_by_designation": rev.get("designation",""),
        "doc_version": self.in_version.text().strip() or "1.0",
        "doc_status": self.cmb_status.currentText(),
        "client_history": self.txt_history.toPlainText().strip(),
        "approved_by": self.cmb_approved.currentText(),
        "approved_by_designation": self._emp_by_name(self.cmb_approved.currentText()).get("designation",""),
        "released_by": self.cmb_released.currentText(),
        "released_by_designation": self._emp_by_name(self.cmb_released.currentText()).get("designation",""),
        "release_date": self.in_release_date.date().toString("dd-MM-yyyy"),
        "selected_tools": self.get_selected_tools(),
        "team_members": self.get_team_members(),
        "limitation": self.txt_limitation.toPlainText().strip().split('\n'),
        
       
}

    def set_data(self, d):
        for name, cmb in [(d.get("prepared_by",""), self.cmb_prepared), (d.get("reviewed_by",""), self.cmb_reviewed)]:
            if name and cmb.findText(name) >= 0: cmb.setCurrentText(name)
        self.in_version.setText(d.get("doc_version","1.0"))
        idx = self.cmb_status.findText(d.get("doc_status","Draft"));
        if idx >= 0: self.cmb_status.setCurrentIndex(idx)
        self.txt_history.setPlainText(d.get("client_history",""))
        self.txt_limitation.setPlainText(d.get("limitation",""))
        for name, cmb in [(d.get("approved_by",""), self.cmb_approved), (d.get("released_by",""), self.cmb_released)]:
            if name and cmb.findText(name) >= 0: cmb.setCurrentText(name)
        if d.get("release_date"):
            self.in_release_date.setDate(QDate.fromString(d["release_date"], "dd-MM-yyyy"))

    def _filter_tools(self, cat):
        tools = self.db.get_tools("" if cat == "All Categories" else cat)
        self.tools_tbl.setRowCount(0); self._tool_ids = []
        for t in tools:
            r = self.tools_tbl.rowCount(); self.tools_tbl.insertRow(r)
            chk = QCheckBox(); chk.setStyleSheet("margin-left:8px"); self.tools_tbl.setCellWidget(r, 0, chk)
            self.tools_tbl.setItem(r, 1, QTableWidgetItem(t["tool_name"]))
            self.tools_tbl.setItem(r, 2, QTableWidgetItem(t["tool_version"]))
            self.tools_tbl.setItem(r, 3, QTableWidgetItem(t["tool_type"]))
            self.tools_tbl.setItem(r, 4, QTableWidgetItem(t["category"]))
            btn = QPushButton("✕"); btn.setObjectName("btn_del")
            btn.clicked.connect(lambda _, tid=t["id"]: self._remove_tool(tid))
            self.tools_tbl.setCellWidget(r, 5, btn); self.tools_tbl.setRowHeight(r, 32)
            self._tool_ids.append(t["id"])

    def _remove_tool(self, tool_id):
        self.db.delete_tool(tool_id); self._filter_tools(self.cmb_tool_cat.currentText())

    def _add_custom_tool(self):
        name = self.in_tool_name.text().strip()
        if not name: return
        self.db.add_tool(name, self.in_tool_ver.text().strip(),
                         self.cmb_tool_type.currentText(), self.cmb_tool_newcat.currentText())
        self.in_tool_name.clear(); self.in_tool_ver.clear(); self._filter_tools(self.cmb_tool_cat.currentText())

    def get_selected_tools(self):
        selected = []
        for row in range(self.tools_tbl.rowCount()):
            chk = self.tools_tbl.cellWidget(row, 0)
            if chk and chk.isChecked():
                selected.append({
                    "tool_name": self.tools_tbl.item(row,1).text() if self.tools_tbl.item(row,1) else "",
                    "tool_version": self.tools_tbl.item(row,2).text() if self.tools_tbl.item(row,2) else "",
                    "tool_type": self.tools_tbl.item(row,3).text() if self.tools_tbl.item(row,3) else "",
                    "category": self.tools_tbl.item(row,4).text() if self.tools_tbl.item(row,4) else "",
                    "tool_id": self._tool_ids[row] if row < len(self._tool_ids) else 0,
                })
        return selected

    def _add_team_member(self):
        name = self.cmb_add_member.currentText().strip()
        if not name: return
        self._add_team_member_from_emp(self._emp_by_name(name))

    def _add_team_member_from_emp(self, emp):
        name = emp.get("name", "")
        for r in range(self.team_list.rowCount()):
            if self.team_list.item(r, 0) and self.team_list.item(r, 0).text() == name: return
        r = self.team_list.rowCount(); self.team_list.insertRow(r)
        self.team_list.setItem(r, 0, QTableWidgetItem(emp.get("name", "")))
        self.team_list.setItem(r, 1, QTableWidgetItem(emp.get("designation", "")))
        self.team_list.setItem(r, 2, QTableWidgetItem(emp.get("email", "")))
        self.team_list.setItem(r, 3, QTableWidgetItem(emp.get("qualifications", "")))
        self.team_list.setItem(r, 4, QTableWidgetItem(emp.get("cert_in_listed", "No")))
        self.team_list.setRowHeight(r, 32)

    def get_team_members(self):
        return [{"name": self._tbl_cell(r,0), "designation": self._tbl_cell(r,1),
                 "email": self._tbl_cell(r,2), "qualifications": self._tbl_cell(r,3),
                 "cert_in_listed": self._tbl_cell(r,4)} for r in range(self.team_list.rowCount())]

    def _tbl_cell(self, row, col):
        item = self.team_list.item(row, col); return item.text().strip() if item else ""


# ═══ ObsTable (unchanged) ═════════════════════════════════════════════════════
class ObsTable(QWidget):
    COLS = ["#","Title","Severity","Affected URL","CVE","Description","Impact","Recommendation",""]
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(6)
        h = QHBoxLayout()
        btn_add = QPushButton("＋  Add Row"); btn_add.setObjectName("btn_add_obs")
        btn_lib = QPushButton("⚡  From Library"); btn_lib.setObjectName("btn_lib")
        btn_add.clicked.connect(self._add_blank); btn_lib.clicked.connect(self._pick_lib)
        h.addWidget(btn_add); h.addWidget(btn_lib); h.addStretch(); v.addLayout(h)
        self.tbl = QTableWidget(0, len(self.COLS)); self.tbl.setHorizontalHeaderLabels(self.COLS)
        hdr = self.tbl.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        for c in [5,6,7]: hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(8, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.setMinimumHeight(400); self.tbl.setWordWrap(True)
        self.tbl.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v.addWidget(self.tbl)

    def _add_blank(self): self._add_row({})
    def _pick_lib(self):
        dlg = ObsLibraryDialog(self.db, self)
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.selected: self._add_row(dlg.selected)

    def _add_row(self, data):
        r = self.tbl.rowCount(); self.tbl.insertRow(r)
        sev_cmb = QComboBox(); sev_cmb.addItems(["Critical","High","Medium","Low","Info"])
        sev = data.get("severity","Medium")
        idx = sev_cmb.findText(sev, Qt.MatchFlag.MatchFixedString)
        if idx >= 0: sev_cmb.setCurrentIndex(idx)
        sev_cmb.currentTextChanged.connect(lambda t, cb=sev_cmb: cb.setStyleSheet(f"color:{SEV_COLORS.get(t,'#e2e8f0')}"))
        sev_cmb.setStyleSheet(f"color:{SEV_COLORS.get(sev,'#e2e8f0')}")
        self.tbl.setItem(r, 0, QTableWidgetItem(str(r+1)))
        self.tbl.setItem(r, 1, QTableWidgetItem(data.get("title","")))
        self.tbl.setCellWidget(r, 2, sev_cmb)
        self.tbl.setItem(r, 3, QTableWidgetItem(data.get("affected_url","")))
        self.tbl.setItem(r, 4, QTableWidgetItem(data.get("cve","")))
        for col, key in [(5,"description"),(6,"impact"),(7,"recommendation")]:
            te = QTextEdit(); te.setPlainText(data.get(key, "")); te.setMinimumHeight(75)
            te.setStyleSheet("background:#2d3748;color:#e2e8f0;border:none;font-size:12px;padding:4px;")
            self.tbl.setCellWidget(r, col, te)
        btn_del = QPushButton("✕"); btn_del.setObjectName("btn_del")
        btn_del.clicked.connect(lambda _, row=r: self._del(row))
        self.tbl.setCellWidget(r, 8, btn_del); self.tbl.setRowHeight(r, 85)

    def _del(self, row):
        self.tbl.removeRow(row)
        for i in range(self.tbl.rowCount()): self.tbl.setItem(i, 0, QTableWidgetItem(str(i+1)))

    def _cell(self, r, c):
        widget = self.tbl.cellWidget(r, c)
        if isinstance(widget, QTextEdit): return widget.toPlainText().strip()
        item = self.tbl.item(r, c); return item.text().strip() if item else ""

    def get_observations(self):
        obs = []
        for r in range(self.tbl.rowCount()):
            wgt = self.tbl.cellWidget(r, 2)
            obs.append({"sr_no": self._cell(r,0), "title": self._cell(r,1),
                        "severity": wgt.currentText() if wgt else "Medium",
                        "affected_url": self._cell(r,3), "cve": self._cell(r,4),
                        "description": self._cell(r,5), "impact": self._cell(r,6),
                        "recommendation": self._cell(r,7)})
        return obs
    
    def get_observations_for_export(self, report_type: str):
        """Get observations with column names mapped for the specific report type."""
        
        # Define affected column name based on report type
        AFFECTED_COLUMN_MAP = {
            'web': 'affected_url',
            'android': 'affected_apk',
            'ios': 'affected_ipa',
            'api': 'affected_endpoint',
            'red_team': 'attack_vector',
            'source_code': 'affected_path',
        }
        
        affected_field = AFFECTED_COLUMN_MAP.get(report_type, 'affected_url')
        
        obs = []
        for r in range(self.tbl.rowCount()):
            wgt = self.tbl.cellWidget(r, 2)
            
            # Build observation with correct affected field
            obs_dict = {
                "sr_no": self._cell(r, 0),
                "title": self._cell(r, 1),
                "severity": wgt.currentText() if wgt else "Medium",
                affected_field: self._cell(r, 3),  # Dynamic field name
                "cve": self._cell(r, 4),
                "description": self._cell(r, 5),
                "impact": self._cell(r, 6),
                "recommendation": self._cell(r, 7),
            }
            obs.append(obs_dict)
        
        return obs

    def set_observations(self, obs_list):
        self.tbl.setRowCount(0)
        for obs in obs_list: self._add_row(obs)


# ═══ ReportDetailsForm — Page2Report metadata (data-source removed) ═══════════
class ReportDetailsForm(QWidget):
    def __init__(self, db: DBManager, parent=None):
        super().__init__(parent); self.db = db; self._build()

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(10)
        lbl = QLabel("Client · Report · Output"); lbl.setObjectName("page_title")
        sub = QLabel("Client info, audit period, report type, contact, and output location.")
        sub.setObjectName("page_sub"); sub.setWordWrap(True)
        v.addWidget(lbl); v.addWidget(sub); v.addWidget(_divider())

        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget(); iv = QVBoxLayout(inner); iv.setSpacing(12); iv.setContentsMargins(4,4,8,4)

        grp1 = QGroupBox("Client & Application"); f1 = QFormLayout(grp1); f1.setSpacing(8)
        f1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.in_client = QLineEdit(); self.in_client.setPlaceholderText("e.g. Acme Bank Pvt. Ltd.")
        self.in_app = QLineEdit(); self.in_app.setPlaceholderText("e.g. Internet Banking Portal")
        self.in_app_type = QComboBox(); self.in_app_type.addItems(["External","Internal"])
        self.in_start = QDateEdit(); self.in_start.setCalendarPopup(True)
        self.in_start.setDate(QDate.currentDate()); self.in_start.setDisplayFormat("dd-MM-yyyy")
        self.in_end = QDateEdit(); self.in_end.setCalendarPopup(True)
        self.in_end.setDate(QDate.currentDate().addDays(14)); self.in_end.setDisplayFormat("dd-MM-yyyy")
        date_h = QHBoxLayout(); date_h.addWidget(self.in_start); date_h.addWidget(QLabel("to")); date_h.addWidget(self.in_end); date_h.addStretch()
        self.scope_input = QTextEdit()
        self.scope_input.setPlaceholderText(
        "Enter each scope item on a new line:\n"
        "https://example.com/login\n"
        "https://example.com/api/*\n"
        "10.0.0.0/24")
        self.scope_input.setMinimumHeight(100)
        self.in_method = QComboBox(); self.in_method.addItems(["Grey Box","Black Box","White Box"])
        f1.addRow("Client Name:", self.in_client); f1.addRow("App Name:", self.in_app)
        f1.addRow("App Type:", self.in_app_type); f1.addRow("Audit Period:", date_h)
        f1.addRow("Scope (URL/IP):", self.scope_input)
        iv.addWidget(grp1)

        grp2 = QGroupBox("Report Settings"); f2 = QFormLayout(grp2); f2.setSpacing(8)
        f2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.cmb_type = QComboBox(); self.cmb_type.addItems(["Web","Api","Android","ios","VA","CA","Sourcecode","DB","CA_Nessus"])
        self.cmb_env = QComboBox(); self.cmb_env.addItems(["Production","UAT"])
        f2.addRow("Report Type:", self.cmb_type); f2.addRow("Environment:", self.cmb_env); iv.addWidget(grp2)

        grp_contact = QGroupBox("Client Contact Person"); fc = QFormLayout(grp_contact); fc.setSpacing(8)
        fc.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.in_contact_name = QLineEdit(); self.in_contact_name.setPlaceholderText("e.g. John Smith")
        self.in_contact_desig = QLineEdit(); self.in_contact_desig.setPlaceholderText("e.g. IT Manager")
        self.in_contact_email = QLineEdit(); self.in_contact_email.setPlaceholderText("e.g. john@client.com")
        fc.addRow("Contact Person:", self.in_contact_name); fc.addRow("Designation:", self.in_contact_desig)
        fc.addRow("Email:", self.in_contact_email); iv.addWidget(grp_contact)

        grp4 = QGroupBox("Files & Output"); f4 = QFormLayout(grp4); f4.setSpacing(8)
        f4.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.pick_poc = FilePicker("Optional: POC screenshots folder", folder=True)
        self.pick_output = FilePicker("Leave blank for auto-named output", save=True)
        f4.addRow("POC Folder:", self.pick_poc); f4.addRow("Output File:", self.pick_output); iv.addWidget(grp4)

        iv.addStretch(); scroll.setWidget(inner); v.addWidget(scroll, stretch=1)

    def audit_period(self):
        return f"{self.in_start.date().toString('dd-MM-yyyy')} - {self.in_end.date().toString('dd-MM-yyyy')}"

    def get_data(self):
        scope_text = self.scope_input.toPlainText().strip()
        scope_list = [s.strip() for s in scope_text.split('\n') if s.strip()]
        
        return {
            "client_name": self.in_client.text().strip(),
            "app_name": self.in_app.text().strip(),
            "app_type": self.in_app_type.currentText(),
            "audit_period": self.audit_period(),
            "scope": scope_list,  # ← Changed to list
            "url": scope_text,
            "method": self.in_method.currentText(),
            "report_type": self.cmb_type.currentText(),
            "environment": self.cmb_env.currentText(),
            "poc_folder": self.pick_poc.text(),
            "output_file": self.pick_output.text(),
            "client_contact_person": self.in_contact_name.text().strip(),
            "client_designation": self.in_contact_desig.text().strip(),
            "client_email": self.in_contact_email.text().strip(),
        }

    def set_data(self, d):
        self.in_client.setText(d.get("client_name","")); self.in_app.setText(d.get("app_name",""))
        idx = self.in_app_type.findText(d.get("app_type","External"));
        if idx >= 0: self.in_app_type.setCurrentIndex(idx)
        self.scope_input.setPlainText(d.get("scope",""))
        idx = self.in_method.findText(d.get("method","Grey Box"));
        if idx >= 0: self.in_method.setCurrentIndex(idx)
        idx = self.cmb_type.findText(d.get("report_type","Web"));
        if idx >= 0: self.cmb_type.setCurrentIndex(idx)
        idx = self.cmb_env.findText(d.get("environment","Production"));
        if idx >= 0: self.cmb_env.setCurrentIndex(idx)
        self.pick_poc.setText(d.get("poc_folder","")); self.pick_output.setText(d.get("output_file",""))
        if d.get("start_date"): self.in_start.setDate(QDate.fromString(d["start_date"], "dd-MM-yyyy"))
        if d.get("end_date"): self.in_end.setDate(QDate.fromString(d["end_date"], "dd-MM-yyyy"))
        self.in_contact_name.setText(d.get("client_contact_person",""))
        self.in_contact_desig.setText(d.get("client_designation",""))
        self.in_contact_email.setText(d.get("client_email",""))
       

    def get_profile_dates(self):
        return {"start_date": self.in_start.date().toString("dd-MM-yyyy"),
                "end_date": self.in_end.date().toString("dd-MM-yyyy")}


# ── Section: General Info (two tabs) ──────────────────────────────────────────
class GeneralInfoSection(QWidget):
    def __init__(self, db: DBManager):
        super().__init__()
        self.page1 = Page1General(db)
        self.details = ReportDetailsForm(db)
        v = QVBoxLayout(self); v.setContentsMargins(22, 18, 22, 18); v.setSpacing(6)
        title = QLabel("General Info"); title.setObjectName("page_title")
        title.setStyleSheet("font-size:22px")
        v.addWidget(title)
        tabs = QTabWidget()
        tabs.addTab(self.details, "Client · Report · Output")
        tabs.addTab(self.page1, "Personnel · Tools · Document")
        v.addWidget(tabs, stretch=1)

    def get_data(self):
        d = self.page1.get_data(); d.update(self.details.get_data()); return d
    def set_data(self, d): self.page1.set_data(d); self.details.set_data(d)
    def get_profile_dates(self): return self.details.get_profile_dates()


# ── Shared generate output (progress + log) ───────────────────────────────────
class GenerateOutput(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(0,0,0,0); v.setSpacing(8)
        self.progress = QProgressBar(); self.progress.setValue(0); self.progress.setFormat("%p%")
        v.addWidget(self.progress)
        self.log = QPlainTextEdit(); self.log.setReadOnly(True)
        self.log.setPlaceholderText("Logs will appear here…"); self.log.setMinimumHeight(180)
        v.addWidget(self.log, stretch=1)
    def append_log(self, m): self.log.appendPlainText(m)
    def set_progress(self, v): self.progress.setValue(v)
    def reset(self): self.log.clear(); self.progress.setValue(0)


def _section_header(title, subtitle):
    box = QWidget(); v = QVBoxLayout(box); v.setContentsMargins(0,0,0,0); v.setSpacing(2)
    t = QLabel(title); t.setObjectName("page_title"); t.setStyleSheet("font-size:22px")
    s = QLabel(subtitle); s.setObjectName("page_sub"); s.setWordWrap(True)
    v.addWidget(t); v.addWidget(s)
    return box


# ── Section: Word from Excel ──────────────────────────────────────────────────
class WordFromExcelSection(QWidget):
    generate_requested = pyqtSignal()
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(22, 18, 22, 18); v.setSpacing(12)
        v.addWidget(_section_header("Word from Excel", "Generate the .docx from a findings workbook. General Info is used for cover/meta."))
        grp = QGroupBox("Excel Source"); pe = QFormLayout(grp); pe.setSpacing(8)
        pe.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.pick_excel = FilePicker("Select single findings .xlsx file")
        pe.addRow("Single File:", self.pick_excel)
        self.chk_batch = QCheckBox("Batch mode — process entire folder")
        self.chk_batch.setStyleSheet("color:#63b3ed;font-weight:500")
        self.chk_batch.toggled.connect(self._toggle_batch)
        pe.addRow("", self.chk_batch)
        self.pnl_batch = QWidget(); pb = QFormLayout(self.pnl_batch); pb.setSpacing(8)
        pb.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.pick_excel_folder = FilePicker("Folder containing .xlsx files", folder=True)
        self.pick_output_folder = FilePicker("Folder to save all reports", folder=True)
        pb.addRow("Excel Folder:", self.pick_excel_folder); pb.addRow("Output Folder:", self.pick_output_folder)
        self.pnl_batch.hide(); pe.addRow("", self.pnl_batch)
        v.addWidget(grp)
        self.btn_generate = QPushButton("⚡  Generate Report"); self.btn_generate.setObjectName("btn_generate")
        self.btn_generate.clicked.connect(self.generate_requested.emit)
        v.addWidget(self.btn_generate, alignment=Qt.AlignmentFlag.AlignLeft)
        self.output = GenerateOutput(); v.addWidget(self.output, stretch=1)

    def _toggle_batch(self, checked):
        self.pick_excel.setVisible(not checked); self.pnl_batch.setVisible(checked)

    def get_excel_data(self):
        return {"excel_file": self.pick_excel.text(), "batch_mode": self.chk_batch.isChecked(),
                "excel_folder": self.pick_excel_folder.text(), "output_folder": self.pick_output_folder.text()}


# ── Section: Manual Word ──────────────────────────────────────────────────────
class ManualWordSection(QWidget):
    generate_requested = pyqtSignal()
    def __init__(self, db: DBManager):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(22, 18, 22, 18); v.setSpacing(12)
        v.addWidget(_section_header("Manual Word", "Type observations directly, then generate. General Info is used for cover/meta."))
        self.obs_table = ObsTable(db); v.addWidget(self.obs_table, stretch=1)
        self.btn_generate = QPushButton("⚡  Generate Report"); self.btn_generate.setObjectName("btn_generate")
        self.btn_generate.clicked.connect(self.generate_requested.emit)
        v.addWidget(self.btn_generate, alignment=Qt.AlignmentFlag.AlignLeft)
        self.output = GenerateOutput(); v.addWidget(self.output, stretch=1)

    def get_observations(self): return self.obs_table.get_observations()

    def get_observations_for_export(self, report_type: str):
        """Get observations with column names mapped for the specific report type."""
        
        AFFECTED_COLUMN_MAP = {
            'web': 'affected_url',
            'android': 'affected_apk',
            'ios': 'affected_ipa',
            'api': 'affected_endpoint',
            'red_team': 'attack_vector',
            'source_code': 'affected_path',
            
        }
        
        affected_field = AFFECTED_COLUMN_MAP.get(report_type, 'affected_url')
        
        obs = []
        for r in range(self.obs_table.tbl.rowCount()):
            wgt = self.obs_table.tbl.cellWidget(r, 2)
            obs.append({
                "sr_no": self.obs_table._cell(r, 0),
                "title": self.obs_table._cell(r, 1),
                "severity": wgt.currentText() if wgt else "Medium",
                affected_field: self.obs_table._cell(r, 3),  # Dynamic field
                "cve": self.obs_table._cell(r, 4),
                "description": self.obs_table._cell(r, 5),
                "impact": self.obs_table._cell(r, 6),
                "recommendation": self.obs_table._cell(r, 7),
            })
        return obs


# ── Section: Generate Excel template ──────────────────────────────────────────
class GenerateExcelSection(QWidget):
    def __init__(self, get_obs_callback=None, get_scope_callback=None, get_limitation_callback=None):
        super().__init__()
        from src.excel_template import REPORT_TYPES
        self._types = REPORT_TYPES
        self.get_obs_callback = get_obs_callback  # Store callback
        self.get_scope_callback = get_scope_callback
        self.get_limitation_callback = get_limitation_callback
        
        v = QVBoxLayout(self)
        v.setContentsMargins(22, 18, 22, 18)
        v.setSpacing(12)
        
        v.addWidget(_section_header("Generate Excel", "Create a blank findings workbook for a given report type."))
        
        grp = QGroupBox("Template")
        f = QFormLayout(grp)
        f.setSpacing(8)
        f.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.cmb_type = QComboBox()
        for k, label in REPORT_TYPES.items():
            self.cmb_type.addItem(label, k)
        
        self.pick_out = FilePicker("Where to save the .xlsx", save=False)
        self.pick_out._save = True
        
        f.addRow("Report Type:", self.cmb_type)
        f.addRow("Save As:", self.pick_out)
        v.addWidget(grp)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        self.btn = QPushButton("📊  Generate Template")
        self.btn.setObjectName("primary")
        self.btn.clicked.connect(self._generate)
        btn_layout.addWidget(self.btn)
        
        self.btn_with_obs = QPushButton("📤  Generate with Observations")
        self.btn_with_obs.setObjectName("btn_save_profile")
        self.btn_with_obs.clicked.connect(self._generate_with_obs)
        btn_layout.addWidget(self.btn_with_obs)

        self.btn_export_poc = QPushButton("🖼️  Export with POCs")
        self.btn_export_poc.setObjectName("btn_lib")
        self.btn_export_poc.clicked.connect(self._export_with_pocs)
        btn_layout.addWidget(self.btn_export_poc)
        
        btn_layout.addStretch()
        v.addLayout(btn_layout)
        
        self.status = QLabel("")
        self.status.setObjectName("page_sub")
        v.addWidget(self.status)
        v.addStretch(1)

    def _generate(self):
        key = self.cmb_type.currentData()
        out = self.pick_out.text()
        if not out:
            out, _ = QFileDialog.getSaveFileName(self, "Save Template", f"Excel_Template_{key}.xlsx", "Excel (*.xlsx)")
            if not out: return
        try:
            from src.excel_template import ExcelTemplateGenerator
            path = ExcelTemplateGenerator().generate(key, out)
            self.status.setText(f"✅ Template created: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Failed", f"Could not create template:\n{e}")

    def _generate_with_obs(self):
        obs = self.get_obs_callback() if self.get_obs_callback else []
        
        if not obs:
            reply = QMessageBox.question(
                self,
                "No Observations",
                "No manual observations found. Export empty template?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        
        key = self.cmb_type.currentData()
        out = self.pick_out.text()
        if not out:
            out, _ = QFileDialog.getSaveFileName(
                self, "Save Template", f"Excel_Template_{key.upper()}.xlsx", "Excel (*.xlsx)"
            )
            if not out:
                return
        
        # Get scope and limitation
        parent = self.parent()
        while parent and not hasattr(parent, 'sec_general'):
            parent = parent.parent()
        
        scope_list = []
        limitation_list = []
        if parent:
            gen_data = parent.sec_general.get_data()
            scope_list = gen_data.get('scope', [])
            limitation_list = gen_data.get('limitation', [])
        
        try:
            from src.excel_template import ExcelTemplateGenerator
            gen = ExcelTemplateGenerator()
            gen.generate(key, out, obs, None, scope_list, limitation_list)  # ← Added scope and limitation
            self.status.setText(f"✅ Template created with {len(obs)} observations: {out}")
            QMessageBox.information(
                self,
                "Success",
                f"✅ Template exported with {len(obs)} observations!\n\nSaved to: {out}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Failed", f"Could not create template:\n{e}")


    def _export_with_pocs(self):
        """Export Excel with POC images embedded from existing folders."""
        
        # Get report type
        key = self.cmb_type.currentData()
        parent = self.parent()
        while parent and not hasattr(parent, 'sec_general'):
            parent = parent.parent()

        if parent:
            gen_data = parent.sec_general.get_data()
            scope_list = gen_data.get('scope', [])
            limitation_list = gen_data.get('limitation', [])
        else:
            scope_list = []
            limitation_list = []
        
       
        
        # Get Excel file path from user
        excel_file, _ = QFileDialog.getOpenFileName(
            self,
            "Select Existing Excel File",
            "",
            "Excel Files (*.xlsx)"
        )
        if not excel_file:
            return
        
        # Get POC folder path
        poc_folder = QFileDialog.getExistingDirectory(
            self,
            "Select POC Folder",
            os.path.dirname(excel_file)
        )
        if not poc_folder:
            return
        
        # Ask for save location
        out, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel with POCs",
            f"Excel_POC_{key.upper()}.xlsx",
            "Excel (*.xlsx)"
        )
        if not out:
            return
        
        try:
            # Read observations from existing Excel
            from src.excel_reader import ExcelReader
            reader = ExcelReader()
            reader.load(excel_file)
            observations = reader.read_observations()

            if not isinstance(observations, list):
                observations = list(observations)
            
            # Generate new Excel with POCs embedded
            from src.excel_template import ExcelTemplateGenerator
            gen = ExcelTemplateGenerator()
            gen.generate(key, out, observations, poc_folder, scope=scope_list, limitation=limitation_list )
            
            QMessageBox.information(
                self,
                "Success",
                f"✅ Excel exported with POCs!\n\nSaved to: {out}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Failed", f"Could not export:\n{e}")

# ── Section: Misc ─────────────────────────────────────────────────────────────
class MiscSection(QWidget):
    theme_changed = pyqtSignal(str)
    save_profile_requested = pyqtSignal()
    load_profile_requested = pyqtSignal()
    def __init__(self, db: DBManager):
        super().__init__(); self.db = db
        v = QVBoxLayout(self); v.setContentsMargins(22, 18, 22, 18); v.setSpacing(12)
        v.addWidget(_section_header("Misc", "Manage employees and the observation library, switch theme, save/load profiles."))
        grp = QGroupBox("Data"); g = QVBoxLayout(grp); g.setSpacing(8)
        b1 = QPushButton("👥  Manage Employees"); b1.clicked.connect(self._employees)
        b2 = QPushButton("📚  Observation Library"); b2.clicked.connect(self._library)
        g.addWidget(b1); g.addWidget(b2); v.addWidget(grp)
        grp2 = QGroupBox("Appearance & Profiles"); f = QFormLayout(grp2); f.setSpacing(8)
        self.cmb_theme = QComboBox()
        self.cmb_theme.currentTextChanged.connect(self.theme_changed.emit)
        f.addRow("Theme:", self.cmb_theme)
        ph = QHBoxLayout()
        bs = QPushButton("💾  Save Profile"); bs.setObjectName("btn_save_profile"); bs.clicked.connect(self.save_profile_requested.emit)
        bl = QPushButton("⬆  Load Profile"); bl.setObjectName("btn_load_profile"); bl.clicked.connect(self.load_profile_requested.emit)
        ph.addWidget(bs); ph.addWidget(bl); ph.addStretch()
        f.addRow("Profiles:", self._wrap(ph)); v.addWidget(grp2)
        v.addStretch(1)

    def _wrap(self, layout):
        w = QWidget(); w.setLayout(layout); return w
    def set_themes(self, names):
        self.cmb_theme.blockSignals(True); self.cmb_theme.clear(); self.cmb_theme.addItems(names); self.cmb_theme.blockSignals(False)
    def _employees(self): EmployeeDialog(self.db, self).exec()
    def _library(self): ObsLibraryDialog(self.db, self).exec()


# ── Section: Credit ───────────────────────────────────────────────────────────
class CreditSection(QWidget):
    def __init__(self):
        super().__init__(); self.setObjectName("creditPage")
        v = QVBoxLayout(self); v.setContentsMargins(40, 40, 40, 40)
        v.addStretch(1)
        t = QLabel(APP_NAME); t.setObjectName("creditTitle"); t.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub = QLabel(f"Security Audit Report Generator  ·  v{APP_VERSION}")
        sub.setObjectName("creditText"); sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body = QLabel("Built for streamlined penetration-test reporting.\n"
                      "Excel findings in, formatted Word reports out.")
        body.setObjectName("creditText"); body.setAlignment(Qt.AlignmentFlag.AlignCenter); body.setWordWrap(True)
        v.addWidget(t); v.addSpacing(8); v.addWidget(sub); v.addSpacing(18); v.addWidget(body)
        v.addStretch(2)


# ── Launcher widgets ──────────────────────────────────────────────────────────
class NavCard(QPushButton):
    def __init__(self, key, title, icon, subtitle):
        super().__init__(); self.key = key
        self.setObjectName("navCard"); self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setMinimumHeight(112); self.setMinimumWidth(330)
        lay = QVBoxLayout(self); lay.setContentsMargins(22, 18, 22, 18); lay.setSpacing(6)
        top = QHBoxLayout(); top.setSpacing(12)
        ic = QLabel(icon); ic.setObjectName("navCardIcon"); _click_through(ic)
        ttl = QLabel(title); ttl.setObjectName("navCardTitle"); _click_through(ttl)
        top.addWidget(ic); top.addWidget(ttl); top.addStretch(); lay.addLayout(top)
        sub = QLabel(subtitle); sub.setObjectName("navCardSub"); _click_through(sub); lay.addWidget(sub)
        self._sh = _shadow(self, blur=24, y=10, alpha=150)
        self._anim = QPropertyAnimation(self._sh, b"blurRadius")
        self._anim.setDuration(150); self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
    def enterEvent(self, e): self._lift(42); super().enterEvent(e)
    def leaveEvent(self, e): self._lift(24); super().leaveEvent(e)
    def _lift(self, to): self._anim.stop(); self._anim.setEndValue(to); self._anim.start()


class SidebarItem(QPushButton):
    def __init__(self, key, title, icon):
        super().__init__(f"   {icon}    {title}"); self.key = key
        self.setObjectName("sideItem"); self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor); self.setMinimumHeight(46)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)


class LauncherView(QWidget):
    section_selected = pyqtSignal(str)
    def __init__(self):
        super().__init__(); self.setObjectName("launcher")
        outer = QVBoxLayout(self); outer.setContentsMargins(60, 46, 60, 46); outer.addStretch(1)
        title = QLabel(APP_NAME); title.setObjectName("launchTitle"); title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub = QLabel("Security Audit Report Generator"); sub.setObjectName("launchSub"); sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(title); outer.addWidget(sub); outer.addSpacing(30)
        grid = QGridLayout(); grid.setSpacing(20); grid.setColumnStretch(0, 1); grid.setColumnStretch(1, 1)
        for i, (key, t, ic, s) in enumerate(SECTIONS):
            card = NavCard(key, t, ic, s)
            card.clicked.connect(lambda _=False, k=key: self.section_selected.emit(k))
            grid.addWidget(card, i // 2, i % 2)
        wrap = QWidget(); wrap.setObjectName("launchWrap"); wrap.setLayout(grid); wrap.setFixedWidth(720)
        hb = QHBoxLayout(); hb.addStretch(1); hb.addWidget(wrap); hb.addStretch(1); outer.addLayout(hb)
        outer.addStretch(2)


class MainView(QWidget):
    section_selected = pyqtSignal(str)
    home_requested = pyqtSignal()
    def __init__(self):
        super().__init__(); self.setObjectName("mainView")
        self._panels = {}; self._items = {}
        h = QHBoxLayout(self); h.setContentsMargins(18, 18, 18, 18); h.setSpacing(16)
        sidebar = QFrame(); sidebar.setObjectName("sidebar"); sidebar.setFixedWidth(238)
        sv = QVBoxLayout(sidebar); sv.setContentsMargins(14, 18, 14, 18); sv.setSpacing(8)
        brand = QLabel(APP_NAME); brand.setObjectName("brand"); brand.setWordWrap(True)
        sv.addWidget(brand); sv.addSpacing(6)
        home = QPushButton("   \u2302    Home"); home.setObjectName("sideHome"); home.setMinimumHeight(44)
        home.setCursor(Qt.CursorShape.PointingHandCursor); home.clicked.connect(self.home_requested.emit)
        sv.addWidget(home); sv.addSpacing(8)
        for key, t, ic, _s in SECTIONS:
            it = SidebarItem(key, t, ic)
            it.clicked.connect(lambda _=False, k=key: self.select(k))
            sv.addWidget(it); self._items[key] = it
        sv.addStretch(1)
        ver = QLabel(f"v{APP_VERSION}"); ver.setObjectName("sideVer"); sv.addWidget(ver)
        _shadow(sidebar, blur=30, y=12, alpha=150)
        self.stack = QStackedWidget(); self.stack.setObjectName("contentStack")
        h.addWidget(sidebar); h.addWidget(self.stack, stretch=1)

    def set_panel(self, key, widget):
        old = self._panels.get(key)
        if old is not None: self.stack.removeWidget(old); old.deleteLater()
        self.stack.addWidget(widget); self._panels[key] = widget

    def select(self, key):
        for k, it in self._items.items(): it.setChecked(k == key)
        w = self._panels.get(key)
        if w is not None: self.stack.setCurrentWidget(w)
        self.section_selected.emit(key)


# ═══ Main shell ═══════════════════════════════════════════════════════════════
class AppShell(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1040, 700); self.resize(1200, 780)
        self.db = DBManager()
        self._thread = self._worker = None
        self._anim_ref = None
        self._theme = "Dark"

        self.central = BackgroundHost(self); self.setCentralWidget(self.central)
        lay = QVBoxLayout(self.central); lay.setContentsMargins(0, 0, 0, 0)
        self.stack = QStackedWidget(); self.stack.setObjectName("rootStack"); lay.addWidget(self.stack)

        self.launcher = LauncherView()
        self.main = MainView()
        self.stack.addWidget(self.launcher); self.stack.addWidget(self.main)

        # build sections
        self.sec_general = GeneralInfoSection(self.db)
        self.sec_excel_tpl = GenerateExcelSection(get_obs_callback=self._get_manual_obs)        
        self.sec_word = WordFromExcelSection()
        self.sec_manual = ManualWordSection(self.db)
        self.sec_misc = MiscSection(self.db)
        self.sec_credit = CreditSection()

        self.main.set_panel("general_info", self.sec_general)
        self.main.set_panel("generate_excel", self.sec_excel_tpl)
        self.main.set_panel("word_from_excel", self.sec_word)
        self.main.set_panel("manual_word", self.sec_manual)
        self.main.set_panel("misc", self.sec_misc)
        self.main.set_panel("credit", self.sec_credit)

        # wiring
        self.launcher.section_selected.connect(self._open_section)
        self.main.home_requested.connect(self._go_home)
        self.sec_word.generate_requested.connect(self._start_word_from_excel)
        self.sec_manual.generate_requested.connect(self._start_manual)
        self.sec_misc.set_themes(list(THEMES.keys()))
        self.sec_misc.theme_changed.connect(self._apply_theme)
        self.sec_misc.save_profile_requested.connect(self._save_profile)
        self.sec_misc.load_profile_requested.connect(self._load_profile)

        self.setStatusBar(QStatusBar()); self._status("Ready")
        self._apply_theme(self._theme)
        self.set_background_image(BACKGROUND_IMAGE)

    # public API
    def set_background_image(self, path):
        pix = QPixmap(path) if path and Path(path).exists() else QPixmap()
        self.central.set_pixmap(pix)

    def _apply_theme(self, name):
        self._theme = name
        self.setStyleSheet(THEMES.get(name, THEMES["Dark"]) + SHELL_QSS)

    # navigation
    def _open_section(self, key):
        self.main.select(key); self._fade_to(1)
    def _go_home(self): self._fade_to(0)

    def _fade_to(self, index):
        self.stack.setCurrentIndex(index)
        w = self.stack.currentWidget()
        eff = QGraphicsOpacityEffect(w); w.setGraphicsEffect(eff)
        anim = QPropertyAnimation(eff, b"opacity"); anim.setDuration(220)
        anim.setStartValue(0.0); anim.setEndValue(1.0); anim.setEasingCurve(QEasingCurve.Type.InOutCubic)
        anim.finished.connect(lambda: w.setGraphicsEffect(None))
        anim.start(); self._anim_ref = anim

    def _status(self, msg): self.statusBar().showMessage(msg, 8000)

    # config assembly (mirrors the old _build_config)
    def _build_config(self, *, excel_file=None, manual_obs=None):
        gi = self.sec_general.get_data()
        cfg = ReportConfig(
            prepared_by=gi["prepared_by"], prepared_by_designation=gi["prepared_by_designation"],
            reviewed_by=gi["reviewed_by"], reviewed_by_designation=gi["reviewed_by_designation"],
            doc_version=gi["doc_version"], doc_status=gi["doc_status"],
            client_history=gi["client_history"], limitation=gi["limitation"],
            client_name=gi["client_name"], app_name=gi["app_name"], app_type=gi["app_type"],
            audit_period=gi["audit_period"], method=gi["method"],
            report_type=gi["report_type"], environment=gi["environment"],
            poc_folder=gi["poc_folder"], output_file=gi["output_file"],
            approved_by=gi["approved_by"], approved_by_designation=gi["approved_by_designation"],
            released_by=gi["released_by"], released_by_designation=gi["released_by_designation"],
            release_date=gi["release_date"], client_contact_person=gi["client_contact_person"],
            client_designation=gi["client_designation"], client_email=gi["client_email"],
            selected_tools=gi["selected_tools"], team_members=gi["team_members"],
            url=gi.get("url", ""),
            scope=gi.get("scope", []), 
        )
        if manual_obs is not None: cfg.manual_observations = manual_obs
        if excel_file is not None: cfg.excel_file = excel_file
        return cfg
    
    def _get_manual_obs(self):
        """Get observations from manual entry section with correct affected column."""
        if hasattr(self.sec_manual, 'get_observations_for_export'):
            report_type = self.sec_general.get_data().get("report_type", "web").lower()
            return self.sec_manual.get_observations_for_export(report_type)
        return []

    # generation
    def _start_word_from_excel(self):
        ex = self.sec_word.get_excel_data()
        if ex["batch_mode"]:
            if not ex["excel_folder"]:
                QMessageBox.warning(self, "Missing Input", "Please select an Excel folder for batch mode."); return
            if not ex["output_folder"]:
                QMessageBox.warning(self, "Missing Input", "Please select an output folder for batch mode."); return
        elif not ex["excel_file"]:
            QMessageBox.warning(self, "Missing Input", "Please select an Excel file (or enable batch mode)."); return
        cfg = self._build_config(excel_file=ex["excel_file"])
        self._run(cfg, self.sec_word, batch=ex["batch_mode"],
                  excel_folder=ex["excel_folder"], output_folder=ex["output_folder"])

    def _start_manual(self):
        obs = self.sec_manual.get_observations()
        if not obs:
            QMessageBox.warning(self, "No Observations", "Please add at least one observation."); return
        cfg = self._build_config(manual_obs=obs)
        self._run(cfg, self.sec_manual, batch=False)

    def _run(self, cfg, section, *, batch=False, excel_folder="", output_folder=""):
        section.btn_generate.setEnabled(False)
        section.output.reset()
        self._status("Generating report…")
        self._thread = QThread()
        if batch:
            self._worker = BatchWorker(cfg, excel_folder, output_folder)
            self._worker.moveToThread(self._thread)
            self._thread.started.connect(self._worker.run)
            self._worker.log.connect(section.output.append_log)
            self._worker.progress.connect(section.output.set_progress)
            self._worker.finished.connect(lambda r, s=section: self._on_batch_finished(r, s, output_folder))
        else:
            self._worker = GeneratorWorker(cfg)
            self._worker.moveToThread(self._thread)
            self._thread.started.connect(self._worker.run)
            self._worker.log.connect(section.output.append_log)
            self._worker.progress.connect(section.output.set_progress)
            self._worker.finished.connect(lambda r, s=section: self._on_finished(r, s))
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    @pyqtSlot()
    def _on_finished(self, result: ReportResult, section):
        section.btn_generate.setEnabled(True)
        if result.success:
            gi = self.sec_general.get_data()
            self.db.save_report_history(gi["client_name"], gi["app_name"], gi["report_type"],
                                        result.output_path, gi["prepared_by"])
            self._status(f"Done ✓  →  {result.output_path}")
            QMessageBox.information(self, "Report Generated",
                f"✅ Report created successfully!\n\n📄 Output:  {result.output_path}\n"
                f"📊 Observations:  {result.observations_count}\n\n⚠️  Please review before sharing.")
        else:
            self._status(f"Error: {result.error}")
            QMessageBox.critical(self, "Generation Failed", f"❌ Report generation failed:\n\n{result.error}")

    def _on_batch_finished(self, result: BatchResult, section, output_folder):
        section.btn_generate.setEnabled(True)
        if result.failed == 0:
            self._status(f"Batch done ✓ — {result.success} reports generated")
            QMessageBox.information(self, "Batch Complete",
                f"✅ All {result.success} reports generated successfully!\n\n📁 Saved to: {output_folder}")
        else:
            self._status(f"Batch done — {result.success} ok, {result.failed} failed")
            errors = "\n".join(f"✗ {f}: {e}" for f, e in result.errors)
            QMessageBox.warning(self, "Batch Complete with Errors",
                f"✅ Success: {result.success}\n❌ Failed:  {result.failed}\n\nErrors:\n{errors}")

    # profiles
    def _full_profile(self):
        gi = self.sec_general.get_data(); gi.update(self.sec_general.get_profile_dates())
        return {"general": gi}

    def _save_profile(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Profile", str(PROFILE_DIR), "JSON Profiles (*.json)")
        if not path: return
        if not path.endswith(".json"): path += ".json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self._full_profile(), f, indent=2)
        self._status(f"Profile saved: {path}")
        QMessageBox.information(self, "Saved", f"✅ Profile saved:\n{path}")

    def _load_profile(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Profile", str(PROFILE_DIR), "JSON Profiles (*.json)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f: data = json.load(f)
            if "general" in data: self.sec_general.set_data(data["general"])
            self._status(f"Profile loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Load Failed", f"Could not load profile:\n{e}")


# THEMES kept minimal here — paste your full THEMES dict back in to restore
# Light / Midnight Blue. Dark is the full one above.
THEMES = {"Dark": DARK}


def launch_gui():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME); app.setApplicationVersion(APP_VERSION)
    try:
        app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)
    except AttributeError:
        pass
    window = AppShell(); window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    launch_gui()
