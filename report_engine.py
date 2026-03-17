"""
report_engine.py
================
Centralized report generation engine.
All report-making logic lives here — shared by GUI, CLI, and any future interface.

Usage:
    from report_engine import ReportEngine, ReportConfig

    config = ReportConfig(
        client_name="Acme Bank Pvt. Ltd.",
        app_name="Internet Banking Portal",
        app_type="External Web Application",
        audit_period="01-01-2026 - 15-01-2026",
        url="https://banking.acme.com",
        method="Grey Box",
        report_type="Web",
        environment="Production",
        excel_file="findings.xlsx",
        poc_folder="poc_screenshots/",
        output_file="Draft_Report_Acme.docx",
    )

    engine = ReportEngine(config)
    engine.run(progress_callback=my_fn, log_callback=my_fn)
"""

from __future__ import annotations

import os
import re
import sys

if getattr(sys, 'frozen', False):
    _ENGINE_DIR = sys._MEIPASS        # inside the .exe bundle
else:
    _ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Inches, Pt, RGBColor
from docxtpl import DocxTemplate

# ---------------------------------------------------------------------------
# Local src imports (excel_reader, poc_finder live in ./src/)
# ---------------------------------------------------------------------------
_SRC = os.path.join(_ENGINE_DIR, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_reader import ExcelReader
from poc_finder import POCFinder


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ReportConfig:
    """All inputs needed to generate one report."""
    client_name: str = ""
    app_name: str = ""
    app_type: str = ""
    audit_period: str = ""
    url: str = ""
    method: str = ""
    report_type: str = "Web"          # Web | Api | Mobile
    environment: str = "Production"   # Production | Uat
    excel_file: str = ""
    poc_folder: str = ""
    output_file: str = ""
    template_base: str = ""  # folder that holds *_template.docx files

    # Derived / auto-filled — leave blank to auto-compute
    high_count: str = ""
    medium_count: str = ""
    low_count: str = ""
    total_count: str = ""

    def template_path(self) -> str:
        mapping = {
            "Web":    "web_template.docx",
            "Api":    "api_template.docx",
            "Mobile": "mobile_template.docx",
        }
        filename = mapping.get(self.report_type.capitalize(), "web_template.docx")
        base = self.template_base or os.path.join(_ENGINE_DIR, "templates")
        return os.path.join(base, filename)

    def resolved_output(self, safe_name: str) -> str:
        if self.output_file:
            base = os.path.basename(self.output_file)
            return base if base.endswith(".docx") else base + ".docx"
        stamp = datetime.now().strftime("%Y%m%d")
        return f"Draft_Report_{safe_name}_{stamp}.docx"


@dataclass
class ReportResult:
    """Returned by ReportEngine.run() — callers inspect this."""
    success: bool
    output_path: str = ""
    error: str = ""
    observations_count: int = 0


# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------

def _clean_text(text) -> str:
    """Escape XML-unsafe chars and strip control characters."""
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def _clean_observations(observations: list[dict]) -> list[dict]:
    return [
        {k: _clean_text(v) for k, v in obs.items()}
        for obs in observations
    ]


# ---------------------------------------------------------------------------
# Engine
# ---------------------------------------------------------------------------

class ReportEngine:
    """
    Single, reusable engine for generating security audit reports.

    Parameters
    ----------
    config : ReportConfig
    progress_callback : callable(int) | None
        Called with values 0-100 to report progress.
    log_callback : callable(str) | None
        Called with human-readable status strings.
    """

    SEVERITY_COLORS = {
        "high":     "EE0000",
        "critical": "EE0000",
        "medium":   "FFC000",
        "low":      "00B050",
    }

    def __init__(
        self,
        config: ReportConfig,
        progress_callback: Optional[Callable[[int], None]] = None,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        self.config = config
        self._progress = progress_callback or (lambda _: None)
        self._log = log_callback or print

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def run(self) -> ReportResult:
        try:
            return self._execute()
        except Exception as exc:
            self._log(f"❌ Fatal error: {exc}")
            return ReportResult(success=False, error=str(exc))

    # ------------------------------------------------------------------
    # Private pipeline
    # ------------------------------------------------------------------

    def _execute(self) -> ReportResult:
        cfg = self.config
        self._progress(0)

        # ── 1. Validate inputs ──────────────────────────────────────────
        self._log("🔍 Validating inputs...")
        if not cfg.excel_file or not os.path.exists(cfg.excel_file):
            raise FileNotFoundError(f"Excel file not found: {cfg.excel_file}")
        if not os.path.exists(cfg.template_path()):
            raise FileNotFoundError(f"Template not found: {cfg.template_path()}")
        self._progress(10)

        # ── 2. Load Excel ───────────────────────────────────────────────
        self._log("📊 Loading Excel file...")
        reader = ExcelReader()
        reader.load(cfg.excel_file)

        index_df      = reader.read_index_sheet()
        scope_df      = reader.read_scope_sheet()
        limitation_df = reader.read_limitation_sheet()
        observations  = reader.read_observations()
        self._progress(25)

        # ── 3. Clean data ───────────────────────────────────────────────
        self._log("🧹 Cleaning data...")
        observations = _clean_observations(observations)
        report_name  = reader.extract_report_name(index_df)
        safe_name    = re.sub(r'[\\/*?:"<>|]', "", report_name)
        stats        = reader.extract_summary_stats(index_df)
        self._progress(35)

        # ── 4. Resolve severity counts ─────────────────────────────────
        high   = cfg.high_count   or self._count_sev(observations, ["high", "critical"]) or stats["high"]
        medium = cfg.medium_count or self._count_sev(observations, ["medium"])           or stats["medium"]
        low    = cfg.low_count    or self._count_sev(observations, ["low"])              or stats["low"]
        total  = cfg.total_count  or str(len(observations))

        self._log(f"📊 Severities — High: {high}, Medium: {medium}, Low: {low}, Total: {total}")

        # ── 5. POC scanner ─────────────────────────────────────────────
        poc_finder = POCFinder()
        if cfg.poc_folder and os.path.exists(cfg.poc_folder):
            poc_finder.scan_folder(cfg.poc_folder)
            self._log(f"🖼️  POC folders found: {len(poc_finder.get_all_vulnerability_folders())}")
        elif cfg.poc_folder:
            self._log(f"⚠️  POC folder not found: {cfg.poc_folder}")
        self._progress(45)

        # ── 6. Build template context ──────────────────────────────────
        context = {
            "App_Name":    _clean_text(cfg.app_name),
            "type":        _clean_text(cfg.app_type),
            "Audit_Period":_clean_text(cfg.audit_period),
            "Client_Name": _clean_text(cfg.client_name),
            "environment": cfg.environment.capitalize(),
            "URL":         _clean_text(cfg.url),
            "high":        high,
            "medium":      medium,
            "low":         low,
            "total":       total,
            "method":      _clean_text(cfg.method),
            "report_name": report_name,
            "date_today":  datetime.today().strftime("%d %B %Y"),
            "observations":observations,
            "scope":       _clean_text(reader.get_cell_value(scope_df, 1, 1, "Full application scope")),
            "limitation":  _clean_text(reader.get_cell_value(limitation_df, 1, 0, "No limitations specified")),
        }

        # ── 7. Render docxtpl template ─────────────────────────────────
        self._log("📝 Rendering template...")
        tpl = DocxTemplate(cfg.template_path())
        tpl.render(context)
        temp_path = "_report_engine_temp.docx"
        tpl.save(temp_path)
        self._progress(60)

        # ── 8. Post-processing ─────────────────────────────────────────
        self._log("🎨 Post-processing: executive summary table...")
        doc = Document(temp_path)
        self._build_exec_table(doc, observations)
        self._remove_empty_rows(doc)
        self._progress(75)

        # ── 9. Insert POC images ───────────────────────────────────────
        if cfg.poc_folder and os.path.exists(cfg.poc_folder) and observations:
            self._log("🖼️  Inserting POC screenshots...")
            self._insert_pocs(doc, observations, poc_finder)
        self._progress(88)

        # ── 10. Save ───────────────────────────────────────────────────
        output_path = cfg.resolved_output(safe_name)
        doc.save(output_path)
        os.remove(temp_path)
        self._progress(100)

        self._log(f"✅ Report saved: {output_path}")
        self._log("⚠️  Please review the report before sharing.")

        return ReportResult(
            success=True,
            output_path=output_path,
            observations_count=len(observations),
        )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _count_sev(observations: list[dict], labels: list[str]) -> str:
        count = sum(
            1 for o in observations
            if o.get("severity", "").lower() in labels
        )
        return str(count) if count else ""

    def _build_exec_table(self, doc: Document, observations: list[dict]):
        """Find the 7-column executive summary table and populate it."""
        exec_table = next(
            (t for t in doc.tables if len(t.columns) == 7), None
        )
        if not exec_table:
            self._log("⚠️  Executive summary table (7 cols) not found — skipping.")
            return

        # Remove all data rows (keep header row 0)
        while len(exec_table.rows) > 1:
            exec_table._tbl.remove(exec_table.rows[1]._tr)

        for item in observations:
            cells = exec_table.add_row().cells
            values = [
                item.get("sr_no", ""),
                item.get("title", ""),
                item.get("severity", ""),
                item.get("affected_url", ""),
                item.get("cve", ""),
                item.get("recommendation", ""),
                "New Observation",
            ]
            for idx, val in enumerate(values):
                cell = cells[idx]
                for p in cell.paragraphs:
                    p.clear()
                para = cell.paragraphs[0]
                text = str(val) if pd.notna(val) else ""
                run = para.add_run(text)
                run.font.size = Pt(10)
                run.font.name = "Calibri"
                run.font.color.rgb = RGBColor(0, 0, 0)

                # Colour severity cell
                if idx == 2:
                    for sev, colour in self.SEVERITY_COLORS.items():
                        if sev in text.lower():
                            shading = parse_xml(
                                f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                                f'w:fill="{colour}"/>'
                            )
                            cell._tc.get_or_add_tcPr().append(shading)
                            break

                para.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if idx == 3
                    else WD_ALIGN_PARAGRAPH.JUSTIFY
                )

    @staticmethod
    def _remove_empty_rows(doc: Document):
        for table in doc.tables:
            empty = [
                i for i, row in enumerate(table.rows)
                if all(c.text.strip() == "" for c in row.cells)
            ]
            for i in reversed(empty):
                table._tbl.remove(table.rows[i]._tr)

    def _insert_pocs(
        self,
        doc: Document,
        observations: list[dict],
        poc_finder: POCFinder,
    ):
        marker = "<!-- POC will be inserted here during post-processing -->"
        obs_index = 0
        i = 0
        while i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            if marker in para.text and obs_index < len(observations):
                vuln_title = observations[obs_index].get("title", "").strip()
                para._element.getparent().remove(para._element)

                if vuln_title:
                    for img_path in poc_finder.get_pocs_by_vulnerability(vuln_title):
                        try:
                            p = doc.paragraphs[i].insert_paragraph_before()
                            p.add_run().add_picture(img_path)
                            i += 1
                        except Exception as e:
                            self._log(f"    ⚠️  Could not add image {img_path}: {e}")
                    self._log(f"    ✓ POCs added for: {vuln_title[:40]}")

                obs_index += 1
            else:
                i += 1