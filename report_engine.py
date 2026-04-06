"""
report_engine.py
================
Centralized report generation engine.
Shared by GUI, CLI, and any future interface.
"""

from __future__ import annotations

import os
import re
import sys

# ── MUST be first — anchor all paths to this file's location ────────────────
if getattr(sys, 'frozen', False):
    _ENGINE_DIR = sys._MEIPASS          # running as .exe
else:
    _ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── src/ imports anchored — works from any CWD or subfolder ─────────────────
_SRC = os.path.join(_ENGINE_DIR, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Optional

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Pt, RGBColor
from docxtpl import DocxTemplate

from excel_reader import ExcelReader
from poc_finder import POCFinder


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class ReportConfig:
    """
    All inputs needed to generate one report.

    Data source modes
    -----------------
    Excel mode  : set excel_file path, leave manual_observations empty
    Manual mode : set manual_observations list, leave excel_file empty
    """

    # ── Report / client info ─────────────────────────────────────────────────
    client_name:    str = ""
    app_name:       str = ""
    app_type:       str = ""
    audit_period:   str = ""
    url:            str = ""
    method:         str = ""
    report_type:    str = "Web"         # Web | Api | Mobile
    environment:    str = "Production"  # Production | Uat

    # ── Cover / front-page fields ────────────────────────────────────────────
    prepared_by:                str = ""
    prepared_by_designation:    str = ""
    reviewed_by:                str = ""
    reviewed_by_designation:    str = ""
    doc_version:                str = "1.0"
    client_history:             str = ""
    limitation:                 str = ""
    # Document approval
    approved_by:                str = ""
    approved_by_designation:    str = ""
    released_by:                str = ""
    released_by_designation:    str = ""
    release_date:               str = ""

    # Client contact
    client_contact_person:  str = ""
    client_designation:     str = ""
    client_email:           str = ""

    # Tools used — list of dicts with keys:
    # tool_name, tool_version, tool_type, category
    selected_tools: list = field(default_factory=list)
    team_members: list = field(default_factory=list)
    # keys per dict: name, designation, email, qualifications, cert_in_listed

    # ── Data source — only ONE should be set at a time ───────────────────────
    excel_file:             str  = ""
    manual_observations:    list = field(default_factory=list)
    # manual_observations keys per dict:
    #   sr_no, title, severity, description, impact,
    #   recommendation, affected_url, cve

    # ── Files ────────────────────────────────────────────────────────────────
    poc_folder:     str = ""
    output_file:    str = ""
    template_base:  str = ""   # blank = auto → _ENGINE_DIR/templates/

    # ── Optional count overrides (blank = auto-calculate) ────────────────────
    high_count:     str = ""
    medium_count:   str = ""
    low_count:      str = ""
    total_count:    str = ""



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
            p = os.path.basename(self.output_file)
            return p if p.endswith(".docx") else p + ".docx"
        stamp = datetime.now().strftime("%Y%m%d")
        return f"Draft_Report_{safe_name}_{stamp}.docx"

    @property
    def is_manual_mode(self) -> bool:
        """True when observations supplied directly — Excel not needed."""
        return bool(self.manual_observations) and not self.excel_file


@dataclass
class ReportResult:
    success:            bool
    output_path:        str = ""
    error:              str = ""
    observations_count: int = 0


# ── Text utilities ────────────────────────────────────────────────────────────

def _clean_text(text) -> str:
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def _clean_observations(observations: list[dict]) -> list[dict]:
    return [{k: _clean_text(v) for k, v in obs.items()} for obs in observations]


def _auto_number(observations: list[dict]) -> list[dict]:
    for i, obs in enumerate(observations, 1):
        if not obs.get("sr_no"):
            obs["sr_no"] = str(i)
    return observations


# ── Engine ────────────────────────────────────────────────────────────────────

class ReportEngine:
    """
    Single reusable engine. Import and call engine.run() from anywhere.

    Parameters
    ----------
    config            : ReportConfig
    progress_callback : callable(int 0-100)
    log_callback      : callable(str)
    """

    SEVERITY_COLORS = {
        "critical": "EE0000",
        "high":     "EE0000",
        "medium":   "FFC000",
        "low":      "00B050",
        "info":     "0070C0",
    }

    def __init__(
        self,
        config:            ReportConfig,
        progress_callback: Optional[Callable[[int], None]] = None,
        log_callback:      Optional[Callable[[str], None]] = None,
    ):
        self.config    = config
        self._progress = progress_callback or (lambda _: None)
        self._log      = log_callback or print

    def run(self) -> ReportResult:
        try:
            return self._execute()
        except Exception as exc:
            import traceback
            self._log(f"❌ Fatal error: {exc}")
            self._log(traceback.format_exc())
            return ReportResult(success=False, error=str(exc))

    def _execute(self) -> ReportResult:
        cfg = self.config
        self._progress(0)

        # ── 1. Validate ───────────────────────────────────────────────────────
        self._log("🔍 Validating inputs...")
        if cfg.is_manual_mode:
            if not cfg.manual_observations:
                raise ValueError("Manual mode: no observations provided.")
        else:
            if not cfg.excel_file or not os.path.exists(cfg.excel_file):
                raise FileNotFoundError(f"Excel file not found: {cfg.excel_file}")

        if not os.path.exists(cfg.template_path()):
            raise FileNotFoundError(f"Template not found: {cfg.template_path()}")
        self._progress(10)

        # ── 2. Load observations ──────────────────────────────────────────────
        scope_text      = ""
        limitation_text = cfg.limitation or "No limitations specified"
        report_name     = cfg.app_name or "Report"
        stats           = {"high": "0", "medium": "0", "low": "0"}

        if cfg.is_manual_mode:
            self._log("📋 Using manual observations...")
            observations = list(cfg.manual_observations)
            safe_name    = re.sub(r'[\\/*?:"<>|]', "", cfg.app_name or "Report")

        else:
            self._log("📊 Loading Excel file...")
            reader        = ExcelReader()
            reader.load(cfg.excel_file)
            index_df      = reader.read_index_sheet()
            scope_df      = reader.read_scope_sheet()
            limitation_df = reader.read_limitation_sheet()
            observations  = reader.read_observations()
            report_name   = reader.extract_report_name(index_df)
            safe_name     = re.sub(r'[\\/*?:"<>|]', "", report_name)
            stats         = reader.extract_summary_stats(index_df)
            scope_text    = _clean_text(
                reader.get_cell_value(scope_df, 1, 1, "Full application scope")
            )
            if not cfg.limitation:
                limitation_text = _clean_text(
                    reader.get_cell_value(limitation_df, 1, 0, "No limitations specified")
                )

        self._progress(25)

        # ── 3. Clean + number ─────────────────────────────────────────────────
        self._log("🧹 Cleaning data...")
        observations = _clean_observations(observations)
        observations = _auto_number(observations)
        self._progress(35)

        # ── 4. Severity counts ────────────────────────────────────────────────
        high   = cfg.high_count   or self._count_sev(observations, ["high", "critical"]) or stats.get("high",   "0")
        medium = cfg.medium_count or self._count_sev(observations, ["medium"])           or stats.get("medium", "0")
        low    = cfg.low_count    or self._count_sev(observations, ["low"])              or stats.get("low",    "0")
        total  = cfg.total_count  or str(len(observations))
        self._log(f"📊 High: {high}  Medium: {medium}  Low: {low}  Total: {total}")

        # ── 5. POC scanner ────────────────────────────────────────────────────
        poc_finder = POCFinder()
        if cfg.poc_folder and os.path.exists(cfg.poc_folder):
            poc_finder.scan_folder(cfg.poc_folder)
            self._log(f"🖼️  POC folders: {len(poc_finder.get_all_vulnerability_folders())}")
        elif cfg.poc_folder:
            self._log(f"⚠️  POC folder not found: {cfg.poc_folder}")
        self._progress(45)

        # ── 6. Build context ──────────────────────────────────────────────────
        context = {
            "App_Name":     _clean_text(cfg.app_name),
            "type":         _clean_text(cfg.app_type),
            "Audit_Period": _clean_text(cfg.audit_period),
            "Client_Name":  _clean_text(cfg.client_name),
            "environment":  cfg.environment.capitalize(),
            "URL":          _clean_text(cfg.url),
            "method":       _clean_text(cfg.method),
            "report_name":  report_name,
            "date_today":   datetime.today().strftime("%d %B %Y"),
            "observations": observations,
            "scope":        scope_text,
            "limitation":   _clean_text(limitation_text),
            "high":         high,
            "medium":       medium,
            "low":          low,
            "total":        total,
            # Cover fields — use {{ prepared_by }}, {{ doc_version }} etc. in .docx template
            "prepared_by":              _clean_text(cfg.prepared_by),
            "prepared_by_designation":  _clean_text(cfg.prepared_by_designation),
            "reviewed_by":              _clean_text(cfg.reviewed_by),
            "reviewed_by_designation":  _clean_text(cfg.reviewed_by_designation),
            "doc_version":              _clean_text(cfg.doc_version),
            "client_history":           _clean_text(cfg.client_history),

            "approved_by":              _clean_text(cfg.approved_by),
            "approved_by_designation":  _clean_text(cfg.approved_by_designation),
            "released_by":              _clean_text(cfg.released_by),
            "released_by_designation":  _clean_text(cfg.released_by_designation),
            "release_date":             _clean_text(cfg.release_date),
            "client_contact_person":    _clean_text(cfg.client_contact_person),
            "client_designation":       _clean_text(cfg.client_designation),
            "client_email":             _clean_text(cfg.client_email),
           
        }
        self._progress(55)

        # ── 7. Render template ────────────────────────────────────────────────
        self._log("📝 Rendering template...")
        tpl       = DocxTemplate(cfg.template_path())
        tpl.render(context)
        temp_path = os.path.join(_ENGINE_DIR, "_render_temp.docx")
        tpl.save(temp_path)
        self._progress(65)

        # ── 8. Post-process ───────────────────────────────────────────────────
        self._log("🎨 Building executive summary table...")
        doc = Document(temp_path)
        self._build_exec_table(doc, observations)

        # Build auditing team table (6-column)
        self._log(f"DEBUG — team_members count: {len(cfg.team_members)}")
        self._log(f"DEBUG — selected_tools count: {len(cfg.selected_tools)}")

        if cfg.team_members:
            self._log("👥 Building auditing team table...")
            self._build_team_table(doc, cfg.team_members)
        else:
            self._log("⚠️  No team members received — skipping team table.")

        if cfg.selected_tools:
            self._log("🔧 Building tools table...")
            self._build_tools_table(doc, cfg.selected_tools)
        else:
            self._log("⚠️  No tools received — skipping tools table.")

        # ── 9. POC images ─────────────────────────────────────────────────────
        if cfg.poc_folder and os.path.exists(cfg.poc_folder) and observations:
            self._log("🖼️  Inserting POC screenshots...")
            self._insert_pocs(doc, observations, poc_finder)
        self._progress(90)

        # ── 10. Save ──────────────────────────────────────────────────────────
        output_path = cfg.resolved_output(safe_name)
        doc.save(output_path)
        try:
            os.remove(temp_path)
        except OSError:
            pass
        self._progress(100)

        self._log(f"✅ Report saved: {output_path}")
        self._log("⚠️  Please review the report carefully before sharing.")

        return ReportResult(
            success=True,
            output_path=output_path,
            observations_count=len(observations),
        )

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _count_sev(observations: list[dict], labels: list[str]) -> str:
        n = sum(1 for o in observations if o.get("severity", "").lower() in labels)
        return str(n) if n else ""

    def _build_exec_table(self, doc: Document, observations: list[dict]):
        exec_table = next((t for t in doc.tables if len(t.columns) == 7), None)
        if not exec_table:
            self._log("⚠️  7-column exec summary table not found — skipping.")
            return
        while len(exec_table.rows) > 1:
            exec_table._tbl.remove(exec_table.rows[1]._tr)
        for item in observations:
            cells  = exec_table.add_row().cells
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
                run  = para.add_run(text)
                run.font.size      = Pt(10)
                run.font.name      = "Calibri"
                run.font.color.rgb = RGBColor(0, 0, 0)
                if idx == 2:
                    for sev, colour in self.SEVERITY_COLORS.items():
                        if sev in text.lower():
                            shading = parse_xml(
                                f'<w:shd xmlns:w="http://schemas.openxmlformats.org/'
                                f'wordprocessingml/2006/main" w:fill="{colour}"/>'
                            )
                            cell._tc.get_or_add_tcPr().append(shading)
                            break
                para.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if idx == 3 else WD_ALIGN_PARAGRAPH.JUSTIFY
                )
    def _build_team_table(self, doc: Document, team_members: list[dict]):
        """
        Find the 6-column auditing team table and populate it.
        Columns: Sr.No | Name | Designation | Email | Qualifications | CERT-In Listed
        """
        team_table = next(
            (t for t in doc.tables
             if len(t.columns) == 6
             and "qualification" in t.rows[0].cells[4].text.strip().lower()),
            None
        )
        if not team_table:
                self._log("⚠️  Audit team table not found — skipping.")
                # Print all tables found for diagnosis
                for i, t in enumerate(doc.tables):
                    h = " | ".join(c.text.strip()[:15] for c in t.rows[0].cells)
                    self._log(f"    Table {i}: {len(t.columns)} cols → {h}")
                return
        self._log(f"✓ Audit team table found — adding {len(team_members)} members")

        # Clear all rows except header
        while len(team_table.rows) > 1:
            team_table._tbl.remove(team_table.rows[1]._tr)

        for i, member in enumerate(team_members, 1):
            cells  = team_table.add_row().cells
            values = [
                str(i),
                member.get("name", ""),
                member.get("designation", ""),
                member.get("email", ""),
                member.get("qualifications", ""),
                member.get("cert_in_listed", "No"),
            ]
            for idx, val in enumerate(values):
                cell = cells[idx]
                for p in cell.paragraphs:
                    p.clear()
                para = cell.paragraphs[0]
                run  = para.add_run(str(val))
                run.font.size      = Pt(10)
                run.font.name      = "Calibri"
                run.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment     = WD_ALIGN_PARAGRAPH.CENTER


    def _build_tools_table(self, doc: Document, tools: list[dict]):
        """
        Find the 4-column tools table and populate it.
        Columns: Sr.No | Tool Name | Version | Open Source / Licensed
        """
        # Find 4-column table that is NOT the exec summary (which is 7 cols)
        tools_table = next(
            (t for t in doc.tables
             if len(t.columns) == 4
             and "name of tool" in t.rows[0].cells[1].text.strip().lower()),
            None
        )
        if not tools_table:
            self._log("⚠️  Tools table not found — skipping.")
            return

        # Clear all rows except header
        while len(tools_table.rows) > 1:
            tools_table._tbl.remove(tools_table.rows[1]._tr)

        for i, tool in enumerate(tools, 1):
            cells  = tools_table.add_row().cells
            values = [
                str(i),
                tool.get("tool_name", ""),
                tool.get("tool_version", ""),
                tool.get("tool_type", ""),
            ]
            for idx, val in enumerate(values):
                cell = cells[idx]
                for p in cell.paragraphs:
                    p.clear()
                para = cell.paragraphs[0]
                run  = para.add_run(str(val))
                run.font.size      = Pt(10)
                run.font.name      = "Calibri"
                run.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER if idx in [0, 2, 3]
                    else WD_ALIGN_PARAGRAPH.LEFT
                )

    @staticmethod
    def _remove_empty_rows(doc: Document):
        for table in doc.tables:
            empty = [i for i, row in enumerate(table.rows)
                     if all(c.text.strip() == "" for c in row.cells)]
            for i in reversed(empty):
                table._tbl.remove(table.rows[i]._tr)

    def _insert_pocs(self, doc: Document, observations: list[dict],
                     poc_finder: POCFinder):
        marker    = "<!-- POC will be inserted here during post-processing -->"
        obs_index = 0
        i         = 0
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
                            self._log(f"    ⚠️  Could not add image: {e}")
                    self._log(f"    ✓ POCs added: {vuln_title[:40]}")
                obs_index += 1
            else:
                i += 1
