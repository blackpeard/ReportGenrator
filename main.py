"""
main.py
=======
CLI wrapper around ReportEngine.
All report logic lives in report_engine.py — this file only handles
argument parsing and user input, then delegates to the engine.

Usage:
    python main.py -s findings.xlsx -p poc_screenshots/ -t web -env production

GUI:
    python gui_app.py
"""

import argparse
import os
import sys

# Support GUI env-var injection (legacy FROM_GUI mode)
FROM_GUI = os.environ.get("FROM_GUI") == "True"


def _get_user_input(field_name: str, default: str = "") -> str:
    val = input(f"  ✏️  Enter {field_name} [{default}]: ").strip()
    return val if val else default


def main():
    parser = argparse.ArgumentParser(description="Security Audit Report Generator (CLI)")

    parser.add_argument("-t", "--type",
                        choices=["web", "api", "mobile"], default="web")
    parser.add_argument("-env", "--environment",
                        choices=["production", "uat"], default="production")
    parser.add_argument("-s", "--sheet", required=True,
                        help="Input Excel file (.xlsx)")
    parser.add_argument("-p", "--poc",
                        help="POC folder path (optional)")
    parser.add_argument("-o", "--output",
                        help="Output filename (optional)")

    args = parser.parse_args()

    # ── Collect metadata (GUI env-vars override interactive prompts) ────
    if FROM_GUI:
        client_name  = os.environ.get("CLIENT_NAME", "")
        app_name     = os.environ.get("APP_NAME", "")
        app_type     = os.environ.get("APP_TYPE", "")
        audit_period = os.environ.get("AUDIT_PERIOD", "")
        url          = os.environ.get("TARGET_URL", "")
        method       = os.environ.get("TEST_METHOD", "")
    else:
        print("\n📝 Please provide report information:")
        client_name  = _get_user_input("Client Name",                   "Test Bank Pvt. Ltd.")
        app_name     = _get_user_input("Application Name",              "Website Name")
        app_type     = _get_user_input("Application Type",              "Internal Web Application")
        audit_period = _get_user_input("Audit Period",                  "17-01-2026 - 21-02-2026")
        url          = _get_user_input("Target URL",                    "https://example.com")
        method       = _get_user_input("Testing Method (Grey/Black Box)","Grey Box")

    # ── Import here so help/error messages work without dependencies ────
    sys.path.insert(0, os.path.dirname(__file__))
    from report_engine import ReportConfig, ReportEngine

    config = ReportConfig(
        client_name  = client_name,
        app_name     = app_name,
        app_type     = app_type,
        audit_period = audit_period,
        url          = url,
        method       = method,
        report_type  = args.type.capitalize(),
        environment  = args.environment.capitalize(),
        excel_file   = args.sheet,
        poc_folder   = args.poc or "",
        output_file  = args.output or "",
    )

    engine = ReportEngine(config)   # uses print() as default log_callback
    result = engine.run()

    sys.exit(0 if result.success else 1)


if __name__ == "__main__":
    main()