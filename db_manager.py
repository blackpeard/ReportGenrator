"""
db_manager.py
=============
SQLite wrapper for Report_tool.
Handles: employees, observation library, report history.

DB file lives at:  Report_tool/data/report_tool.db
"""

from __future__ import annotations
import os
import sys
import sqlite3
from datetime import datetime
from typing import Optional

# ── Resolve data/ folder same way as report_engine resolves templates/ ──────
if getattr(sys, 'frozen', False):
    _BASE_DIR = sys._MEIPASS
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_BASE_DIR, "data", "report_tool.db")
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)


# ── Schema ───────────────────────────────────────────────────────────────────

_SCHEMA = """
CREATE TABLE IF NOT EXISTS employees (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    designation TEXT NOT NULL DEFAULT '',
    email       TEXT NOT NULL DEFAULT '',
    department  TEXT NOT NULL DEFAULT '',
    qualifications      TEXT NOT NULL DEFAULT '',
    cert_in_listed      TEXT NOT NULL DEFAULT 'No',
    active      INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS obs_library (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    title           TEXT NOT NULL,
    severity        TEXT NOT NULL DEFAULT 'Medium',
    description     TEXT NOT NULL DEFAULT '',
    impact          TEXT NOT NULL DEFAULT '',
    recommendation  TEXT NOT NULL DEFAULT '',
    affected_url    TEXT NOT NULL DEFAULT '',
    cve             TEXT NOT NULL DEFAULT '',
    ref             TEXT NOT NULL DEFAULT '',
    category        TEXT NOT NULL DEFAULT 'General'
);

CREATE TABLE IF NOT EXISTS report_history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    client      TEXT NOT NULL DEFAULT '',
    app_name    TEXT NOT NULL DEFAULT '',
    report_type TEXT NOT NULL DEFAULT '',
    output_path TEXT NOT NULL DEFAULT '',
    prepared_by TEXT NOT NULL DEFAULT '',
    created_at  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS tools (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tool_name   TEXT NOT NULL,
    tool_version TEXT NOT NULL DEFAULT '',
    tool_type   TEXT NOT NULL DEFAULT 'Open Source',
    category    TEXT NOT NULL DEFAULT 'General',
    description TEXT NOT NULL DEFAULT '',
    active      INTEGER NOT NULL DEFAULT 1
);
"""

# ── Seed data — common web vulnerabilities ───────────────────────────────────

_SEED_OBSERVATIONS = [
    ("SQL Injection", "High",
     "The application is vulnerable to SQL Injection attacks. User-supplied input is not properly sanitised before being included in SQL queries, allowing an attacker to manipulate the database queries.",
     "An attacker could read sensitive data from the database, modify or delete data, execute administration operations on the database, and in some cases issue commands to the operating system.",
     "Use parameterised queries or prepared statements. Implement input validation and whitelist acceptable characters. Apply the principle of least privilege to database accounts.",
     "", "CVE-2021-27101", "https://owasp.org/www-community/attacks/SQL_Injection", "Injection"),

    ("Cross-Site Scripting (XSS)", "Medium",
     "The application does not properly encode user-supplied data before rendering it in the browser, allowing attackers to inject malicious client-side scripts.",
     "Attackers can steal session cookies, redirect users to malicious websites, deface the application, or perform actions on behalf of the victim user.",
     "Encode all user-supplied output using context-aware encoding. Implement a Content Security Policy (CSP) header. Validate and sanitise all input on the server side.",
     "", "", "https://owasp.org/www-community/attacks/xss/", "Injection"),

    ("Insecure Direct Object Reference (IDOR)", "High",
     "The application exposes internal implementation objects such as database records or files directly. An attacker can manipulate these references to access unauthorised data.",
     "Unauthorised access to other users' data, potential data breach, privilege escalation, and violation of data privacy regulations.",
     "Implement proper access control checks for every object access. Use indirect references mapped to the actual database keys. Log all access control failures.",
     "", "", "https://owasp.org/www-community/attacks/Insecure_Direct_Object_Reference", "Access Control"),

    ("Broken Authentication", "High",
     "The application's authentication mechanism is improperly implemented, allowing attackers to compromise passwords, keys, or session tokens.",
     "Account takeover, identity theft, unauthorised access to sensitive functionality, and potential full system compromise.",
     "Implement multi-factor authentication. Use strong password policies. Ensure session tokens are invalidated after logout. Implement account lockout after failed attempts.",
     "", "", "https://owasp.org/Top10/A07_2021-Identification_and_Authentication_Failures/", "Authentication"),

    ("Sensitive Data Exposure", "High",
     "The application transmits or stores sensitive data without adequate protection. Sensitive information such as passwords, credit card numbers, or personal data is exposed.",
     "Data breach, financial loss, reputational damage, and regulatory non-compliance (GDPR, PCI-DSS).",
     "Encrypt all sensitive data at rest and in transit using strong algorithms. Disable caching of sensitive data. Implement proper key management.",
     "", "", "https://owasp.org/Top10/A02_2021-Cryptographic_Failures/", "Cryptography"),

    ("Security Misconfiguration", "Medium",
     "The application, framework, or server is insecurely configured. Default credentials, unnecessary features, or verbose error messages are present.",
     "Attackers can exploit misconfigured systems to gain unauthorised access, gather sensitive information, or disrupt services.",
     "Implement a hardening process. Remove unnecessary features and frameworks. Review and update configurations regularly. Disable directory listing.",
     "", "", "https://owasp.org/Top10/A05_2021-Security_Misconfiguration/", "Configuration"),

    ("Cross-Site Request Forgery (CSRF)", "Medium",
     "The application does not validate that requests are intentionally made by the authenticated user. An attacker can trick users into performing unintended actions.",
     "Attackers can force authenticated users to perform state-changing requests, resulting in unauthorised fund transfers, email changes, or data modification.",
     "Implement CSRF tokens in all state-changing requests. Verify the Origin and Referer headers. Use the SameSite cookie attribute.",
     "", "", "https://owasp.org/www-community/attacks/csrf", "Access Control"),

    ("Using Components with Known Vulnerabilities", "Medium",
     "The application uses third-party libraries, frameworks, or components with known security vulnerabilities.",
     "Exploitation of known vulnerabilities in outdated components can lead to data loss, server compromise, or full application takeover.",
     "Regularly audit third-party components. Subscribe to security advisories. Implement a patch management process. Remove unused dependencies.",
     "", "", "https://owasp.org/Top10/A06_2021-Vulnerable_and_Outdated_Components/", "Configuration"),

    ("Insufficient Logging and Monitoring", "Low",
     "The application does not maintain adequate logs of security-relevant events, making it difficult to detect attacks or investigate security incidents.",
     "Delayed detection of breaches, inability to perform forensic analysis, failure to meet compliance requirements.",
     "Implement comprehensive logging for authentication, access control, and input validation failures. Establish a log monitoring process. Protect logs from tampering.",
     "", "", "https://owasp.org/Top10/A09_2021-Security_Logging_and_Monitoring_Failures/", "Logging"),

    ("XML External Entity Injection (XXE)", "High",
     "The application parses XML input containing a reference to an external entity, which can be used to disclose internal files or perform server-side request forgery.",
     "Disclosure of confidential files, server-side request forgery, denial of service, and remote code execution in some configurations.",
     "Disable XML external entity processing in the XML parser. Validate and sanitise all XML input. Use less complex data formats such as JSON where possible.",
     "", "CVE-2021-44228", "https://owasp.org/www-community/vulnerabilities/XML_External_Entity_(XXE)_Processing", "Injection"),

    ("Clickjacking", "Low",
     "The application can be embedded in an iframe on a malicious website, tricking users into performing unintended actions by clicking on invisible elements.",
     "Users can be tricked into performing unintended actions such as liking pages, making purchases, or changing account settings.",
     "Implement X-Frame-Options header set to DENY or SAMEORIGIN. Use the Content-Security-Policy frame-ancestors directive.",
     "", "", "https://owasp.org/www-community/attacks/Clickjacking", "Configuration"),

    ("Weak Password Policy", "Medium",
     "The application does not enforce a strong password policy, allowing users to set easily guessable passwords.",
     "Accounts with weak passwords are vulnerable to brute force and dictionary attacks, leading to account takeover.",
     "Enforce minimum password length of 12 characters. Require a mix of uppercase, lowercase, numbers, and special characters. Implement password history and expiry policies.",
     "", "", "https://owasp.org/www-community/controls/Blocking_Brute_Force_Attacks", "Authentication"),

    ("Missing HTTP Security Headers", "Low",
     "The application's HTTP responses are missing important security headers that help protect against common web attacks.",
     "Increased attack surface for XSS, clickjacking, MIME sniffing, and other client-side attacks.",
     "Implement security headers: Content-Security-Policy, X-Frame-Options, X-Content-Type-Options, Strict-Transport-Security, Referrer-Policy.",
     "", "", "https://owasp.org/www-project-secure-headers/", "Configuration"),

    ("Directory Traversal", "High",
     "The application allows users to access files and directories outside the intended web root directory by manipulating file path variables.",
     "Unauthorised access to sensitive files such as configuration files, source code, or system files containing credentials.",
     "Validate and sanitise all file path inputs. Use a whitelist of allowed files or directories. Implement proper access controls at the OS level.",
     "", "", "https://owasp.org/www-community/attacks/Path_Traversal", "Injection"),

    ("Open Redirect", "Low",
     "The application accepts user-controlled input that specifies a link to an external site and uses it to redirect users without validation.",
     "Phishing attacks, credential theft, and reputation damage when the application is used to redirect users to malicious sites.",
     "Avoid using redirects and forwards where possible. If used, validate the destination URL against a whitelist of allowed URLs.",
     "", "", "https://owasp.org/www-community/attacks/Unvalidated_Redirects_and_Forwards", "Access Control"),
]

_SEED_EMPLOYEES = [
    ("Pushpendra Bharambe", "Partner", "pushpendra.bharambe@nangia.com", "Cyber Security"),
    ("Asif Balasinor", "Associate Director", "asif.balasinor@nangia.com", "Cyber Security"),
    ("Neelkanth Gawde", "Manager", "neelkanth.gawde@nangia.com", "Cyber Security"),
    ("Milind Fegade", "Manager", "milind.fegade@nangia.com", "Cyber Security"),
    ("Bhupendra Singh Sisodiya", "Senior Consultant", "bhupendra.singh@nangiaglobal.com", "Cyber Security"),
    ("Gaurav Belwal", "Senior Consultant", "gaurav.belwal@nangiaglobal.com", "Cyber Security"),
    ("Satyanarayan", "Consultant", "satya.sangwal@nangiaglobal.com", "Cyber Security"),
]

_SEED_TOOLS = [
    # Web
    ("Burp Suite Professional", "2024.x", "Licensed",    "Web",        ),
    ("OWASP ZAP",               "2.15",   "Open Source", "Web",        ),
    ("Nikto",                   "2.1.6",  "Open Source", "Web",        ),
    ("SQLMap",                  "1.8",    "Open Source", "Web",         ),
    ("WFuzz",                   "3.1",    "Open Source", "Web",         ),
    ("Gobuster",                "3.6",    "Open Source", "Web",        ),
    # API
    ("Postman",                 "11.x",   "Licensed",    "API",        ),
    ("Insomnia",                "9.x",    "Open Source", "API",         ),
    ("ffuf",                    "2.1",    "Open Source", "API",         ),
    # Mobile
    ("MobSF",                   "4.x",    "Open Source", "Mobile",     ),
    ("Frida",                   "16.x",   "Open Source", "Mobile",      ),
    ("Objection",               "1.11",   "Open Source", "Mobile",      ),
    ("APKTool",                 "2.9",    "Open Source", "Mobile",     ),
    # Source Code
    ("SonarQube",               "10.x",   "Open Source", "Source Code", ),
    ("Semgrep",                 "1.x",    "Open Source", "Source Code", ),
    ("Bandit",                  "1.7",    "Open Source", "Source Code", ),
    ("Checkmarx",               "2024",   "Licensed",    "Source Code", ),
    # Red Team
    ("Metasploit Framework",    "6.x",    "Open Source", "Red Team",    ),
    ("Cobalt Strike",           "4.x",    "Licensed",    "Red Team",  ),
    ("BloodHound",              "4.x",    "Open Source", "Red Team",    ),
    ("Mimikatz",                "2.x",    "Open Source", "Red Team",   ),
    ("Responder",               "3.x",    "Open Source", "Red Team",   ),
    # Internal / Network
    ("Nmap",                    "7.95",   "Open Source", "Internal",   ),
    ("Nessus",                  "10.x",   "Licensed",    "Internal",   ),
    ("OpenVAS",                 "22.x",   "Open Source", "Internal",    ),
    ("Wireshark",               "4.x",    "Open Source", "Internal",   ),
    ("Netcat",                  "1.x",    "Open Source", "Internal",    ),
    # General
    ("Kali Linux",              "2024.x", "Open Source", "General",    ),
    ("Ncrack",                  "0.7",    "Open Source", "General",    ),
    ("Hydra",                   "9.5",    "Open Source", "General",     ),
]


# ── Database manager class ───────────────────────────────────────────────────

class DBManager:
    """
    Thin SQLite wrapper.
    Usage:
        db = DBManager()
        employees = db.get_employees()
        obs = db.search_observations("sql")
    """

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        return conn

    def _init_db(self):
        """Create tables and seed default data if first run."""
        with self._connect() as conn:
            conn.executescript(_SCHEMA)
            # Seed only if tables are empty
            if conn.execute("SELECT COUNT(*) FROM obs_library").fetchone()[0] == 0:
                conn.executemany(
                    """INSERT INTO obs_library
                       (title, severity, description, impact, recommendation,
                        affected_url, cve, ref, category)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    _SEED_OBSERVATIONS,
                )
            if conn.execute("SELECT COUNT(*) FROM employees").fetchone()[0] == 0:
                conn.executemany(
                    "INSERT INTO employees (name, designation, email, department) VALUES (?,?,?,?)",
                    _SEED_EMPLOYEES,
                )
            if conn.execute("SELECT COUNT(*) FROM tools").fetchone()[0] == 0:
                conn.executemany(
                    "INSERT INTO tools (tool_name, tool_version, tool_type, category) VALUES (?,?,?,?)",
                    _SEED_TOOLS,
                )

    # ── Employees ────────────────────────────────────────────────────────────

    def get_employees(self, active_only: bool = True) -> list[dict]:
        with self._connect() as conn:
            q = "SELECT * FROM employees"
            if active_only:
                q += " WHERE active=1"
            q += " ORDER BY name"
            return [dict(r) for r in conn.execute(q).fetchall()]

    def add_employee(self, name: str, designation: str, email: str = "", department: str = "", qualifications: str = "", cert_in_listed: str = "No") -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """INSERT INTO employees
                (name, designation, email, department, qualifications, cert_in_listed)
                VALUES (?,?,?,?,?,?)""",
                (name, designation, email, department, qualifications, cert_in_listed),
            )
            return cur.lastrowid

    def update_employee(self, emp_id: int, name: str, designation: str,
                        email: str = "", department: str = "",):
        with self._connect() as conn:
            conn.execute(
                "UPDATE employees SET name=?, designation=?, email=?, department=? WHERE id=?",
                (name, designation, email, department, emp_id),
            )

    def delete_employee(self, emp_id: int):
        with self._connect() as conn:
            conn.execute("UPDATE employees SET active=0 WHERE id=?", (emp_id,))

    # ── Observation library ───────────────────────────────────────────────────

    def get_all_observations(self) -> list[dict]:
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM obs_library ORDER BY severity DESC, title"
            ).fetchall()]

    def search_observations(self, query: str) -> list[dict]:
        like = f"%{query}%"
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                """SELECT * FROM obs_library
                   WHERE title LIKE ? OR category LIKE ? OR description LIKE ?
                   ORDER BY severity DESC, title""",
                (like, like, like),
            ).fetchall()]

    def get_observation_by_id(self, obs_id: int) -> Optional[dict]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM obs_library WHERE id=?", (obs_id,)
            ).fetchone()
            return dict(row) if row else None

    def add_observation(self, title: str, severity: str, description: str,
                        impact: str, recommendation: str, affected_url: str = "",
                        cve: str = "", ref: str = "", category: str = "General") -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """INSERT INTO obs_library
                   (title, severity, description, impact, recommendation,
                    affected_url, cve, ref, category)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (title, severity, description, impact, recommendation,
                 affected_url, cve, ref, category),
            )
            return cur.lastrowid

    def update_observation(self, obs_id: int, title: str, severity: str,
                           description: str, impact: str, recommendation: str,
                           affected_url: str = "", cve: str = "",
                           ref: str = "", category: str = "General"):
        with self._connect() as conn:
            conn.execute(
                """UPDATE obs_library SET title=?, severity=?, description=?,
                   impact=?, recommendation=?, affected_url=?, cve=?,
                   ref=?, category=? WHERE id=?""",
                (title, severity, description, impact, recommendation,
                 affected_url, cve, ref, category, obs_id),
            )

    def delete_observation(self, obs_id: int):
        with self._connect() as conn:
            conn.execute("DELETE FROM obs_library WHERE id=?", (obs_id,))

    def get_categories(self) -> list[str]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT category FROM obs_library ORDER BY category"
            ).fetchall()
            return [r[0] for r in rows]

      # ── Tools ────────────────────────────────────────────────────────   

    def get_tools(self, category: str = "") -> list[dict]:
        with self._connect() as conn:
            if category and category != "All":
                rows = conn.execute(
                    "SELECT * FROM tools WHERE active=1 AND category=? ORDER BY category, tool_name",
                    (category,)
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM tools WHERE active=1 ORDER BY category, tool_name"
                ).fetchall()
            return [dict(r) for r in rows]

    def get_tool_categories(self) -> list[str]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT category FROM tools WHERE active=1 ORDER BY category"
            ).fetchall()
            return [r[0] for r in rows]

    def add_tool(self, name: str, version: str, tool_type: str,
                category: str, description: str = "") -> int:
        with self._connect() as conn:
            cur = conn.execute(
                "INSERT INTO tools (tool_name, tool_version, tool_type, category, description) VALUES (?,?,?,?,?)",
                (name, version, tool_type, category, description),
            )
            return cur.lastrowid

    def delete_tool(self, tool_id: int):
        with self._connect() as conn:
            conn.execute("UPDATE tools SET active=0 WHERE id=?", (tool_id,))

    # ── Report history ────────────────────────────────────────────────────────

    def save_report_history(self, client: str, app_name: str, report_type: str,
                            output_path: str, prepared_by: str):
        with self._connect() as conn:
            conn.execute(
                """INSERT INTO report_history
                   (client, app_name, report_type, output_path, prepared_by, created_at)
                   VALUES (?,?,?,?,?,?)""",
                (client, app_name, report_type, output_path, prepared_by,
                 datetime.now().strftime("%Y-%m-%d %H:%M")),
            )

    def get_report_history(self, limit: int = 50) -> list[dict]:
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM report_history ORDER BY created_at DESC LIMIT ?",
                (limit,),
            ).fetchall()]
