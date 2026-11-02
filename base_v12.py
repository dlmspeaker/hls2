#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hayleys Life Sciences – Service Agreement Portal v12.0
Enterprise Edition + AI Technician Recommendation + Lab Analytics + Reminders + Customer Email + 3‑Year Reports
Enhanced by Red Axis AI Dev - FULLY FIXED
"""

import csv
import json
import os
import sqlite3
import threading
import webbrowser
import hashlib
import secrets
import logging
import io
import zipfile
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from http.server import BaseHTTPRequestHandler, HTTPServer
import urllib.parse
import openpyxl
from collections import Counter

# ==================== CONFIG ====================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Where the database, backups and logs live. Docker sets HLS_DATA_DIR=/data (a persistent volume).
DATA_DIR = os.environ.get("HLS_DATA_DIR", SCRIPT_DIR)
os.makedirs(DATA_DIR, exist_ok=True)
DB_NAME = os.path.join(DATA_DIR, "hayleys_redaxis_v12.db")
CSV_FILE = os.path.join(SCRIPT_DIR, "Full_Categorized_HiTech_Data (version 1).csv")
HTML_FILE = os.path.join(SCRIPT_DIR, "super_dashboard_v12.html")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
LOGS_DIR = os.path.join(DATA_DIR, "logs")
PORT = 8080

for dir_path in [BACKUP_DIR, LOGS_DIR]:
    os.makedirs(dir_path, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, "server.log")),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# For PDF reports
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    logger.warning("fpdf or matplotlib not installed. PDF reports will be disabled.")

# For email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# For scheduler
try:
    from apscheduler.schedulers.background import BackgroundScheduler
    SCHEDULER_AVAILABLE = True
except ImportError:
    SCHEDULER_AVAILABLE = False
    logger.warning("apscheduler not installed. Reminders will be disabled.")


class DatabaseManager:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_database()
        self._create_indexes()
        self._backup_count = 0
        if SCHEDULER_AVAILABLE:
            self._start_scheduler()

    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _auto_backup(self, reason: str = "auto"):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"auto_backup_{reason}_{timestamp}.zip"
            filepath = os.path.join(BACKUP_DIR, filename)
            zip_data = self.export_full_backup_zip()
            with open(filepath, 'wb') as f:
                f.write(zip_data)
            logger.info(f"Auto-backup created: {filepath}")
            self._cleanup_old_backups(keep=20)
            return filename
        except Exception as e:
            logger.error(f"Auto-backup failed: {e}")
            return None

    def _cleanup_old_backups(self, keep=20):
        try:
            files = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("auto_backup_") and f.endswith(".zip")],
                           key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f)))
            while len(files) > keep:
                old = files.pop(0)
                os.remove(os.path.join(BACKUP_DIR, old))
                logger.info(f"Removed old backup: {old}")
        except Exception as e:
            logger.warning(f"Cleanup old backups failed: {e}")

    def _save_uploaded_csv_as_backup(self, customer: str, csv_data: str):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"company_csv_{customer}_{timestamp}.csv"
            filepath = os.path.join(BACKUP_DIR, filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(csv_data)
            logger.info(f"Uploaded CSV saved as backup: {filepath}")
        except Exception as e:
            logger.error(f"Failed to save uploaded CSV backup: {e}")

    # ---------- INIT ----------
    def _init_database(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        # Customers
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                contact_person TEXT,
                email TEXT,
                phone TEXT,
                region TEXT DEFAULT 'Western',
                service_agreement_start TEXT,
                service_agreement_end TEXT,
                agreement_type TEXT DEFAULT 'Basic',
                created_at TEXT
            )
        ''')
        # Labs
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS labs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT NOT NULL,
                lab_name TEXT NOT NULL,
                contact_person TEXT,
                phone TEXT,
                email TEXT,
                location TEXT,
                created_at TEXT,
                FOREIGN KEY (customer_name) REFERENCES customers(name) ON DELETE CASCADE,
                UNIQUE(customer_name, lab_name)
            )
        ''')
        # Instruments
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS instruments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                serial_number TEXT UNIQUE NOT NULL,
                machine_type TEXT NOT NULL,
                model TEXT,
                owner_customer TEXT NOT NULL,
                lab_id INTEGER,
                division TEXT DEFAULT 'GENERAL LAB',
                parent_serial TEXT,
                is_secondary INTEGER DEFAULT 0,
                description TEXT,
                status TEXT DEFAULT 'Active',
                warranty_end TEXT,
                last_service TEXT,
                next_service_due TEXT,
                created_at TEXT,
                FOREIGN KEY (owner_customer) REFERENCES customers(name) ON DELETE CASCADE,
                FOREIGN KEY (lab_id) REFERENCES labs(id) ON DELETE SET NULL,
                FOREIGN KEY (parent_serial) REFERENCES instruments(serial_number) ON DELETE CASCADE
            )
        ''')
        # Service tickets
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS service_tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_no TEXT,
                service_date TEXT,
                serial_number TEXT,
                job_actions TEXT,
                technicians TEXT,
                service_notes TEXT,
                service_type TEXT DEFAULT 'Routine',
                created_at TEXT,
                FOREIGN KEY (serial_number) REFERENCES instruments(serial_number) ON DELETE CASCADE
            )
        ''')
        # Customer-level agreements
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS service_agreements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT NOT NULL,
                agreement_number TEXT UNIQUE,
                start_date TEXT,
                end_date TEXT,
                agreement_type TEXT,
                total_amount REAL,
                status TEXT DEFAULT 'Active',
                created_at TEXT,
                FOREIGN KEY (customer_name) REFERENCES customers(name) ON DELETE CASCADE
            )
        ''')
        # Machine-level agreements
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS machine_agreements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                serial_number TEXT NOT NULL,
                agreement_number TEXT UNIQUE,
                start_date TEXT,
                end_date TEXT,
                agreement_type TEXT,
                total_amount REAL,
                status TEXT DEFAULT 'Active',
                created_at TEXT,
                FOREIGN KEY (serial_number) REFERENCES instruments(serial_number) ON DELETE CASCADE
            )
        ''')
        # Users
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'viewer',
                full_name TEXT,
                last_login TEXT,
                created_at TEXT
            )
        ''')
        # System logs
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS system_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                type TEXT NOT NULL,
                message TEXT NOT NULL,
                user TEXT,
                created_at TEXT
            )
        ''')
        # Transfer log
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transfer_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                serial_number TEXT NOT NULL,
                from_customer TEXT,
                to_customer TEXT,
                from_lab TEXT,
                to_lab TEXT,
                transferred_by TEXT,
                transferred_at TEXT,
                reason TEXT
            )
        ''')
        # --- NEW TABLES for Customer Engagement ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS customer_preferences (
                customer_name TEXT PRIMARY KEY,
                preferred_contact_method TEXT DEFAULT 'email',
                email TEXT,
                phone TEXT,
                service_reminder_days INTEGER DEFAULT 7,
                receive_email_reminders INTEGER DEFAULT 1,
                FOREIGN KEY (customer_name) REFERENCES customers(name) ON DELETE CASCADE
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS service_surveys (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                service_ticket_id INTEGER,
                customer_name TEXT,
                rating INTEGER,
                feedback TEXT,
                nps_score INTEGER,
                sent_at TEXT,
                responded_at TEXT,
                FOREIGN KEY (service_ticket_id) REFERENCES service_tickets(id) ON DELETE CASCADE,
                FOREIGN KEY (customer_name) REFERENCES customers(name) ON DELETE CASCADE
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reminder_settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_recipients TEXT,
                reminder_days_before INTEGER DEFAULT 7,
                send_to_customers INTEGER DEFAULT 0
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reminder_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                reminder_type TEXT,
                serial_number TEXT,
                customer_name TEXT,
                sent_at TEXT,
                recipients TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS machine_health (
                serial_number TEXT PRIMARY KEY,
                health_score REAL,
                last_updated TEXT,
                FOREIGN KEY (serial_number) REFERENCES instruments(serial_number) ON DELETE CASCADE
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales_opportunities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT,
                machine_serial TEXT,
                opportunity_type TEXT,
                description TEXT,
                value REAL,
                status TEXT DEFAULT 'Open',
                created_at TEXT,
                FOREIGN KEY (customer_name) REFERENCES customers(name) ON DELETE CASCADE
            )
        ''')

        # Default users
        cursor.execute("SELECT COUNT(*) FROM users")
        if cursor.fetchone()[0] == 0:
            now = datetime.now().isoformat()
            users = [
                ('admin', hashlib.sha256("admin123".encode()).hexdigest(), 'admin', 'Administrator', now),
                ('service', hashlib.sha256("service@1".encode()).hexdigest(), 'service', 'Service Manager', now),
                ('sales', hashlib.sha256("sales123".encode()).hexdigest(), 'sales', 'Sales Manager', now)
            ]
            for u in users:
                cursor.execute(
                    "INSERT INTO users (username, password_hash, role, full_name, created_at) VALUES (?,?,?,?,?)",
                    u
                )

        # Default reminder settings
        cursor.execute("SELECT COUNT(*) FROM reminder_settings")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO reminder_settings (email_recipients, reminder_days_before, send_to_customers) VALUES (?, ?, ?)",
                           ("service@hayleys.com, sales@hayleys.com", 7, 0))

        conn.commit()
        conn.close()
        self._import_csv()

    def _create_indexes(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_instruments_owner ON instruments(owner_customer)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_instruments_parent ON instruments(parent_serial)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_instruments_lab ON instruments(lab_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_services_serial ON service_tickets(serial_number)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_services_date ON service_tickets(service_date)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_customers_name ON customers(name)")
        conn.commit()
        conn.close()

    def _import_csv(self):
        if not os.path.exists(CSV_FILE):
            logger.warning(f"CSV not found: {CSV_FILE}")
            return
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM instruments")
        if cursor.fetchone()[0] > 0:
            conn.close()
            return
        logger.info("Importing initial CSV...")
        try:
            with open(CSV_FILE, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    customer = row.get("Customer", "").strip()
                    if not customer:
                        continue
                    cursor.execute("INSERT OR IGNORE INTO customers (name, created_at) VALUES (?, ?)",
                                   (customer, datetime.now().isoformat()))
                    serial = row.get("Serial Number", "").strip()
                    if not serial:
                        serial = f"IMP-{int(datetime.now().timestamp())}"
                    division = row.get("Division", "").strip() or "GENERAL LAB"
                    machine_type = row.get("Machine", "").strip() or "Unknown"
                    model = row.get("Model", "").strip()
                    cursor.execute('''
                        INSERT OR IGNORE INTO instruments
                        (serial_number, machine_type, model, owner_customer, division, is_secondary, status, created_at)
                        VALUES (?,?,?,?,?,?,?,?)
                    ''', (serial, machine_type, model, customer, division, 0, 'Active', datetime.now().isoformat()))
                    service_date_raw = row.get("Service Date", "").strip()
                    if service_date_raw:
                        try:
                            dt = datetime.strptime(service_date_raw, "%B %d, %Y")
                            service_date = dt.strftime("%Y-%m-%d")
                        except:
                            service_date = service_date_raw
                        cursor.execute('''
                            INSERT INTO service_tickets
                            (job_no, service_date, serial_number, job_actions, technicians, created_at)
                            VALUES (?,?,?,?,?,?)
                        ''', (row.get("JOB NO", ""), service_date, serial, row.get("Job", ""), row.get("Tech_1", ""), datetime.now().isoformat()))
                        cursor.execute("UPDATE instruments SET last_service = ? WHERE serial_number = ?", (service_date, serial))
                        try:
                            last_date = datetime.strptime(service_date, "%Y-%m-%d")
                            next_due = last_date + timedelta(days=180)
                            cursor.execute("UPDATE instruments SET next_service_due = ? WHERE serial_number = ?", (next_due.strftime("%Y-%m-%d"), serial))
                        except:
                            pass
                conn.commit()
        except Exception as e:
            logger.error(f"CSV import error: {e}")
        finally:
            conn.close()
        logger.info("CSV import finished")

    # ---------- COMPANY CSV EXPORT ----------
    def export_company_csv(self, customer_name: str) -> str:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM instruments WHERE owner_customer=?", (customer_name,))
        machines = [dict(r) for r in cursor.fetchall()]
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow([
            "Service Date", "Division", "JOB NO", "Customer", "Supplier",
            "Machine", "Model", "Serial Number", "Job",
            "Contact Person", "Tech_1", "Tech_2", "Tech_3", "Tech_4", "Tech_5",
            "Old Serial Number", "Parent Serial Number", "Lab"
        ])
        for m in machines:
            cursor.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date", (m['serial_number'],))
            services = [dict(r) for r in cursor.fetchall()]
            if services:
                for s in services:
                    techs = s.get('technicians', '').split(',') if s.get('technicians') else []
                    tech_cols = [techs[i] if i < len(techs) else '' for i in range(5)]
                    writer.writerow([
                        s.get('service_date', ''),
                        m.get('division', ''),
                        s.get('job_no', ''),
                        customer_name,
                        '',
                        m.get('machine_type', ''),
                        m.get('model', ''),
                        m['serial_number'],
                        s.get('job_actions', ''),
                        '',
                        *tech_cols,
                        '',   # Old Serial Number
                        m.get('parent_serial', ''),
                        ''   # Lab
                    ])
            else:
                writer.writerow([
                    '',
                    m.get('division', ''),
                    '',
                    customer_name,
                    '',
                    m.get('machine_type', ''),
                    m.get('model', ''),
                    m['serial_number'],
                    '',
                    '',
                    '', '', '', '', '',
                    '',
                    m.get('parent_serial', ''),
                    ''
                ])
        conn.close()
        return output.getvalue()

    # ---------- ATOMIC COMPANY CSV UPDATE ----------
    def upload_company_csv(self, customer_name: str, csv_data: str, username: str) -> Dict:
        self._save_uploaded_csv_as_backup(customer_name, csv_data)
        backup_zip = self.export_full_backup_zip()
        backup_filename = f"pre_update_{customer_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        with open(backup_path, 'wb') as f:
            f.write(backup_zip)

        conn = self._get_connection()
        cursor = conn.cursor()
        added_machines = 0
        updated_machines = 0
        renamed_machines = 0
        added_services = 0
        errors = []

        try:
            conn.execute("BEGIN TRANSACTION")
            cursor.execute("INSERT OR IGNORE INTO customers (name, created_at) VALUES (?, ?)",
                           (customer_name, datetime.now().isoformat()))

            pending_links = {}
            reader = csv.DictReader(io.StringIO(csv_data), skipinitialspace=True)
            rows = list(reader)
            for row in rows:
                new_serial = row.get("Serial Number", "").strip()
                if not new_serial:
                    continue

                old_serial = row.get("Old Serial Number", "").strip()
                parent_serial = row.get("Parent Serial Number", "").strip()
                machine_type = row.get("Machine", "").strip() or "Unknown"
                model = row.get("Model", "").strip()
                division = row.get("Division", "").strip() or "GENERAL LAB"

                lab_name = row.get("Lab", "").strip()
                lab_id = None
                if lab_name:
                    cursor.execute("SELECT id FROM labs WHERE customer_name = ? AND lab_name = ?", (customer_name, lab_name))
                    lab_row = cursor.fetchone()
                    if lab_row:
                        lab_id = lab_row['id']
                    else:
                        cursor.execute("INSERT INTO labs (customer_name, lab_name, created_at) VALUES (?,?,?)",
                                       (customer_name, lab_name, datetime.now().isoformat()))
                        lab_id = cursor.lastrowid

                techs = []
                for i in range(1, 10):
                    tech_col = f"Tech_{i}"
                    if tech_col in row and row[tech_col].strip():
                        techs.append(row[tech_col].strip())
                technicians = ",".join(techs) if techs else row.get("Tech_1", "").strip()

                if old_serial and old_serial != new_serial:
                    cursor.execute("SELECT * FROM instruments WHERE serial_number=?", (old_serial,))
                    old_exists = cursor.fetchone()
                    if not old_exists:
                        errors.append(f"Row: Old serial '{old_serial}' not found. Skipped.")
                        continue

                    cursor.execute("SELECT * FROM instruments WHERE serial_number=?", (new_serial,))
                    new_exists = cursor.fetchone()

                    temp_serial = f"__TEMP_RENAME_{old_serial}_{int(datetime.now().timestamp())}"

                    cursor.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (temp_serial, old_serial))
                    cursor.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (temp_serial, old_serial))
                    cursor.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (temp_serial, old_serial))
                    cursor.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (temp_serial, old_serial))

                    if new_exists:
                        temp2 = f"__TEMP_SWAP_{new_serial}_{int(datetime.now().timestamp())}"
                        cursor.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (temp2, new_serial))
                        cursor.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (temp2, new_serial))
                        cursor.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (temp2, new_serial))
                        cursor.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (temp2, new_serial))
                        cursor.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (old_serial, temp2))
                        cursor.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (old_serial, temp2))
                        cursor.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (old_serial, temp2))
                        cursor.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (old_serial, temp2))
                        cursor.execute("UPDATE instruments SET machine_type=?, model=?, division=?, lab_id=? WHERE serial_number=?",
                                       (machine_type, model, division, lab_id, new_serial))
                        renamed_machines += 1
                    else:
                        cursor.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (new_serial, temp_serial))
                        cursor.execute("UPDATE instruments SET machine_type=?, model=?, division=?, lab_id=? WHERE serial_number=?",
                                       (machine_type, model, division, lab_id, new_serial))
                        renamed_machines += 1

                    if parent_serial:
                        # linked after all rows are loaded, so row order in the CSV does not matter
                        pending_links[new_serial] = parent_serial

                else:
                    cursor.execute("SELECT * FROM instruments WHERE serial_number=?", (new_serial,))
                    existing = cursor.fetchone()
                    if existing:
                        cursor.execute('''
                            UPDATE instruments SET machine_type=?, model=?, division=?, lab_id=?
                            WHERE serial_number=?
                        ''', (machine_type, model, division, lab_id, new_serial))
                        updated_machines += 1
                    else:
                        cursor.execute('''
                            INSERT INTO instruments
                            (serial_number, machine_type, model, owner_customer, division, is_secondary, status, lab_id, created_at)
                            VALUES (?,?,?,?,?,?,?,?,?)
                        ''', (new_serial, machine_type, model, customer_name, division, 0, 'Active', lab_id, datetime.now().isoformat()))
                        added_machines += 1

                    if parent_serial:
                        # linked after all rows are loaded, so row order in the CSV does not matter
                        pending_links[new_serial] = parent_serial

                service_date_raw = row.get("Service Date", "").strip()
                if service_date_raw:
                    try:
                        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%B %d, %Y", "%d/%m/%Y"):
                            try:
                                dt = datetime.strptime(service_date_raw, fmt)
                                service_date = dt.strftime("%Y-%m-%d")
                                break
                            except:
                                continue
                        else:
                            service_date = service_date_raw
                    except:
                        service_date = service_date_raw

                    job_no = row.get("JOB NO", "").strip()
                    cursor.execute("SELECT id FROM service_tickets WHERE serial_number=? AND service_date=? AND job_no=?",
                                   (new_serial, service_date, job_no))
                    if not cursor.fetchone():
                        cursor.execute('''
                            INSERT INTO service_tickets
                            (job_no, service_date, serial_number, job_actions, technicians, created_at)
                            VALUES (?,?,?,?,?,?)
                        ''', (job_no, service_date, new_serial, row.get("Job", ""), technicians, datetime.now().isoformat()))
                        added_services += 1
                        cursor.execute("UPDATE instruments SET last_service=? WHERE serial_number=?", (service_date, new_serial))
                        try:
                            last_date = datetime.strptime(service_date, "%Y-%m-%d")
                            next_due = last_date + timedelta(days=180)
                            cursor.execute("UPDATE instruments SET next_service_due=? WHERE serial_number=?", (next_due.strftime("%Y-%m-%d"), new_serial))
                        except:
                            pass

            # second pass: attach secondary machines to their parent machine
            linked_secondaries = 0
            for child, par in pending_links.items():
                prow = cursor.execute("SELECT is_secondary, lab_id FROM instruments WHERE serial_number=?", (par,)).fetchone()
                if child == par:
                    errors.append(f"Serial '{child}' cannot be its own parent. Skipped link.")
                elif not prow:
                    errors.append(f"Parent serial '{par}' for '{child}' not found. Machine kept as primary.")
                elif prow['is_secondary'] == 1 or par in pending_links:
                    errors.append(f"Parent '{par}' is itself a secondary. '{child}' kept as primary.")
                else:
                    cursor.execute("UPDATE instruments SET is_secondary=1, parent_serial=?, lab_id=COALESCE(?, lab_id) WHERE serial_number=?", (par, prow['lab_id'], child))
                    linked_secondaries += 1
            conn.commit()
            msg = f"✅ Added {added_machines}, updated {updated_machines}, renamed {renamed_machines}, added {added_services} services, linked {linked_secondaries} secondary machines."
            if errors:
                msg += f" ⚠️ Errors: {'; '.join(errors[:5])}"
            self._add_log('COMPANY_CSV_UPDATE', f"Customer {customer_name}: {msg}", user=username)
            return {
                "status": "success",
                "message": msg,
                "backup_file": backup_filename,
                "details": {
                    "added": added_machines,
                    "updated": updated_machines,
                    "renamed": renamed_machines,
                    "services": added_services,
                    "secondaries_linked": linked_secondaries,
                    "errors": errors
                }
            }
        except Exception as e:
            conn.rollback()
            self._add_log('COMPANY_CSV_ERROR', f"Customer {customer_name}: {str(e)}", user=username)
            return {"status": "error", "message": f"❌ Update failed: {str(e)}", "backup_file": backup_filename}
        finally:
            conn.close()

    # ---------- FULL BACKUP (ZIP) ----------
    def export_full_backup_zip(self) -> bytes:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            tables = {
                "customers": "SELECT * FROM customers",
                "instruments": "SELECT * FROM instruments",
                "service_tickets": "SELECT * FROM service_tickets",
                "labs": "SELECT * FROM labs",
                "service_agreements": "SELECT * FROM service_agreements",
                "machine_agreements": "SELECT * FROM machine_agreements",
                "transfer_log": "SELECT * FROM transfer_log",
                "system_logs": "SELECT * FROM system_logs",
                "customer_preferences": "SELECT * FROM customer_preferences",
                "service_surveys": "SELECT * FROM service_surveys",
                "reminder_history": "SELECT * FROM reminder_history",
                "machine_health": "SELECT * FROM machine_health",
                "sales_opportunities": "SELECT * FROM sales_opportunities"
            }
            conn = self._get_connection()
            cursor = conn.cursor()
            for table_name, query in tables.items():
                cursor.execute(query)
                rows = cursor.fetchall()
                if not rows:
                    continue
                col_names = [description[0] for description in cursor.description]
                output = io.StringIO()
                writer = csv.writer(output)
                writer.writerow(col_names)
                for row in rows:
                    writer.writerow(row)
                zipf.writestr(f"{table_name}_{timestamp}.csv", output.getvalue())
            conn.close()
        zip_buffer.seek(0)
        return zip_buffer.getvalue()

    def save_backup_to_server(self) -> str:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"hayleys_backup_{timestamp}.zip"
        filepath = os.path.join(BACKUP_DIR, filename)
        zip_data = self.export_full_backup_zip()
        with open(filepath, 'wb') as f:
            f.write(zip_data)
        logger.info(f"Manual backup saved: {filepath}")
        return filepath

    # ---------- LAB MANAGEMENT ----------
    def get_labs(self, customer_name: str) -> List[Dict]:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM labs WHERE customer_name = ? ORDER BY lab_name", (customer_name,))
        labs = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return labs

    def create_lab(self, data: Dict) -> Dict:
        self._auto_backup("create_lab")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO labs (customer_name, lab_name, contact_person, phone, email, location, created_at)
                VALUES (?,?,?,?,?,?,?)
            ''', (data['customer_name'], data['lab_name'], data.get('contact_person',''), data.get('phone',''),
                  data.get('email',''), data.get('location',''), datetime.now().isoformat()))
            conn.commit()
            self._add_log('LAB_ADD', f"Added lab {data['lab_name']} for {data['customer_name']}")
            return {"status": "success", "message": "Lab created"}
        except sqlite3.IntegrityError:
            return {"status": "error", "message": "Lab name already exists for this customer"}
        finally:
            conn.close()

    def update_lab(self, data: Dict) -> Dict:
        self._auto_backup("update_lab")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE labs SET lab_name=?, contact_person=?, phone=?, email=?, location=?
                WHERE id=? AND customer_name=?
            ''', (data['lab_name'], data.get('contact_person',''), data.get('phone',''),
                  data.get('email',''), data.get('location',''), data['id'], data['customer_name']))
            conn.commit()
            self._add_log('LAB_UPDATE', f"Updated lab {data['lab_name']}")
            return {"status": "success", "message": "Lab updated"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def delete_lab(self, lab_id: int, customer_name: str) -> Dict:
        self._auto_backup("delete_lab")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE instruments SET lab_id = NULL WHERE lab_id = ?", (lab_id,))
            cursor.execute("DELETE FROM labs WHERE id = ? AND customer_name = ?", (lab_id, customer_name))
            conn.commit()
            self._add_log('LAB_DELETE', f"Deleted lab id {lab_id}")
            return {"status": "success", "message": "Lab deleted"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def assign_machine_to_lab(self, serial_number: str, lab_id: Optional[int]) -> Dict:
        self._auto_backup("assign_machine_lab")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE instruments SET lab_id = ? WHERE serial_number = ?", (lab_id, serial_number))
            conn.commit()
            self._add_log('MACHINE_ASSIGN_LAB', f"Machine {serial_number} assigned to lab {lab_id}")
            return {"status": "success", "message": "Machine assigned to lab"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    # ---------- LAB SUMMARY (with region) ----------
    def get_labs_summary(self, customer: str = None, division: str = None) -> List[Dict]:
        conn = self._get_connection()
        cursor = conn.cursor()
        query = """
            SELECT
                l.id, l.lab_name, l.customer_name, l.location, c.region as customer_region,
                COUNT(i.id) as total_machines
            FROM labs l
            LEFT JOIN customers c ON l.customer_name = c.name
            LEFT JOIN instruments i ON l.id = i.lab_id
            WHERE 1=1
        """
        params = []
        if customer:
            query += " AND l.customer_name = ?"
            params.append(customer)
        if division:
            query += " AND i.division = ?"
            params.append(division)
        query += " GROUP BY l.id ORDER BY l.customer_name, l.lab_name"
        cursor.execute(query, params)
        labs = []
        for row in cursor.fetchall():
            lab = dict(row)
            breakdown_query = """
                SELECT
                    machine_type,
                    COUNT(*) as count,
                    division
                FROM instruments
                WHERE lab_id = ?
            """
            breakdown_params = [lab['id']]
            if division:
                breakdown_query += " AND division = ?"
                breakdown_params.append(division)
            breakdown_query += " GROUP BY machine_type, division"
            cursor.execute(breakdown_query, breakdown_params)
            breakdown = [dict(r) for r in cursor.fetchall()]

            type_counts = {}
            div_counts = {}
            for b in breakdown:
                type_counts[b['machine_type']] = type_counts.get(b['machine_type'], 0) + b['count']
                div_counts[b['division']] = div_counts.get(b['division'], 0) + b['count']

            lab['type_breakdown'] = [{'type': k, 'count': v} for k, v in type_counts.items()]
            lab['division_breakdown'] = [{'division': k, 'count': v} for k, v in div_counts.items() if v > 0]
            labs.append(lab)

        conn.close()
        return labs

    # ---------- MACHINE TRANSFER ----------
    def transfer_machine(self, data: Dict) -> Dict:
        self._auto_backup("transfer_machine")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT owner_customer, lab_id FROM instruments WHERE serial_number = ?", (data['serial_number'],))
            current = cursor.fetchone()
            if not current:
                return {"status": "error", "message": "Machine not found"}
            from_customer = current['owner_customer']
            from_lab_id = current['lab_id']
            to_customer = data['to_customer']
            to_lab_id = data.get('to_lab_id') or None

            cursor.execute("UPDATE instruments SET owner_customer = ?, lab_id = ? WHERE serial_number = ?",
                           (to_customer, to_lab_id, data['serial_number']))
            cursor.execute('''
                INSERT INTO transfer_log (serial_number, from_customer, to_customer, from_lab, to_lab, transferred_by, transferred_at, reason)
                VALUES (?,?,?,?,?,?,?,?)
            ''', (data['serial_number'], from_customer, to_customer, str(from_lab_id), str(to_lab_id),
                  data.get('transferred_by', 'system'), datetime.now().isoformat(), data.get('reason', '')))
            conn.commit()
            self._add_log('MACHINE_TRANSFER', f"Machine {data['serial_number']} transferred from {from_customer} to {to_customer}")
            return {"status": "success", "message": f"Machine transferred to {to_customer}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def get_transfer_log(self, limit=100) -> List[Dict]:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM transfer_log ORDER BY transferred_at DESC LIMIT ?", (limit,))
        logs = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return logs

    # ---------- MACHINE-LEVEL AGREEMENTS ----------
    def get_machine_agreements(self, serial_number: str) -> List[Dict]:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM machine_agreements WHERE serial_number = ? ORDER BY created_at DESC", (serial_number,))
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    def add_machine_agreement(self, data: Dict) -> Dict:
        self._auto_backup("add_machine_agreement")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            agreement_no = f"MA-{datetime.now().strftime('%Y%m%d')}-{secrets.token_hex(3)[:4].upper()}"
            cursor.execute('''
                INSERT INTO machine_agreements (serial_number, agreement_number, start_date, end_date, agreement_type, total_amount, status, created_at)
                VALUES (?,?,?,?,?,?,?,?)
            ''', (data['serial_number'], agreement_no, data.get('start_date'), data.get('end_date'),
                  data.get('agreement_type'), float(data.get('total_amount',0)), 'Active', datetime.now().isoformat()))
            conn.commit()
            self._add_log('MACHINE_AGREEMENT_ADD', f"Added agreement {agreement_no} for {data['serial_number']}")
            return {"status": "success", "message": "Agreement added"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    # ---------- CSV UPLOAD (BULK) ----------
    def upload_csv_data(self, csv_content: str) -> Dict:
        self._save_uploaded_csv_as_backup("bulk_upload", csv_content)
        self._auto_backup("bulk_csv_upload")
        conn = self._get_connection()
        cursor = conn.cursor()
        added_instruments = 0
        updated_instruments = 0
        added_services = 0
        try:
            reader = csv.DictReader(io.StringIO(csv_content), skipinitialspace=True)
            for row in reader:
                customer = row.get("Customer", "").strip()
                if not customer:
                    continue
                cursor.execute("INSERT OR IGNORE INTO customers (name, created_at) VALUES (?, ?)",
                               (customer, datetime.now().isoformat()))
                serial = row.get("Serial Number", "").strip()
                if not serial:
                    continue
                machine_type = row.get("Machine", "").strip() or "Unknown"
                model = row.get("Model", "").strip()
                division = row.get("Division", "").strip() or "GENERAL LAB"
                cursor.execute("SELECT * FROM instruments WHERE serial_number = ?", (serial,))
                existing = cursor.fetchone()
                if existing:
                    cursor.execute('''
                        UPDATE instruments SET machine_type=?, model=?, owner_customer=?, division=?
                        WHERE serial_number=?
                    ''', (machine_type, model, customer, division, serial))
                    updated_instruments += 1
                else:
                    cursor.execute('''
                        INSERT INTO instruments (serial_number, machine_type, model, owner_customer, division, is_secondary, status, created_at)
                        VALUES (?,?,?,?,?,?,?,?)
                    ''', (serial, machine_type, model, customer, division, 0, 'Active', datetime.now().isoformat()))
                    added_instruments += 1
                service_date_raw = row.get("Service Date", "").strip()
                if service_date_raw:
                    try:
                        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%B %d, %Y", "%d/%m/%Y"):
                            try:
                                dt = datetime.strptime(service_date_raw, fmt)
                                service_date = dt.strftime("%Y-%m-%d")
                                break
                            except:
                                continue
                        else:
                            service_date = service_date_raw
                    except:
                        service_date = service_date_raw
                    job_no = row.get("JOB NO", "").strip()
                    cursor.execute("SELECT id FROM service_tickets WHERE serial_number=? AND service_date=? AND job_no=?",
                                   (serial, service_date, job_no))
                    if not cursor.fetchone():
                        cursor.execute('''
                            INSERT INTO service_tickets (job_no, service_date, serial_number, job_actions, technicians, created_at)
                            VALUES (?,?,?,?,?,?)
                        ''', (job_no, service_date, serial, row.get("Job", ""), row.get("Tech_1", ""), datetime.now().isoformat()))
                        added_services += 1
                        cursor.execute("UPDATE instruments SET last_service = ? WHERE serial_number = ?", (service_date, serial))
                        try:
                            last_date = datetime.strptime(service_date, "%Y-%m-%d")
                            next_due = last_date + timedelta(days=180)
                            cursor.execute("UPDATE instruments SET next_service_due = ? WHERE serial_number = ?", (next_due.strftime("%Y-%m-%d"), serial))
                        except:
                            pass
            conn.commit()
            self._add_log('CSV_UPLOAD', f"Uploaded CSV: {added_instruments} new, {updated_instruments} updated, {added_services} services")
            return {"status": "success", "message": f"Processed: {added_instruments} new, {updated_instruments} updated, {added_services} services"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    # ---------- SEARCH / FILTER ----------
    def search_instruments(self, filters: Dict) -> List[Dict]:
        conn = self._get_connection()
        cursor = conn.cursor()
        query = "SELECT i.*, c.name as customer_name, l.lab_name FROM instruments i LEFT JOIN customers c ON i.owner_customer=c.name LEFT JOIN labs l ON i.lab_id=l.id WHERE 1=1"
        params = []
        if filters.get('search'):
            search = f"%{filters['search']}%"
            query += " AND (i.serial_number LIKE ? OR i.machine_type LIKE ? OR i.model LIKE ? OR i.owner_customer LIKE ?)"
            params.extend([search, search, search, search])
        if filters.get('division'):
            query += " AND i.division = ?"
            params.append(filters['division'])
        if filters.get('customer'):
            query += " AND i.owner_customer = ?"
            params.append(filters['customer'])
        query += " ORDER BY i.owner_customer, i.is_secondary"
        cursor.execute(query, params)
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    # ---------- EXPORT EXCEL ----------
    def export_customer_excel(self, customer_name: str) -> bytes:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Customer Details"
        customer = self.get_customer_detail(customer_name)['customer']
        ws1.append(["Customer Name", customer.get('name', '')])
        ws1.append(["Contact Person", customer.get('contact_person', '')])
        ws1.append(["Phone", customer.get('phone', '')])
        ws1.append(["Email", customer.get('email', '')])
        ws1.append(["Region", customer.get('region', '')])
        ws1.append(["Agreement Type", customer.get('agreement_type', '')])
        ws1.append(["Agreement Start", customer.get('service_agreement_start', '')])
        ws1.append(["Agreement End", customer.get('service_agreement_end', '')])

        ws2 = wb.create_sheet("Machines & Service")
        ws2.append(["Serial Number", "Machine Type", "Model", "Division", "Lab", "Status", "Last Service", "Next Service Due", "Warranty End"])
        instruments = self.get_all_instruments()
        for inst in instruments:
            if inst['owner_customer'] == customer_name:
                lab_name = ""
                if inst.get('lab_id'):
                    cur = self._get_connection().cursor()
                    cur.execute("SELECT lab_name FROM labs WHERE id=?", (inst['lab_id'],))
                    lab_row = cur.fetchone()
                    if lab_row:
                        lab_name = lab_row['lab_name']
                    cur.close()
                ws2.append([inst['serial_number'], inst['machine_type'], inst['model'], inst['division'],
                            lab_name, inst['status'], inst.get('last_service', ''), inst.get('next_service_due', ''), inst.get('warranty_end', '')])
        ws3 = wb.create_sheet("Service History")
        ws3.append(["Serial Number", "Service Date", "Job No", "Actions", "Technician(s)", "Type", "Notes"])
        for inst in instruments:
            if inst['owner_customer'] == customer_name:
                tickets = self._get_service_tickets(inst['serial_number'])
                for t in tickets:
                    ws3.append([inst['serial_number'], t['service_date'], t.get('job_no',''), t.get('job_actions',''),
                                t.get('technicians',''), t.get('service_type',''), t.get('service_notes','')])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def export_machine_report(self, serial_number: str) -> bytes:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Machine Service Report"
        conn = self._get_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM instruments WHERE serial_number=?", (serial_number,))
        machine = dict(cur.fetchone() or {})
        cur.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date", (serial_number,))
        services = [dict(r) for r in cur.fetchall()]
        conn.close()
        ws.append(["Machine Service Report"])
        ws.append([])
        ws.append(["Serial Number", machine.get('serial_number', '')])
        ws.append(["Machine Type", machine.get('machine_type', '')])
        ws.append(["Model", machine.get('model', '')])
        ws.append(["Customer", machine.get('owner_customer', '')])
        ws.append(["Division", machine.get('division', '')])
        ws.append(["Status", machine.get('status', '')])
        ws.append(["Warranty End", machine.get('warranty_end', '')])
        ws.append(["Last Service", machine.get('last_service', '')])
        ws.append(["Next Service Due", machine.get('next_service_due', '')])
        ws.append([])
        ws.append(["Service History"])
        ws.append(["Date", "Job No", "Actions", "Technician(s)", "Type", "Notes"])
        for s in services:
            ws.append([s['service_date'], s.get('job_no',''), s.get('job_actions',''),
                       s.get('technicians',''), s.get('service_type',''), s.get('service_notes','')])
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def _get_service_tickets(self, serial: str):
        conn = self._get_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date", (serial,))
        res = [dict(r) for r in cur.fetchall()]
        conn.close()
        return res

    # ---------- MODEL STATISTICS ----------
    def get_model_statistics(self, customer: str = None) -> Dict:
        conn = self._get_connection()
        cursor = conn.cursor()
        base_query = "FROM instruments i"
        params = []
        if customer:
            base_query += " WHERE i.owner_customer = ?"
            params.append(customer)

        query = f"SELECT machine_type, model, COUNT(*) as count {base_query} GROUP BY machine_type, model ORDER BY count DESC"
        cursor.execute(query, params)
        model_counts = [dict(r) for r in cursor.fetchall()]

        query2 = f"SELECT division, COUNT(*) as count {base_query} GROUP BY division"
        cursor.execute(query2, params)
        division_counts = [dict(r) for r in cursor.fetchall()]

        query3 = f"""
            SELECT i.machine_type, AVG((
                SELECT COUNT(*) FROM service_tickets s WHERE s.serial_number = i.serial_number
            )) as avg_services
            {base_query}
            GROUP BY i.machine_type
            ORDER BY avg_services DESC
        """
        cursor.execute(query3, params)
        service_avg = [dict(r) for r in cursor.fetchall()]

        conn.close()
        return {
            "model_counts": model_counts,
            "division_counts": division_counts,
            "service_avg": service_avg
        }

    # ---------- AI TECHNICIAN RECOMMENDATION (ADVANCED) ----------
    def recommend_technicians_advanced(self, serial_number: str, service_date: str = None, service_type: str = None, customer: str = None) -> Dict:
        conn = self._get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT machine_type, model, owner_customer FROM instruments WHERE serial_number=?", (serial_number,))
        machine = cursor.fetchone()
        if not machine:
            conn.close()
            return {'teams': [], 'weekend_warning': False}

        machine_type = machine['machine_type']
        owner_customer = customer or machine['owner_customer']

        # Get all technicians
        cursor.execute("SELECT DISTINCT technicians FROM service_tickets WHERE technicians IS NOT NULL AND technicians != ''")
        all_techs = [r['technicians'] for r in cursor.fetchall()]
        tech_list = []
        for t in all_techs:
            tech_list.extend([x.strip() for x in t.split(',') if x.strip()])
        tech_list = list(set(tech_list))

        tech_scores = {}
        for tech in tech_list:
            cursor.execute("SELECT COUNT(*) as count FROM service_tickets WHERE technicians LIKE ? AND serial_number = ?", (f"%{tech}%", serial_number))
            familiarity = cursor.fetchone()['count']
            cursor.execute("""
                SELECT COUNT(*) as count FROM service_tickets s
                JOIN instruments i ON s.serial_number = i.serial_number
                WHERE s.technicians LIKE ? AND i.machine_type = ?
            """, (f"%{tech}%", machine_type))
            type_familiarity = cursor.fetchone()['count']
            machine_score = familiarity * 0.4 + type_familiarity * 0.6

            cursor.execute("""
                SELECT COUNT(*) as positive FROM service_tickets 
                WHERE technicians LIKE ? AND (service_notes LIKE '%good%' OR service_notes LIKE '%excellent%')
            """, (f"%{tech}%",))
            positive = cursor.fetchone()['positive']
            cursor.execute("SELECT COUNT(*) as total FROM service_tickets WHERE technicians LIKE ?", (f"%{tech}%",))
            total = cursor.fetchone()['total'] or 1
            performance = positive / total if total > 0 else 0.5

            tech_scores[tech] = {
                'machine_score': machine_score,
                'performance': performance,
                'familiarity': familiarity + type_familiarity
            }

        import itertools
        teams = []
        for team in itertools.combinations(tech_list, 2):
            t1, t2 = team
            s1 = tech_scores.get(t1, {})
            s2 = tech_scores.get(t2, {})
            avg_score = (s1.get('machine_score', 0) + s2.get('machine_score', 0)) / 2
            avg_perf = (s1.get('performance', 0) + s2.get('performance', 0)) / 2
            cursor.execute("""
                SELECT COUNT(*) as count FROM service_tickets 
                WHERE technicians LIKE ? AND technicians LIKE ?
            """, (f"%{t1}%", f"%{t2}%"))
            together = cursor.fetchone()['count']
            bonus = min(together / 5, 0.3)
            final_score = (avg_score * 0.6 + avg_perf * 0.4) * (1 + bonus)
            teams.append({
                'team': [t1, t2],
                'score': round(final_score, 2),
                'details': {
                    't1_score': s1.get('machine_score', 0),
                    't2_score': s2.get('machine_score', 0),
                    'together': together,
                    'machine_familiarity': s1.get('familiarity',0)+s2.get('familiarity',0)
                }
            })

        teams.sort(key=lambda x: x['score'], reverse=True)
        weekend_warning = False
        if service_date:
            try:
                weekday = datetime.strptime(service_date, "%Y-%m-%d").weekday()
                if weekday >= 5:
                    weekend_warning = True
            except:
                pass

        conn.close()
        return {'teams': teams, 'weekend_warning': weekend_warning}

    # ---------- CUSTOMER EMAIL SUMMARY ----------
    def send_customer_service_summary(self, customer_name: str, custom_message: str = "") -> Dict:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE name=?", (customer_name,))
        customer = cursor.fetchone()
        if not customer:
            conn.close()
            return {"status": "error", "message": "Customer not found"}
        if not customer['email']:
            conn.close()
            return {"status": "error", "message": "Customer has no email address on file"}

        cursor.execute("SELECT * FROM instruments WHERE owner_customer=?", (customer_name,))
        machines = [dict(r) for r in cursor.fetchall()]
        for m in machines:
            cursor.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date DESC", (m['serial_number'],))
            m['services'] = [dict(r) for r in cursor.fetchall()]
            cursor.execute("SELECT * FROM machine_agreements WHERE serial_number=? AND status='Active'", (m['serial_number'],))
            m['agreements'] = [dict(r) for r in cursor.fetchall()]

        cursor.execute("SELECT * FROM service_agreements WHERE customer_name=? AND status='Active'", (customer_name,))
        customer_agreements = [dict(r) for r in cursor.fetchall()]
        conn.close()

        subject = f"Service Summary for {customer_name}"
        body = self._build_customer_summary_email(customer, machines, customer_agreements, custom_message)
        success = self._send_email(customer['email'], subject, body)
        if success:
            self._add_log('CUSTOMER_EMAIL', f"Sent service summary to {customer_name} ({customer['email']})")
            return {"status": "success", "message": f"Email sent to {customer['email']}"}
        else:
            return {"status": "error", "message": "Failed to send email"}

    def _build_customer_summary_email(self, customer, machines, customer_agreements, custom_message):
        html = f"""
        <html>
        <head><style>body{{font-family:Arial;}} .header{{background:#10b981;padding:15px;color:white;}} .card{{border:1px solid #ddd;padding:15px;border-radius:8px;margin:10px 0;}}</style></head>
        <body>
        <div class="header"><h2>Hayleys Life Sciences – Service Summary</h2></div>
        <p>Dear {customer['name']},</p>
        <p>Please find below a summary of your registered machines and their service status.</p>
        """
        if custom_message:
            html += f"<p><strong>Message:</strong> {custom_message}</p>"
        html += f"<p><strong>Contact Person:</strong> {customer.get('contact_person', 'N/A')}<br><strong>Phone:</strong> {customer.get('phone', 'N/A')}</p>"

        for m in machines:
            html += f"""
            <div class="card">
                <h3>{m['machine_type']} - {m['model'] or 'N/A'}</h3>
                <p><strong>Serial:</strong> {m['serial_number']}<br>
                <strong>Division:</strong> {m['division']}<br>
                <strong>Status:</strong> {m['status']}<br>
                <strong>Last Service:</strong> {m['last_service'] or 'Never'}<br>
                <strong>Next Service Due:</strong> {m['next_service_due'] or 'Not set'}</p>
                <h4>Service History ({len(m['services'])} records)</h4>
                <ul>
            """
            for s in m['services'][:5]:
                html += f"<li>{s['service_date']} – {s['job_actions']} (Technician: {s['technicians']})</li>"
            if len(m['services']) > 5:
                html += f"<li>... and {len(m['services'])-5} more.</li>"
            html += "</ul>"
            if m['agreements']:
                html += "<h4>Active Machine Agreements</h4><ul>"
                for a in m['agreements']:
                    html += f"<li>{a['agreement_type']} – {a['start_date']} to {a['end_date']} (Amount: LKR {a['total_amount']})</li>"
                html += "</ul>"
            html += "</div>"

        if customer_agreements:
            html += "<h3>Customer Service Agreements</h3><ul>"
            for a in customer_agreements:
                html += f"<li>{a['agreement_number']} – {a['agreement_type']} ({a['start_date']} to {a['end_date']})</li>"
            html += "</ul>"

        html += """
        <p>If you have any questions, please contact your service representative.</p>
        <p>Thank you for choosing Hayleys Life Sciences.</p>
        <p><small>This is an automated summary. Please do not reply directly to this email.</small></p>
        </body>
        </html>
        """
        return html

    def _send_email(self, to, subject, body_html):
        SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
        SMTP_PORT = int(os.environ.get('SMTP_PORT', 587))
        SMTP_USER = os.environ.get('SMTP_USER', 'your-email@gmail.com')
        SMTP_PASS = os.environ.get('SMTP_PASS', 'your-password')

        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = SMTP_USER
            msg['To'] = to
            msg.attach(MIMEText(body_html, 'html'))
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.sendmail(SMTP_USER, to, msg.as_string())
            return True
        except Exception as e:
            logger.error(f"Email send failed: {e}")
            return False

    # ---------- SCHEDULER & REMINDERS ----------
    def _start_scheduler(self):
        if not SCHEDULER_AVAILABLE:
            return
        self.scheduler = BackgroundScheduler()
        self.scheduler.add_job(self._daily_reminder_job, 'cron', hour=8, minute=0)
        self.scheduler.start()

    def _daily_reminder_job(self):
        if not SCHEDULER_AVAILABLE:
            return
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT email_recipients, reminder_days_before, send_to_customers FROM reminder_settings LIMIT 1")
        settings = cursor.fetchone()
        if not settings:
            return
        recipients = settings['email_recipients']
        days_before = settings['reminder_days_before'] or 7
        send_to_customers = settings['send_to_customers']
        conn.close()

        # Service due
        conn = self._get_connection()
        cursor = conn.cursor()
        limit_date = (datetime.now() + timedelta(days=days_before)).strftime("%Y-%m-%d")
        cursor.execute("""
            SELECT i.*, c.name as customer_name, c.email as customer_email
            FROM instruments i
            JOIN customers c ON i.owner_customer = c.name
            WHERE i.status='Active' AND i.next_service_due IS NOT NULL
              AND i.next_service_due <= ? AND i.next_service_due > date('now')
        """, (limit_date,))
        due_services = [dict(r) for r in cursor.fetchall()]
        for machine in due_services:
            subject = f"⚠️ Service Due Alert: {machine['machine_type']} - {machine['serial_number']}"
            body = self._build_service_reminder_email(machine, days_before)
            self._send_email(recipients, subject, body)
            self._log_reminder('service_due', machine['serial_number'], machine['customer_name'], recipients)
            if send_to_customers and machine.get('customer_email'):
                self._send_email(machine['customer_email'], subject, body)
        conn.close()

        # Agreement expiry
        conn = self._get_connection()
        cursor = conn.cursor()
        limit_date = (datetime.now() + timedelta(days=days_before)).strftime("%Y-%m-%d")
        cursor.execute("""
            SELECT sa.*, c.name as customer_name, c.email as customer_email
            FROM service_agreements sa
            JOIN customers c ON sa.customer_name = c.name
            WHERE sa.status='Active' AND sa.end_date IS NOT NULL
              AND sa.end_date <= ? AND sa.end_date > date('now')
        """, (limit_date,))
        expiring = [dict(r) for r in cursor.fetchall()]
        for agreement in expiring:
            subject = f"📄 Agreement Expiring: {agreement['agreement_number']} - {agreement['customer_name']}"
            body = self._build_agreement_reminder_email(agreement, days_before)
            self._send_email(recipients, subject, body)
            self._log_reminder('agreement_expiry', agreement['agreement_number'], agreement['customer_name'], recipients)
            if send_to_customers and agreement.get('customer_email'):
                self._send_email(agreement['customer_email'], subject, body)
        conn.close()

    def _build_service_reminder_email(self, machine, days):
        return f"""
        <html>
        <head><style>body{{font-family:Arial;}} .card{{border:1px solid #ddd;padding:15px;border-radius:8px;}}</style></head>
        <body>
        <h2>🔧 Service Reminder</h2>
        <p>The following machine is due for service within {days} days.</p>
        <div class="card">
            <p><strong>Serial Number:</strong> {machine['serial_number']}</p>
            <p><strong>Machine Type:</strong> {machine['machine_type']}</p>
            <p><strong>Model:</strong> {machine['model'] or 'N/A'}</p>
            <p><strong>Customer:</strong> {machine['customer_name']}</p>
            <p><strong>Next Service Due:</strong> {machine['next_service_due']}</p>
            <p><strong>Last Service:</strong> {machine['last_service'] or 'Never'}</p>
            <p><strong>Division:</strong> {machine['division']}</p>
        </div>
        <p>Please schedule service promptly. <a href="http://localhost:8080">Login to system</a> to assign a team.</p>
        </body>
        </html>
        """

    def _build_agreement_reminder_email(self, agreement, days):
        return f"""
        <html>
        <head><style>body{{font-family:Arial;}} .card{{border:1px solid #ddd;padding:15px;border-radius:8px;}}</style></head>
        <body>
        <h2>📄 Service Agreement Expiring</h2>
        <p>The following service agreement will expire within {days} days.</p>
        <div class="card">
            <p><strong>Agreement:</strong> {agreement['agreement_number']}</p>
            <p><strong>Customer:</strong> {agreement['customer_name']}</p>
            <p><strong>Type:</strong> {agreement['agreement_type']}</p>
            <p><strong>Start Date:</strong> {agreement['start_date']}</p>
            <p><strong>End Date:</strong> {agreement['end_date']}</p>
            <p><strong>Total Amount:</strong> LKR {agreement['total_amount'] or 0}</p>
        </div>
        <p>Please contact the customer for renewal. <a href="http://localhost:8080">Login to system</a> to update.</p>
        </body>
        </html>
        """

    def _log_reminder(self, reminder_type, serial, customer, recipients):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO reminder_history (reminder_type, serial_number, customer_name, sent_at, recipients) VALUES (?,?,?,?,?)",
                       (reminder_type, serial, customer, datetime.now().isoformat(), recipients))
        conn.commit()
        conn.close()

    def get_reminder_history(self, limit=50):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM reminder_history ORDER BY sent_at DESC LIMIT ?", (limit,))
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    # ---------- CUSTOMER 360 (FOR SALES) ----------
    def get_customer_360(self, customer_name: str) -> Dict:
        detail = self.get_customer_detail(customer_name)
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM sales_opportunities WHERE customer_name=? AND status='Open'", (customer_name,))
        opportunities = [dict(r) for r in cursor.fetchall()]
        cursor.execute("""
            SELECT AVG(rating) as avg_rating, COUNT(*) as total_feedback
            FROM service_surveys
            WHERE customer_name=?
        """, (customer_name,))
        satisfaction = cursor.fetchone()
        conn.close()
        detail['opportunities'] = opportunities
        detail['satisfaction'] = dict(satisfaction) if satisfaction else {'avg_rating': None, 'total_feedback': 0}
        return detail

    # ---------- CORE CRUD (Existing) ----------
    def reset_database(self) -> Dict:
        self._auto_backup("reset_database")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("PRAGMA foreign_keys = OFF")
            for table in ['transfer_log', 'machine_agreements', 'service_tickets', 'instruments', 'labs', 'service_agreements', 'customers', 'system_logs', 'users', 'customer_preferences', 'service_surveys', 'reminder_settings', 'reminder_history', 'machine_health', 'sales_opportunities']:
                cursor.execute(f"DROP TABLE IF EXISTS {table}")
            cursor.execute("PRAGMA foreign_keys = ON")
            conn.commit()
            conn.close()
            self._init_database()
            return {"status": "success", "message": "Database reset and CSV re-imported"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def get_all_customers(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT c.*,
                   COUNT(DISTINCT CASE WHEN i.is_secondary=0 THEN i.id END) as primary_count,
                   COUNT(DISTINCT CASE WHEN i.is_secondary=1 THEN i.id END) as secondary_count
            FROM customers c LEFT JOIN instruments i ON c.name=i.owner_customer
            GROUP BY c.id ORDER BY c.name
        ''')
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    def get_customer_detail(self, name: str):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE name=?", (name,))
        customer = dict(cursor.fetchone() or {})
        cursor.execute("SELECT * FROM service_agreements WHERE customer_name=? ORDER BY created_at DESC", (name,))
        agreements = [dict(r) for r in cursor.fetchall()]
        cursor.execute("SELECT * FROM instruments WHERE owner_customer=? AND is_secondary=0 ORDER BY created_at DESC", (name,))
        primaries = [dict(r) for r in cursor.fetchall()]
        total_secondary = 0
        total_services = 0
        pending = 0
        for p in primaries:
            cursor.execute("SELECT * FROM instruments WHERE parent_serial=? AND is_secondary=1", (p['serial_number'],))
            secs = [dict(r) for r in cursor.fetchall()]
            p['secondary_machines'] = secs
            total_secondary += len(secs)
            cursor.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date DESC", (p['serial_number'],))
            p['service_history'] = [dict(r) for r in cursor.fetchall()]
            total_services += len(p['service_history'])
            if p.get('next_service_due') and p['next_service_due'] <= datetime.now().strftime("%Y-%m-%d"):
                pending += 1
            for s in secs:
                if s.get('parent_serial'):
                    cursor.execute("SELECT * FROM machine_agreements WHERE serial_number = ? AND status='Active'", (s['parent_serial'],))
                    s['primary_agreements'] = [dict(r) for r in cursor.fetchall()]
                else:
                    s['primary_agreements'] = []
                cursor.execute("SELECT * FROM service_tickets WHERE serial_number=? ORDER BY service_date DESC", (s['serial_number'],))
                s['service_history'] = [dict(r) for r in cursor.fetchall()]
                total_services += len(s['service_history'])
                if s.get('next_service_due') and s['next_service_due'] <= datetime.now().strftime("%Y-%m-%d"):
                    pending += 1
        cursor.execute("SELECT * FROM labs WHERE customer_name=? ORDER BY lab_name", (name,))
        labs = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return {
            "customer": customer,
            "agreements": agreements,
            "primary_machines": primaries,
            "labs": labs,
            "total_secondary": total_secondary,
            "total_services": total_services,
            "pending_services": pending
        }

    def create_customer(self, data):
        self._auto_backup("create_customer")
        conn = self._get_connection()
        try:
            conn.execute('''
                INSERT INTO customers (name, contact_person, phone, email, region, service_agreement_start, service_agreement_end, agreement_type, created_at)
                VALUES (?,?,?,?,?,?,?,?,?)
            ''', (data['name'], data.get('contact_person',''), data.get('phone',''), data.get('email',''),
                  data.get('region','Western'), data.get('service_agreement_start',''), data.get('service_agreement_end',''),
                  data.get('agreement_type','Basic'), datetime.now().isoformat()))
            conn.commit()
            self._add_log('CUSTOMER_ADD', f"Added {data['name']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def update_customer(self, data):
        self._auto_backup("update_customer")
        conn = self._get_connection()
        try:
            conn.execute('''
                UPDATE customers SET contact_person=?, phone=?, email=?, region=?,
                service_agreement_start=?, service_agreement_end=?, agreement_type=?
                WHERE name=?
            ''', (data.get('contact_person',''), data.get('phone',''), data.get('email',''),
                  data.get('region','Western'), data.get('service_agreement_start',''),
                  data.get('service_agreement_end',''), data.get('agreement_type','Basic'), data['name']))
            conn.commit()
            self._add_log('CUSTOMER_UPDATE', f"Updated {data['name']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def delete_customer(self, data):
        self._auto_backup("delete_customer")
        conn = self._get_connection()
        try:
            conn.execute("DELETE FROM customers WHERE name=?", (data['name'],))
            conn.commit()
            self._add_log('CUSTOMER_DELETE', f"Deleted {data['name']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def add_service_agreement(self, data):
        self._auto_backup("add_service_agreement")
        conn = self._get_connection()
        try:
            agreement_no = f"SA-{datetime.now().strftime('%Y%m%d')}-{secrets.token_hex(3)[:4].upper()}"
            conn.execute('''
                INSERT INTO service_agreements (customer_name, agreement_number, start_date, end_date, agreement_type, total_amount, status, created_at)
                VALUES (?,?,?,?,?,?,?,?)
            ''', (data['customer_name'], agreement_no, data.get('start_date'), data.get('end_date'),
                  data.get('agreement_type'), float(data.get('total_amount',0)), 'Active', datetime.now().isoformat()))
            conn.execute('''
                UPDATE customers SET service_agreement_start=?, service_agreement_end=?, agreement_type=?
                WHERE name=?
            ''', (data.get('start_date'), data.get('end_date'), data.get('agreement_type'), data['customer_name']))
            conn.commit()
            self._add_log('AGREEMENT_ADD', f"Added {agreement_no}")
            return True
        except:
            return False
        finally:
            conn.close()

    def get_all_instruments(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM instruments ORDER BY owner_customer, is_secondary")
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    def add_primary_instrument(self, data):
        self._auto_backup("add_primary_instrument")
        conn = self._get_connection()
        try:
            cur = conn.cursor()
            cur.execute("SELECT 1 FROM instruments WHERE serial_number=?", (data['serial_number'],))
            if cur.fetchone():
                return False
            cur.execute('''
                INSERT INTO instruments (serial_number, machine_type, model, owner_customer, division, is_secondary, status, warranty_end, lab_id, created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            ''', (data['serial_number'], data['machine_type'], data.get('model',''), data['owner_customer'],
                  data.get('division','GENERAL LAB'), 0, data.get('status','Active'), data.get('warranty_end',''),
                  data.get('lab_id'), datetime.now().isoformat()))
            conn.commit()
            self._add_log('MACHINE_ADD', f"Added primary {data['serial_number']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def update_instrument(self, data):
        self._auto_backup("update_instrument")
        conn = self._get_connection()
        try:
            old_serial = data['old_serial_number']
            new_serial = data.get('serial_number', old_serial).strip()
            if new_serial != old_serial:
                cur = conn.cursor()
                cur.execute("SELECT 1 FROM instruments WHERE serial_number=?", (new_serial,))
                if cur.fetchone():
                    return False
                conn.execute("UPDATE instruments SET serial_number=? WHERE serial_number=?", (new_serial, old_serial))
                conn.execute("UPDATE service_tickets SET serial_number=? WHERE serial_number=?", (new_serial, old_serial))
                conn.execute("UPDATE machine_agreements SET serial_number=? WHERE serial_number=?", (new_serial, old_serial))
                conn.execute("UPDATE instruments SET parent_serial=? WHERE parent_serial=?", (new_serial, old_serial))
                conn.execute('''
                    UPDATE instruments SET machine_type=?, model=?, division=?, status=?, warranty_end=?, lab_id=?
                    WHERE serial_number=?
                ''', (data.get('machine_type'), data.get('model'), data.get('division','GENERAL LAB'),
                      data.get('status','Active'), data.get('warranty_end',''), data.get('lab_id'), new_serial))
            else:
                conn.execute('''
                    UPDATE instruments SET machine_type=?, model=?, division=?, status=?, warranty_end=?, lab_id=?
                    WHERE serial_number=?
                ''', (data.get('machine_type'), data.get('model'), data.get('division','GENERAL LAB'),
                      data.get('status','Active'), data.get('warranty_end',''), data.get('lab_id'), old_serial))
            conn.commit()
            self._add_log('MACHINE_UPDATE', f"Updated {old_serial}")
            return True
        except:
            return False
        finally:
            conn.close()

    def delete_instrument(self, data):
        self._auto_backup("delete_instrument")
        conn = self._get_connection()
        try:
            conn.execute("DELETE FROM instruments WHERE serial_number=?", (data['serial_number'],))
            conn.commit()
            self._add_log('MACHINE_DELETE', f"Deleted {data['serial_number']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def add_secondary_instrument(self, data: Dict) -> Dict:
        self._auto_backup("add_secondary_instrument")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT owner_customer, is_secondary, lab_id FROM instruments WHERE serial_number = ?", (data['parent_serial'],))
            parent = cursor.fetchone()
            if not parent:
                return {"status": "error", "message": "Parent machine not found"}
            if parent['is_secondary'] == 1:
                return {"status": "error", "message": "Cannot attach to a secondary machine. Select a primary."}
            cursor.execute("SELECT 1 FROM instruments WHERE serial_number = ?", (data['serial_number'],))
            if cursor.fetchone():
                return {"status": "error", "message": f"Serial {data['serial_number']} already exists"}
            cursor.execute('''
                INSERT INTO instruments
                (serial_number, machine_type, model, owner_customer, division, parent_serial, is_secondary, description, status, lab_id, created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            ''', (data['serial_number'], data['machine_type'], data.get('model', ''),
                  parent['owner_customer'], data.get('division', 'GENERAL LAB'),
                  data['parent_serial'], 1, data.get('description', ''), 'Active', parent['lab_id'], datetime.now().isoformat()))
            conn.commit()
            self._add_log('MACHINE_ADD', f"Added secondary {data['serial_number']} under {data['parent_serial']}")
            return {"status": "success", "message": f"Secondary {data['serial_number']} added"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def convert_primary_to_secondary(self, serial: str, target: str) -> Dict:
        self._auto_backup("convert_primary_to_secondary")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT * FROM instruments WHERE serial_number = ?", (serial,))
            src = cursor.fetchone()
            if not src:
                return {"status": "error", "message": "Source not found"}
            if src['is_secondary']:
                return {"status": "error", "message": "Already secondary"}
            cursor.execute("SELECT COUNT(*) FROM instruments WHERE parent_serial = ?", (serial,))
            if cursor.fetchone()[0] > 0:
                return {"status": "error", "message": "Machine has secondary children. Remove them first."}
            cursor.execute("SELECT * FROM instruments WHERE serial_number = ?", (target,))
            tgt = cursor.fetchone()
            if not tgt or tgt['is_secondary']:
                return {"status": "error", "message": "Target must be a primary machine"}
            if tgt['owner_customer'] != src['owner_customer']:
                return {"status": "error", "message": "Different customers"}
            cursor.execute("UPDATE instruments SET is_secondary=1, parent_serial=?, lab_id=? WHERE serial_number=?", (target, tgt['lab_id'], serial))
            conn.commit()
            self._add_log('MACHINE_CONVERT', f"Converted {serial} to secondary under {target}")
            return {"status": "success", "message": f"{serial} is now secondary under {target}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def promote_secondary_to_primary(self, serial: str) -> Dict:
        self._auto_backup("promote_secondary_to_primary")
        conn = self._get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT * FROM instruments WHERE serial_number = ?", (serial,))
            m = cursor.fetchone()
            if not m or not m['is_secondary']:
                return {"status": "error", "message": "Not a secondary machine"}
            cursor.execute("SELECT COUNT(*) FROM instruments WHERE parent_serial = ?", (serial,))
            if cursor.fetchone()[0] > 0:
                return {"status": "error", "message": "Machine has its own secondaries. Remove them first."}
            cursor.execute("UPDATE instruments SET is_secondary=0, parent_serial=NULL WHERE serial_number=?", (serial,))
            conn.commit()
            self._add_log('MACHINE_PROMOTE', f"Promoted {serial} to primary")
            return {"status": "success", "message": f"{serial} is now primary"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            conn.close()

    def add_service_ticket(self, data):
        self._auto_backup("add_service_ticket")
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO service_tickets (serial_number, job_no, service_date, job_actions, technicians, service_notes, service_type, created_at)
                VALUES (?,?,?,?,?,?,?,?)
            ''', (data['serial_number'], data.get('job_no'), data.get('service_date'), data.get('job_actions'),
                  data.get('technicians'), data.get('service_notes'), data.get('service_type','Routine'), datetime.now().isoformat()))
            service_date = data.get('service_date')
            if service_date:
                try:
                    next_due = (datetime.strptime(service_date, "%Y-%m-%d") + timedelta(days=180)).strftime("%Y-%m-%d")
                except:
                    next_due = (datetime.now() + timedelta(days=180)).strftime("%Y-%m-%d")
                cursor.execute("UPDATE instruments SET last_service=?, next_service_due=? WHERE serial_number=?", (service_date, next_due, data['serial_number']))
            conn.commit()
            self._add_log('SERVICE_ADD', f"Service for {data['serial_number']}")
            return True
        except Exception as e:
            logger.error(f"Service add error: {e}")
            return False
        finally:
            conn.close()

    def update_service_ticket(self, data):
        self._auto_backup("update_service_ticket")
        conn = self._get_connection()
        try:
            conn.execute('''
                UPDATE service_tickets SET service_date=?, job_actions=?, technicians=?, service_notes=?, service_type=?
                WHERE id=?
            ''', (data.get('service_date'), data.get('job_actions'), data.get('technicians'),
                  data.get('service_notes'), data.get('service_type'), data['id']))
            if data.get('service_date'):
                try:
                    next_due = (datetime.strptime(data['service_date'], "%Y-%m-%d") + timedelta(days=180)).strftime("%Y-%m-%d")
                    conn.execute("UPDATE instruments SET next_service_due=? WHERE serial_number=(SELECT serial_number FROM service_tickets WHERE id=?)",
                                 (next_due, data['id']))
                except:
                    pass
            conn.commit()
            self._add_log('SERVICE_UPDATE', f"Updated ticket {data['id']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def delete_service_ticket(self, data):
        self._auto_backup("delete_service_ticket")
        conn = self._get_connection()
        try:
            conn.execute("DELETE FROM service_tickets WHERE id=?", (data['id'],))
            conn.commit()
            self._add_log('SERVICE_DELETE', f"Deleted ticket {data['id']}")
            return True
        except:
            return False
        finally:
            conn.close()

    def get_due_services(self, days=30):
        conn = self._get_connection()
        cursor = conn.cursor()
        limit = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
        cursor.execute('''
            SELECT i.*, c.name as customer_name FROM instruments i
            JOIN customers c ON i.owner_customer=c.name
            WHERE i.next_service_due IS NOT NULL AND i.next_service_due <= ? AND i.status='Active'
            ORDER BY i.next_service_due
        ''', (limit,))
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    def get_expiring_agreements(self, days=30):
        conn = self._get_connection()
        cursor = conn.cursor()
        limit = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
        cursor.execute('''
            SELECT sa.*, c.name as customer_name FROM service_agreements sa
            JOIN customers c ON sa.customer_name=c.name
            WHERE sa.end_date IS NOT NULL AND sa.end_date <= ? AND sa.status='Active'
        ''', (limit,))
        res = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return res

    def get_technician_performance(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT technicians, COUNT(*) as service_count FROM service_tickets
            WHERE technicians IS NOT NULL AND technicians!='' GROUP BY technicians ORDER BY service_count DESC
        ''')
        ranking = [dict(r) for r in cursor.fetchall()]
        total = sum(r['service_count'] for r in ranking)
        recommendations = []
        if ranking:
            recommendations.append({"type":"success","title":"🏆 Top Performer","message":f"{ranking[0]['technicians']} with {ranking[0]['service_count']} services"})
        due = self.get_due_services(days=0)
        if due:
            top_due = due[0]
            recommendations.append({"type":"warning","title":"⚠️ Overdue Service","message":f"Machine {top_due['serial_number']} (customer: {top_due['customer_name']}) is overdue for service."})
        conn.close()
        return {"total_technicians":len(ranking),"total_services":total,"technician_rankings":ranking,"ai_recommendations":recommendations}

    def get_dashboard_stats(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        stats = {
            "total_customers": cursor.execute("SELECT COUNT(*) FROM customers").fetchone()[0],
            "primary_assets": cursor.execute("SELECT COUNT(*) FROM instruments WHERE is_secondary=0").fetchone()[0],
            "secondary_assets": cursor.execute("SELECT COUNT(*) FROM instruments WHERE is_secondary=1").fetchone()[0],
            "total_service_tickets": cursor.execute("SELECT COUNT(*) FROM service_tickets").fetchone()[0],
            "pending_services": cursor.execute("SELECT COUNT(*) FROM instruments WHERE next_service_due IS NOT NULL AND next_service_due <= date('now') AND status='Active'").fetchone()[0],
            "active_agreements": cursor.execute("SELECT COUNT(*) FROM service_agreements WHERE status='Active'").fetchone()[0]
        }
        conn.close()
        return stats

    def get_service_trend(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        data = []
        for i in range(1,13):
            cnt = cursor.execute("SELECT COUNT(*) FROM service_tickets WHERE strftime('%m', service_date)=?", (f"{i:02d}",)).fetchone()[0]
            data.append(cnt)
        conn.close()
        return {"labels":months, "data":data}

    def get_top_customers(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT c.name, COUNT(s.id) as service_count FROM customers c
            LEFT JOIN instruments i ON c.name=i.owner_customer
            LEFT JOIN service_tickets s ON i.serial_number=s.serial_number
            GROUP BY c.name ORDER BY service_count DESC LIMIT 10
        ''')
        rows = cursor.fetchall()
        conn.close()
        return {"labels":[r[0] for r in rows], "data":[r[1] for r in rows]}

    # ---------- TECHNICIAN SUMMARY FOR REPORTS ----------
    def get_technician_summary(self, technician: str, years: int = 3) -> Dict:
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as total FROM service_tickets WHERE technicians LIKE ?", (f"%{technician}%",))
        total = cursor.fetchone()['total']
        current_year = datetime.now().year
        start_year = current_year - years + 1
        cursor.execute("""
            SELECT strftime('%Y', service_date) as year, COUNT(*) as count
            FROM service_tickets
            WHERE technicians LIKE ? AND service_date IS NOT NULL AND strftime('%Y', service_date) >= ?
            GROUP BY year ORDER BY year
        """, (f"%{technician}%", str(start_year)))
        yearly = [dict(r) for r in cursor.fetchall()]
        cursor.execute("""
            SELECT i.machine_type, COUNT(s.id) as count
            FROM service_tickets s JOIN instruments i ON s.serial_number = i.serial_number
            WHERE s.technicians LIKE ? AND strftime('%Y', s.service_date) >= ?
            GROUP BY i.machine_type ORDER BY count DESC
        """, (f"%{technician}%", str(start_year)))
        machine_types = [dict(r) for r in cursor.fetchall()]
        cursor.execute("""
            SELECT i.owner_customer, COUNT(s.id) as count
            FROM service_tickets s JOIN instruments i ON s.serial_number = i.serial_number
            WHERE s.technicians LIKE ? AND strftime('%Y', s.service_date) >= ?
            GROUP BY i.owner_customer ORDER BY count DESC LIMIT 10
        """, (f"%{technician}%", str(start_year)))
        top_customers = [dict(r) for r in cursor.fetchall()]
        cursor.execute("""
            SELECT s.*, i.machine_type, i.owner_customer
            FROM service_tickets s JOIN instruments i ON s.serial_number = i.serial_number
            WHERE s.technicians LIKE ? ORDER BY s.service_date DESC LIMIT 20
        """, (f"%{technician}%",))
        recent = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return {
            "technician": technician,
            "total_services": total,
            "yearly_breakdown": yearly,
            "machine_type_breakdown": machine_types,
            "top_customers": top_customers,
            "recent_services": recent,
            "report_period": f"{start_year} - {current_year}"
        }

    def generate_technician_pdf(self, technician: str, years: int = 3) -> bytes:
        if not PDF_AVAILABLE:
            raise RuntimeError("PDF libraries not installed. Run: pip install fpdf matplotlib")
        data = self.get_technician_summary(technician, years)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, f"Technician Performance Report: {technician}", ln=True, align='C')
        pdf.set_font("Arial", "", 10)
        pdf.cell(200, 8, f"Period: {data['report_period']} | Total Services: {data['total_services']}", ln=True, align='C')
        pdf.ln(5)
        if data['yearly_breakdown']:
            fig, ax = plt.subplots(figsize=(6, 3))
            years = [str(item['year']) for item in data['yearly_breakdown']]
            counts = [item['count'] for item in data['yearly_breakdown']]
            ax.bar(years, counts, color='#10b981')
            ax.set_title('Services per Year')
            ax.set_ylabel('Count')
            ax.set_xlabel('Year')
            plt.tight_layout()
            img = io.BytesIO()
            plt.savefig(img, format='png')
            plt.close()
            img.seek(0)
            img_path = f"/tmp/{technician}_yearly.png"
            with open(img_path, 'wb') as f:
                f.write(img.getvalue())
            pdf.image(img_path, x=10, y=pdf.get_y()+2, w=190)
            pdf.ln(40)
        if data['machine_type_breakdown']:
            fig, ax = plt.subplots(figsize=(5, 4))
            labels = [item['machine_type'] for item in data['machine_type_breakdown']]
            sizes = [item['count'] for item in data['machine_type_breakdown']]
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            ax.set_title('Machine Type Distribution')
            plt.tight_layout()
            img = io.BytesIO()
            plt.savefig(img, format='png')
            plt.close()
            img.seek(0)
            img_path = f"/tmp/{technician}_pie.png"
            with open(img_path, 'wb') as f:
                f.write(img.getvalue())
            pdf.add_page()
            pdf.cell(200, 10, "Machine Type Distribution", ln=True, align='C')
            pdf.image(img_path, x=10, y=pdf.get_y()+5, w=190)
            pdf.ln(50)
        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(100, 10, "Top Customers", ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.cell(80, 8, "Customer", border=1)
        pdf.cell(40, 8, "Services", border=1, ln=True)
        for item in data['top_customers']:
            pdf.cell(80, 8, item['owner_customer'], border=1)
            pdf.cell(40, 8, str(item['count']), border=1, ln=True)
        pdf.ln(5)
        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(100, 10, "Recent Services", ln=True)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(25, 8, "Date", border=1)
        pdf.cell(40, 8, "Machine", border=1)
        pdf.cell(40, 8, "Customer", border=1)
        pdf.cell(40, 8, "Actions", border=1, ln=True)
        pdf.set_font("Arial", "", 8)
        for s in data['recent_services']:
            pdf.cell(25, 8, s['service_date'], border=1)
            pdf.cell(40, 8, s['machine_type'], border=1)
            pdf.cell(40, 8, s['owner_customer'], border=1)
            pdf.cell(40, 8, s['job_actions'][:30], border=1, ln=True)
        return pdf.output(dest='S').encode('latin1')

    def generate_technician_excel(self, technician: str, years: int = 3) -> bytes:
        data = self.get_technician_summary(technician, years)
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Summary"
        ws1.append(["Technician", technician])
        ws1.append(["Total Services", data['total_services']])
        ws1.append(["Report Period", data['report_period']])
        ws1.append([])
        ws1.append(["Year", "Count"])
        for item in data['yearly_breakdown']:
            ws1.append([item['year'], item['count']])
        ws1.append([])
        ws1.append(["Machine Type", "Count"])
        for item in data['machine_type_breakdown']:
            ws1.append([item['machine_type'], item['count']])
        ws2 = wb.create_sheet("Top Customers")
        ws2.append(["Customer", "Services"])
        for item in data['top_customers']:
            ws2.append([item['owner_customer'], item['count']])
        ws3 = wb.create_sheet("Recent Services")
        ws3.append(["Date", "Machine Type", "Customer", "Job No", "Actions"])
        for s in data['recent_services']:
            ws3.append([s['service_date'], s['machine_type'], s['owner_customer'], s.get('job_no',''), s.get('job_actions','')])
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    # ========== NEW: Advanced Analytics Endpoints ==========

    def get_yearly_trend(self, year: str = None, technician: str = None) -> Dict:
        """Get yearly service trend data for analytics page"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT strftime('%m', service_date) as month, COUNT(*) as count
            FROM service_tickets
            WHERE service_date IS NOT NULL
        """
        params = []
        
        if year:
            query += " AND strftime('%Y', service_date) = ?"
            params.append(year)
        if technician:
            query += " AND technicians LIKE ?"
            params.append(f"%{technician}%")
        
        query += " GROUP BY month ORDER BY month"
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        # Fill in missing months
        month_data = {str(i).zfill(2): 0 for i in range(1, 13)}
        for row in rows:
            month_data[row['month']] = row['count']
        
        return {
            "labels": ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
            "data": [month_data[str(i).zfill(2)] for i in range(1, 13)]
        }

    def get_model_distribution_by_year(self, year: str = None) -> Dict:
        """Get model distribution for analytics page"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT i.machine_type, COUNT(*) as count
            FROM instruments i
            LEFT JOIN service_tickets s ON i.serial_number = s.serial_number
            WHERE 1=1
        """
        params = []
        
        if year:
            query += " AND strftime('%Y', s.service_date) = ?"
            params.append(year)
        
        query += " GROUP BY i.machine_type ORDER BY count DESC LIMIT 10"
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        return {
            "labels": [row['machine_type'] for row in rows],
            "data": [row['count'] for row in rows]
        }

    def get_tech_performance_by_year(self, year: str = None) -> Dict:
        """Get technician performance for analytics page"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT technicians, COUNT(*) as count
            FROM service_tickets
            WHERE technicians IS NOT NULL AND technicians != ''
        """
        params = []
        
        if year:
            query += " AND strftime('%Y', service_date) = ?"
            params.append(year)
        
        query += " GROUP BY technicians ORDER BY count DESC LIMIT 10"
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        return {
            "labels": [row['technicians'] for row in rows],
            "data": [row['count'] for row in rows]
        }

    def get_all_technicians(self) -> List[str]:
        """Get list of all unique technicians"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT DISTINCT technicians 
            FROM service_tickets 
            WHERE technicians IS NOT NULL AND technicians != ''
        """)
        
        techs = set()
        for row in cursor.fetchall():
            for t in row['technicians'].split(','):
                if t.strip():
                    techs.add(t.strip())
        
        conn.close()
        return sorted(list(techs))

    # ========== NEW: Enhanced Customer Filtering ==========

    def get_customers_by_region(self, region: str = None) -> List[Dict]:
        """Get customers filtered by region"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = "SELECT * FROM customers"
        params = []
        
        if region:
            query += " WHERE region = ?"
            params.append(region)
        
        query += " ORDER BY name"
        
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def get_machines_by_type(self, machine_type: str = None, customer: str = None) -> List[Dict]:
        """Get machines filtered by type and customer"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT i.*, c.name as customer_name, l.lab_name 
            FROM instruments i 
            LEFT JOIN customers c ON i.owner_customer = c.name 
            LEFT JOIN labs l ON i.lab_id = l.id 
            WHERE 1=1
        """
        params = []
        
        if machine_type:
            query += " AND i.machine_type LIKE ?"
            params.append(f"%{machine_type}%")
        
        if customer:
            query += " AND i.owner_customer = ?"
            params.append(customer)
        
        query += " ORDER BY i.owner_customer"
        
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    # ========== NEW: Export Analytics ==========

    def export_analytics_excel(self, year: str = None, technician: str = None) -> bytes:
        """Export analytics data to Excel"""
        wb = openpyxl.Workbook()
        
        # Yearly trend
        ws1 = wb.active
        ws1.title = "Yearly Trend"
        trend = self.get_yearly_trend(year, technician)
        ws1.append(["Month", "Services"])
        for label, data in zip(trend['labels'], trend['data']):
            ws1.append([label, data])
        
        # Model distribution
        ws2 = wb.create_sheet("Model Distribution")
        models = self.get_model_distribution_by_year(year)
        ws2.append(["Model", "Count"])
        for label, data in zip(models['labels'], models['data']):
            ws2.append([label, data])
        
        # Tech performance
        ws3 = wb.create_sheet("Technician Performance")
        techs = self.get_tech_performance_by_year(year)
        ws3.append(["Technician", "Services"])
        for label, data in zip(techs['labels'], techs['data']):
            ws3.append([label, data])
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def export_analytics_pdf(self, year: str = None, technician: str = None) -> bytes:
        """Export analytics data to PDF"""
        if not PDF_AVAILABLE:
            raise RuntimeError("PDF libraries not installed")
        
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, "Analytics Report", ln=True, align='C')
        
        # Yearly trend
        trend = self.get_yearly_trend(year, technician)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(200, 10, "Yearly Trend", ln=True)
        pdf.set_font("Arial", "", 10)
        
        # Create bar chart
        fig, ax = plt.subplots(figsize=(6, 3))
        ax.bar(trend['labels'], trend['data'], color='#10b981')
        ax.set_title('Services by Month')
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png')
        plt.close()
        img.seek(0)
        img_path = f"/tmp/analytics_trend.png"
        with open(img_path, 'wb') as f:
            f.write(img.getvalue())
        pdf.image(img_path, x=10, y=pdf.get_y()+2, w=190)
        pdf.ln(30)
        
        # Model distribution
        models = self.get_model_distribution_by_year(year)
        if models['labels']:
            fig, ax = plt.subplots(figsize=(5, 4))
            ax.pie(models['data'], labels=models['labels'], autopct='%1.1f%%')
            ax.set_title('Model Distribution')
            plt.tight_layout()
            img = io.BytesIO()
            plt.savefig(img, format='png')
            plt.close()
            img.seek(0)
            img_path = f"/tmp/analytics_models.png"
            with open(img_path, 'wb') as f:
                f.write(img.getvalue())
            pdf.add_page()
            pdf.cell(200, 10, "Model Distribution", ln=True, align='C')
            pdf.image(img_path, x=10, y=pdf.get_y()+2, w=190)
        
        # Tech performance
        techs = self.get_tech_performance_by_year(year)
        if techs['labels']:
            fig, ax = plt.subplots(figsize=(6, 3))
            ax.bar(techs['labels'], techs['data'], color='#f59e0b')
            ax.set_title('Technician Performance')
            plt.tight_layout()
            img = io.BytesIO()
            plt.savefig(img, format='png')
            plt.close()
            img.seek(0)
            img_path = f"/tmp/analytics_techs.png"
            with open(img_path, 'wb') as f:
                f.write(img.getvalue())
            pdf.add_page()
            pdf.cell(200, 10, "Technician Performance", ln=True, align='C')
            pdf.image(img_path, x=10, y=pdf.get_y()+2, w=190)
        
        return pdf.output(dest='S').encode('latin1')

    # ========== NEW: Change Password ==========

    def change_password(self, username: str, old_password: str, new_password: str) -> Dict:
        """Change user password"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        old_hash = hashlib.sha256(old_password.encode()).hexdigest()
        cursor.execute("SELECT * FROM users WHERE username = ? AND password_hash = ?", (username, old_hash))
        user = cursor.fetchone()
        
        if not user:
            conn.close()
            return {"status": "error", "message": "Current password is incorrect"}
        
        new_hash = hashlib.sha256(new_password.encode()).hexdigest()
        cursor.execute("UPDATE users SET password_hash = ? WHERE username = ?", (new_hash, username))
        conn.commit()
        conn.close()
        
        self._add_log('PASSWORD_CHANGE', f"User {username} changed password")
        return {"status": "success", "message": "Password updated successfully"}

    # ========== NEW: User Management ==========

    def list_users(self) -> List[Dict]:
        """Get list of all users"""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, role, full_name, last_login, created_at FROM users ORDER BY username")
        users = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return users

    def create_user(self, data: Dict) -> Dict:
        """Create a new user"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM users WHERE username = ?", (data['username'],))
        if cursor.fetchone():
            conn.close()
            return {"status": "error", "message": "Username already exists"}
        
        password_hash = hashlib.sha256(data['password'].encode()).hexdigest() if data.get('password') else hashlib.sha256("changeme".encode()).hexdigest()
        cursor.execute("""
            INSERT INTO users (username, password_hash, role, full_name, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (data['username'], password_hash, data.get('role', 'viewer'), data.get('full_name', ''), datetime.now().isoformat()))
        conn.commit()
        conn.close()
        
        self._add_log('USER_CREATE', f"Created user {data['username']}")
        return {"status": "success", "message": "User created"}

    def update_user(self, data: Dict) -> Dict:
        """Update an existing user"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        if data.get('password'):
            password_hash = hashlib.sha256(data['password'].encode()).hexdigest()
            cursor.execute("""
                UPDATE users SET password_hash = ?, role = ?, full_name = ?
                WHERE id = ? AND username != 'admin'
            """, (password_hash, data.get('role', 'viewer'), data.get('full_name', ''), data['id']))
        else:
            cursor.execute("""
                UPDATE users SET role = ?, full_name = ?
                WHERE id = ? AND username != 'admin'
            """, (data.get('role', 'viewer'), data.get('full_name', ''), data['id']))
        
        conn.commit()
        conn.close()
        
        self._add_log('USER_UPDATE', f"Updated user {data['username']}")
        return {"status": "success", "message": "User updated"}

    def delete_user(self, user_id: int) -> Dict:
        """Delete a user"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT username FROM users WHERE id = ?", (user_id,))
        user = cursor.fetchone()
        if not user:
            conn.close()
            return {"status": "error", "message": "User not found"}
        
        if user['username'] == 'admin':
            conn.close()
            return {"status": "error", "message": "Cannot delete admin user"}
        
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        conn.close()
        
        self._add_log('USER_DELETE', f"Deleted user {user['username']}")
        return {"status": "success", "message": "User deleted"}

    # ---------- LOGGING ----------
    def authenticate(self, username, password):
        conn = self._get_connection()
        cursor = conn.cursor()
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        cursor.execute("SELECT * FROM users WHERE username=? AND password_hash=?", (username, pwd_hash))
        user = cursor.fetchone()
        if user:
            cursor.execute("UPDATE users SET last_login=? WHERE username=?", (datetime.now().isoformat(), username))
            conn.commit()
            user_dict = dict(user)
            user_dict.pop('password_hash', None)
            conn.close()
            return user_dict
        conn.close()
        return None

    def get_recent_logs(self, limit=100):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM system_logs ORDER BY timestamp DESC LIMIT ?", (limit,))
        logs = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return logs

    def _add_log(self, typ, msg, user="system"):
        conn = self._get_connection()
        try:
            conn.execute("INSERT INTO system_logs (timestamp, type, message, user, created_at) VALUES (?,?,?,?,?)",
                         (datetime.now().isoformat(), typ, msg, user, datetime.now().isoformat()))
            conn.commit()
        except:
            pass
        finally:
            conn.close()

    def add_log(self, data):
        conn = self._get_connection()
        try:
            conn.execute("INSERT INTO system_logs (timestamp, type, message, user, created_at) VALUES (?,?,?,?,?)",
                         (data.get('timestamp', datetime.now().isoformat()), data['type'], data['message'], data.get('user','system'), datetime.now().isoformat()))
            conn.commit()
            return True
        except:
            return False
        finally:
            conn.close()


class SuperHandler(BaseHTTPRequestHandler):
    db = DatabaseManager(DB_NAME)

    def _send_json(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data, default=str).encode())

    def _send_file(self, content, filename, content_type):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(content)

    def _send_html(self):
        if os.path.exists(HTML_FILE):
            with open(HTML_FILE, 'r', encoding='utf-8') as f:
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.end_headers()
                self.wfile.write(f.read().encode())
        else:
            self._send_json({"error": "HTML missing"}, 404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        
        if path == "/" or path == "/index.html":
            self._send_html()
        elif path == "/api/stats":
            self._send_json(self.db.get_dashboard_stats())
        elif path == "/api/customers/all":
            self._send_json(self.db.get_all_customers())
        elif path.startswith("/api/customers/") and path.endswith("/detail"):
            name = urllib.parse.unquote(path.split('/')[3])
            self._send_json(self.db.get_customer_detail(name))
        elif path == "/api/instruments":
            filters = {
                'search': query.get('search', [''])[0],
                'division': query.get('division', [''])[0],
                'customer': query.get('customer', [''])[0]
            }
            self._send_json(self.db.search_instruments(filters))
        elif path == "/api/tech/performance":
            self._send_json(self.db.get_technician_performance())
        elif path == "/api/tech/recommend":
            serial = query.get('serial', [''])[0]
            self._send_json(self.db.recommend_technicians_advanced(serial))
        elif path == "/api/tech/list":
            self._send_json(self.db.get_all_technicians())
        elif path == "/api/logs/recent":
            self._send_json(self.db.get_recent_logs())
        elif path == "/api/transfer/log":
            self._send_json(self.db.get_transfer_log())
        elif path == "/api/charts/service-trend":
            self._send_json(self.db.get_service_trend())
        elif path == "/api/charts/top-customers":
            self._send_json(self.db.get_top_customers())
        elif path == "/api/charts/model-stats":
            customer = query.get('customer', [''])[0]
            self._send_json(self.db.get_model_statistics(customer))
        elif path == "/api/alerts/due-services":
            self._send_json(self.db.get_due_services())
        elif path == "/api/alerts/expiring-agreements":
            self._send_json(self.db.get_expiring_agreements())
        elif path == "/api/reminders/due":
            days = int(query.get('days', ['30'])[0])
            self._send_json(self.db.get_due_services(days))
        elif path == "/api/reminder-history":
            limit = int(query.get('limit', ['50'])[0])
            self._send_json(self.db.get_reminder_history(limit))
        elif path.startswith("/api/customer-360/"):
            customer = urllib.parse.unquote(path.split('/')[3])
            self._send_json(self.db.get_customer_360(customer))
        # ========== ANALYTICS ENDPOINTS ==========
        elif path == "/api/analytics/yearly-trend":
            year = query.get('year', [''])[0]
            technician = query.get('technician', [''])[0]
            self._send_json(self.db.get_yearly_trend(year, technician))
        elif path == "/api/analytics/model-distribution":
            year = query.get('year', [''])[0]
            self._send_json(self.db.get_model_distribution_by_year(year))
        elif path == "/api/analytics/tech-performance":
            year = query.get('year', [''])[0]
            self._send_json(self.db.get_tech_performance_by_year(year))
        # ========== USER MANAGEMENT ==========
        elif path == "/api/users/list":
            if self.headers.get('X-User-Role') == 'admin':
                self._send_json(self.db.list_users())
            else:
                self._send_json({"error": "Unauthorized"}, 403)
        # ========== EXPORT ANALYTICS ==========
        elif path == "/api/export/analytics-excel":
            year = query.get('year', [''])[0]
            technician = query.get('technician', [''])[0]
            excel_data = self.db.export_analytics_excel(year, technician)
            self._send_file(excel_data, "analytics_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        elif path == "/api/export/analytics-pdf":
            year = query.get('year', [''])[0]
            technician = query.get('technician', [''])[0]
            pdf_data = self.db.export_analytics_pdf(year, technician)
            self._send_file(pdf_data, "analytics_report.pdf", "application/pdf")
        # ========== TECHNICIAN REPORTS ==========
        elif path.startswith("/api/export/technician-pdf/"):
            technician = urllib.parse.unquote(path.split('/')[4])
            try:
                pdf_data = self.db.generate_technician_pdf(technician)
                self._send_file(pdf_data, f"{technician}_report.pdf", "application/pdf")
            except Exception as e:
                self._send_json({"error": str(e)}, 500)
        elif path.startswith("/api/export/technician-excel/"):
            technician = urllib.parse.unquote(path.split('/')[4])
            excel_data = self.db.generate_technician_excel(technician)
            self._send_file(excel_data, f"{technician}_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # ========== BACKUP ==========
        elif path == "/api/admin/backup":
            zip_data = self.db.export_full_backup_zip()
            self._send_file(zip_data, f"hayleys_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", "application/zip")
        # ========== COMPANY CSV EXPORT ==========
        elif path == "/api/export/company-csv":
            customer = query.get('customer', [''])[0]
            if not customer:
                self._send_json({"error": "Missing customer"}, 400)
                return
            csv_data = self.db.export_company_csv(customer)
            self._send_file(csv_data.encode('utf-8-sig'), f"{customer}_template.csv", "text/csv")
        # ========== CUSTOMER EXCEL EXPORT ==========
        elif path.startswith("/api/export/customer/"):
            name = urllib.parse.unquote(path.split('/')[4])
            excel_data = self.db.export_customer_excel(name)
            self._send_file(excel_data, f"{name}_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # ========== MACHINE EXCEL EXPORT ==========
        elif path.startswith("/api/export/machine/"):
            serial = urllib.parse.unquote(path.split('/')[4])
            excel_data = self.db.export_machine_report(serial)
            self._send_file(excel_data, f"{serial}_service_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # ========== LABS SUMMARY ==========
        elif path == "/api/labs/summary":
            customer = query.get('customer', [''])[0]
            division = query.get('division', [''])[0]
            self._send_json(self.db.get_labs_summary(customer, division))
        # ========== CUSTOMERS BY REGION ==========
        elif path == "/api/customers/by-region":
            region = query.get('region', [''])[0]
            self._send_json(self.db.get_customers_by_region(region))
        # ========== MACHINES BY TYPE ==========
        elif path == "/api/machines/by-type":
            machine_type = query.get('machine_type', [''])[0]
            customer = query.get('customer', [''])[0]
            self._send_json(self.db.get_machines_by_type(machine_type, customer))
        else:
            self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        data = json.loads(self.rfile.read(length)) if length > 0 else {}
        path = self.path.split('?')[0]

        if path == "/api/login":
            user = self.db.authenticate(data.get('username',''), data.get('password',''))
            if user:
                self._send_json({"status":"success","token":secrets.token_hex(32),"user":user})
            else:
                self._send_json({"status":"error","message":"Invalid credentials"},401)
        # ========== CHANGE PASSWORD ==========
        elif path == "/api/change-password":
            result = self.db.change_password(data.get('username'), data.get('old_password'), data.get('new_password'))
            self._send_json(result)
        # ========== USER MANAGEMENT ==========
        elif path == "/api/users/create":
            if data.get('requester_role') != 'admin':
                self._send_json({"status": "error", "message": "Unauthorized"}, 403)
            else:
                self._send_json(self.db.create_user(data))
        elif path == "/api/users/update":
            if data.get('requester_role') != 'admin':
                self._send_json({"status": "error", "message": "Unauthorized"}, 403)
            else:
                self._send_json(self.db.update_user(data))
        elif path == "/api/users/delete":
            if data.get('requester_role') != 'admin':
                self._send_json({"status": "error", "message": "Unauthorized"}, 403)
            else:
                self._send_json(self.db.delete_user(data['id']))
        elif path == "/api/admin/reset-database":
            if data.get('username') == 'admin':
                self._send_json(self.db.reset_database())
            else:
                self._send_json({"status":"error","message":"Unauthorized"},403)
        elif path == "/api/customers/create":
            self._send_json({"status":"success" if self.db.create_customer(data) else "error"})
        elif path == "/api/customers/update":
            self._send_json({"status":"success" if self.db.update_customer(data) else "error"})
        elif path == "/api/customers/delete":
            self._send_json({"status":"success" if self.db.delete_customer(data) else "error"})
        elif path == "/api/agreements/add":
            self._send_json({"status":"success" if self.db.add_service_agreement(data) else "error"})
        elif path == "/api/instruments/add":
            self._send_json({"status":"success" if self.db.add_primary_instrument(data) else "error"})
        elif path == "/api/instruments/add-secondary":
            self._send_json(self.db.add_secondary_instrument(data))
        elif path == "/api/instruments/update":
            self._send_json({"status":"success" if self.db.update_instrument(data) else "error"})
        elif path == "/api/instruments/delete":
            self._send_json({"status":"success" if self.db.delete_instrument(data) else "error"})
        elif path == "/api/instruments/convert-to-secondary":
            self._send_json(self.db.convert_primary_to_secondary(data.get('serial_number'), data.get('target_parent_serial')))
        elif path == "/api/instruments/promote-to-primary":
            self._send_json(self.db.promote_secondary_to_primary(data.get('serial_number')))
        elif path == "/api/instruments/transfer":
            self._send_json(self.db.transfer_machine(data))
        elif path == "/api/service/add":
            self._send_json({"status":"success" if self.db.add_service_ticket(data) else "error"})
        elif path == "/api/service/update":
            self._send_json({"status":"success" if self.db.update_service_ticket(data) else "error"})
        elif path == "/api/service/delete":
            self._send_json({"status":"success" if self.db.delete_service_ticket(data) else "error"})
        elif path == "/api/logs/add":
            self._send_json({"status":"success" if self.db.add_log(data) else "error"})
        elif path == "/api/upload-csv":
            self._send_json(self.db.upload_csv_data(data.get('csv_data', '')))
        elif path == "/api/labs/create":
            self._send_json(self.db.create_lab(data))
        elif path == "/api/labs/update":
            self._send_json(self.db.update_lab(data))
        elif path == "/api/labs/delete":
            self._send_json(self.db.delete_lab(data['lab_id'], data['customer_name']))
        elif path == "/api/machine-agreement/add":
            self._send_json(self.db.add_machine_agreement(data))
        elif path == "/api/upload-company-csv":
            result = self.db.upload_company_csv(data.get('customer_name'), data.get('csv_data'), data.get('username', 'system'))
            self._send_json(result)
        elif path == "/api/send-customer-email":
            result = self.db.send_customer_service_summary(data.get('customer_name'), data.get('message', ''))
            self._send_json(result)
        else:
            self._send_json({"error":"Unknown"},404)


def start_server():
    server = HTTPServer(("", PORT), SuperHandler)
    print("\n" + "="*70)
    print(" HAYLEYS LIFE SCIENCES - ENTERPRISE PORTAL v12.0")
    print(" Powered by Red Axis AI | Full Integrated System")
    print("="*70)
    print(f" Server: http://localhost:{PORT}")
    print(f" Login:  admin / admin123")
    print("         service / service@1")
    print("="*70)
    print(" ALL FEATURES ACTIVE + Enhanced Intelligence")
    print("="*70 + "\n")
    server.serve_forever()


if __name__ == "__main__":
    threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    try:
        start_server()
    except KeyboardInterrupt:
        print("\n Server stopped.")