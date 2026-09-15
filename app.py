from flask import Flask, request, redirect, session, send_file, render_template_string
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
import os
from io import BytesIO
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# =========================================================
# APP SETTINGS
# =========================================================

app = Flask(__name__)

app.secret_key = "village-system-secret-key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(BASE_DIR, "village_system.db")

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
FAMILY_UPLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, "family")
IMPORTANT_FOLDER = os.path.join(UPLOAD_FOLDER, "important")
PROFILE_FOLDER = os.path.join(UPLOAD_FOLDER, "profile")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(FAMILY_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMPORTANT_FOLDER, exist_ok=True)
os.makedirs(PROFILE_FOLDER, exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


# =========================================================
# ALLOWED FILES
# =========================================================

IMAGE_EXTENSIONS = {
    "jpg",
    "jpeg",
    "png",
    "webp"
}

DOCUMENT_EXTENSIONS = {
    "pdf",
    "xlsx",
    "xls",
    "csv",
    "docx",
    "doc"
}

ALL_EXTENSIONS = IMAGE_EXTENSIONS | DOCUMENT_EXTENSIONS


def allowed_file(filename):
    if not filename:
        return False

    if "." not in filename:
        return False

    extension = filename.rsplit(".", 1)[1].lower()

    return extension in ALL_EXTENSIONS


def file_extension(filename):
    if "." not in filename:
        return ""

    return filename.rsplit(".", 1)[1].lower()


# =========================================================
# DATABASE
# =========================================================

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def create_database():

    conn = get_db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS villages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS village_info (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_id INTEGER UNIQUE,
            family_count INTEGER DEFAULT 0,
            male_count INTEGER DEFAULT 0,
            female_count INTEGER DEFAULT 0,
            population INTEGER DEFAULT 0,
            houses INTEGER DEFAULT 0,
            revenue REAL DEFAULT 0,
            mobile TEXT DEFAULT '',
            other_info TEXT DEFAULT '',
            FOREIGN KEY(village_id)
            REFERENCES villages(id)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS family_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_id INTEGER NOT NULL,
            family_id TEXT NOT NULL,
            member_name TEXT NOT NULL,
            age INTEGER DEFAULT 0,
            gender TEXT DEFAULT '',
            marital_status TEXT DEFAULT '',
            education TEXT DEFAULT '',
            caste TEXT DEFAULT '',
            voter_id_no TEXT DEFAULT '',
            ration_card_no TEXT DEFAULT '',
            aadhaar_no TEXT DEFAULT '',
            assessment_number TEXT DEFAULT '',
            land_details TEXT DEFAULT '',
            govt_schemes TEXT DEFAULT '',
            mobile TEXT DEFAULT '',
            pension_id TEXT DEFAULT '',
            address TEXT DEFAULT '',
            remarks TEXT DEFAULT '',
            photo TEXT DEFAULT '',
            FOREIGN KEY(village_id)
            REFERENCES villages(id)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS family_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL,
            document_type TEXT NOT NULL,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            FOREIGN KEY(member_id)
            REFERENCES family_members(id)
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS important_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            detail TEXT DEFAULT '',
            detail_date TEXT DEFAULT '',
            remarks TEXT DEFAULT '',
            file_name TEXT DEFAULT '',
            stored_name TEXT DEFAULT '',
            FOREIGN KEY(village_id)
            REFERENCES villages(id)
        )
    """)

    conn.commit()

        # =====================================================
    # ADD NEW FAMILY MEMBER DETAILS
    # =====================================================

    new_columns = [
        ("father_husband_name", "TEXT DEFAULT ''"),
        ("mother_wife_name", "TEXT DEFAULT ''"),
        ("date_of_birth", "TEXT DEFAULT ''"),
        ("relationship", "TEXT DEFAULT ''"),
        ("occupation", "TEXT DEFAULT ''"),
        ("disability", "TEXT DEFAULT ''"),
        ("blood_group", "TEXT DEFAULT ''"),
        ("bank_account_no", "TEXT DEFAULT ''"),
        ("ifsc_code", "TEXT DEFAULT ''"),
        ("gas_connection", "TEXT DEFAULT ''"),
        ("electricity_connection", "TEXT DEFAULT ''"),
        ("electricity_service_no", "TEXT DEFAULT ''"),
        ("bike_details", "TEXT DEFAULT ''"),
        ("car_details", "TEXT DEFAULT ''"),
        ("other_vehicle_details", "TEXT DEFAULT ''")
    ]

    for column_name, column_type in new_columns:
        try:
            conn.execute(
                f"ALTER TABLE family_members ADD COLUMN {column_name} {column_type}"
            )
        except sqlite3.OperationalError:
            pass

    conn.commit()

    # Add photo column to old database
    try:
        conn.execute("""
            ALTER TABLE family_members
            ADD COLUMN photo TEXT DEFAULT ''
        """)
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # Create Gurrampalem automatically
    village = conn.execute("""
        SELECT id
        FROM villages
        WHERE name = ?
    """, ("Gurrampalem",)).fetchone()


    conn.commit()

    # -----------------------------------------------------
    # Create Gurrampalem automatically
    # -----------------------------------------------------

    village = conn.execute("""
        SELECT id
        FROM villages
        WHERE name = ?
    """, ("Gurrampalem",)).fetchone()

    if village is None:

        cursor = conn.execute("""
            INSERT INTO villages(name)
            VALUES(?)
        """, ("Gurrampalem",))

        village_id = cursor.lastrowid

        conn.execute("""
            INSERT INTO village_info(village_id)
            VALUES(?)
        """, (village_id,))

        conn.commit()

    conn.close()


create_database()

def create_admin_user():
    conn = get_db()

    admin_user = os.environ.get("ADMIN_USER")
    admin_password = os.environ.get("ADMIN_PASSWORD")

    if admin_user and admin_password:
        existing_admin = conn.execute(
            "SELECT id FROM users WHERE user_id = ?",
            (admin_user,)
        ).fetchone()

        if existing_admin is None:
            conn.execute(
                "INSERT INTO users (user_id, password_hash, role) VALUES (?, ?, ?)",
                (
                    admin_user,
                    generate_password_hash(admin_password),
                    "admin"
                )
            )
            conn.commit()

    conn.close()


create_admin_user()


# =========================================================
# LOGIN HELPER
# =========================================================

def logged_in():
    return session.get("logged_in") is True


# =========================================================
# HTML HEADER
# =========================================================

STYLE = """
<style>

/* ===== FAMILY CARDS DESIGN ===== */

.family-cards-card {
    padding: 28px;
}

.family-cards-card h2 {
    color: #1e3a8a;
    font-size: 30px;
    margin-bottom: 25px;
}

.family-cards-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 22px;
}

.family-card {
    background: linear-gradient(145deg, #ffffff, #f1f7ff);
    border: 1px solid #dbeafe;
    border-radius: 20px;
    padding: 22px;
    box-shadow: 0 8px 25px rgba(30, 64, 175, 0.12);
    transition: 0.25s;
}

.family-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 14px 30px rgba(30, 64, 175, 0.20);
}

.family-card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding-bottom: 15px;
    border-bottom: 1px solid #dbeafe;
}

.family-label {
    display: block;
    color: #64748b;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 1px;
}

.family-card-header h3 {
    margin: 5px 0 0;
    color: #123a7a;
    font-size: 24px;
}

.member-count {
    background: #dbeafe;
    color: #1d4ed8;
    padding: 9px 14px;
    border-radius: 20px;
    font-weight: 700;
}

.family-card-members {
    margin-top: 18px;
}

.family-member-row {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 12px;
    margin-bottom: 10px;
    background: white;
    border-radius: 14px;
    border: 1px solid #e5e7eb;
}

.family-member-row img {
    flex-shrink: 0;
    border-radius: 50%;
    object-fit: cover;
    border: 3px solid #bfdbfe;
    box-shadow: 0 3px 10px rgba(0,0,0,0.12);
}

.family-no-photo {
    width: 55px;
    height: 55px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #e2e8f0;
    border-radius: 50%;
    font-size: 28px;
}

.family-member-info {
    display: flex;
    flex-direction: column;
    gap: 5px;
}

.family-member-info a {
    color: #2563eb;
    font-size: 17px;
    font-weight: 700;
    text-decoration: none;
}

.family-member-info span {
    color: #64748b;
    font-size: 14px;
}

.view-family-btn {
    display: block;
    margin-top: 18px;
    padding: 12px;
    text-align: center;
    border-radius: 12px;
    background: linear-gradient(135deg, #2563eb, #1d4ed8);
    color: white;
    text-decoration: none;
    font-weight: 700;
    box-shadow: 0 5px 15px rgba(37, 99, 235, 0.25);
    transition: 0.25s;
}

.view-family-btn:hover {
    transform: translateY(-2px);
}

.no-family-card {
    padding: 30px;
    text-align: center;
    color: #64748b;
}

@media (max-width: 800px) {
    .family-cards-grid {
        grid-template-columns: 1fr;
    }
}

/* ===== FAMILY TABLE DESIGN ===== */

.family-table-card {
    overflow: hidden;
}

.family-table-card h2 {
    color: #123a7a;
    font-size: 28px;
    margin-bottom: 20px;
}

.family-table-card .table-wrap {
    overflow-x: auto;
    border-radius: 16px;
}

.family-table-card table {
    width: 100%;
    min-width: 1800px;
    border-collapse: separate;
    border-spacing: 0;
    background: white;
}

.family-table-card th {
    padding: 14px 12px;
    background: #123a7a;
    color: white;
    font-weight: 700;
    white-space: nowrap;
}

.family-table-card td {
    padding: 12px;
    border-bottom: 1px solid #e2e8f0;
    text-align: center;
    white-space: nowrap;
}

.family-table-card tr:hover td {
    background: #f0f9ff;
}

.family-table-card td img {
    border: 3px solid #dbeafe;
    box-shadow: 0 3px 10px rgba(0,0,0,0.15);
}

.stats-card {
    display: flex;
    align-items: center;
    gap: 18px;

    background: rgba(255, 255, 255, 0.95);

    padding: 22px 28px;
    margin: 25px auto;

    border-radius: 18px;

    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.20);

    max-width: 350px;
}

.stats-icon {
    font-size: 42px;
}

.stats-title {
    color: #475569;
    font-size: 15px;
    font-weight: bold;
}

.stats-number {
    color: #123a7a;
    font-size: 30px;
    font-weight: 800;
}

body {
    margin: 0;
    min-height: 100vh;

    background-image:
        linear-gradient(
            rgba(0, 0, 0, 0.20),
            rgba(0, 0, 0, 0.20)
        ),
        url("/static/village.jpg");

    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    background-repeat: no-repeat;

    font-family: Arial, sans-serif;
}

.container {
    width: 92%;
    max-width: 1200px;

    margin: 30px auto;

    padding: 10px;
}

.card {
    background: rgba(255, 255, 255, 0.94);

    border-radius: 24px;

    padding: 28px;

    margin-bottom: 25px;

    box-shadow:
        0 10px 30px rgba(0, 0, 0, 0.15);

    border: 1px solid rgba(255, 255, 255, 0.8);

    backdrop-filter: blur(8px);
}



h1 {
    color: #1e3a8a;
}

h2 {
    color: #334155;
}

input,
select,
textarea {
    width: 100%;
    padding: 11px;
    margin-top: 6px;
    margin-bottom: 12px;
    box-sizing: border-box;
    border: 1px solid #aaa;
    border-radius: 6px;
    font-size: 15px;
}

textarea {
    min-height: 100px;
}

button,
.btn {
    display: inline-block;
    padding: 11px 17px;
    border: none;
    border-radius: 6px;
    text-decoration: none;
    cursor: pointer;
    margin: 4px;
    font-size: 15px;
}

.blue {
    background: #2563eb;
    color: white;
}

.green {
    background: #16a34a;
    color: white;
}

.red {
    background: #dc2626;
    color: white;
}

.orange {
    background: #f59e0b;
    color: black;
}

.gray {
    background: #64748b;
    color: white;
}

.purple {
    background: #7c3aed;
    color: white;
}

.section {
    background: #eff6ff;
    border-left: 5px solid #2563eb;
    padding: 18px;
    margin-top: 20px;
    border-radius: 8px;
}

.grid {
    display: grid;
    grid-template-columns: repeat(4, 180px);
    gap: 10px;
    justify-content: center;
    width: 100%;
    margin: 0 auto 20px auto;
}

.stat {
    background: white;
    padding: 20px;
    border-radius: 12px;
    box-shadow: 0 2px 8px #bbb;
    text-align: center;
}

/* ===== MODERN VILLAGE DESIGN ===== */

.village-header {
    background: rgba(255,255,255,0.90);
    border-radius: 24px;
    padding: 25px 30px;
    margin-bottom: 25px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.18);
    backdrop-filter: blur(8px);
}

.village-title {
    font-size: 42px;
    font-weight: 800;
    color: #123a7a;
    margin: 0;
}

.village-subtitle {
    color: #64748b;
    font-size: 18px;
    margin-top: 8px;
}

.modern-stat {
    width: 180px;
    height: 120px;
    min-width: 180px;
    min-height: 120px;
    padding: 15px;
    box-sizing: border-box;

    border-radius: 18px;
    text-align: center;
    color: white;

    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
}

.modern-stat:hover {
    transform: translateY(-6px);
    box-shadow: 0 15px 30px rgba(0,0,0,0.25);
}

.modern-stat-icon {
    font-size: 38px;
    margin-bottom: 8px;
}

.modern-stat-number {
    font-size: 42px;
    font-weight: 800;
    color: white;
    line-height: 1.1;
}

.modern-stat-title {
    font-size: 18px;
    font-weight: 700;
    color: white;
    margin-top: 8px;
}


/* ===== 4 COLOR STAT CARDS ===== */

.grid .modern-stat:nth-child(1) {
    background: #ef4444;
}

.grid .modern-stat:nth-child(2) {
    background: #22c55e;
}

.grid .modern-stat:nth-child(3) {
    background: #f59e0b;
}

.grid .modern-stat:nth-child(4) {
    background: #3b82f6;
}

.modern-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    box-sizing: border-box;

    width: 185px;
    height: 60px;
    min-width: 185px;
    min-height: 60px;
    padding: 8px 10px;
    margin: 4px;

    border-radius: 13px;
    color: white;
    text-decoration: none;

    font-size: 16px;
    font-weight: 700;
    line-height: 1.2;
    text-align: center;

    transition: 0.25s;
}


.options-card {
    display: grid;
    grid-template-columns: repeat(6, 185px);
    justify-content: center;
    gap: 10px;
}

.options-card h2,
.options-card p {
    grid-column: 1 / -1;
    width: 100%;
    text-align: center;
}

.options-card .modern-btn {
    width: 185px !important;
    height: 60px !important;
    min-width: 185px !important;
    min-height: 60px !important;

    margin: 0 !important;
    padding: 8px 10px !important;

    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    box-sizing: border-box !important;
}

    padding: 8px 12px;
    margin: 6px;

    border-radius: 13px;
    color: white;
    text-decoration: none;
}

.modern-btn:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 18px rgba(0,0,0,0.20);
}

.btn-family {
    background: linear-gradient(135deg,#16a34a,#22c55e);
}

.btn-info {
    background: linear-gradient(135deg,#2563eb,#3b82f6);
}

.btn-important {
    background: linear-gradient(135deg,#7c3aed,#9333ea);
}

.btn-excel {
    background: linear-gradient(135deg,#f59e0b,#f97316);
}

.btn-word {
    background: linear-gradient(135deg,#ef4444,#dc2626);
}

.btn-back {
    background: linear-gradient(135deg,#64748b,#475569);
}

.stat-number {
    font-size: 30px;
    font-weight: bold;
    color: #1d4ed8;
}

.table-wrap {
    overflow-x: auto;
}

table {
    border-collapse: collapse;
    width: 100%;
    min-width: 1100px;
    background: white;
}

th,
td {
    border: 1px solid #bbb;
    padding: 8px;
    text-align: center;
}

th {
    background: #dbeafe;
}

.file-box {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    padding: 12px;
    border-radius: 8px;
    margin-bottom: 10px;
}

.small {
    color: #64748b;
    font-size: 13px;
}

/* ===== MODERN VILLAGE BACKGROUND ===== */

body {
    margin: 0;
    padding: 0;
    font-family: Arial, sans-serif;
    min-height: 100vh;
    background:
    linear-gradient(
        135deg,
        #e0f2fe 0%,
        #f8fafc 50%,
        #dbeafe 100%
    );

    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    background-repeat: no-repeat;

    color: #1e293b;
}

.container {
    width: 92%;
    max-width: 1200px;
    margin: 30px auto;
    padding: 10px;
}

.card {
    background: rgba(255, 255, 255, 0.94);
    border-radius: 24px;
    padding: 28px;
    margin-bottom: 25px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
    border: 1px solid rgba(255, 255, 255, 0.8);
    backdrop-filter: blur(8px);
}

@media (max-width: 700px) {
    .container {
        width: 94%;
        margin: 15px auto;
    }

    .village-title {
        font-size: 30px;
    }

    .modern-stat {
        padding: 18px;
    }

    .modern-stat-number {
        font-size: 28px;
    }

    .modern-btn {
        display: block;
        text-align: center;
        margin: 10px 0;
    }
}

/* ===== MODERN DASHBOARD ===== */

.dashboard-title {
    text-align: center;
    font-size: 38px;
    font-weight: 800;
    color: #123a7a;
    margin-bottom: 8px;
}

.dashboard-subtitle {
    text-align: center;
    color: #64748b;
    font-size: 17px;
    margin-bottom: 30px;
}

.dashboard-stat {
    background: rgba(255,255,255,0.95);
    border-radius: 24px;
    padding: 25px;
    margin-bottom: 25px;
    text-align: center;
    box-shadow: 0 10px 30px rgba(0,0,0,0.15);
}

.dashboard-stat-icon {
    font-size: 45px;
}

.dashboard-stat-title {
    font-size: 17px;
    font-weight: 600;
    color: #475569;
    margin-top: 8px;
}

.dashboard-stat-number {
    font-size: 40px;
    font-weight: 800;
    color: #2563eb;
    margin-top: 5px;
}

.village-dashboard-card {
    background: rgba(255,255,255,0.95);
    border-radius: 24px;
    padding: 25px;
    margin-bottom: 20px;
    box-shadow: 0 10px 28px rgba(0,0,0,0.15);
    transition: 0.25s;
}

.village-dashboard-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 16px 35px rgba(0,0,0,0.20);
}

.village-dashboard-name {
    font-size: 25px;
    font-weight: 800;
    color: #123a7a;
    margin-bottom: 20px;
}

.dashboard-btn {
    display: inline-block;
    padding: 12px 18px;
    margin: 5px;
    border-radius: 12px;
    color: white;
    text-decoration: none;
    font-weight: 700;
    transition: 0.2s;
}

.dashboard-btn:hover {
    transform: translateY(-2px);
    box-shadow: 0 7px 16px rgba(0,0,0,0.18);
}

.dashboard-open {
    background: linear-gradient(135deg,#2563eb,#3b82f6);
}

.dashboard-important {
    background: linear-gradient(135deg,#7c3aed,#9333ea);
}

.dashboard-edit {
    background: linear-gradient(135deg,#f59e0b,#f97316);
}

.dashboard-delete {
    background: linear-gradient(135deg,#ef4444,#dc2626);
}

.dashboard-add-btn {
    display: block;
    width: fit-content;
    margin: 0 auto 30px auto;
    padding: 14px 28px;
    border-radius: 14px;
    background: linear-gradient(135deg,#16a34a,#22c55e);
    color: white;
    text-decoration: none;
    font-size: 17px;
    font-weight: 800;
    box-shadow: 0 8px 20px rgba(0,0,0,0.15);
    transition: 0.25s;
}

.dashboard-add-btn:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 25px rgba(0,0,0,0.20);
}

/* ===== BEAUTIFUL ADD VILLAGE PAGE ===== */

.add-village-card {
    max-width: 600px;
    margin: 40px auto;
    padding: 40px;
    background: rgba(255,255,255,0.97);
    border-radius: 30px;
    box-shadow: 0 20px 50px rgba(0,0,0,0.18);
    text-align: center;
}

.add-village-icon {
    font-size: 65px;
    margin-bottom: 10px;
}

.add-village-title {
    margin: 0;
    font-size: 36px;
    font-weight: 800;
    color: #123a7a;
}

.add-village-subtitle {
    color: #64748b;
    font-size: 16px;
    margin: 10px 0 30px;
}

.add-village-label {
    display: block;
    text-align: left;
    font-size: 17px;
    font-weight: 700;
    color: #334155;
    margin-bottom: 8px;
}

.add-village-input {
    width: 100%;
    box-sizing: border-box;
    padding: 16px;
    border: 2px solid #dbeafe;
    border-radius: 14px;
    font-size: 17px;
    outline: none;
    margin-bottom: 20px;
}

.add-village-input:focus {
    border-color: #2563eb;
    box-shadow: 0 0 0 4px rgba(37,99,235,0.12);
}

.add-village-save {
    width: 100%;
    padding: 16px;
    border: none;
    border-radius: 14px;
    background: linear-gradient(135deg,#16a34a,#22c55e);
    color: white;
    font-size: 18px;
    font-weight: 800;
    cursor: pointer;
    transition: 0.25s;
}

.add-village-save:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 25px rgba(22,163,74,0.30);
}

.add-village-back {
    display: block;
    margin-top: 16px;
    padding: 14px;
    border-radius: 14px;
    background: linear-gradient(135deg,#64748b,#475569);
    color: white;
    text-decoration: none;
    font-weight: 700;
}

/* ===== VILLAGE INFORMATION FORM ===== */

.village-edit-card form {
    margin-top: 25px;
}

.village-edit-card form label {
    display: block;
    margin-top: 18px;
    margin-bottom: 7px;
    font-size: 16px;
    font-weight: 700;
    color: #334155;
}

.village-edit-card form input,
.village-edit-card form textarea {
    width: 100%;
    box-sizing: border-box;
    padding: 14px 16px;
    border: 2px solid #dbeafe;
    border-radius: 13px;
    background: #f8fafc;
    font-size: 16px;
    outline: none;
    transition: 0.25s;
}

.village-edit-card form input:focus,
.village-edit-card form textarea:focus {
    border-color: #2563eb;
    background: white;
    box-shadow: 0 0 0 4px rgba(37,99,235,0.12);
}

.village-edit-card form textarea {
    min-height: 130px;
    resize: vertical;
}

.village-edit-card form button {
    margin-top: 25px;
    padding: 15px 28px;
    border: none;
    border-radius: 14px;
    background: linear-gradient(135deg,#16a34a,#22c55e);
    color: white;
    font-size: 17px;
    font-weight: 800;
    cursor: pointer;
    box-shadow: 0 8px 20px rgba(22,163,74,0.25);
    transition: 0.25s;

}

.back-village-btn {
    display: inline-block;
    margin-top: 15px;
    padding: 14px 22px;
    border-radius: 14px;
    background: linear-gradient(135deg, #64748b, #475569);
    color: white;
    text-decoration: none;
    font-weight: 700;
    font-size: 16px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.15);
    transition: 0.25s;
}

.back-village-btn:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 25px rgba(0,0,0,0.22);
}

.village-edit-card form button:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 25px rgba(22,163,74,0.30);
}

.info-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0 25px;
}

.info-grid label {
    margin-top: 18px;
}

.info-grid input {
    width: 100%;
}

@media (max-width: 700px) {
    .info-grid {
        grid-template-columns: 1fr;
    }
}

</style>
"""


# =========================================================
# LOGIN
# =========================================================

# =========================================================
# LOGIN
# =========================================================

@app.route("/", methods=["GET", "POST"])
def login():
    error = False

    if request.method == "POST":

        user_id = request.form.get("user_id", "").strip()
        password = request.form.get("password", "")

        conn = get_db()

        user = conn.execute(
            "SELECT * FROM users WHERE user_id = ? AND active = 1",
            (user_id,)
        ).fetchone()

        conn.close()

        if user and check_password_hash(user["password_hash"], password):
            session["logged_in"] = True
            session["user_id"] = user["user_id"]
            session["role"] = user["role"]
            return redirect("/dashboard")

        error = True
    return render_template_string(
        STYLE + """
        <div class="container">
            <div class="card"
                 style="max-width:430px;margin:80px auto;text-align:center;">

                <h1>🏠 Village Information</h1>

                <h2>Secure Login</h2>

                <form method="POST">

                    <input type="text"
                           name="user_id"
                           placeholder="User ID"
                           required>

                    <input type="password"
                           name="password"
                           placeholder="Password"
                           required>

                    <button class="blue" type="submit">
                        🔐 Login
                    </button>

                </form>

                {% if error %}
                    <p style="color:red;">
                        Invalid login details
                    </p>
                {% endif %}

            </div>
        </div>
        """
    )

    # Wrong login
    error = True

    return render_template_string(
        STYLE + """

        <style>

        /* ===== PREMIUM SECURITY LOGIN ===== */

        .login-page {
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            overflow: hidden;

            background:
                radial-gradient(
                    circle at 15% 20%,
                    rgba(37, 99, 235, 0.35),
                    transparent 35%
                ),
                radial-gradient(
                    circle at 85% 75%,
                    rgba(14, 165, 233, 0.25),
                    transparent 35%
                ),
                linear-gradient(
                    135deg,
                    #020617,
                    #071a3d,
                    #0b2d63
                );
        }

        /* Background security circles */

        .security-circle {
            position: absolute;
            border: 1px solid rgba(96, 165, 250, 0.15);
            border-radius: 50%;
            pointer-events: none;
        }

        .circle-one {
            width: 500px;
            height: 500px;
            right: -180px;
            top: -120px;
        }

        .circle-two {
            width: 350px;
            height: 350px;
            left: -160px;
            bottom: -100px;
        }

        /* Background lock */

        .background-lock {
            position: absolute;
            right: 7%;
            top: 50%;
            transform: translateY(-50%);
            font-size: 230px;
            opacity: 0.055;
            filter: blur(1px);
            pointer-events: none;
        }

        .background-shield {
            position: absolute;
            left: 5%;
            top: 18%;
            font-size: 180px;
            opacity: 0.045;
            pointer-events: none;
        }

        /* Login card */

        .premium-login-card {
            width: 430px;
            max-width: calc(100% - 30px);
            padding: 42px 38px 30px;

            background: rgba(255, 255, 255, 0.97);

            border-radius: 28px;

            box-shadow:
                0 30px 80px rgba(0, 0, 0, 0.45),
                0 0 40px rgba(37, 99, 235, 0.18);

            position: relative;
            z-index: 5;

            text-align: center;

            border: 1px solid rgba(255, 255, 255, 0.7);
        }

        /* Lock logo */

        .premium-lock {
            width: 92px;
            height: 92px;

            margin: 0 auto 18px;

            display: flex;
            align-items: center;
            justify-content: center;

            border-radius: 50%;

            font-size: 46px;

            background:
                linear-gradient(
                    145deg,
                    #2563eb,
                    #0b3b9e
                );

            box-shadow:
                0 12px 30px rgba(37, 99, 235, 0.4),
                inset 0 1px 1px rgba(255,255,255,0.4);

            border: 4px solid rgba(147, 197, 253, 0.7);
        }

        .premium-login-card h1 {
            margin: 8px 0 5px;
            color: #123a7a;
            font-size: 32px;
            font-weight: 800;
        }

        .secure-title {
            color: #64748b;
            font-size: 17px;
            margin-bottom: 28px;
        }

        /* Security line */

        .security-line {
            display: flex;
            align-items: center;
            gap: 10px;
            justify-content: center;
            margin: 10px 0 25px;
            color: #64748b;
            font-size: 14px;
        }

        .security-line::before,
        .security-line::after {
            content: "";
            height: 1px;
            width: 55px;
            background: #cbd5e1;
        }

        /* Labels */

        .login-label {
            display: block;
            text-align: left;
            margin-bottom: 8px;
            color: #173b69;
            font-weight: 700;
            font-size: 16px;
        }

        /* Inputs */

        .login-input {
            width: 100%;
            box-sizing: border-box;

            padding: 16px 18px;

            margin-bottom: 20px;

            border: 1px solid #cbd5e1;
            border-radius: 13px;

            font-size: 16px;

            outline: none;

            background: #ffffff;

            transition: 0.25s;
        }

        .login-input:focus {
            border-color: #2563eb;

            box-shadow:
                0 0 0 4px rgba(37, 99, 235, 0.12);
        }

        /* Login button */

        .premium-login-btn {
            width: 100%;

            padding: 16px;

            margin-top: 5px;

            border: none;
            border-radius: 14px;

            background:
                linear-gradient(
                    135deg,
                    #2563eb,
                    #0b3b9e
                );

            color: white;

            font-size: 18px;
            font-weight: 800;

            cursor: pointer;

            box-shadow:
                0 12px 25px rgba(37, 99, 235, 0.35);

            transition: 0.25s;
        }

        .premium-login-btn:hover {
            transform: translateY(-2px);

            box-shadow:
                0 16px 32px rgba(37, 99, 235, 0.45);
        }

        .premium-login-btn:active {
            transform: translateY(0);
        }

        /* Error */

        .login-error {
            background: #fef2f2;
            color: #b91c1c;

            border: 1px solid #fecaca;

            border-radius: 10px;

            padding: 10px;

            margin-bottom: 18px;

            font-size: 14px;
            font-weight: 600;
        }

        /* Authorized users */

        .authorized {
            margin-top: 22px;

            color: #64748b;

            font-size: 14px;
        }

        /* Footer */

        .login-footer {
            margin-top: 24px;

            color: #64748b;

            font-size: 14px;
        }

        /* Mobile */

        @media (max-width: 600px) {

            .premium-login-card {
                padding: 32px 24px 25px;
            }

            .premium-login-card h1 {
                font-size: 27px;
            }

            .background-lock {
                font-size: 150px;
            }

            .background-shield {
                font-size: 120px;
            }

        }

        </style>


        <div class="login-page">

            <div class="security-circle circle-one"></div>
            <div class="security-circle circle-two"></div>

            <div class="background-lock">
                🔒
            </div>

            <div class="background-shield">
                🛡️
            </div>


            <div class="premium-login-card">

                <div class="premium-lock">
                    🔐
                </div>

                <h1>
                    Village Information
                </h1>

                <div class="secure-title">
                    Secure Login
                </div>

                <div class="security-line">
                    🛡️ Protected Access
                </div>


                {% if error %}

                <div class="login-error">
                    ⚠️ Invalid User ID or Password
                </div>

                {% endif %}


                <form method="POST">

                    <label class="login-label">
                        User ID
                    </label>

                    <input
                        class="login-input"
                        type="text"
                        name="user_id"
                        placeholder="Enter User ID"
                        autocomplete="username"
                        required
                    >


                    <label class="login-label">
                        Password
                    </label>

                    <input
                        class="login-input"
                        type="password"
                        name="password"
                        placeholder="Enter Password"
                        autocomplete="current-password"
                        required
                    >


                    <button
                        class="premium-login-btn"
                        type="submit"
                    >
                        🔐 Login Securely
                    </button>

                </form>


                <div class="authorized">
                    🛡️ Authorized Users Only
                </div>

                <div class="login-footer">
                    © Village Information System
                </div>

            </div>

        </div>

        """,
        error=error
    )

    # =====================================================
    # NEW LOGIN PAGE
    # =====================================================

    return render_template_string(
        STYLE + """

        <style>

        .login-page::before {
    content: "🔒";
    position: absolute;
    font-size: 280px;
    opacity: 0.06;
    color: white;
    right: 5%;
    bottom: -40px;
    transform: rotate(-10deg);
}

.login-page::after {
    content: "🛡️";
    position: absolute;
    font-size: 220px;
    opacity: 0.05;
    color: white;
    left: 4%;
    top: 8%;
}

        .login-page {
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;

    background:
        linear-gradient(rgba(10, 35, 80, 0.82), rgba(20, 65, 130, 0.88)),
        radial-gradient(circle at top, #3b82f6, #07152f 70%);

    position: relative;
    overflow: hidden;
}

        .login-card {
            width: 100%;
            max-width: 430px;
            background: rgba(255,255,255,0.97);
            padding: 40px;
            border-radius: 24px;
            box-shadow: 0 15px 45px rgba(0,0,0,0.18);
            text-align: center;
            box-sizing: border-box;
        }

        .login-icon {
            width: 85px;
            height: 85px;
            margin: 0 auto 18px;
            border-radius: 50%;
            background: #2563eb;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 42px;
            box-shadow: 0 8px 20px rgba(37,99,235,0.30);
        }

        .login-title {
            margin: 0;
            color: #1e3a8a;
            font-size: 30px;
        }

        .login-subtitle {
            margin: 8px 0 30px;
            color: #64748b;
            font-size: 16px;
        }

        .login-label {
            display: block;
            text-align: left;
            margin-bottom: 7px;
            color: #334155;
            font-weight: bold;
        }

        .login-input-box {
            position: relative;
            margin-bottom: 20px;
        }

        .login-input {
            width: 100%;
            height: 52px;
            padding: 0 16px 0 48px;
            border: 1px solid #cbd5e1;
            border-radius: 12px;
            font-size: 16px;
            box-sizing: border-box;
            outline: none;
        }

        .login-input:focus {
            border-color: #2563eb;
            box-shadow: 0 0 0 3px rgba(37,99,235,0.12);
        }

        .input-icon {
            position: absolute;
            left: 16px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 20px;
        }

        .password-toggle {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            border: none;
            background: transparent;
            cursor: pointer;
            font-size: 19px;
            padding: 5px;
        }

        .login-button {
            width: 100%;
            height: 52px;
            border: none;
            border-radius: 12px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: white;
            font-size: 17px;
            font-weight: bold;
            cursor: pointer;
            margin-top: 5px;
            box-shadow: 0 8px 18px rgba(37,99,235,0.25);
        }

        .login-button:hover {
            transform: translateY(-1px);
            box-shadow: 0 10px 22px rgba(37,99,235,0.32);
        }

        .login-footer {
            margin-top: 25px;
            color: #64748b;
            font-size: 13px;
        }

        @media (max-width: 500px) {

            .login-card {
                padding: 30px 22px;
            }

            .login-title {
                font-size: 25px;
            }

        }

        </style>


        <div class="login-page">

            <div class="login-card">

                <div class="login-icon">
                    🏠
                </div>

                <h1 class="login-title">
                    Village Information
                </h1>

                <p class="login-subtitle">
                    Secure Login
                </p>


                <form method="POST">

                    <label class="login-label">
                        User ID
                    </label>

                    <div class="login-input-box">

                        <span class="input-icon">
                            👤
                        </span>

                        <input
                            class="login-input"
                            type="text"
                            name="user_id"
                            placeholder="Enter User ID"
                            autocomplete="username"
                            required
                        >

                    </div>


                    <label class="login-label">
                        Password
                    </label>

                    <div class="login-input-box">

                        <span class="input-icon">
                            🔐
                        </span>

                        <input
                            class="login-input"
                            id="loginPassword"
                            type="password"
                            name="password"
                            placeholder="Enter Password"
                            autocomplete="current-password"
                            required
                        >

                        <button
                            type="button"
                            class="password-toggle"
                            onclick="togglePassword()"
                            id="passwordButton"
                        >
                            👁️
                        </button>

                    </div>


                    <button
                        class="login-button"
                        type="submit"
                    >
                        🔓 Login
                    </button>

                </form>


                <div class="login-footer">
                    © Village Information System
                </div>

            </div>

        </div>


        <script>

        function togglePassword() {

            const password =
                document.getElementById("loginPassword");

            const button =
                document.getElementById("passwordButton");

            if (password.type === "password") {

                password.type = "text";
                button.innerHTML = "🙈";

            } else {

                password.type = "password";
                button.innerHTML = "👁️";

            }

        }

        </script>

    )

    .login-subtitle {
            margin: 8px 0 30px;
            color: #64748b;
            font-size: 16px;
        }

.login-label {
            display: block;
            text-align: left;
            margin-bottom: 7px;
            color: #334155;
            font-weight: bold;
        }

.login-input-box {
            position: relative;
            margin-bottom: 20px;
        }

.login-input {
            width: 100%;
            height: 52px;
            padding: 0 16px 0 48px;
            border: 1px solid #cbd5e1;
            border-radius: 12px;
            font-size: 16px;
            box-sizing: border-box;
            outline: none;
        }

.login-input:focus {
            border-color: #2563eb;
            box-shadow: 0 0 0 3px rgba(37,99,235,0.12);
        }

        .input-icon {
            position: absolute;
            left: 16px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 20px;
        }

password-toggle {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            border: none;
            background: transparent;
            cursor: pointer;
            font-size: 19px;
            padding: 5px;
        }

.login-button {
            width: 100%;
            height: 52px;
            border: none;
            border-radius: 12px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: white;
            font-size: 17px;
            font-weight: bold;
            cursor: pointer;
            margin-top: 5px;
            box-shadow: 0 8px 18px rgba(37,99,235,0.25);
        }

        .login-button:hover {
            transform: translateY(-1px);
            box-shadow: 0 10px 22px rgba(37,99,235,0.32);
        }

        .login-footer {
            margin-top: 25px;
            color: #64748b;
            font-size: 13px;
        }

        @media (max-width: 500px) {

            .login-card {
                padding: 30px 22px;
            }

            .login-title {
                font-size: 25px;
            }

        }

        </style>

        <div class="login-page">

            <div class="login-card">

                <div class="login-icon">
                    🏠
                </div>

                <h1 class="login-title">
                    Village Information
                </h1>

                <p class="login-subtitle">
                    Secure Login
                </p>

                <form method="POST">

                    <label class="login-label">
                        User ID
                    </label>

                    <div class="login-input-box">

                        <span class="input-icon">
                            👤
                        </span>

                        <input
                            class="login-input"
                            type="text"
                            name="user_id"
                            placeholder="Enter User ID"
                            autocomplete="username"
                            required
                        >

                    </div>


                    <label class="login-label">
                        Password
                    </label>

                    <div class="login-input-box">

                        <span class="input-icon">
                            🔐
                        </span>

                        <input
                            class="login-input"
                            id="loginPassword"
                            type="password"
                            name="password"
                            placeholder="Enter Password"
                            autocomplete="current-password"
                            required
                        >

                        <button
                            type="button"
                            class="password-toggle"
                            onclick="togglePassword()"
                            id="passwordButton"
                        >
                            👁️
                        </button>

                    </div>


                    <button
                        class="login-button"
                        type="submit"
                    >
                        🔓 Login
                    </button>

                </form>

                <div class="login-footer">
                    © Village Information System
                </div>

            </div>

        </div>


        <script>

        function togglePassword() {

            const password =
                document.getElementById("loginPassword");

            const button =
                document.getElementById("passwordButton");

            if (password.type === "password") {

                password.type = "text";
                button.innerHTML = "🙈";

            } else {

                password.type = "password";
                button.innerHTML = "👁️";

            }

        }

        </script>
        """
    )
# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
def dashboard():

    if not logged_in():
        return redirect("/")

    conn = get_db()

    villages = conn.execute("""
        SELECT *
        FROM villages
        `ORDER BY name`
    """).fetchall()

    # Total statistics
    population = conn.execute("""
        SELECT COALESCE(SUM(population), 0)
        FROM village_info
    """).fetchone()[0]

    houses = conn.execute("""
        SELECT COALESCE(SUM(houses), 0)
        FROM village_info
    """).fetchone()[0]

    family_members = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
    """).fetchone()[0]

    important_details = conn.execute("""
        SELECT COUNT(*)
        FROM important_details
    """).fetchone()[0]

    village_count = len(villages)

    conn.close()

    return render_template_string(
        STYLE + """

<style>

body {
    margin: 0;
    font-family: Arial, sans-serif;
    background: #eef5fb;
}

.dashboard-wrapper {
    display: flex;
    min-height: 100vh;
}

.sidebar {
    width: 260px;
    background: linear-gradient(180deg, #092a55, #12284a);
    color: white;
    padding-top: 20px;
    position: fixed;
    left: 0;
    top: 0;
    bottom: 0;
}

.sidebar-title {
    text-align: center;
    font-size: 28px;
    font-weight: bold;
    padding: 15px;
    margin-bottom: 20px;
}

.sidebar-menu {
    padding: 0;
    margin: 0;
}

.sidebar-menu a {
    display: block;
    color: white;
    text-decoration: none;
    padding: 18px 25px;
    font-size: 18px;
    font-weight: bold;
}

.sidebar-menu a:hover {
    background: #1477df;
}

.main-content {
    margin-left: 260px;
    width: calc(100% - 260px);
}

.top-header {
    background: linear-gradient(90deg, #063b78, #0877dc);
    color: white;
    padding: 22px 35px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.village-name {
    font-size: 34px;
    font-weight: bold;
}

.village-subtitle {
    font-size: 18px;
    margin-top: 5px;
}

.admin-box {
    background: rgba(0,0,0,0.18);
    padding: 15px 25px;
    border-radius: 15px;
    font-size: 18px;
    font-weight: bold;
}

.content {
    padding: 25px;
}

.stats {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 22px;
    margin-bottom: 25px;
}

.stat-card {
    background: white;
    border-radius: 15px;
    padding: 25px;
    text-align: center;
    box-shadow: 0 3px 12px rgba(0,0,0,0.12);
}

.stat-icon {
    font-size: 45px;
}

.stat-number {
    font-size: 38px;
    font-weight: bold;
    margin: 8px 0;
}

.stat-title {
    font-size: 19px;
    font-weight: bold;
}

.options-card {
    background: white;
    border-radius: 18px;
    padding: 25px;
    width: 100%;
    max-width: none;
    box-sizing: border-box;
    box-shadow: 0 3px 15px rgba(0,0,0,0.12);
}

.options-title {
    color: #17467d;
    font-size: 30px;
    border-bottom: 2px solid #ddd;
    padding-bottom: 15px;
}

.options-grid {
    display: grid !important;
    grid-template-columns: repeat(4, minmax(220px, 1fr)) !important;
    gap: 16px;
    width: 100% !important;
    max-width: none !important;
    margin: 0 auto;
    box-sizing: border-box;
}

.options-card {
    width: 100% !important;
    max-width: none !important;
    box-sizing: border-box;
}

.option-btn {
    color: white;
    text-decoration: none;
    padding: 18px 12px;
    border-radius: 12px;
    text-align: center;
    font-size: 17px;
    font-weight: bold;
    display: block;
}

.option-btn:hover {
    opacity: 0.88;
    transform: translateY(-2px);
}

.green { background: #16ad57; }
.blue { background: #0879e6; }
.purple { background: #7737e8; }
.orange { background: #f58a0b; }
.red { background: #ed3030; }
.teal { background: #079eaa; }
.gold { background: #d39b00; }
.gray { background: #50627a; }

.footer {
    margin-top: 25px;
    background: #dff7e9;
    padding: 22px;
    border-radius: 15px;
    text-align: center;
    color: #13733b;
    font-size: 18px;
    font-weight: bold;
}

@media (max-width: 1000px) {

    .stats {
        grid-template-columns: repeat(2, 1fr);
    }

    .options-grid {
        grid-template-columns: repeat(2, 1fr);
    }
}

@media (max-width: 650px) {

    .sidebar {
        width: 75px;
    }

    .sidebar-title {
        font-size: 0;
    }

    .sidebar-menu a {
        font-size: 0;
        text-align: center;
        padding: 18px 5px;
    }

    .main-content {
        margin-left: 75px;
        width: calc(100% - 75px);
    }

    .stats {
        grid-template-columns: 1fr;
    }

    .options-grid {
        grid-template-columns: 1fr;
    }

    .top-header {
        padding: 15px;
    }

    .village-name {
        font-size: 24px;
    }
}

</style>


<div class="dashboard-wrapper">

    <!-- SIDEBAR -->

    <div class="sidebar">

        <div class="sidebar-title">
            🏠 Village
        </div>

        <div class="sidebar-menu">

            <a href="/dashboard">🏠 Dashboard</a>

            {% if villages %}
            <a href="/village/{{ villages[0]['id'] }}">
                🏢 Village Information
            </a>

            <a href="/families/{{ villages[0]['id'] }}">
                👥 Family Members
            </a>

            <a href="/important/{{ villages[0]['id'] }}">
                📋 Important Details
            </a>

            <a href="/members-reports/{{ villages[0]['id'] }}">
                📊 Reports
            </a>

            <a href="/export-excel/{{ villages[0]['id'] }}">
                📥 Export / Download
            </a>
            {% endif %}

            <a href="/logout">
                🚪 Logout
            </a>

        </div>

    </div>


    <!-- MAIN CONTENT -->

    <div class="main-content">

        <!-- HEADER -->

        <div class="top-header">

            <div>
                <div class="village-name">
                    🏠
                    {% if villages %}
                        {{ villages[0]["name"] }}
                    {% else %}
                        Village Information
                    {% endif %}
                </div>

                <div class="village-subtitle">
                    Our Village • Our People • Our Future
                </div>
            </div>

            <div class="admin-box">
                👤 Admin
            </div>

        </div>


        <div class="content">


            <!-- STATISTICS -->

            <div class="stats">

                <div class="stat-card">
                    <div class="stat-icon">👥</div>
                    <div class="stat-number">
                        {{ population }}
                    </div>
                    <div class="stat-title">
                        Population
                    </div>
                </div>


                <div class="stat-card">
                    <div class="stat-icon">🏠</div>
                    <div class="stat-number">
                        {{ houses }}
                    </div>
                    <div class="stat-title">
                        Houses
                    </div>
                </div>


                <div class="stat-card">
                    <div class="stat-icon">👨‍👩‍👧</div>
                    <div class="stat-number">
                        {{ family_members }}
                    </div>
                    <div class="stat-title">
                        Family Members
                    </div>
                </div>


                <div class="stat-card">
                    <div class="stat-icon">📋</div>
                    <div class="stat-number">
                        {{ important_details }}
                    </div>
                    <div class="stat-title">
                        Important Details
                    </div>
                </div>

            </div>


            <!-- VILLAGE OPTIONS -->

            <div class="options-card">

                <h2 class="options-title">
                    ▦ Village Options
                </h2>

                <div class="options-grid">

                    {% if villages %}

                    <a class="option-btn green"
                       href="/families/{{ villages[0]['id'] }}">
                        👨‍👩‍👧 Family Information
                    </a>

                    <a class="option-btn blue"
                       href="/family-members-report/{{ villages[0]['id'] }}">
                        👤 Members
                    </a>

                    <a class="option-btn purple"
                       href="/members-reports/{{ villages[0]['id'] }}">
                        📋 Members Reports
                    </a>

                    <a class="option-btn purple"
                       href="/father-husband-report/{{ villages[0]['id'] }}">
                        👨 Father / Husband Names
                    </a>

                    <a class="option-btn orange"
                       href="/mother-wife-report/{{ villages[0]['id'] }}">
                        👩 Mother / Wife Names
                    </a>

                    <a class="option-btn teal"
                       href="/date-of-birth-report/{{ villages[0]['id'] }}">
                        📅 Date of Birth
                    </a>

                    <a class="option-btn purple"
                       href="/relationship-report/{{ villages[0]['id'] }}">
                        🔗 Relationship
                    </a>

                    <a class="option-btn orange"
                       href="/occupation-report/{{ villages[0]['id'] }}">
                        💼 Occupation
                    </a>

                    <a class="option-btn red"
                       href="/disability-report/{{ villages[0]['id'] }}">
                        ♿ Disability
                    </a>

                    <a class="option-btn teal"
                       href="/blood-group-report/{{ villages[0]['id'] }}">
                        💧 Blood Group
                    </a>

                    <a class="option-btn purple"
                       href="/bank-account-report/{{ villages[0]['id'] }}">
                        🏦 Bank Accounts
                    </a>

                    <a class="option-btn blue"
                       href="/ifsc-report/{{ villages[0]['id'] }}">
                        🏦 IFSC Codes
                    </a>

                    <a class="option-btn green"
                       href="/voter-id-report/{{ villages[0]['id'] }}">
                        🪪 Voter IDs
                    </a>

                    <a class="option-btn purple"
                       href="/ration-card-report/{{ villages[0]['id'] }}">
                        🪪 Ration Cards
                    </a>

                    <a class="option-btn gold"
                       href="/family-reports/{{ villages[0]['id'] }}">
                        📊 Family Reports
                    </a>

                    <a class="option-btn gray"
                       href="/village/{{ villages[0]['id'] }}">
                        ℹ Village Information
                    </a>

                    <a class="option-btn red"
                       href="/important/{{ villages[0]['id'] }}">
                        📁 Important Details
                    </a>

                    <a class="option-btn blue"
                       href="/export-excel/{{ villages[0]['id'] }}">
                        📗 Village Excel
                    </a>

                    <a class="option-btn purple"
                       href="/export-word/{{ villages[0]['id'] }}">
                        📘 Village Word
                    </a>

                    {% endif %}

                </div>

            </div>


            <!-- FOOTER -->

            <div class="footer">
                🌱 “A Strong Village Builds a Stronger Nation”
                <br>
                Gurrampalem Village Information System
            </div>

        </div>

    </div>

</div>

""",
        villages=villages,
        population=population,
        houses=houses,
        family_members=family_members,
        important_details=important_details
    )


# =========================================================
# ADD VILLAGE
# =========================================================

@app.route("/add-village", methods=["GET", "POST"])
def add_village():

    if not logged_in():
        return redirect("/")

    if request.method == "POST":

        name = request.form.get(
            "village_name",
            ""
        ).strip()

        if not name:
            return "Village name required."

        conn = get_db()

        try:

            cursor = conn.execute("""
                INSERT INTO villages(name)
                VALUES(?)
            """, (name,))

            village_id = cursor.lastrowid

            conn.execute("""
                INSERT INTO village_info(village_id)
                VALUES(?)
            """, (village_id,))

            conn.commit()

        except sqlite3.IntegrityError:

            conn.close()

            return "Village already exists."

        conn.close()

        return redirect("/dashboard")

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card">

<div class="add-village-card">

    <div class="add-village-icon">
        🏡
    </div>

    <h1 class="add-village-title">
        Add New Village
    </h1>

    <p class="add-village-subtitle">
        Add a new village to your Village Information System
    </p>

    <form method="POST">

        <label class="add-village-label">
            🏘️ Village Name
        </label>

        <input
            class="add-village-input"
            type="text"
            name="village_name"
            placeholder="Enter village name"
            required
        >

        <button
            class="add-village-save"
            type="submit"
        >
            ➕ Save Village
        </button>

    </form>

    <a
        class="add-village-back"
        href="/dashboard"
    >
        ← Back to Dashboard
    </a>

</div>
        """
    )




# =========================================================
# VILLAGE PAGE
# =========================================================

@app.route("/village/<int:village_id>")
def village_page(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    info = conn.execute("""
        SELECT *
        FROM village_info
        WHERE village_id = ?
    """, (village_id,)).fetchone()

    member_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
    """, (village_id,)).fetchone()[0]

    important_count = conn.execute("""
        SELECT COUNT(*)
        FROM important_details
        WHERE village_id = ?
    """, (village_id,)).fetchone()[0]

    family_option_count = conn.execute("""
        SELECT COUNT(DISTINCT family_id)
        FROM family_members
        WHERE village_id = ?
    """, (village_id,)).fetchone()[0]

    father_husband_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND father_husband_name IS NOT NULL
        AND TRIM(father_husband_name) != ''
    """, (village_id,)).fetchone()[0]

    mother_wife_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND mother_wife_name IS NOT NULL
        AND TRIM(mother_wife_name) != ''
    """, (village_id,)).fetchone()[0]

    dob_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND date_of_birth IS NOT NULL
        AND TRIM(date_of_birth) != ''
    """, (village_id,)).fetchone()[0]

    relationship_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND relationship IS NOT NULL
        AND TRIM(relationship) != ''
    """, (village_id,)).fetchone()[0]

    occupation_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND occupation IS NOT NULL
        AND TRIM(occupation) != ''
    """, (village_id,)).fetchone()[0]

    disability_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND disability IS NOT NULL
        AND TRIM(disability) != ''
    """, (village_id,)).fetchone()[0]

    blood_group_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND blood_group IS NOT NULL
        AND TRIM(blood_group) != ''
    """, (village_id,)).fetchone()[0]

    bank_account_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND bank_account_no IS NOT NULL
        AND TRIM(bank_account_no) != ''
    """, (village_id,)).fetchone()[0]

    ifsc_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND ifsc_code IS NOT NULL
        AND TRIM(ifsc_code) != ''
    """, (village_id,)).fetchone()[0]

    voter_id_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND voter_id_no IS NOT NULL
        AND TRIM(voter_id_no) != ''
    """, (village_id,)).fetchone()[0]

    ration_card_count = conn.execute("""
        SELECT COUNT(*)
        FROM family_members
        WHERE village_id = ?
        AND ration_card_no IS NOT NULL
        AND TRIM(ration_card_no) != ''
    """, (village_id,)).fetchone()[0]

    if village is None:
        return "Village not found."

    return render_template_string(
        STYLE + """
        <div class="container">

        <!-- VILLAGE HEADER -->
        <div class="village-header">

            <h1 class="village-title">
                🏡 {{ village["name"] }}
            </h1>

            <div class="village-subtitle">
                Our Village • Our People • Our Future
            </div>

        </div>


<!-- STATISTICS -->
<div class="grid">

    <!-- POPULATION -->
    <div class="modern-stat">
        <div class="modern-stat-icon">👥</div>

        <div class="modern-stat-number">
            {{ info["population"] if info else 0 }}
        </div>

        <div class="modern-stat-title">
            Population
        </div>
    </div>


    <!-- HOUSES -->
    <div class="modern-stat">
        <div class="modern-stat-icon">🏠</div>

        <div class="modern-stat-number">
            {{ info["houses"] if info else 0 }}
        </div>

        <div class="modern-stat-title">
            Houses
        </div>
    </div>


    <!-- FAMILY MEMBERS -->
    <div class="modern-stat"
         onclick="window.location.href='/members-reports/{{ village['id'] }}'"
         style="cursor:pointer;">

        <div class="modern-stat-icon">👤</div>

        <div class="modern-stat-number">
            {{ member_count }}
        </div>

        <div class="modern-stat-title">
            Family Members
        </div>
    </div>


    <!-- IMPORTANT DETAILS -->
    <div class="modern-stat">
        <div class="modern-stat-icon">📋</div>

        <div class="modern-stat-number">
            {{ important_count }}
        </div>

        <div class="modern-stat-title">
            Important Details
        </div>
    </div>

</div>


        <!-- VILLAGE OPTIONS -->

        <div class="options-card">

            <h2 class="options-title">
                📌 Village Options
            </h2>

            <p class="village-subtitle">
                Manage and view village information easily
            </p>


            <a
                class="modern-btn btn-family"
                href="/families/{{ village['id'] }}"
                >
                👨‍👩‍👧 Family Information
                <span class="option-count">{{ info["family_count"] if info else 0
                }}</span>
            </a>

            <a
               class="modern-btn btn-family"
               href="/family-members-report/{{ village['id'] }}"
           >
              👤 Members
            </a>

            <a
                class="modern-btn btn-important"
                href="/members-reports/{{ village['id'] }}"
            >
                📋 Members Reports
                <span class="option-count">{{ member_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/father-husband-report/{{ village['id'] }}"
            >
               👨 Father / Husband Names
               <span class="option-count">{{ father_husband_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/mother-wife-report/{{ village['id'] }}"
            >
                👩 Mother / Wife Names
                <span class="option-count">{{ mother_wife_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/date-of-birth-report/{{ village['id'] }}"
            >
                🎂 Date of Birth
                <span class="option-count">{{ dob_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/relationship-report/{{ village['id'] }}"
            >
                🔗 Relationship
                <span class="option-count">{{ relationship_count }}</span>
            </a> 

            <a
                class="modern-btn btn-important"
                href="/occupation-report/{{ village['id'] }}"
            >
                💼 Occupation
                <span class="option-count">{{ occupation_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/disability-report/{{ village['id'] }}"
            >
                ♿ Disability
                <span class="option-count">{{ disability_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/blood-group-report/{{ village['id'] }}"
            >
                🩸 Blood Group
                <span class="option-count">{{ blood_group_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/bank-account-report/{{ village['id'] }}"
            >
                🏦 Bank Accounts
                <span class="option-count">{{ bank_account_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/ifsc-code-report/{{ village['id'] }}"
            >
                🏛️ IFSC Codes
                <span class="option-count">{{ ifsc_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/voter-id-report/{{ village['id'] }}"
            >
                🪪 Voter IDs
                <span class="option-count">{{ voter_id_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/ration-card-report/{{ village['id'] }}"
            >
                📇 Ration Cards
                <span class="option-count">{{ ration_card_count }}</span>
            </a>

            <a
                class="modern-btn btn-important"
                href="/family-reports/{{ village['id'] }}"
            >
                📋 Family Reports
                <span class="option-count">{{ family_option_count }}</span>
            </a>


            <a
                class="modern-btn btn-info"
                href="/edit-information/{{ village['id'] }}"
            >
                ✏️ Village Information
            </a>


            <a
                class="modern-btn btn-important"
                href="/important/{{ village['id'] }}"
            >
                📁 Important Details
            </a>


            <a
                class="modern-btn btn-excel"
                href="/export-excel/{{ village['id'] }}"
            >
                📊 Village Excel
            </a>


            <a
                class="modern-btn btn-word"
                href="/export-word/{{ village['id'] }}"
            >
                📄 Village Word
            </a>


            <a
                class="modern-btn btn-back"
                href="/dashboard"
            >
                ← Dashboard
            </a>

        </div>

    </div>

    """,
    village=village,
    info=info,
    member_count=member_count,
    important_count=important_count,
    family_option_count=family_option_count,
    father_husband_count=father_husband_count,
    mother_wife_count=mother_wife_count,
    dob_count=dob_count,
    relationship_count=relationship_count,
    occupation_count=occupation_count,
    disability_count=disability_count,
    blood_group_count=blood_group_count,
    bank_account_count=bank_account_count,
    ifsc_count=ifsc_count,
    voter_id_count=voter_id_count,
    ration_card_count=ration_card_count
)

# =========================================================
# FAMILY REPORTS
# =========================================================

@app.route("/family-reports/<int:village_id>")
def family_reports(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if village is None:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        ORDER BY family_id, member_name
    """, (village_id,)).fetchall()

    conn.close()

    # GROUP MEMBERS BY FAMILY
    family_groups = {}

    for member in members:

        family_id = member["family_id"]

        if family_id not in family_groups:
            family_groups[family_id] = []

        family_groups[family_id].append(member)

    return render_template_string(
        STYLE + """

        <div class="container">

            <!-- REPORT HEADER -->

            <div class="card">

                <h1>
                    📋 Family Detailed Report
                </h1>

                <h2>
                    🏡 {{ village["name"] }}
                </h2>

                <div style="
                    display:flex;
                    gap:15px;
                    flex-wrap:wrap;
                    margin-top:20px;
                ">

                    <div class="card" style="flex:1;min-width:180px;">
                        <h3>👨‍👩‍👧 Families</h3>
                        <h1>{{ family_groups|length }}</h1>
                    </div>

                    <div class="card" style="flex:1;min-width:180px;">
                        <h3>👤 Members</h3>
                        <h1>{{ members|length }}</h1>
                    </div>

                </div>

                <br>

                <button
                    class="btn blue"
                    onclick="window.print()"
                >
                    🖨️ Print Report
                </button>

                <a
                    class="btn gray"
                    href="/village/{{ village['id'] }}"
                >
                    ← Back to Village
                </a>

            </div>


            <!-- FAMILY REPORTS -->

            {% for family_id, family_members in family_groups.items() %}

            <div class="card">

                <h2>
                    🏠 Family ID:
                    {{ family_id }}
                </h2>

                <p>
                    👨‍👩‍👧 Total Members:
                    <b>{{ family_members|length }}</b>
                </p>


                {% for member in family_members %}

                <div class="card"
                     style="
                        margin-top:20px;
                        border-left:6px solid #2563eb;
                     ">

                    <!-- MEMBER HEADER -->

                    <h2>
                        👤 {{ member["member_name"] }}
                    </h2>

                    {% if member["photo"] %}

                    <img
                        src="/profile-photo/{{ member['id'] }}"
                        style="
                            width:100px;
                            height:100px;
                            object-fit:cover;
                            border-radius:50%;
                            border:4px solid #2563eb;
                            margin-bottom:15px;
                        "
                    >

                    {% endif %}


                    <!-- PERSONAL DETAILS -->

                    <h3>
                        👤 Personal Details
                    </h3>

                    <table>

                        <tr>
                            <th>Member Name</th>
                            <td>{{ member["member_name"] }}</td>
                        </tr>

                        <tr>
                            <th>Father / Husband Name</th>
                            <td>{{ member["father_husband_name"] }}</td>
                        </tr>

                        <tr>
                            <th>Mother / Wife Name</th>
                            <td>{{ member["mother_wife_name"] }}</td>
                        </tr>

                        <tr>
                            <th>Date of Birth</th>
                            <td>{{ member["date_of_birth"] }}</td>
                        </tr>

                        <tr>
                            <th>Age</th>
                            <td>{{ member["age"] }}</td>
                        </tr>

                        <tr>
                            <th>Relationship</th>
                            <td>{{ member["relationship"] }}</td>
                        </tr>

                        <tr>
                            <th>Gender</th>
                            <td>{{ member["gender"] }}</td>
                        </tr>

                        <tr>
                            <th>Marital Status</th>
                            <td>{{ member["marital_status"] }}</td>
                        </tr>

                        <tr>
                            <th>Education</th>
                            <td>{{ member["education"] }}</td>
                        </tr>

                        <tr>
                            <th>Occupation</th>
                            <td>{{ member["occupation"] }}</td>
                        </tr>

                        <tr>
                            <th>Disability</th>
                            <td>{{ member["disability"] }}</td>
                        </tr>

                        <tr>
                            <th>Blood Group</th>
                            <td>{{ member["blood_group"] }}</td>
                        </tr>

                    </table>


                    <!-- IDENTITY DETAILS -->

                    <h3>
                        🪪 Identity Details
                    </h3>

                    <table>

                        <tr>
                            <th>Voter ID</th>
                            <td>{{ member["voter_id_no"] }}</td>
                        </tr>

                        <tr>
                            <th>Ration Card</th>
                            <td>{{ member["ration_card_no"] }}</td>
                        </tr>

                        <tr>
                            <th>Aadhaar</th>
                            <td>{{ member["aadhaar_no"] }}</td>
                        </tr>

                        <tr>
                            <th>Assessment Number</th>
                            <td>{{ member["assessment_number"] }}</td>
                        </tr>

                        <tr>
                            <th>Pension ID</th>
                            <td>{{ member["pension_id"] }}</td>
                        </tr>

                    </table>


                    <!-- BANK DETAILS -->

                    <h3>
                        🏦 Bank Details
                    </h3>

                    <table>

                        <tr>
                            <th>Bank Account No</th>
                            <td>{{ member["bank_account_no"] }}</td>
                        </tr>

                        <tr>
                            <th>IFSC Code</th>
                            <td>{{ member["ifsc_code"] }}</td>
                        </tr>

                    </table>


                    <!-- PROPERTY DETAILS -->

                    <h3>
                        🏡 Property Details
                    </h3>

                    <table>

                        <tr>
                            <th>Land Details</th>
                            <td>{{ member["land_details"] }}</td>
                        </tr>

                        <tr>
                            <th>Bike Details</th>
                            <td>{{ member["bike_details"] }}</td>
                        </tr>

                        <tr>
                            <th>Car Details</th>
                            <td>{{ member["car_details"] }}</td>
                        </tr>

                        <tr>
                            <th>Other Vehicle</th>
                            <td>{{ member["other_vehicle_details"] }}</td>
                        </tr>

                    </table>


                    <!-- UTILITY DETAILS -->

                    <h3>
                        ⚡ Utility Details
                    </h3>

                    <table>

                        <tr>
                            <th>Gas Connection</th>
                            <td>{{ member["gas_connection"] }}</td>
                        </tr>

                        <tr>
                            <th>Electricity Connection</th>
                            <td>{{ member["electricity_connection"] }}</td>
                        </tr>

                        <tr>
                            <th>Electricity Service No</th>
                            <td>{{ member["electricity_service_no"] }}</td>
                        </tr>

                    </table>


                    <!-- GOVERNMENT DETAILS -->

                    <h3>
                        🏛️ Government / Scheme Details
                    </h3>

                    <table>

                        <tr>
                            <th>Government Schemes</th>
                            <td>{{ member["govt_schemes"] }}</td>
                        </tr>

                        <tr>
                            <th>Pension ID</th>
                            <td>{{ member["pension_id"] }}</td>
                        </tr>

                    </table>


                    <!-- CONTACT DETAILS -->

                    <h3>
                        📱 Contact & Address
                    </h3>

                    <table>

                        <tr>
                            <th>Mobile</th>
                            <td>{{ member["mobile"] }}</td>
                        </tr>

                        <tr>
                            <th>Address</th>
                            <td>{{ member["address"] }}</td>
                        </tr>

                        <tr>
                            <th>Remarks</th>
                            <td>{{ member["remarks"] }}</td>
                        </tr>

                    </table>

                </div>

                {% endfor %}

            </div>

            {% endfor %}


            <a
                class="btn gray"
                href="/village/{{ village['id'] }}"
            >
                ← Back to Village
            </a>

        </div>

        <style>

            table {
                width:100%;
                border-collapse:collapse;
                margin-bottom:25px;
            }

            th,
            td {
                padding:12px;
                border:1px solid #ddd;
                text-align:left;
            }

            th {
                width:30%;
                background:#f1f5f9;
            }

            h3 {
                margin-top:25px;
                color:#1e3a8a;
            }

            @media print {

                .btn {
                    display:none !important;
                }

                body {
                    background:white !important;
                }

                .card {
                    box-shadow:none !important;
                }

            }

        </style>

        """,

        village=village,
        members=members,
        family_groups=family_groups
    )

# =========================================================
# FAMILY LIST
# =========================================================

@app.route("/families/<int:village_id>")
def connist(village_id):

    if not logged_in():
        return redirect("/")

    search = request.args.get("search", "").strip()

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if village is None:
        conn.close()
        return "Village not found."

    if search:

        like = "%" + search + "%"

        members = conn.execute("""
            SELECT *
            FROM family_members
            WHERE village_id = ?
            AND (
                member_name LIKE ?
                OR family_id LIKE ?
                OR mobile LIKE ?
                OR voter_id_no LIKE ?
                OR ration_card_no LIKE ?
                OR aadhaar_no LIKE ?
                OR assessment_number LIKE ?
                OR pension_id LIKE ?
                OR caste LIKE ?
            )
            ORDER BY family_id, member_name
        """, (
            village_id,
            like,
            like,
            like,
            like,
            like,
            like,
            like,
            like,
            like
        )).fetchall()

    else:

        members = conn.execute("""
            SELECT *
            FROM family_members
            WHERE village_id = ?
            ORDER BY family_id, member_name
        """, (village_id,)).fetchall()

    conn.close()

        # GROUP MEMBERS BY FAMILY ID
    family_groups = {}

    for member in members:
        family_id = member["family_id"]

        if family_id not in family_groups:
            family_groups[family_id] = []

        family_groups[family_id].append(member)

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card family-table-card">

                <h1>
                    👨‍👩‍👧 Family Information
                </h1>

                <h2>
                    🏡 {{ village["name"] }}
                </h2>

                <a
                    class="btn green"
                    href="/add-family/{{ village['id'] }}"
                >
                    ➕ Add Family Member
                </a>

                <a
                    class="btn blue"
                    href="/export-family-excel/{{ village['id'] }}"
                >
                    📊 Family Excel
                </a>

                <a
                    class="btn purple"
                    href="/export-family-word/{{ village['id'] }}"
                >
                    📄 Family Word
                </a>

                <a
                    class="btn gray"
                    href="/village/{{ village['id'] }}"
                >
                    ← Back to Village
                </a>

            </div>


            <div class="card">

                <h2>
                    🔎 Search Family Member
                </h2>

                <form method="GET">

                    <input
                        type="text"
                        name="search"
                        value="{{ search }}"
                        placeholder="Search Name / Family ID / Mobile / Voter ID / Aadhaar"
                    >

                    <button
                        class="blue"
                        type="submit"
                    >
                        🔍 Search
                    </button>

                    <a
                        class="btn gray"
                        href="/families/{{ village['id'] }}"
                    >
                        Clear
                    </a>

                </form>

            </div>


            <div class="card family-table-card">

                <div class="card family-cards-card">

    <h2>
        👨‍👩‍👧 Family Members
    </h2>

    <div class="family-cards-grid">

        {% for family_id, family_members in family_groups.items() %}

        <div class="family-card">

            <div class="family-card-header">

                <div>
                    <span class="family-label">
                        FAMILY ID
                    </span>

                    <h3>
                        🏠 {{ family_id }}
                    </h3>
                </div>

                <div class="member-count">
                    👨‍👩‍👧 {{ family_members|length }}
                </div>

            </div>

            <div class="family-card-members">

                {% for member in family_members %}

                <div class="family-member-row">

                    {% if member["photo"] %}

                    <img
                        src="/profile-photo/{{ member['id'] }}"
                        width="55"
                        height="55"
                        style="object-fit:cover;border-radius:50%;"
                    >

                    {% else %}

                    <div class="family-no-photo">
                        👤
                    </div>

                    {% endif %}

                    <div class="family-member-info">

                        <a
                            href="/family-profile/{{ member['id'] }}"
                        >
                            {{ member["member_name"] }}
                        </a>

                        <span>
                            {{ member["age"] }} Years •
                            {{ member["gender"] }}
                        </span>

                    </div>

                </div>

                {% endfor %}

            </div>

            <a
                class="view-family-btn"
                href="/family/{{ village['id'] }}/{{ family_id }}"
            >
                👁️ View Family
            </a>

        </div>

        {% else %}

        <div class="no-family-card">
            <h3>👨‍👩‍👧 No Family Members Found</h3>
            <p>
                Click "Add Family Member" to add a new member.
            </p>
        </div>

        {% endfor %}

    </div>

</div>

        """,
        village=village,
        members=members,
        family_groups=family_groups,
        search=search
    )

@app.route("/family-members-report/<int:village_id>")
def family_members_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found"

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        ORDER BY family_id, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>
                    👤 Members
                </h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - All Family Members
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Occupation</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>

                                <td>
                                    {{ member["family_id"] }}
                                </td>

                                <td>
                                    <strong>
                                        {{ member["member_name"] }}
                                    </strong>
                                </td>

                                <td>
                                    {{ member["age"] or "-" }}
                                </td>

                                <td>
                                    {{ member["gender"] or "-" }}
                                </td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>
                                    {{ member["occupation"] or "-" }}
                                </td>

                                <td>
                                    {{ member["mobile"] or "-" }}
                                </td>

                                <td>

                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>

                                </td>

                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        👤 No family members found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>

        """
        ,
        village=village,
        members=members
    )

# =========================================================
# MEMBERS REPORTS
# =========================================================

@app.route("/members-reports/<int:village_id>")
def members_reports(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        ORDER BY family_id, member_name
    """, (village_id,)).fetchall()

    conn.close()

    all_members = members

    selected_gender = request.args.get("gender", "").strip().title()

    if selected_gender not in ("Male", "Female"):
        selected_gender = ""

    total_members = len(all_members)

    total_families = len(
    set(member["family_id"] for member in all_members)
    )

    male_count = sum(
    1 for member in all_members
    if (member["gender"] or "").strip().lower() == "male"
    )

    female_count = sum(
    1 for member in all_members
    if (member["gender"] or "").strip().lower() == "female"
    )

    if selected_gender:
            members = [
            member for member in all_members
            if (member["gender"] or "").strip().lower()
            == selected_gender.lower()
        ]

    report_title = "Members Summary Report"

    if selected_gender:
        report_title = selected_gender + " Members Report"

    return render_template_string(
        STYLE + """
        

        <div class="container">

            <div class="card">

                <h1>📋 Members Reports</h1>

                <p  class="village-subtitle">
                    {{ village["name"] }} - {{ report_title }}
                </p>

                <div style="
                    display:flex;
                    gap:15px;
                    flex-wrap:wrap;
                    margin:20px 0;
                ">

                    <div
                            class="card"
                            style="flex:1; min-width:160px; cursor:pointer;"
                            onclick="window.location.href='/members-reports/{{ village['id'] }}'"
                    >
                            <h3>👤 Total Members</h3>
                            <h1>{{ total_members }}</h1>
                    </div>

                    <div
                            class="card"
                            style="flex:1; min-width:160px; cursor:pointer;"
                            onclick="window.location.href='/families/{{ village['id'] }}'"
                    >
                            <h3>👨‍👩‍👧 Total Families</h3>
                            <h1>{{ total_families }}</h1>
                    </div>

                    <div
                            class="card"
                            style="flex:1; min-width:160px; cursor:pointer;"
                            onclick="window.location.href='/members-reports/{{ village['id'] }}?gender=Male'"
                    >
                            <h3>👨 Male</h3>
                            <h1>{{ male_count }}</h1>
                    </div>

                    <div
                            class="card"
                            style="flex:1; min-width:160px; cursor:pointer;"
                            onclick="window.location.href='/members-reports/{{ village['id'] }}?gender=Female'"
                    >
                            <h3>👩 Female</h3>
                            <h1>{{ female_count }}</h1>
                     </div>

                </div>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Father / Husband</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Occupation</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                            <td>
                                <a href="/family-profile/{{ member['id'] }}"
                                style="
                                color:#2563eb;
                                font-weight:800;
                                text-decoration:none;
                                font-size:17px;
                                ">
                                {{ member["member_name"] }}
                                </a>
                             </td>

                                <td>
                                    {{ member["father_husband_name"] or "-" }}
                                </td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>
                                    {{ member["occupation"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        👤 No family members found for this village.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members,
        total_members=total_members,
        total_families=total_families,
        male_count=male_count,
        female_count=female_count,
        report_title=report_title,
        selected_gender=selected_gender
    )

# =========================================================
# FATHER / HUSBAND NAMES REPORT
# =========================================================

@app.route("/father-husband-report/<int:village_id>")
def father_husband_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND father_husband_name IS NOT NULL
        AND TRIM(father_husband_name) != ''
        ORDER BY father_husband_name, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>👨 Father / Husband Names</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Father / Husband Names List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Father / Husband Name</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>
                                    {{ member["father_husband_name"] }}
                                </td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        👨 No Father / Husband names found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# MOTHER / WIFE NAMES REPORT
# =========================================================

@app.route("/mother-wife-report/<int:village_id>")
def mother_wife_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND mother_wife_name IS NOT NULL
        AND TRIM(mother_wife_name) != ''
        ORDER BY mother_wife_name, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>👩 Mother / Wife Names</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Mother / Wife Names List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Mother / Wife Name</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>
                                    {{ member["mother_wife_name"] }}
                                </td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        👩 No Mother / Wife names found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# DATE OF BIRTH REPORT
# =========================================================

@app.route("/date-of-birth-report/<int:village_id>")
def date_of_birth_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND date_of_birth IS NOT NULL
        AND TRIM(date_of_birth) != ''
        ORDER BY date_of_birth, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🎂 Date of Birth</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Date of Birth List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Date of Birth</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>
                                    {{ member["date_of_birth"] }}
                                </td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        🎂 No Date of Birth details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# RELATIONSHIP REPORT
# =========================================================

@app.route("/relationship-report/<int:village_id>")
def relationship_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND relationship IS NOT NULL
        AND TRIM(relationship) != ''
        ORDER BY relationship, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🔗 Relationship</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Relationship List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Relationship</th>
                                <th>Father / Husband</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["relationship"] }}</td>

                                <td>
                                    {{ member["father_husband_name"] or "-" }}
                                </td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        🔗 No relationship details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# OCCUPATION REPORT
# =========================================================

@app.route("/occupation-report/<int:village_id>")
def occupation_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND occupation IS NOT NULL
        AND TRIM(occupation) != ''
        ORDER BY occupation, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>💼 Occupation</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Occupation List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Occupation</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["occupation"] }}</td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        💼 No occupation details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# DISABILITY REPORT
# =========================================================

@app.route("/disability-report/<int:village_id>")
def disability_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND disability IS NOT NULL
        AND TRIM(disability) != ''
        ORDER BY disability, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>♿ Disability</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Disability Details
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Disability</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["disability"] }}</td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        ♿ No disability details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>💼 Occupation

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# BLOOD GROUP REPORT
# =========================================================

@app.route("/blood-group-report/<int:village_id>")
def blood_group_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND blood_group IS NOT NULL
        AND TRIM(blood_group) != ''
        ORDER BY blood_group, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🩸 Blood Group</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Blood Group List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Blood Group</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Relationship</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["blood_group"] }}</td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>
                                    {{ member["relationship"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        🩸 No blood group details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# BANK ACCOUNTS REPORT
# =========================================================

@app.route("/bank-account-report/<int:village_id>")
def bank_account_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND bank_account_no IS NOT NULL
        AND TRIM(bank_account_no) != ''
        ORDER BY member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🏦 Bank Accounts</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Bank Account List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Bank Account No</th>
                                <th>IFSC Code</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["bank_account_no"] }}</td>

                                <td>
                                    {{ member["ifsc_code"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}Bank Accounts

                    <div class="empty-box">
                        🏦 No bank account details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# IFSC CODE REPORT
# =========================================================

@app.route("/ifsc-code-report/<int:village_id>")
def ifsc_code_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND ifsc_code IS NOT NULL
        AND TRIM(ifsc_code) != ''
        ORDER BY ifsc_code, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🏛️ IFSC Codes</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - IFSC Code List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>IFSC Code</th>
                                <th>Bank Account No</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["ifsc_code"] }}</td>

                                <td>
                                    {{ member["bank_account_no"] or "-" }}
                                </td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        🏛️ No IFSC code details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# VOTER ID REPORT
# =========================================================

@app.route("/voter-id-report/<int:village_id>")
def voter_id_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND voter_id_no IS NOT NULL
        AND TRIM(voter_id_no) != ''
        ORDER BY voter_id_no, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>🪪 Voter IDs</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Voter ID List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Voter ID</th>
                                <th>Age</th>
                                <th>Gender</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["voter_id_no"] }}</td>

                                <td>{{ member["age"] or "-" }}</td>

                                <td>{{ member["gender"] or "-" }}</td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        🪪 No voter ID details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# RATION CARD REPORT
# =========================================================

@app.route("/ration-card-report/<int:village_id>")
def ration_card_report(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if not village:
        conn.close()
        return "Village not found."

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        AND ration_card_no IS NOT NULL
        AND TRIM(ration_card_no) != ''
        ORDER BY ration_card_no, member_name
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <div class="card">

                <h1>📇 Ration Cards</h1>

                <p class="village-subtitle">
                    {{ village["name"] }} - Ration Card List
                </p>

                {% if members %}

                <div style="overflow-x:auto;">

                    <table class="data-table">

                        <thead>
                            <tr>
                                <th>Family ID</th>
                                <th>Member Name</th>
                                <th>Ration Card No</th>
                                <th>Mobile</th>
                                <th>Details</th>
                            </tr>
                        </thead>

                        <tbody>

                        {% for member in members %}

                            <tr>
                                <td>{{ member["family_id"] }}</td>

                                <td>
                                    <strong>{{ member["member_name"] }}</strong>
                                </td>

                                <td>{{ member["ration_card_no"] }}</td>

                                <td>{{ member["mobile"] or "-" }}</td>

                                <td>
                                    <a
                                        href="/family-profile/{{ member['id'] }}"
                                        class="modern-btn btn-info"
                                    >
                                        👁️ View
                                    </a>
                                </td>
                            </tr>

                        {% endfor %}

                        </tbody>

                    </table>

                </div>

                {% else %}

                    <div class="empty-box">
                        📇 No ration card details found.
                    </div>

                {% endif %}

                <br>

                <a
                    href="/village/{{ village['id'] }}"
                    class="modern-btn btn-back"
                >
                    ← Back to Village
                </a>

            </div>

        </div>
        """,
        village=village,
        members=members
    )

# =========================================================
# FAMILY FULL PAGE
# =========================================================

@app.route("/family/<int:village_id>/<family_id>")
def family_full_page(village_id, family_id):

        if not logged_in():
            return redirect("/")

        conn = get_db()

        village = conn.execute("""
            SELECT *
            FROM villages
            WHERE id = ?
        """, (village_id,)).fetchone()

        members = conn.execute("""
            SELECT *
            FROM family_members
            WHERE village_id = ?
            AND family_id = ?
            ORDER BY member_name
        """, (village_id, family_id)).fetchall()

        conn.close()

        return render_template_string(
            STYLE + """

            <style>

.family-full-header {
    background: white;
    border-radius: 20px;
    padding: 30px;
    margin-bottom: 25px;
    box-shadow: 0 5px 20px rgba(0,0,0,0.12);
    border-left: 6px solid #0879e6;
}

.family-full-header h1 {
    color: #17467d;
    font-size: 32px;
    margin-bottom: 10px;
}

.family-full-header h2 {
    color: #333;
    font-size: 22px;
    margin-bottom: 18px;
}

.family-id-badge,
.family-member-count {
    display: inline-block;
    background: #e8f1ff;
    color: #17467d;
    padding: 10px 18px;
    border-radius: 20px;
    margin-right: 10px;
    font-weight: bold;
}

.family-detail-card {
    background: white;
    border-radius: 20px;
    padding: 25px;
    margin-bottom: 25px;
    box-shadow: 0 5px 20px rgba(0,0,0,0.12);
}

.family-profile-header {
    display: grid;
    grid-template-columns: 220px 1fr;
    gap: 25px;
    align-items: center;
    margin-bottom: 25px;
}

.family-profile-photo {
    width: 200px;
    height: 200px;
    object-fit: cover;
    border-radius: 18px;
    border: 4px solid #e3edff;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

.family-profile-placeholder {
    width: 200px;
    height: 200px;
    border-radius: 18px;
    background: #eef4ff;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 70px;
}

.family-profile-header h2 {
    font-size: 26px;
    color: #17467d;
}

.family-detail-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
}

.family-detail-item {
    background: #f6f9fd;
    border: 1px solid #dce8f5;
    border-radius: 12px;
    padding: 15px;
    min-height: 65px;
}

.family-detail-item span {
    display: block;
    color: #667085;
    font-size: 13px;
    font-weight: bold;
    margin-bottom: 7px;
}

.family-detail-item b {
    color: #172b4d;
    font-size: 16px;
    word-break: break-word;
}

.family-detail-item.full-width {
    grid-column: 1 / -1;
}

.family-actions {
    display: flex;
    gap: 12px;
    margin-top: 22px;
}

.family-edit-btn,
.family-delete-btn {
    display: inline-block;
    padding: 12px 25px;
    border-radius: 10px;
    color: white !important;
    text-decoration: none;
    font-weight: bold;
    font-size: 16px;
}

.family-edit-btn {
    background: #0879e6;
}

.family-delete-btn {
    background: #ed3030;
}

.family-edit-btn:hover,
.family-delete-btn:hover {
    opacity: 0.85;
    transform: translateY(-2px);
}

@media (max-width: 900px) {
    .family-profile-header {
        grid-template-columns: 1fr;
        text-align: center;
    }

    .family-profile-photo,
    .family-profile-placeholder {
        margin: auto;
    }

    .family-detail-grid {
        grid-template-columns: repeat(2, 1fr);
    }
}

@media (max-width: 600px) {
    .family-detail-grid {
        grid-template-columns: 1fr;
    }

    .family-profile-photo,
    .family-profile-placeholder {
        width: 160px;
        height: 160px;
    }
}

</style>
        
<div class="container">

            <div class="card family-full-header">

                <h1>👨‍👩‍👧 Family Information</h1>

                <h2>🏡 {{ village["name"] }}</h2>

                <div class="family-id-badge">
                    🆔 Family ID: {{ family_id }}
                </div>

                <div class="family-member-count">
                    👨‍👩‍👧 Members: {{ members|length }}
                </div>

            </div>


            {% for m in members %}

            <div class="card family-detail-card">

                <div class="family-profile-header">

                    {% if m["photo"] %}

                    <img
                        src="/profile-photo/{{ m['id'] }}"
                        class="family-profile-photo"
                    >

                    {% else %}

                    <div class="family-profile-placeholder">
                        👤
                    </div>

                    {% endif %}

                    <div>

                        <h2>

                            <a

                               href="/family-profile/{{ m['id'] }}"
                               style="text-decoration:none; color:inherit;"
                            >
                            👤 {{ m["member_name"] }}
                            </a>
                        </h2>

                        <p>
                            Family ID: <b>{{ m["family_id"] }}</b>
                        </p>

                    </div>

                </div>


                <div class="family-detail-grid">

                    <div class="family-detail-item">
                        <span>Age</span>
                        <b>{{ m["age"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Gender</span>
                        <b>{{ m["gender"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Marital Status</span>
                        <b>{{ m["marital_status"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Education</span>
                        <b>{{ m["education"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Caste</span>
                        <b>{{ m["caste"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Voter ID</span>
                        <b>{{ m["voter_id_no"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Ration Card</span>
                        <b>{{ m["ration_card_no"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Aadhaar</span>
                        <b>{{ m["aadhaar_no"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Assessment Number</span>
                        <b>{{ m["assessment_number"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Mobile</span>
                        <b>{{ m["mobile"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Pension ID</span>
                        <b>{{ m["pension_id"] }}</b>
                    </div>

                    <div class="family-detail-item">
                        <span>Land Details</span>
                        <b>{{ m["land_details"] }}</b>
                    </div>

                    <div class="family-detail-item full-width">
                        <span>Government Schemes</span>
                        <b>{{ m["govt_schemes"] }}</b>
                    </div>

                    <div class="family-detail-item full-width">
                        <span>Address</span>
                        <b>{{ m["address"] }}</b>
                    </div>

                    <div class="family-detail-item full-width">
                        <span>Remarks</span>
                        <b>{{ m["remarks"] }}</b>
                    </div>

                </div>


                <div class="family-actions">

                    <a
                        href="/edit-family/{{ m['id'] }}"
                        class="family-edit-btn"
                    >
                        ✏️ Edit
                    </a>

                    <a
                        href="/delete-family/{{ m['id'] }}"
                        class="family-delete-btn"
                        onclick="return confirm('Delete this family member?')"
                    >
                        🗑️ Delete
                    </a>

                </div>

            </div>

            {% endfor %}


            <a
                class="btn gray"
                href="/families/{{ village['id'] }}"
            >
                ← Back to Family Information
            </a>

        </div>
        """,
        village=village,
        members=members,
        family_id=family_id
    )

# =========================================================
# ADD FAMILY MEMBER
# =========================================================
@app.route("/add-family/<int:village_id>", methods=["GET", "POST"])
def add_family(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if village is None:
        conn.close()
        return "Village not found."

    if request.method == "POST":

        family_id = request.form.get("family_id", "").strip()
        member_name = request.form.get("member_name", "").strip()
        age = request.form.get("age", "").strip()
        gender = request.form.get("gender", "").strip()

        marital_status = request.form.get(
            "marital_status", ""
        ).strip()

        education = request.form.get(
            "education", ""
        ).strip()

        caste = request.form.get(
            "caste", ""
        ).strip()

        voter_id_no = request.form.get(
            "voter_id_no", ""
        ).strip()

        ration_card_no = request.form.get(
            "ration_card_no", ""
        ).strip()

        aadhaar_no = request.form.get(
            "aadhaar_no", ""
        ).strip()

        assessment_number = request.form.get(
            "assessment_number", ""
        ).strip()

        land_details = request.form.get(
            "land_details", ""
        ).strip()

        govt_schemes = request.form.get(
            "govt_schemes", ""
        ).strip()

        mobile = request.form.get(
            "mobile", ""
        ).strip()

        pension_id = request.form.get(
            "pension_id", ""
        ).strip()

        address = request.form.get(
            "address", ""
        ).strip()

        remarks = request.form.get(
            "remarks", ""
        ).strip()

        father_husband_name = request.form.get(
            "father_husband_name", ""
        ).strip()

        mother_wife_name = request.form.get(
            "mother_wife_name", ""
        ).strip()

        date_of_birth = request.form.get(
            "date_of_birth", ""
        ).strip()

        relationship = request.form.get(
            "relationship", ""
        ).strip()

        occupation = request.form.get(
            "occupation", ""
        ).strip()

        disability = request.form.get(
            "disability", ""
        ).strip()

        blood_group = request.form.get(
            "blood_group", ""
        ).strip()

        bank_account_no = request.form.get(
            "bank_account_no", ""
        ).strip()

        ifsc_code = request.form.get(
            "ifsc_code", ""
        ).strip()

        gas_connection = request.form.get(
            "gas_connection", ""
        ).strip()

        electricity_connection = request.form.get(
            "electricity_connection", ""
        ).strip()

        electricity_service_no = request.form.get(
            "electricity_service_no", ""
        ).strip()

        bike_details = request.form.get(
            "bike_details", ""
        ).strip()

        car_details = request.form.get(
            "car_details", ""
        ).strip()

        other_vehicle_details = request.form.get(
            "other_vehicle_details", ""
        ).strip()

        # =================================================
        # PROFILE PHOTO
        # =================================================

        photo = request.files.get("photo")

        photo_name = ""

        if photo and photo.filename:

            original_name = secure_filename(
                photo.filename
            )

            extension = file_extension(
                original_name
            )

            allowed_photo_extensions = [
                "jpg",
                "jpeg",
                "png",
                "webp"
            ]

            if extension not in allowed_photo_extensions:
                conn.close()
                return (
                    "Only JPG, JPEG, PNG and WEBP "
                    "photos are allowed."
                )

            photo_name = (
                str(village_id)
                + "_"
                + str(abs(hash(original_name)))
                + "."
                + extension
            )

            photo_path = os.path.join(
                UPLOAD_FOLDER,
                photo_name
            )

            photo.save(photo_path)

        # =================================================
        # SAVE FAMILY MEMBER
        # =================================================

        conn.execute("""
            INSERT INTO family_members
            (
                village_id,
                family_id,
                member_name,
                age,
                gender,
                marital_status,
                education,
                caste,
                voter_id_no,
                ration_card_no,
                aadhaar_no,
                assessment_number,
                land_details,
                govt_schemes,
                mobile,
                pension_id,
                address,
                remarks,
                photo,
                father_husband_name,
                mother_wife_name,
                date_of_birth,
                relationship,
                occupation,
                disability,
                blood_group,
                bank_account_no,
                ifsc_code,
                gas_connection,
                electricity_connection,
                electricity_service_no,
                bike_details,
                car_details,
                other_vehicle_details
            )
            VALUES
           
            (
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?
            )
        """, (
            village_id,
            family_id,
            member_name,
            age,
            gender,
            marital_status,
            education,
            caste,
            voter_id_no,
            ration_card_no,
            aadhaar_no,
            assessment_number,
            land_details,
            govt_schemes,
            mobile,
            pension_id,
            address,
            remarks,
            photo_name,
            father_husband_name,
            mother_wife_name,
            date_of_birth,
            relationship,
            occupation,
            disability,
            blood_group,
            bank_account_no,
            ifsc_code,
            gas_connection,
            electricity_connection,
            electricity_service_no,
            bike_details,
            car_details,
            other_vehicle_details
        ))

        conn.commit()
        conn.close()

        return redirect(
            "/families/" + str(village_id)
        )

    # =================================================
    # ADD FAMILY FORM
    # =================================================

    return render_template_string(
        STYLE + """
        
        <div class="container">

            <div class="card">

                <h1>
                    ➕ Add Family Member
                </h1>

                <h2>
                    🏡 {{ village["name"] }}
                </h2>

                <form
                    method="POST"
                    enctype="multipart/form-data"
                >

                    <label>Family ID</label>

                    <input
                        type="text"
                        name="family_id"
                        placeholder="Family ID"
                        required
                    >

                    <label>Member Name</label>

                    <input
                        type="text"
                        name="member_name"
                        placeholder="Member Name"
                        required
                    >

                    <label>Father / Husband Name</label>
                    <input
                        type="text"
                        name="father_husband_name"
                        placeholder="Father / Husband Name"
                    >

                    <label>Mother / Wife Name</label>
                    <input
                        type="text"
                        name="mother_wife_name"
                        placeholder="Mother / Wife Name"
                    >

                    <label>Date of Birth</label>
                    <input
                        type="date"
                        name="date_of_birth"
                    >

                    <label>Relationship</label>
                    <input
                        type="text"
                        name="relationship"
                        placeholder="Relationship"
                    >

                    <label>Age</label>

                    <input
                        type="number"
                        name="age"
                        placeholder="Age"
                    >

                    <label>Gender</label>

                    <select name="gender">

                        <option value="">
                            Select Gender
                        </option>

                        <option value="Male">
                            Male
                        </option>

                        <option value="Female">
                            Female
                        </option>

                        <option value="Other">
                            Other
                        </option>

                    </select>

                    <label>Marital Status</label>

                    <select name="marital_status">

                        <option value="">
                            Select Marital Status
                        </option>

                        <option value="Married">
                            Married
                        </option>

                        <option value="Unmarried">
                            Unmarried
                        </option>

                        <option value="Widow">
                            Widow
                        </option>

                        <option value="Widower">
                            Widower
                        </option>

                        <option value="Divorced">
                            Divorced
                        </option>

                    </select>

                    <label>Education</label>

                    <input
                        type="text"
                        name="education"
                        placeholder="Education"
                    >

                    <label>Occupation</label>
                    <input
                        type="text"
                        name="occupation"
                        placeholder="Occupation"
                    >

                    <label>Disability</label>
                    <input
                        type="text"
                        name="disability"
                        placeholder="Disability"
                    >

                    <label>Blood Group</label>
                    <input
                        type="text"
                        name="blood_group"
                        placeholder="Blood Group"
                    >

                    <label>Caste</label>

                    <input
                        type="text"
                        name="caste"
                        placeholder="Caste"
                    >

                    <label>Voter ID No</label>

                    <input
                        type="text"
                        name="voter_id_no"
                        placeholder="Voter ID No"
                    >

                    <label>Ration Card No</label>

                    <input
                        type="text"
                        name="ration_card_no"
                        placeholder="Ration Card No"
                    >

                    <label>Aadhaar No</label>

                    <input
                        type="text"
                        name="aadhaar_no"
                        placeholder="Aadhaar No"
                    >

                    <label>Bank Account No</label>
                    <input
                        type="text"
                        name="bank_account_no"
                        placeholder="Bank Account Number"
                    >

                    <label>IFSC Code</label>
                    <input
                        type="text"
                        name="ifsc_code"
                        placeholder="IFSC Code"
                    >

                    <label>Assessment Number</label>

                    <input
                        type="text"
                        name="assessment_number"
                        placeholder="Assessment Number"
                    >

                    <label>Land Details</label>

                    <textarea
                        name="land_details"
                        placeholder="Land Details"
                    ></textarea>

                    <label>Government Schemes</label>

                    <textarea
                        name="govt_schemes"
                        placeholder="Government Schemes"
                    ></textarea>

                    <label>Mobile</label>

                    <input
                        type="text"
                        name="mobile"
                        placeholder="Mobile Number"
                    >

                    <label>Gas Connection</label>
                    <select name="gas_connection">
                        <option value="">Select</option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                    </select>

                    <label>Electricity Connection</label>
                    <select name="electricity_connection">
                        <option value="">Select</option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                    </select>

                    <label>Electricity Service Number</label>
                    <input
                        type="text"
                        name="electricity_service_no"
                        placeholder="Electricity Service Number"
                    >

                    <label>Pension ID</label>

                    <input
                        type="text"
                        name="pension_id"
                        placeholder="Pension ID"
                    >

                    <label>Bike Details</label>
                    <input
                        type="text"
                        name="bike_details"
                        placeholder="Bike Yes/No, Registration No, Brand/Model"
                    >

                    <label>Car Details</label>
                    <input
                        type="text"
                        name="car_details"
                        placeholder="Car Yes/No, Registration No, Brand/Model"
                    >

                    <label>Other Vehicle Details</label>
                    <textarea
                        name="other_vehicle_details"
                        placeholder="Other Vehicle Details"
                    ></textarea>

                    <label>Address</label>

                    <textarea
                        name="address"
                        placeholder="Address"
                    ></textarea>

                    <label>Remarks</label>

                    <textarea
                        name="remarks"
                        placeholder="Remarks"
                    ></textarea>

                    <label>
                        📷 Profile Photo
                    </label>

                    <input
                        type="file"
                        name="photo"
                        accept=".jpg,.jpeg,.png,.webp"
                    >

                    <br>

                    <button
                        class="btn green"
                        type="submit"
                    >
                        💾 Save Family Member
                    </button>

                </form>

            </div>

            <a
                class="btn gray"
                href="/families/{{ village['id'] }}"
            >
                ← Back to Family
            </a>

        </div>
        """,
        village=village
    )

# =========================================================
# COMPLETE FAMILY MEMBER PROFILE
# =========================================================

@app.route("/family-profile/<int:member_id>")
def family_profile(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT
            family_members.*,
            villages.name AS village_name
        FROM family_members
        JOIN villages
        ON family_members.village_id = villages.id
        WHERE family_members.id = ?
    """, (member_id,)).fetchone()

    if member is None:
        conn.close()
        return "Member not found."

    files = conn.execute("""
        SELECT *
        FROM family_files
        WHERE member_id = ?
        ORDER BY id DESC
    """, (member_id,)).fetchall()
    family_members = conn.execute("""
    SELECT *
    FROM family_members
    WHERE family_id = ?
    ORDER BY id
    """, (member["family_id"],)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """

        <div class="container">

            <!-- PROFILE HEADER -->
            <div class="card">

                <div style="
                    display:flex;
                    align-items:center;
                    gap:25px;
                    flex-wrap:wrap;
                ">

                    {% if member["photo"] %}

                    <img
                        src="/profile-photo/{{ member['id'] }}"
                        style="
                            width:180px;
                            height:180px;
                            object-fit:cover;
                            border-radius:50%;
                            border:5px solid #2563eb;
                        "
                    >

                    {% else %}

                    <div style="
                        width:150px;
                        height:150px;
                        border-radius:50%;
                        background:#e2e8f0;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        font-size:70px;
                    ">
                        👤
                    </div>

                    {% endif %}

                    <div>

                        <h1>
                            👤 {{ member["member_name"] }}
                        </h1>

                        <h3>
                            Family ID:
                            {{ member["family_id"] }}
                        </h3>

                        <p>
                            🏠 Village:
                            <b>{{ member["village_name"] }}</b>
                        </p>

                    </div>

                </div>

            </div>

            <!-- PROFILE ACTION BUTTONS -->
            <div class="card" style="text-align:center;">

            <a href="/edit-family/{{ member['id'] }}"
                class="modern-btn btn-info"
                style="margin:5px;">
                ✏️ Edit Member
            </a>

            <a href="/family-profile-pdf/{{ member['id'] }}"
                class="modern-btn btn-info"
                style="margin:5px;">
                📄 Member Report PDF
            </a>

            <a href="/village/{{ member['village_id'] }}"
                class="modern-btn btn-back"
                style="margin:5px;">
                ← Back to Village
            </a>

             </div>


            <!-- PERSONAL INFORMATION -->
            <div class="card">

            <h2>👤 Personal Information</h2>

            <table class="table">

                <tr>
                    <th>S.No</th>
                    <th>Details</th>
                    <th>Information</th>
                </tr>

                <tr>
                    <td>1</td>
                    <th>Member Name</th>
                    <td>{{ member["member_name"] }}</td>
                </tr>

                <tr>
                    <td>2</td>
                    <th>Family ID</th>
                    <td>{{ member["family_id"] }}</td>
                </tr>

                <tr>
                    <td>3</td>
                    <th>Age</th>
                    <td>{{ member["age"] }}</td>
                </tr>

                <tr>
                    <td>4</td>
                    <th>Gender</th>
                    <td>{{ member["gender"] }}</td>
                </tr>

                <tr>
                    <td>5</td>
                    <th>Marital Status</th>
                    <td>{{ member["marital_status"] }}</td>
                </tr>

                <tr>
                    <td>6</td>
                    <th>Father / Husband Name</th>
                    <td>{{ member["father_husband_name"] }}</td>
                </tr>

                <tr>
                    <td>7</td>
                    <th>Mother / Wife Name</th>
                    <td>{{ member["mother_wife_name"] }}</td>
                </tr>

                <tr>
                    <td>8</td>
                    <th>Date of Birth</th>
                    <td>{{ member["date_of_birth"] }}</td>
                </tr>

                <tr>
                    <td>9</td>
                    <th>Relationship</th>
                    <td>{{ member["relationship"] }}</td>
                </tr>

                <tr>
                    <td>10</td>
                    <th>Occupation</th>
                    <td>{{ member["occupation"] }}</td>
                </tr>

                <tr>
                    <td>11</td>
                    <th>Disability</th>
                    <td>{{ member["disability"] }}</td>
                </tr>

                <tr>
                    <td>12</td>
                    <th>Blood Group</th>
                    <td>{{ member["blood_group"] }}</td>
                </tr>

                <tr>
                    <td>13</td>
                    <th>Bank Account No</th>
                    <td>{{ member["bank_account_no"] }}</td>
                </tr>

                <tr>
                    <td>14</td>
                    <th>IFSC Code</th>
                    <td>{{ member["ifsc_code"] }}</td>
                </tr>

                <tr>
                    <td>15</td>
                    <th>Education</th>
                    <td>{{ member["education"] }}</td>
                </tr>

                <tr>
                    <td>16</td>
                    <th>Caste</th>
                    <td>{{ member["caste"] }}</td>
                </tr>

            </table>

            </div>


            <!-- IDENTITY DETAILS -->
            <div class="card">

                <h2>🪪 Identity Details</h2>

                <table>

                    <tr>
                        <th>Voter ID No</th>
                        <td>{{ member["voter_id_no"] }}</td>
                    </tr>

                    <tr>
                        <th>Ration Card No</th>
                        <td>{{ member["ration_card_no"] }}</td>
                    </tr>

                    <tr>
                        <th>Aadhaar No</th>
                        <td>{{ member["aadhaar_no"] }}</td>
                    </tr>

                    <tr>
                        <th>Assessment Number</th>
                        <td>{{ member["assessment_number"] }}</td>
                    </tr>

                    <tr>
                        <th>Pension ID</th>
                        <td>{{ member["pension_id"] }}</td>
                    </tr>

                </table>

            </div>


            <!-- LAND & GOVERNMENT SCHEMES -->
            <div class="card">

                <h2>🌾 Land & Government Schemes</h2>

                <table>

                    <tr>
                        <th>Land Details</th>
                        <td>{{ member["land_details"] }}</td>
                    </tr>

                    <tr>
                        <th>Government Schemes</th>
                        <td>{{ member["govt_schemes"] }}</td>
                    </tr>

                </table>

            </div>

                        <!-- HOUSE & VEHICLE DETAILS -->
            <div class="card">

                <h2>🏠 Utilities & Vehicle Details</h2>

                <table>

                    <tr>
                        <th>Gas Connection</th>
                        <td>{{ member["gas_connection"] }}</td>
                    </tr>

                    <tr>
                        <th>Electricity Connection</th>
                        <td>{{ member["electricity_connection"] }}</td>
                    </tr>

                    <tr>
                        <th>Electricity Service Number</th>
                        <td>{{ member["electricity_service_no"] }}</td>
                    </tr>

                    <tr>
                        <th>Bike Details</th>
                        <td>{{ member["bike_details"] }}</td>
                    </tr>

                    <tr>
                        <th>Car Details</th>
                        <td>{{ member["car_details"] }}</td>
                    </tr>

                    <tr>
                        <th>Other Vehicle Details</th>
                        <td>{{ member["other_vehicle_details"] }}</td>
                    </tr>

                </table>

            </div>


            <!-- CONTACT -->
            <div class="card">

                <h2>📞 Contact & Address</h2>

                <table>

                    <tr>
                        <th>Mobile</th>
                        <td>{{ member["mobile"] }}</td>
                    </tr>

                    <tr>
                        <th>Address</th>
                        <td>{{ member["address"] }}</td>
                    </tr>

                    <tr>
                        <th>Remarks</th>
                        <td>{{ member["remarks"] }}</td>
                    </tr>

                </table>

            </div>


            <!-- DOCUMENTS -->
            <div class="card">

                <h2>📎 Uploaded Documents</h2>

                {% if files %}

                    {% for f in files %}

                    <div style="
                        background:#f8fafc;
                        padding:15px;
                        margin:10px 0;
                        border-radius:10px;
                        border:1px solid #ddd;
                    ">

                        <b>
                            📄 {{ f["document_type"] }}
                        </b>

                        <br>

                        {{ f["original_name"] }}

                        <br><br>

                        <a
                            class="btn blue"
                            href="/family-file-view/{{ f['id'] }}"
                        >
                            👁️ View
                        </a>

                        <a
                            class="btn green"
                            href="/family-file-download/{{ f['id'] }}"
                        >
                            📥 Download
                        </a>

                    </div>

                    {% endfor %}

                {% else %}

                    <p>
                        📭 No documents uploaded.
                    </p>

                {% endif %}

            </div>


            <!-- ACTIONS -->
            <div class="card">

                <a
                    class="btn orange"
                    href="/edit-family/{{ member['id'] }}"
                >
                    ✏️ Edit Profile
                </a>

                <a
                   class="btn green"
                   href="/family-profile-pdf/{{ member['id'] }}"
                >
                   📄 Download PDF
                </a>

                <button
                  class="btn blue"
                  onclick="window.print()"
                >
                  🖨️ Print
                </button>

                <a
                    class="btn blue"
                    href="/family-files/{{ member['id'] }}"
                >
                    📎 Manage Files
                </a>

                <a
                    class="btn gray"
                    href="/families/{{ member['village_id'] }}"
                >
                    ← Back to Family List
                </a>

            </div>

        </div>

        """,
        member=member,
        files=files
    )

# =========================================================
# DELETE FAMILY
# =========================================================

# =========================================================
# DELETE FAMILY
# =========================================================

@app.route("/delete-family/<int:member_id>")
def delete_family(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT *
        FROM family_members
        WHERE id = ?
    """, (member_id,)).fetchone()

    if member is None:
        conn.close()
        return "Member not found."

    village_id = member["village_id"]

    files = conn.execute("""
        SELECT stored_name
        FROM family_files
        WHERE member_id = ?
    """, (member_id,)).fetchall()

    for f in files:

        path = os.path.join(
            FAMILY_UPLOAD_FOLDER,
            f["stored_name"]
        )

        if os.path.exists(path):
            os.remove(path)

    conn.execute("""
        DELETE FROM family_files
        WHERE member_id = ?
    """, (member_id,))

    conn.execute("""
        DELETE FROM family_members
        WHERE id = ?
    """, (member_id,))

    conn.commit()
    conn.close()

    return redirect(
        "/families/" + str(village_id)
    )


# =========================================================
# FAMILY FILES PAGE
# =========================================================

@app.route(
    "/family-files/<int:member_id>",
    methods=["GET", "POST"]
)
def family_files(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT
            family_members.*,
            villages.name AS village_name
        FROM family_members
        JOIN villages
        ON family_members.village_id = villages.id
        WHERE family_members.id = ?
    """, (member_id,)).fetchone()

    if member is None:
        conn.close()
        return "Member not found."

    if request.method == "POST":

        document_type = request.form.get(
            "document_type",
            ""
        )

        uploaded = request.files.get("file")

        if not uploaded or uploaded.filename == "":
            conn.close()
            return "Please select a file."

        if not allowed_file(uploaded.filename):
            conn.close()
            return "Invalid file format."

        original_name = secure_filename(
            uploaded.filename
        )

        extension = file_extension(
            original_name
        )

        stored_name = (
            str(member_id)
            + "_"
            + str(abs(hash(original_name)))
            + "."
            + extension
        )

        path = os.path.join(
            FAMILY_UPLOAD_FOLDER,
            stored_name
        )

        uploaded.save(path)

        conn.execute("""
            INSERT INTO family_files
            (
                member_id,
                document_type,
                original_name,
                stored_name
            )
            VALUES(?,?,?,?)
        """, (
            member_id,
            document_type,
            original_name,
            stored_name
        ))

        conn.commit()
        conn.close()

        return redirect(
            "/family-files/" + str(member_id)
        )

    files = conn.execute("""
        SELECT *
        FROM family_files
        WHERE member_id = ?
        ORDER BY id DESC
    """, (member_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card">

                <h1>
                    📎 Scan / Upload Files
                </h1>

                <h2>
                    {{ member["member_name"] }}
                </h2>

                <form
                    method="POST"
                    enctype="multipart/form-data"
                >

                    <label>
                        Document Type
                    </label>

                    <select name="document_type" required>

                        <option value="">
                            Select Document
                        </option>

                        <option>Voter ID</option>
                        <option>Ration Card</option>
                        <option>Aadhaar Card</option>
                        <option>Assessment Number</option>
                        <option>Land Document</option>
                        <option>Government Scheme</option>
                        <option>Pension Document</option>
                        <option>Photo</option>
                        <option>Other Document</option>

                    </select>

                    <label>
                        Select Scan / File
                    </label>

                    <input
                        type="file"
                        name="file"
                        accept=".jpg,.jpeg,.png,.webp,.pdf,.xlsx,.xls,.csv,.doc,.docx"
                        required
                    >

                    <button class="green">
                        📤 Upload File
                    </button>

                </form>

            </div>


            <div class="card">

                <h2>
                    Uploaded Files
                </h2>

                {% for f in files %}

                <div class="file-box">

                    <b>
                        {{ f["document_type"] }}
                    </b>

                    <br>

                    {{ f["original_name"] }}

                    <br><br>

                    <a
                        class="btn blue"
                        href="/family-file-view/{{ f['id'] }}"
                        target="_blank"
                    >
                        👁️ View
                    </a>

                    <a
                        class="btn green"
                        href="/family-file-download/{{ f['id'] }}"
                    >
                        ⬇️ Download
                    </a>

                    <a
                        class="btn red"
                        href="/family-file-delete/{{ f['id'] }}"
                        onclick="return confirm('Delete this file?')"
                    >
                        🗑️ Delete
                    </a>

                </div>

                {% else %}

                <p>
                    No files uploaded.
                </p>

                {% endfor %}

            </div>


            <a
                class="btn gray"
                href="/families/{{ member['village_id'] }}"
            >
                ← Back to Family
            </a>

        </div>
        """,
        member=member,
        files=files
    )

# =========================================================
# RAW FAMILY FILE
# =========================================================

@app.route("/family-file-raw/<int:file_id>")
def family_file_raw(file_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    f = conn.execute("""
        SELECT *
        FROM family_files
        WHERE id = ?
    """, (file_id,)).fetchone()

    conn.close()

    if f is None:
        return "File not found."

    path = os.path.join(
        FAMILY_UPLOAD_FOLDER,
        f["stored_name"]
    )

    if not os.path.exists(path):
        return "File missing."

    return send_file(
        path,
        as_attachment=False
    )


# =========================================================
# VIEW / PREVIEW FAMILY FILE
# =========================================================

@app.route("/family-file-view/<int:file_id>")
def family_file_view(file_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    f = conn.execute("""
        SELECT *
        FROM family_files
        WHERE id = ?
    """, (file_id,)).fetchone()

    conn.close()

    if f is None:
        return "File not found."

    path = os.path.join(
        FAMILY_UPLOAD_FOLDER,
        f["stored_name"]
    )

    if not os.path.exists(path):
        return "File missing."

    extension = file_extension(
        f["original_name"]
    ).lower()

    # -----------------------------------------------------
    # IMAGE PREVIEW
    # -----------------------------------------------------

    if extension in ["jpg", "jpeg", "png", "webp"]:

        return render_template_string("""
        <!DOCTYPE html>
        <html>
        <head>

            <title>Image Preview</title>

            <style>

                body {
    margin: 0;
    padding: 0;

    font-family: Arial, sans-serif;

    min-height: 100vh;

    background:
        linear-gradient(
            rgba(240, 248, 255, 0.82),
            rgba(220, 252, 231, 0.82)
        ),
        linear-gradient(
            135deg,
            #dbeafe,
            #dcfce7,
            #fef3c7
        );

    background-attachment: fixed;
    background-size: cover;

    color: #1e293b;
}

                .preview-card {
                    background: white;
                    padding: 25px;
                    border-radius: 18px;
                    box-shadow: 0 5px 20px rgba(0,0,0,0.15);
                    max-width: 1000px;
                    margin: auto;
                }

                img {
                    max-width: 100%;
                    max-height: 75vh;
                    border-radius: 12px;
                    object-fit: contain;
                }

                h2 {
                    color: #24446b;
                }

                .btn {
                    display: inline-block;
                    padding: 10px 18px;
                    margin: 10px 5px;
                    border-radius: 8px;
                    text-decoration: none;
                    color: white;
                    font-weight: bold;
                    background: #2563eb;
                }

            </style>

        </head>

        <body>

            <div class="preview-card">

                <h2>🖼️ {{ filename }}</h2>

                <img
                    src="{{ image_url }}"
                    alt="Document Preview"
                >

                <br>

                <a
                    class="btn"
                    href="/family-file-download/{{ file_id }}"
                >
                    ⬇️ Download
                </a>

            </div>

        </body>
        </html>
        """,
        filename=f["original_name"],
        image_url="/family-file-raw/" + str(file_id),
        file_id=file_id
        )

    # -----------------------------------------------------
    # PDF PREVIEW
    # -----------------------------------------------------

    if extension == "pdf":

        return render_template_string("""
        <!DOCTYPE html>
        <html>

        <head>

            <title>PDF Preview</title>

            <style>

                body {
                    margin: 0;
                    min-height: 100vh;

                    background-imaege:
                    lineare-gradient(
                        rgba(0, 0, 0, 0.18),
                        rgba(0, 0, 0, 0.18)
                        ),
                    
                    url("/static/village.jpg");

                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                    background-repeat: no-repeat;

                    font-family: Arial, sans-serif;
                    }

                    }background: #eef6ff;
                    font-family: Arial, sans-serif;
                }

                .topbar {
                    background: white;
                    padding: 15px;
                    text-align: center;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.12);
                }

                h2 {
                    color: #24446b;
                }

                iframe {
                    width: 100%;
                    height: 88vh;
                    border: none;
                }

                .btn {
                    display: inline-block;
                    padding: 9px 18px;
                    border-radius: 8px;
                    text-decoration: none;
                    color: white;
                    background: #16a34a;
                    font-weight: bold;
                }

            </style>

        </head>

        <body>

            <div class="topbar">

                <h2>📄 {{ filename }}</h2>

                <a
                    class="btn"
                    href="/family-file-download/{{ file_id }}"
                >
                    ⬇️ Download PDF
                </a>

            </div>

            <iframe
                src="/family-file-raw/{{ file_id }}"
            ></iframe>

        </body>

        </html>
        """,
        filename=f["original_name"],
        file_id=file_id
        )

    # -----------------------------------------------------
    # OTHER FILE TYPES
    # -----------------------------------------------------

    return render_template_string("""
    <!DOCTYPE html>

    <html>

    <head>

        <title>File</title>

        <style>

            body {
                font-family: Arial;
                text-align: center;
                background: #eef6ff;
                padding: 60px;
            }

            .card {
                background: white;
                max-width: 600px;
                margin: auto;
                padding: 40px;
                border-radius: 18px;
                box-shadow: 0 5px 20px rgba(0,0,0,0.15);
            }

            .btn {
                display: inline-block;
                padding: 12px 20px;
                margin-top: 20px;
                background: #16a34a;
                color: white;
                text-decoration: none;
                border-radius: 8px;
                font-weight: bold;
            }

        </style>

    </head>

    <body>

        <div class="card">

            <h1>📎 {{ filename }}</h1>

            <p>
                This file type cannot be previewed
                directly in the browser.
            </p>

            <a
                class="btn"
                href="/family-file-download/{{ file_id }}"
            >
                ⬇️ Download File
            </a>

        </div>

    </body>

    </html>
    """,
    filename=f["original_name"],
    file_id=file_id
    )


# =========================================================
# DOWNLOAD FAMILY FILE
# =========================================================

@app.route("/family-file-download/<int:file_id>")
def family_file_download(file_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    f = conn.execute("""
        SELECT *
        FROM family_files
        WHERE id = ?
    """, (file_id,)).fetchone()

    conn.close()

    if f is None:
        return "File not found."

    path = os.path.join(
        FAMILY_UPLOAD_FOLDER,
        f["stored_name"]
    )

    return send_file(
        path,
        as_attachment=True,
        download_name=f["original_name"]
    )

# =========================================================
# PROFILE PHOTO VIEW
# =========================================================

@app.route("/profile-photo/<int:member_id>")
def profile_photo(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT photo
        FROM family_members
        WHERE id = ?
    """, (member_id,)).fetchone()

    conn.close()

    if member is None or not member["photo"]:
        return "Photo not found."

    photo_path = os.path.join(UPLOAD_FOLDER, member["photo"])

    if not os.path.exists(photo_path):
        return "Photo file not found."

    return send_file(photo_path)

# =========================================================
# EDIT FAMILY MEMBER
# =========================================================

@app.route("/edit-family/<int:member_id>", methods=["GET", "POST"])
def edit_family(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT *
        FROM family_members
        WHERE id = ?
    """, (member_id,)).fetchone()

    if member is None:
        conn.close()
        return "Family member not found."

    if request.method == "POST":

        family_id = request.form.get("family_id", "")
        member_name = request.form.get("member_name", "")
        age = request.form.get("age", "")
        gender = request.form.get("gender", "")
        marital_status = request.form.get("marital_status", "")
        education = request.form.get("education", "")
        caste = request.form.get("caste", "")
        voter_id_no = request.form.get("voter_id_no", "")
        ration_card_no = request.form.get("ration_card_no", "")
        aadhaar_no = request.form.get("aadhaar_no", "")
        assessment_number = request.form.get("assessment_number", "")
        land_details = request.form.get("land_details", "")
        govt_schemes = request.form.get("govt_schemes", "")
        mobile = request.form.get("mobile", "")
        pension_id = request.form.get("pension_id", "")
        address = request.form.get("address", "")
        remarks = request.form.get("remarks", "")

        # Existing photo
        photo_name = member["photo"]

        # New photo
        photo = request.files.get("photo")

        if photo and photo.filename:

            allowed = {"jpg", "jpeg", "png", "webp"}

            original_name = secure_filename(photo.filename)

            if "." in original_name:
                ext = original_name.rsplit(".", 1)[1].lower()
            else:
                ext = ""

            if ext not in allowed:
                conn.close()
                return "Only JPG, JPEG, PNG and WEBP photos are allowed."

            # Delete old photo
            if photo_name:
                old_photo = os.path.join(
                    UPLOAD_FOLDER,
                    photo_name
                )

                if os.path.exists(old_photo):
                    try:
                        os.remove(old_photo)
                    except:
                        pass

            # New unique photo name
            import uuid

            photo_name = (
                "family_"
                + str(member_id)
                + "_"
                + uuid.uuid4().hex
                + "."
                + ext
            )

            photo.save(
                os.path.join(
                    UPLOAD_FOLDER,
                    photo_name
                )
            )

        conn.execute("""
            UPDATE family_members
            SET
                family_id = ?,
                member_name = ?,
                age = ?,
                gender = ?,
                marital_status = ?,
                education = ?,
                caste = ?,
                voter_id_no = ?,
                ration_card_no = ?,
                aadhaar_no = ?,
                assessment_number = ?,
                land_details = ?,
                govt_schemes = ?,
                mobile = ?,
                pension_id = ?,
                address = ?,
                remarks = ?,
                photo = ?
            WHERE id = ?
        """, (
            family_id,
            member_name,
            age,
            gender,
            marital_status,
            education,
            caste,
            voter_id_no,
            ration_card_no,
            aadhaar_no,
            assessment_number,
            land_details,
            govt_schemes,
            mobile,
            pension_id,
            address,
            remarks,
            photo_name,
            member_id
        ))

        conn.commit()
        village_id = member["village_id"]

        conn.close()

        return redirect(
            "/families/" + str(village_id)
        )

    conn.close()

    return render_template_string(STYLE + """

    <div class="container">

        <div class="card">

            <h1>✏️ Edit Family Member</h1>

            {% if member["photo"] %}
                <div style="text-align:center;margin-bottom:20px;">

                    <img
                        src="/profile-photo/{{ member['id'] }}"
                        style="
                        width:120px;
                        height:120px;
                        object-fit:cover;
                        border-radius:50%;
                        border:5px solid #dbeafe;
                        box-shadow:0 4px 15px rgba(0,0,0,0.15);
                        "
                    >

                    <p>
                        <b>Current Profile Photo</b>
                    </p>

                </div>
            {% endif %}

            <form
                method="POST"
                enctype="multipart/form-data"
            >

                <label>Family ID</label>
                <input
                    type="text"
                    name="family_id"
                    value="{{ member['family_id'] or '' }}"
                    required
                >

                <label>Member Name</label>
                <input
                    type="text"
                    name="member_name"
                    value="{{ member['member_name'] or '' }}"
                    required
                >

                <label>Age</label>
                <input
                    type="number"
                    name="age"
                    value="{{ member['age'] or '' }}"
                >

                <label>Gender</label>
                <select name="gender">
                    <option value="">Select Gender</option>
                    <option value="Male"
                        {% if member["gender"] == "Male" %}selected{% endif %}>
                        Male
                    </option>
                    <option value="Female"
                        {% if member["gender"] == "Female" %}selected{% endif %}>
                        Female
                    </option>
                    <option value="Other"
                        {% if member["gender"] == "Other" %}selected{% endif %}>
                        Other
                    </option>
                </select>

                <label>Marital Status</label>
                <select name="marital_status">
                    <option value="">Select Status</option>
                    <option value="Married"
                        {% if member["marital_status"] == "Married" %}selected{% endif %}>
                        Married
                    </option>
                    <option value="Unmarried"
                        {% if member["marital_status"] == "Unmarried" %}selected{% endif %}>
                        Unmarried
                    </option>
                    <option value="Widow"
                        {% if member["marital_status"] == "Widow" %}selected{% endif %}>
                        Widow
                    </option>
                    <option value="Widower"
                        {% if member["marital_status"] == "Widower" %}selected{% endif %}>
                        Widower
                    </option>
                    <option value="Divorced"
                        {% if member["marital_status"] == "Divorced" %}selected{% endif %}>
                        Divorced
                    </option>
                </select>

                <label>Education</label>
                <input
                    type="text"
                    name="education"
                    value="{{ member['education'] or '' }}"
                >

                <label>Caste</label>
                <input
                    type="text"
                    name="caste"
                    value="{{ member['caste'] or '' }}"
                >

                <label>Voter ID No</label>
                <input
                    type="text"
                    name="voter_id_no"
                    value="{{ member['voter_id_no'] or '' }}"
                >

                <label>Ration Card No</label>
                <input
                    type="text"
                    name="ration_card_no"
                    value="{{ member['ration_card_no'] or '' }}"
                >

                <label>Aadhaar No</label>
                <input
                    type="text"
                    name="aadhaar_no"
                    value="{{ member['aadhaar_no'] or '' }}"
                >

                <label>Assessment Number</label>
                <input
                    type="text"
                    name="assessment_number"
                    value="{{ member['assessment_number'] or '' }}"
                >

                <label>Land Details</label>
                <textarea
                    name="land_details"
                    rows="3"
                >{{ member['land_details'] or '' }}</textarea>

                <label>Government Schemes</label>
                <textarea
                    name="govt_schemes"
                    rows="3"
                >{{ member['govt_schemes'] or '' }}</textarea>

                <label>Mobile</label>
                <input
                    type="text"
                    name="mobile"
                    value="{{ member['mobile'] or '' }}"
                >

                <label>Pension ID</label>
                <input
                    type="text"
                    name="pension_id"
                    value="{{ member['pension_id'] or '' }}"
                >

                <label>Address</label>
                <textarea
                    name="address"
                    rows="3"
                >{{ member['address'] or '' }}</textarea>

                <label>Remarks</label>
                <textarea
                    name="remarks"
                    rows="3"
                >{{ member['remarks'] or '' }}</textarea>

                <label>Change Profile Photo</label>

                <input
                    type="file"
                    name="photo"
                    accept=".jpg,.jpeg,.png,.webp"
                >

                <br><br>

                <button
                    type="submit"
                    class="btn blue"
                >
                    💾 Update Family Member
                </button>

                <a
                    class="btn gray"
                    href="/families/{{ member['village_id'] }}"
                >
                    ← Cancel
                </a>

            </form>

        </div>

    </div>

    """, member=member)

# =========================================================
# FAMILY MEMBER PROFILE PDF
# =========================================================

@app.route("/family-profile-pdf/<int:member_id>")
def family_profile_pdf(member_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    member = conn.execute("""
        SELECT
            family_members.*,
            villages.name AS village_name
        FROM family_members
        JOIN villages
        ON family_members.village_id = villages.id
        WHERE family_members.id = ?
    """, (member_id,)).fetchone()

    if member is None:
        conn.close()
        return "Member not found."

    files = conn.execute("""
        SELECT *
        FROM family_files
        WHERE member_id = ?
        ORDER BY id DESC
    """, (member_id,)).fetchall()

    conn.close()

    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle
    )
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    from io import BytesIO

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=35,
        leftMargin=35,
        topMargin=35,
        bottomMargin=35
    )

    styles = getSampleStyleSheet()

    title_style = styles["Title"]
    title_style.alignment = TA_CENTER

    story = []

    story.append(
        Paragraph(
            "FAMILY MEMBER PROFILE",
            title_style
        )
    )

    story.append(Spacer(1, 15))

    story.append(
        Paragraph(
            "<b>Village:</b> " +
            str(member["village_name"]),
            styles["Normal"]
        )
    )

    story.append(
        Paragraph(
            "<b>Family ID:</b> " +
            str(member["family_id"]),
            styles["Normal"]
        )
    )

    story.append(Spacer(1, 15))

    data = [
        ["Personal Details", ""],

        ["Member Name",
         str(member["member_name"])],

        ["Age",
         str(member["age"])],

        ["Gender",
         str(member["gender"])],

        ["Marital Status",
         str(member["marital_status"])],

        ["Education",
         str(member["education"])],

        ["Caste",
         str(member["caste"])],

        ["Voter ID No",
         str(member["voter_id_no"])],

        ["Ration Card No",
         str(member["ration_card_no"])],

        ["Aadhaar No",
         str(member["aadhaar_no"])],

        ["Assessment Number",
         str(member["assessment_number"])],

        ["Land Details",
         str(member["land_details"])],

        ["Government Schemes",
         str(member["govt_schemes"])],

        ["Mobile",
         str(member["mobile"])],

        ["Pension ID",
         str(member["pension_id"])],

        ["Address",
         str(member["address"])],

        ["Remarks",
         str(member["remarks"])]
    ]

    table = Table(
        data,
        colWidths=[160, 340]
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.black
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),

            (
                "PADDING",
                (0, 0),
                (-1, -1),
                7
            )
        ])
    )

    story.append(table)

    story.append(Spacer(1, 20))

    story.append(
        Paragraph(
            "<b>Uploaded Documents</b>",
            styles["Heading2"]
        )
    )

    if files:

        for f in files:

            story.append(
                Paragraph(
                    "• " +
                    str(f["document_type"]) +
                    " - " +
                    str(f["original_name"]),
                    styles["Normal"]
                )
            )

            story.append(Spacer(1, 5))

    else:

        story.append(
            Paragraph(
                "No documents uploaded.",
                styles["Normal"]
            )
        )

    doc.build(story)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=(
            "Family_Profile_" +
            str(member["member_name"]) +
            ".pdf"
        ),
        mimetype="application/pdf"
    )

# =========================================================
# DELETE FAMILY FILE
# =========================================================

@app.route("/family-file-delete/<int:file_id>")
def family_file_delete(file_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    f = conn.execute("""
        SELECT *
        FROM family_files
        WHERE id = ?
    """, (file_id,)).fetchone()

    if f is None:
        conn.close()
        return "File not found."

    member_id = f["member_id"]

    path = os.path.join(
        FAMILY_UPLOAD_FOLDER,
        f["stored_name"]
    )

    if os.path.exists(path):
        os.remove(path)

    conn.execute("""
        DELETE FROM family_files
        WHERE id = ?
    """, (file_id,))

    conn.commit()
    conn.close()

    return redirect(
        "/family-files/" + str(member_id)
    )


# =========================================================
# IMPORTANT DETAILS
# =========================================================

@app.route(
    "/important/<int:village_id>",
    methods=["GET", "POST"]
)
def important(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if village is None:
        conn.close()
        return "Village not found."

    if request.method == "POST":

        title = request.form.get(
            "title",
            ""
        ).strip()

        detail = request.form.get(
            "detail",
            ""
        ).strip()

        detail_date = request.form.get(
            "detail_date",
            ""
        )

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        uploaded = request.files.get("file")

        file_name = ""
        stored_name = ""

        if uploaded and uploaded.filename:

            if not allowed_file(uploaded.filename):
                conn.close()
                return "Invalid file format."

            file_name = secure_filename(
                uploaded.filename
            )

            extension = file_extension(
                file_name
            )

            stored_name = (
                str(village_id)
                + "_important_"
                + str(abs(hash(file_name)))
                + "."
                + extension
            )

            uploaded.save(
                os.path.join(
                    IMPORTANT_FOLDER,
                    stored_name
                )
            )

        if not title:
            conn.close()
            return "Title is required."

        conn.execute("""
            INSERT INTO important_details
            (
                village_id,
                title,
                detail,
                detail_date,
                remarks,
                file_name,
                stored_name
            )
            VALUES(?,?,?,?,?,?,?)
        """, (
            village_id,
            title,
            detail,
            detail_date,
            remarks,
            file_name,
            stored_name
        ))

        conn.commit()

        return redirect(
            "/important/" + str(village_id)
        )

    details = conn.execute("""
        SELECT *
        FROM important_details
        WHERE village_id = ?
        ORDER BY id DESC
    """, (village_id,)).fetchall()

    conn.close()

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card">

                <h1>
                    📁 Important Details
                </h1>

                <h2>
                    🏡 {{ village["name"] }}
                </h2>

                <form
                    method="POST"
                    enctype="multipart/form-data"
                >

                    <label>
                        Important Title
                    </label>

                    <input
                        name="title"
                        placeholder="Example: Road Work"
                        required
                    >

                    <label>
                        Date
                    </label>

                    <input
                        type="date"
                        name="detail_date"
                    >

                    <label>
                        Important Details
                    </label>

                    <textarea
                        name="detail"
                        placeholder="Write important details here..."
                    ></textarea>

                    <label>
                        Remarks
                    </label>

                    <textarea
                        name="remarks"
                        placeholder="Remarks"
                    ></textarea>

                    <label>
                        Attach Excel / PDF / Document
                    </label>

                    <input
                        type="file"
                        name="file"
                        accept=".pdf,.xlsx,.xls,.csv,.doc,.docx,.jpg,.jpeg,.png,.webp"
                    >

                    <button class="green">
                        💾 Save Important Detail
                    </button>

                </form>

            </div>


            <div class="card">

                <h2>
                    📌 Saved Important Details
                </h2>

                {% for d in details %}

                <div class="file-box">

                    <h3>
                        {{ d["title"] }}
                    </h3>

                    <b>
                        Date:
                    </b>

                    {{ d["detail_date"] }}

                    <p>
                        {{ d["detail"] }}
                    </p>

                    <b>
                        Remarks:
                    </b>

                    <p>
                        {{ d["remarks"] }}
                    </p>

                    {% if d["stored_name"] %}

                    <p>
                        📎 {{ d["file_name"] }}
                    </p>

                    <a
                        class="btn blue"
                        href="/important-file/{{ d['id'] }}"
                        target="_blank"
                    >
                        👁️ View
                    </a>

                    <a
                        class="btn green"
                        href="/important-download/{{ d['id'] }}"
                    >
                        ⬇️ Download
                    </a>

                    {% endif %}

                    <a
                        class="btn orange"
                        href="/edit-important/{{ d['id'] }}"
                    >
                        ✏️ Edit
                    </a>

                    <a
                        class="btn red"
                        href="/delete-important/{{ d['id'] }}"
                        onclick="return confirm('Delete this detail?')"
                    >
                        🗑️ Delete
                    </a>

                </div>

                {% else %}

                <p>
                    No important details added.
                </p>

                {% endfor %}

            </div>


            <a
                class="btn gray"
                href="/village/{{ village['id'] }}"
            >
                ← Back to Village
            </a>

            <a
                class="btn gray"
                href="/dashboard"
            >
               🏠 Back to Dashboard
            </a>

        </div>
        """,
        village=village,
        details=details
    )


# =========================================================
# VIEW IMPORTANT FILE
# =========================================================

@app.route("/important-file/<int:detail_id>")
def important_file(detail_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    d = conn.execute("""
        SELECT *
        FROM important_details
        WHERE id = ?
    """, (detail_id,)).fetchone()

    conn.close()

    if d is None:
        return "File not found."

    path = os.path.join(
        IMPORTANT_FOLDER,
        d["stored_name"]
    )

    if not os.path.exists(path):
        return "File missing."

    return send_file(
        path,
        as_attachment=False,
        download_name=d["file_name"]
    )


# =========================================================
# DOWNLOAD IMPORTANT FILE
# =========================================================

@app.route("/important-download/<int:detail_id>")
def important_download(detail_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    d = conn.execute("""
        SELECT *
        FROM important_details
        WHERE id = ?
    """, (detail_id,)).fetchone()

    conn.close()

    if d is None:
        return "File not found."

    path = os.path.join(
        IMPORTANT_FOLDER,
        d["stored_name"]
    )

    return send_file(
        path,
        as_attachment=True,
        download_name=d["file_name"]
    )


# =========================================================
# DELETE IMPORTANT DETAIL
# =========================================================

@app.route("/delete-important/<int:detail_id>")
def delete_important(detail_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    d = conn.execute("""
        SELECT *
        FROM important_details
        WHERE id = ?
    """, (detail_id,)).fetchone()

    if d is None:
        conn.close()
        return "Detail not found."

    village_id = d["village_id"]

    if d["stored_name"]:

        path = os.path.join(
            IMPORTANT_FOLDER,
            d["stored_name"]
        )

        if os.path.exists(path):
            os.remove(path)

    conn.execute("""
        DELETE FROM important_details
        WHERE id = ?
    """, (detail_id,))

    conn.commit()
    conn.close()

    return redirect(
        "/important/" + str(village_id)
    )


# =========================================================
# EDIT IMPORTANT DETAIL
# =========================================================

@app.route(
    "/edit-important/<int:detail_id>",
    methods=["GET", "POST"]
)
def edit_important(detail_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    d = conn.execute("""
        SELECT *
        FROM important_details
        WHERE id = ?
    """, (detail_id,)).fetchone()

    if d is None:
        conn.close()
        return "Detail not found."

    if request.method == "POST":

        title = request.form.get(
            "title",
            ""
        ).strip()

        detail = request.form.get(
            "detail",
            ""
        ).strip()

        detail_date = request.form.get(
            "detail_date",
            ""
        )

        remarks = request.form.get(
            "remarks",
            ""
        ).strip()

        uploaded = request.files.get("file")

        file_name = d["file_name"]
        stored_name = d["stored_name"]

        if uploaded and uploaded.filename:

            if not allowed_file(uploaded.filename):
                conn.close()
                return "Invalid file format."

            if stored_name:

                old_path = os.path.join(
                    IMPORTANT_FOLDER,
                    stored_name
                )

                if os.path.exists(old_path):
                    os.remove(old_path)

            file_name = secure_filename(
                uploaded.filename
            )

            extension = file_extension(
                file_name
            )

            stored_name = (
                str(d["village_id"])
                + "_important_"
                + str(abs(hash(file_name)))
                + "."
                + extension
            )

            uploaded.save(
                os.path.join(
                    IMPORTANT_FOLDER,
                    stored_name
                )
            )

        conn.execute("""
            UPDATE important_details
            SET
                title = ?,
                detail = ?,
                detail_date = ?,
                remarks = ?,
                file_name = ?,
                stored_name = ?
            WHERE id = ?
        """, (
            title,
            detail,
            detail_date,
            remarks,
            file_name,
            stored_name,
            detail_id
        ))

        conn.commit()
        conn.close()

        return redirect(
            "/important/"
            + str(d["village_id"])
        )

    conn.close()

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card">

                <h1>
                    ✏️ Edit Important Detail
                </h1>

                <form
                    method="POST"
                    enctype="multipart/form-data"
                >

                    <label>
                        Title
                    </label>

                    <input
                        name="title"
                        value="{{ d['title'] }}"
                        required
                    >

                    <label>
                        Date
                    </label>

                    <input
                        type="date"
                        name="detail_date"
                        value="{{ d['detail_date'] }}"
                    >

                    <label>
                        Details
                    </label>

                    <textarea name="detail">{{ d["detail"] }}</textarea>

                    <label>
                        Remarks
                    </label>

                    <textarea name="remarks">{{ d["remarks"] }}</textarea>

                    <label>
                        Replace File
                    </label>

                    <input
                        type="file"
                        name="file"
                        accept=".pdf,.xlsx,.xls,.csv,.doc,.docx,.jpg,.jpeg,.png,.webp"
                    >

                    <button class="green">
                        💾 Update
                    </button>

                </form>

                <a
                    class="btn gray"
                    href="/important/{{ d['village_id'] }}"
                >
                    Back
                </a>

            </div>

        </div>
        """,
        d=d
    )


# =========================================================
# EDIT VILLAGE INFORMATION
# =========================================================

@app.route(
    "/edit-information/<int:village_id>",
    methods=["GET", "POST"]
)
def edit_information(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    info = conn.execute("""
        SELECT *
        FROM village_info
        WHERE village_id = ?
    """, (village_id,)).fetchone()

    if info is None:

        conn.execute("""
            INSERT INTO village_info(village_id)
            VALUES(?)
        """, (village_id,))

        conn.commit()

        info = conn.execute("""
            SELECT *
            FROM village_info
            WHERE village_id = ?
        """, (village_id,)).fetchone()

    if request.method == "POST":

        try:
            family_count = int(
                request.form.get(
                    "family_count",
                    0
                ) or 0
            )
        except ValueError:
            family_count = 0

        try:
            male_count = int(
                request.form.get(
                    "male_count",
                    0
                ) or 0
            )
        except ValueError:
            male_count = 0

        try:
            female_count = int(
                request.form.get(
                    "female_count",
                    0
                ) or 0
            )
        except ValueError:
            female_count = 0

        population = male_count + female_count

        try:
            houses = int(
                request.form.get(
                    "houses",
                    0
                ) or 0
            )
        except ValueError:
            houses = 0

        try:
            revenue = float(
                request.form.get(
                    "revenue",
                    0
                ) or 0
            )
        except ValueError:
            revenue = 0

        mobile = request.form.get(
            "mobile",
            ""
        )

        other_info = request.form.get(
            "other_info",
            ""
        )

        conn.execute("""
            UPDATE village_info
            SET
                family_count = ?,
                male_count = ?,
                female_count = ?,
                population = ?,
                houses = ?,
                revenue = ?,
                mobile = ?,
                other_info = ?
            WHERE village_id = ?
        """, (
            family_count,
            male_count,
            female_count,
            population,
            houses,
            revenue,
            mobile,
            other_info,
            village_id
        ))

        conn.commit()
        conn.close()

        return redirect(
            "/village/" + str(village_id)
        )

    conn.close()

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card village-edit-card">

                <h1>
                    ✏️ Village Information
                </h1>

                <h2>
                    {{ village["name"] }}
                </h2>

                <form method="POST" enctype="multipart/form-data">

    <div class="info-grid">

        <div class="info-field">
            <label>👨‍👩‍👧 Family Count</label>
            <input
                type="number"
                name="family_count"
                value="{{ info['family_count'] }}"
            >
        </div>

        <div class="info-field">
            <label>👨 Male Count</label>
            <input
                type="number"
                name="male_count"
                value="{{ info['male_count'] }}"
            >
        </div>

        <div class="info-field">
            <label>👩 Female Count</label>
            <input
                type="number"
                name="female_count"
                value="{{ info['female_count'] }}"
            >
        </div>

        <div class="info-field">
            <label>🏠 Houses</label>
            <input
                type="number"
                name="houses"
                value="{{ info['houses'] }}"
            >
        </div>

        <div class="info-field">
            <label>💰 Revenue</label>
            <input
                type="number"
                step="0.01"
                name="revenue"
                value="{{ info['revenue'] }}"
            >
        </div>

        <div class="info-field">
            <label>📱 Mobile</label>
            <input
                type="text"
                name="mobile"
                value="{{ info['mobile'] }}"
            >
        </div>

    </div>

    <div class="info-field full-field">
        <label>📝 Other Information</label>

        <textarea
            name="other_info"
            placeholder="Enter other information..."
        >{{ info["other_info"] }}</textarea>
    </div>

    <button class="green" type="submit">
        💾 Save Information
    </button>

    <a href="/village/{{ village['id'] }}" class="back-village-btn">
    ← Back to Village
    </a>
    
</form>

            </div>

        </div>
        """,
        village=village,
        info=info
    )


# =========================================================
# EDIT VILLAGE NAME
# =========================================================

@app.route(
    "/edit-village/<int:village_id>",
    methods=["GET", "POST"]
)
def edit_village(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if request.method == "POST":

        name = request.form.get(
            "village_name",
            ""
        ).strip()

        try:

            conn.execute("""
                UPDATE villages
                SET name = ?
                WHERE id = ?
            """, (
                name,
                village_id
            ))

            conn.commit()

        except sqlite3.IntegrityError:

            conn.close()

            return "Village already exists."

        conn.close()

        return redirect("/dashboard")

    conn.close()

    return render_template_string(
        STYLE + """
        <div class="container">

            <div class="card">

                <h1>
                    Edit Village
                </h1>

                <form method="POST">

                    <input
                        name="village_name"
                        value="{{ village['name'] }}"
                        required
                    >

                    <button class="green">
                        Update
                    </button>

                </form>

                <a
                    class="btn gray"
                    href="/dashboard"
                >
                    Back
                </a>

            </div>

        </div>
        """,
        village=village
    )


# =========================================================
# DELETE VILLAGE
# =========================================================

@app.route("/delete-village/<int:village_id>")
def delete_village(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT *
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    if village and village["name"] == "Gurrampalem":

        conn.close()

        return "Gurrampalem cannot be deleted."

    conn.execute("""
        DELETE FROM family_members2
        WHERE village_id = ?
    """, (village_id,))

    conn.execute("""
        DELETE FROM village_info
        WHERE village_id = ?
    """, (village_id,))

    conn.execute("""
        DELETE FROM important_details
        WHERE village_id = ?
    """, (village_id,))

    conn.execute("""
        DELETE FROM villages
        WHERE id = ?
    """, (village_id,))

    conn.commit()
    conn.close()

    return redirect("/dashboard")


# =========================================================
# VILLAGE EXCEL
# =========================================================

@app.route("/export-excel/<int:village_id>")
def export_excel(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT
            villages.name,
            village_info.family_count,
            village_info.male_count,
            village_info.female_count,
            village_info.population,
            village_info.houses,
            village_info.revenue,
            village_info.mobile,
            village_info.other_info
        FROM villages
        LEFT JOIN village_info
        ON villages.id = village_info.village_id
        WHERE villages.id = ?
    """, (village_id,)).fetchone()

    conn.close()

    if village is None:
        return "Village not found."

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Village Information"

    rows = [
        ("Village Name", village["name"]),
        ("Family Count", village["family_count"]),
        ("Male Count", village["male_count"]),
        ("Female Count", village["female_count"]),
        ("Population", village["population"]),
        ("Houses", village["houses"]),
        ("Revenue", village["revenue"]),
        ("Mobile", village["mobile"]),
        ("Other Information", village["other_info"])
    ]

    for row_number, row in enumerate(rows, 1):

        sheet.cell(
            row=row_number,
            column=1,
            value=row[0]
        )

        sheet.cell(
            row=row_number,
            column=2,
            value=row[1]
        )

    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["B"].width = 50

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=(
            village["name"]
            + "_Village.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# =========================================================
# FAMILY EXCEL
# =========================================================

@app.route("/export-family-excel/<int:village_id>")
def export_family_excel(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT name
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        ORDER BY family_id, member_name
    """, (village_id,)).fetchall()

    conn.close()

    workbook = Workbook()

    sheet = workbook.active

    headers = [
        "Family ID",
        "Member Name",
        "Age",
        "Gender",
        "Marital Status",
        "Education",
        "Caste",
        "Voter ID",
        "Ration Card",
        "Aadhaar",
        "Assessment Number",
        "Land Details",
        "Govt Schemes",
        "Mobile",
        "Pension ID",
        "Address",
        "Remarks"
    ]

    for col, header in enumerate(headers, 1):

        sheet.cell(
            row=1,
            column=col,
            value=header
        )

    row_number = 2

    for m in members:

        values = [
            m["family_id"],
            m["member_name"],
            m["age"],
            m["gender"],
            m["marital_status"],
            m["education"],
            m["caste"],
            m["voter_id_no"],
            m["ration_card_no"],
            m["aadhaar_no"],
            m["assessment_number"],
            m["land_details"],
            m["govt_schemes"],
            m["mobile"],
            m["pension_id"],
            m["address"],
            m["remarks"]
        ]

        for col, value in enumerate(values, 1):

            sheet.cell(
                row=row_number,
                column=col,
                value=value
            )

        row_number += 1

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=(
            village["name"]
            + "_Family.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# =========================================================
# FAMILY WORD
# =========================================================

@app.route("/export-family-word/<int:village_id>")
def export_family_word(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT name
        FROM villages
        WHERE id = ?
    """, (village_id,)).fetchone()

    members = conn.execute("""
        SELECT *
        FROM family_members
        WHERE village_id = ?
        ORDER BY family_id, member_name
    """, (village_id,)).fetchall()

    conn.close()

    document = Document()

    document.add_heading(
        "Family Information",
        level=1
    )

    document.add_paragraph(
        "Village: "
        + str(village["name"])
    )

    for m in members:

        document.add_heading(
            "Family ID: "
            + str(m["family_id"]),
            level=2
        )

        fields = [
            ("Member Name", m["member_name"]),
            ("Age", m["age"]),
            ("Gender", m["gender"]),
            ("Marital Status", m["marital_status"]),
            ("Education", m["education"]),
            ("Caste", m["caste"]),
            ("Voter ID", m["voter_id_no"]),
            ("Ration Card", m["ration_card_no"]),
            ("Aadhaar", m["aadhaar_no"]),
            ("Assessment Number", m["assessment_number"]),
            ("Land Details", m["land_details"]),
            ("Govt Schemes", m["govt_schemes"]),
            ("Mobile", m["mobile"]),
            ("Pension ID", m["pension_id"]),
            ("Address", m["address"]),
            ("Remarks", m["remarks"])
        ]

        for label, value in fields:

            document.add_paragraph(
                str(label)
                + ": "
                + str(value or "")
            )

    output = BytesIO()

    document.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=(
            village["name"]
            + "_Family.docx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# =========================================================
# VILLAGE WORD
# =========================================================

@app.route("/export-pdf/<int:village_id>")
def export_pdf(village_id):

    if not logged_in():
        return redirect("/")

    conn = get_db()

    village = conn.execute("""
        SELECT
            villages.name,
            village_info.family_count,
            village_info.male_count,
            village_info.female_count,
            village_info.population,
            village_info.houses,
            village_info.revenue,
            village_info.mobile,
            village_info.other_info
        FROM villages
        LEFT JOIN village_info
        ON villages.id = village_info.village_id
        WHERE villages.id = ?
    """, (village_id,)).fetchone()

    conn.close()

    if village is None:
        return "Village not found."

    output = BytesIO()

    pdf = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    story = []

    story.append(
        Paragraph(
            "Village Information",
            styles["Title"]
        )
    )

    story.append(Spacer(1, 15))

    rows = [
        ("Village Name", village["name"]),
        ("Family Count", village["family_count"]),
        ("Male Count", village["male_count"]),
        ("Female Count", village["female_count"]),
        ("Total Population", village["population"]),
        ("Houses", village["houses"]),
        ("Revenue", village["revenue"]),
        ("Mobile", village["mobile"]),
        ("Other Information", village["other_info"])
    ]

    for label, value in rows:

        text = (
            "<b>"
            + str(label)
            + ":</b> "
            + str(value or "")
        )

        story.append(
            Paragraph(
                text,
                styles["Normal"]
            )
        )

        story.append(
            Spacer(1, 8)
        )

    pdf.build(story)

    output.seek(0)

    filename = (
        str(village["name"])
        + "_Village.pdf"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session        .clear()

    return        redirect("/")


# =========================================================
# START APPLICATION
# =========================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
