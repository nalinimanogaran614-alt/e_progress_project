from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import mysql.connector
from werkzeug.security import generate_password_hash, check_password_hash
import config
import re
import os
import json
from datetime import datetime, date
from werkzeug.utils import secure_filename
from io import BytesIO, StringIO
import csv
import uuid
import tempfile
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
app = Flask(__name__)
app.secret_key = config.SECRET_KEY
def get_db_connection():
    return mysql.connector.connect(
        host=config.DB_HOST,
        user=config.DB_USER,
        password=config.DB_PASSWORD,
        database=config.DB_NAME,
        port=config.DB_PORT
    )
SEMESTER_SCHEMA_READY = False
def _infer_year_from_class(class_name):
    text = (class_name or "").lower()
    if "1st" in text or "first" in text:
        return 1
    if "2nd" in text or "second" in text:
        return 2
    if "3rd" in text or "third" in text:
        return 3
    return 1
def _year_label_for_class(class_name):
    """Return a readable academic-year label for a class name."""
    year_no = _infer_year_from_class(class_name)
    suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(year_no, 'th')
    return f'{year_no}{suffix} Year'
def _infer_course_type(course_name):
    """Infer UG/PG from the course/program name.
    The program type is never entered manually. Known PG degree keywords are
    checked first. Legacy/unknown names remain UG-compatible.
    """
    raw = (course_name or '').upper().strip()
    pg_patterns = (
        r'(?<![A-Z0-9])MCA(?![A-Z0-9])',
        r'(?<![A-Z0-9])M\.?SC(?![A-Z0-9])',
        r'(?<![A-Z0-9])M\.?COM(?![A-Z0-9])',
        r'(?<![A-Z0-9])MBA(?![A-Z0-9])',
        r'(?<![A-Z0-9])M\.?E(?![A-Z0-9])',
        r'(?<![A-Z0-9])M\.?TECH(?![A-Z0-9])',
        r'(?<![A-Z0-9])MA(?![A-Z0-9])',
        r'(?<![A-Z0-9])MS(?![A-Z0-9])',
    )
    if any(re.search(pattern, raw) for pattern in pg_patterns):
        return "PG"
    if re.search(r'\bMASTER(?:S)?\b|\bPOST\s*GRADUATE\b|\bPOSTGRADUATE\b', raw):
        return "PG"
    return "UG"
def _department_code_from_name(course_name):
    """Generate the legacy department_code automatically from course name."""
    raw = (course_name or '').upper()
    known = ["M.TECH", "MTECH", "MCA", "MBA", "M.COM", "MCOM", "M.SC", "MSC", "M.E", "ME",
             "B.TECH", "BTECH", "B.E", "BE", "BCA", "BBA", "B.COM", "BCOM", "B.SC", "BSC"]
    for code in known:
        if re.search(r'(?<![A-Z0-9])' + re.escape(code) + r'(?![A-Z0-9])', raw):
            return re.sub(r'[^A-Z0-9]', '', code)
    words = re.findall(r'[A-Z]+', raw)
    if not words:
        return 'COURSE'
    initials = ''.join(w[0] for w in words if w)
    return (initials or words[0])[:20]
def _course_type_for_department(cursor, department_id):
    cursor.execute("SELECT department_name FROM departments WHERE id=%s", (department_id,))
    row = cursor.fetchone() or {}
    return _infer_course_type(row.get("department_name", ""))
def _ensure_column(cursor, table_name, column_name, definition):
    """Create a column if it is missing from an existing table."""
    cursor.execute(
        """SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s AND COLUMN_NAME = %s""",
        (table_name, column_name),
    )
    if not cursor.fetchone():
        cursor.execute(f"ALTER TABLE `{table_name}` ADD COLUMN `{column_name}` {definition}")
def ensure_application_schema():
    """Create the legacy app tables and columns used by the admin dashboard."""
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(100) NOT NULL UNIQUE,
                password TEXT NOT NULL,
                role VARCHAR(50) NOT NULL DEFAULT 'student',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS departments (
                id INT AUTO_INCREMENT PRIMARY KEY,
                department_name VARCHAR(150) NOT NULL,
                department_code VARCHAR(50) DEFAULT NULL,
                hod_name VARCHAR(150) DEFAULT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS classes (
                id INT AUTO_INCREMENT PRIMARY KEY,
                department_id INT NOT NULL,
                class_name VARCHAR(100) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faculty (
                id INT AUTO_INCREMENT PRIMARY KEY,
                user_id INT NOT NULL UNIQUE,
                faculty_name VARCHAR(150) NOT NULL,
                email VARCHAR(150) DEFAULT NULL,
                phone VARCHAR(30) DEFAULT NULL,
                department_id INT DEFAULT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id INT AUTO_INCREMENT PRIMARY KEY,
                user_id INT DEFAULT NULL UNIQUE,
                register_number VARCHAR(100) NOT NULL UNIQUE,
                student_name VARCHAR(150) NOT NULL,
                email VARCHAR(150) DEFAULT NULL,
                phone VARCHAR(30) DEFAULT NULL,
                department_id INT DEFAULT NULL,
                class_id INT DEFAULT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS subjects (
                id INT AUTO_INCREMENT PRIMARY KEY,
                subject_code VARCHAR(100) NOT NULL UNIQUE,
                subject_name VARCHAR(200) NOT NULL,
                department_id INT DEFAULT NULL,
                class_id INT DEFAULT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS class_subjects (
                id INT AUTO_INCREMENT PRIMARY KEY,
                class_id INT NOT NULL,
                subject_id INT NOT NULL,
                UNIQUE KEY uq_class_subject (class_id, subject_id)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS marks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                student_id INT NOT NULL,
                subject_id INT NOT NULL,
                internal DECIMAL(6,2) DEFAULT NULL,
                external DECIMAL(6,2) DEFAULT NULL,
                total DECIMAL(6,2) DEFAULT NULL,
                grade VARCHAR(10) DEFAULT NULL,
                result VARCHAR(20) DEFAULT NULL,
                appreciation VARCHAR(200) DEFAULT NULL,
                ca1 DECIMAL(6,2) DEFAULT NULL,
                ca2 DECIMAL(6,2) DEFAULT NULL,
                assignment DECIMAL(6,2) DEFAULT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uq_marks_student_subject (student_id, subject_id)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ep_faculty_details (
                faculty_id INT PRIMARY KEY,
                gender VARCHAR(20) DEFAULT NULL,
                dob DATE DEFAULT NULL,
                address TEXT DEFAULT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ep_student_details (
                student_id INT PRIMARY KEY,
                gender VARCHAR(20) DEFAULT NULL,
                dob DATE DEFAULT NULL,
                address TEXT DEFAULT NULL
            )
        """)
        _ensure_column(cursor, 'departments', 'hod_name', 'VARCHAR(150) DEFAULT NULL')
        _ensure_column(cursor, 'departments', 'department_code', 'VARCHAR(50) DEFAULT NULL')
        _ensure_column(cursor, 'departments', 'department_name', 'VARCHAR(150) NOT NULL')
        _ensure_column(cursor, 'faculty', 'class_id', 'INT DEFAULT NULL')
        conn.commit()
    finally:
        cursor.close(); conn.close()
def save_department_batch(cur, department_id, course_type, start_year):
    """Create/update the department academic batch from START YEAR only.
    UG = 3 academic years, PG = 2 academic years.
    The admin never supplies an end year.
    """
    try:
        start = int(str(start_year).strip())
    except (TypeError, ValueError):
        raise ValueError('Start Year must be a valid year, for example 2026.')
    if start < 2000 or start > 2100:
        raise ValueError('Start Year must be between 2000 and 2100.')
    expected = 3 if course_type == 'UG' else 2
    end = start + expected
    batch = f'{start}-{end}'
    cur.execute("""
        INSERT INTO ep_department_batches
            (department_id,course_type,start_year,end_year,course_duration)
        VALUES(%s,%s,%s,%s,%s)
        ON DUPLICATE KEY UPDATE
            course_type=VALUES(course_type),
            start_year=VALUES(start_year),
            end_year=VALUES(end_year),
            course_duration=VALUES(course_duration)
    """, (department_id, course_type, start, end, batch))
    for y in range(start, end):
        name = f'{y}-{y+1}'
        cur.execute(
            "INSERT IGNORE INTO ep_academic_years(year_name,is_active) VALUES(%s,0)",
            (name,)
        )
    return start, end, batch
def sync_class_semesters(cur):
    """Automatically derive semester mappings from department batch + class year.
    No manual Semester Management is required. Existing data is preserved.
    """
    cur.execute("SELECT id,department_id,class_name FROM classes ORDER BY id")
    classes = cur.fetchall()
    for c in classes:
        class_id = c["id"] if isinstance(c, dict) else c[0]
        department_id = c["department_id"] if isinstance(c, dict) else c[1]
        class_name = c["class_name"] if isinstance(c, dict) else c[2]
        cur.execute("SELECT course_type,start_year,end_year FROM ep_department_batches WHERE department_id=%s LIMIT 1", (department_id,))
        batch = cur.fetchone()
        if not batch:
            continue
        course_type = batch["course_type"] if isinstance(batch, dict) else batch[0]
        start_year = batch["start_year"] if isinstance(batch, dict) else batch[1]
        year_no = _infer_year_from_class(class_name)
        max_years = 3 if course_type == "UG" else 2
        if year_no > max_years:
            continue
        target_start = int(start_year) + year_no - 1
        year_name = f"{target_start}-{target_start+1}"
        cur.execute("SELECT id FROM ep_academic_years WHERE year_name=%s LIMIT 1", (year_name,))
        yr = cur.fetchone()
        if not yr:
            continue
        academic_year_id = yr["id"] if isinstance(yr, dict) else yr[0]
        max_sem = 6 if course_type == "UG" else 4
        for sem in ((year_no * 2) - 1, year_no * 2):
            if sem > max_sem:
                continue
            cur.execute("""
                INSERT INTO ep_class_semesters
                    (class_id,academic_year_id,course_type,year_no,semester_no,section)
                VALUES(%s,%s,%s,%s,%s,'A')
                ON DUPLICATE KEY UPDATE course_type=VALUES(course_type),year_no=VALUES(year_no)
            """, (class_id, academic_year_id, course_type, year_no, sem))
def ensure_semester_schema():
    global SEMESTER_SCHEMA_READY
    if SEMESTER_SCHEMA_READY:
        return
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        try:
            for col in ("department_id", "class_id"):
                cur.execute("""SELECT COLUMN_TYPE, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS
                              WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='subjects' AND COLUMN_NAME=%s""", (col,))
                meta = cur.fetchone()
                if meta and str(meta.get("IS_NULLABLE", "YES")).upper() == "NO":
                    cur.execute(f"ALTER TABLE subjects MODIFY `{col}` {meta['COLUMN_TYPE']} NULL")
        except Exception:
            pass
        try:
            for col in ("ca1", "ca2", "assignment", "external", "internal", "total", "grade", "result", "appreciation"):
                cur.execute("""SELECT COLUMN_TYPE, IS_NULLABLE
                              FROM INFORMATION_SCHEMA.COLUMNS
                              WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='marks' AND COLUMN_NAME=%s""", (col,))
                meta = cur.fetchone()
                if meta and str(meta.get("IS_NULLABLE", "YES")).upper() == "NO":
                    cur.execute(f"ALTER TABLE marks MODIFY `{col}` {meta['COLUMN_TYPE']} NULL")
        except Exception:
            pass
        try:
            for col in ("grade", "result", "appreciation"):
                cur.execute("""SELECT COLUMN_TYPE, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS
                              WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='marks' AND COLUMN_NAME=%s""", (col,))
                meta = cur.fetchone()
                if meta and str(meta.get("IS_NULLABLE", "YES")).upper() == "NO":
                    cur.execute(f"ALTER TABLE marks MODIFY `{col}` {meta['COLUMN_TYPE']} NULL")
        except Exception:
            pass
        cur.execute("""
            CREATE TABLE IF NOT EXISTS result_settings (
                id INT PRIMARY KEY,
                visibility TINYINT(1) NOT NULL DEFAULT 0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    ON UPDATE CURRENT_TIMESTAMP
            )
        """)
        cur.execute("""
            INSERT INTO result_settings (id, visibility)
            VALUES (1, 0)
            ON DUPLICATE KEY UPDATE id = id
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_academic_years (
                id INT AUTO_INCREMENT PRIMARY KEY,
                year_name VARCHAR(20) NOT NULL UNIQUE,
                is_active TINYINT(1) NOT NULL DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_department_batches (
                id INT AUTO_INCREMENT PRIMARY KEY,
                department_id INT NOT NULL UNIQUE,
                course_type ENUM('UG','PG') NOT NULL,
                start_year INT NOT NULL,
                end_year INT NOT NULL,
                course_duration VARCHAR(20) NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_semesters (
                id INT AUTO_INCREMENT PRIMARY KEY,
                course_type ENUM('UG','PG') NOT NULL,
                semester_no INT NOT NULL,
                year_no INT NOT NULL,
                semester_label VARCHAR(40) NOT NULL,
                UNIQUE KEY uq_ep_sem (course_type, semester_no)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_class_semesters (
                id INT AUTO_INCREMENT PRIMARY KEY,
                class_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                course_type VARCHAR(5) NOT NULL,
                year_no INT NOT NULL,
                semester_no INT NOT NULL,
                section VARCHAR(30) DEFAULT 'A',
                UNIQUE KEY uq_ep_class_year_sem (class_id, academic_year_id, semester_no)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_subject_semesters (
                id INT AUTO_INCREMENT PRIMARY KEY,
                subject_id INT NOT NULL,
                class_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                UNIQUE KEY uq_ep_sub_sem (subject_id, academic_year_id, semester_no)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_faculty_assignments (
                id INT AUTO_INCREMENT PRIMARY KEY,
                faculty_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                class_id INT NOT NULL,
                subject_id INT NOT NULL,
                section VARCHAR(30) DEFAULT 'A',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE KEY uq_ep_fac_assign (faculty_id, academic_year_id, semester_no, class_id, subject_id)
            )
        """)
        try:
            cur.execute("SHOW INDEX FROM ep_subject_semesters WHERE Key_name='uq_ep_sub_sem'")
            if cur.fetchall():
                cur.execute("ALTER TABLE ep_subject_semesters DROP INDEX uq_ep_sub_sem")
            cur.execute("ALTER TABLE ep_subject_semesters ADD UNIQUE KEY uq_ep_sub_sem (subject_id,class_id,academic_year_id,semester_no)")
        except mysql.connector.Error as e:
            if getattr(e, 'errno', None) not in (1061, 1091):
                raise
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_student_semesters (
                id INT AUTO_INCREMENT PRIMARY KEY,
                student_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                year_no INT NOT NULL,
                class_id INT NOT NULL,
                section VARCHAR(30) DEFAULT 'A',
                status VARCHAR(20) DEFAULT 'Current',
                UNIQUE KEY uq_ep_student_sem (student_id, academic_year_id, semester_no)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_faculty_details (
                faculty_id INT PRIMARY KEY, gender VARCHAR(20) NULL, dob DATE NULL, address TEXT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_student_details (
                student_id INT PRIMARY KEY, gender VARCHAR(20) NULL, dob DATE NULL, address TEXT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_question_papers (
                id INT AUTO_INCREMENT PRIMARY KEY,
                faculty_id INT NOT NULL,
                subject_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                cia VARCHAR(10) NOT NULL,
                file_path TEXT,
                paper_json LONGTEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE KEY uq_ep_qp (faculty_id, subject_id, academic_year_id, semester_no, cia)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_assignment_marks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                student_id INT NOT NULL,
                subject_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                marks DECIMAL(6,2) NOT NULL DEFAULT 0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uq_ep_assignment_mark (student_id,subject_id,academic_year_id,semester_no)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ep_question_marks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                student_id INT NOT NULL,
                subject_id INT NOT NULL,
                academic_year_id INT NOT NULL,
                semester_no INT NOT NULL,
                exam_type VARCHAR(10) NOT NULL,
                question_no VARCHAR(30) NOT NULL,
                question_text TEXT NULL,
                max_marks DECIMAL(6,2) NOT NULL DEFAULT 0,
                obtained_marks DECIMAL(6,2) NULL,
                choice VARCHAR(5) NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uq_ep_qmark (student_id,subject_id,academic_year_id,semester_no,exam_type,question_no)
            )
        """)
        for ctype, count in (("UG", 6), ("PG", 4)):
            for sem in range(1, count + 1):
                year_no = (sem + 1) // 2
                label = f"Semester {sem}"
                cur.execute("""
                    INSERT INTO ep_semesters(course_type,semester_no,year_no,semester_label)
                    VALUES(%s,%s,%s,%s)
                    ON DUPLICATE KEY UPDATE semester_label=VALUES(semester_label)
                """, (ctype, sem, year_no, label))
        names=[]
        try:
            cur.execute("SELECT year_name FROM academic_years ORDER BY year_name")
            names=[r["year_name"] for r in cur.fetchall() if r.get("year_name")]
        except Exception:
            names=[]
        if not names:
            now=datetime.now()
            start=now.year if now.month >= 6 else now.year-1
            names=[f"{start}-{start+1}"]
        for name in names:
            cur.execute("INSERT IGNORE INTO ep_academic_years(year_name,is_active) VALUES(%s,0)", (name,))
        cur.execute("SELECT COUNT(*) AS n FROM ep_academic_years WHERE is_active=1")
        if int(cur.fetchone()["n"]) == 0:
            cur.execute("UPDATE ep_academic_years SET is_active=0")
            cur.execute("SELECT id FROM ep_academic_years ORDER BY id DESC LIMIT 1")
            row=cur.fetchone()
            if row:
                cur.execute("UPDATE ep_academic_years SET is_active=1 WHERE id=%s", (row["id"],))
        cur.execute("SELECT id,year_name FROM ep_academic_years ORDER BY id")
        years=cur.fetchall()
        active_id=next((r["id"] for r in years if r.get("year_name")), None)
        cur.execute("SELECT id FROM ep_academic_years WHERE is_active=1 LIMIT 1")
        ar=cur.fetchone()
        active_id=ar["id"] if ar else active_id
        sync_class_semesters(cur)
        cur.execute("SELECT id,class_id FROM subjects")
        for sub in cur.fetchall():
            cur.execute("SELECT academic_year_id,year_no,semester_no FROM ep_class_semesters WHERE class_id=%s ORDER BY academic_year_id,semester_no",(sub["class_id"],))
            cms=cur.fetchall()
            for cm in cms:
                cur.execute("""
                    INSERT INTO ep_subject_semesters(subject_id,class_id,academic_year_id,semester_no)
                    VALUES(%s,%s,%s,%s)
                    ON DUPLICATE KEY UPDATE class_id=VALUES(class_id)
                """,(sub["id"],sub["class_id"],cm["academic_year_id"],cm["semester_no"]))
        cur.execute("SELECT id,class_id FROM students")
        for st in cur.fetchall():
            cur.execute("SELECT academic_year_id,year_no,semester_no,section FROM ep_class_semesters WHERE class_id=%s ORDER BY academic_year_id,semester_no",(st["class_id"],))
            cms=cur.fetchall()
            for cm in cms:
                cur.execute("""
                    INSERT INTO ep_student_semesters(student_id,academic_year_id,semester_no,year_no,class_id,section,status)
                    VALUES(%s,%s,%s,%s,%s,%s,'Current')
                    ON DUPLICATE KEY UPDATE year_no=VALUES(year_no),class_id=VALUES(class_id),section=VALUES(section)
                """,(st["id"],cm["academic_year_id"],cm["semester_no"],cm["year_no"],st["class_id"],cm["section"] or 'A'))
        conn.commit()
        SEMESTER_SCHEMA_READY=True
    finally:
        cur.close(); conn.close()
@app.before_request
def _ensure_multi_semester_layer():
    if request.endpoint not in {"static"}:
        try:
            ensure_application_schema()
            ensure_semester_schema()
            conn = get_db_connection()
            cur = conn.cursor(dictionary=True)
            try:
                auto_rollover_academic_year(cur)
                cleanup_passed_out_students(cur)
                sync_current_active_semesters(cur)
                conn.commit()
            finally:
                cur.close()
                conn.close()
        except Exception:
            app.logger.exception("Database initialization failed during request handling.")
def _active_academic_start_year(cursor):
    ay = get_active_academic_year(cursor)
    if not ay:
        return None
    try:
        return int(str(ay["year_name"]).split("-")[0])
    except Exception:
        return None
def cleanup_passed_out_students(cur):
    """Remove only expired student operational data.
    Faculty, departments and Subject Master are never deleted. A student is
    considered passed out when the department batch end year is before the
    currently active academic year's start year.
    """
    active_start = _active_academic_start_year(cur)
    if not active_start:
        return 0
    cur.execute("""
        SELECT s.id
        FROM students s
        INNER JOIN classes c ON c.id=s.class_id
        INNER JOIN ep_department_batches b ON b.department_id=c.department_id
        WHERE b.end_year < %s
    """, (active_start,))
    ids=[r['id'] for r in cur.fetchall()]
    if not ids:
        return 0
    placeholders=','.join(['%s']*len(ids))
    cur.execute(f"SELECT user_id FROM students WHERE id IN ({placeholders}) AND user_id IS NOT NULL", ids)
    user_ids=[r['user_id'] for r in cur.fetchall()]
    for table,col in [('ep_student_semesters','student_id'),('ep_question_marks','student_id'),
                      ('ep_assignment_marks','student_id'),('marks','student_id')]:
        try:
            cur.execute(f"DELETE FROM {table} WHERE {col} IN ({placeholders})", ids)
        except mysql.connector.Error:
            pass
    try:
        cur.execute(f"DELETE FROM ep_student_details WHERE student_id IN ({placeholders})", ids)
    except mysql.connector.Error:
        pass
    try:
        cur.execute("""
            DELETE a FROM ep_faculty_assignments a
            INNER JOIN classes c ON c.id=a.class_id
            INNER JOIN ep_department_batches b ON b.department_id=c.department_id
            INNER JOIN ep_academic_years ay ON ay.id=a.academic_year_id
            WHERE b.end_year < %s AND CAST(SUBSTRING_INDEX(ay.year_name,'-',1) AS UNSIGNED) < %s
        """, (active_start, active_start))
        cur.execute("""
            DELETE q FROM ep_question_papers q
            INNER JOIN ep_academic_years ay ON ay.id=q.academic_year_id
            WHERE CAST(SUBSTRING_INDEX(ay.year_name,'-',1) AS UNSIGNED) < %s
        """, (active_start,))
    except mysql.connector.Error:
        pass
    cur.execute(f"DELETE FROM students WHERE id IN ({placeholders})", ids)
    if user_ids:
        user_placeholders=','.join(['%s']*len(user_ids))
        cur.execute(f"DELETE FROM users WHERE id IN ({user_placeholders}) AND role='student'", user_ids)
    return len(ids)
def auto_rollover_academic_year(cur):
    """Automatically set the current academic year; rollover happens every June 1."""
    today = date.today()
    start = today.year if (today.month, today.day) >= (6, 1) else today.year - 1
    year_name = f"{start}-{start + 1}"
    cur.execute("INSERT IGNORE INTO ep_academic_years(year_name,is_active) VALUES(%s,0)", (year_name,))
    cur.execute("UPDATE ep_academic_years SET is_active=0")
    cur.execute("UPDATE ep_academic_years SET is_active=1 WHERE year_name=%s", (year_name,))
    return year_name
def sync_current_active_semesters(cur):
    """Expose only the semesters belonging to each department's current batch year."""
    active = get_active_academic_year(cur)
    if not active:
        return
    try:
        active_start = int(str(active['year_name']).split('-')[0])
    except Exception:
        return
    active_id = int(active['id'])
    cur.execute("SELECT id,department_id,class_name FROM classes ORDER BY id")
    for c in cur.fetchall():
        class_id=c['id'] if isinstance(c,dict) else c[0]
        dep_id=c['department_id'] if isinstance(c,dict) else c[1]
        class_name=c['class_name'] if isinstance(c,dict) else c[2]
        cur.execute("SELECT course_type,start_year,end_year FROM ep_department_batches WHERE department_id=%s LIMIT 1", (dep_id,))
        b=cur.fetchone()
        if not b: continue
        ctype=b['course_type'] if isinstance(b,dict) else b[0]
        start=int(b['start_year'] if isinstance(b,dict) else b[1]); end=int(b['end_year'] if isinstance(b,dict) else b[2])
        current_year=active_start-start+1
        max_year=3 if ctype=='UG' else 2
        if not (start <= active_start < end and 1 <= current_year <= max_year): continue
        if _infer_year_from_class(class_name) != current_year: continue
        for sem in (current_year*2-1,current_year*2):
            cur.execute("""INSERT INTO ep_class_semesters
                (class_id,academic_year_id,course_type,year_no,semester_no,section)
                VALUES(%s,%s,%s,%s,%s,'A')
                ON DUPLICATE KEY UPDATE course_type=VALUES(course_type),year_no=VALUES(year_no),section=VALUES(section)""",
                (class_id,active_id,ctype,current_year,sem))
def enforce_active_assignment_year(cur, academic_year_id):
    active=get_active_academic_year(cur)
    if not active or int(active['id']) != int(academic_year_id):
        raise ValueError('Assignments can be created only for the active academic year.')
def get_academic_years(cursor):
    cursor.execute("SELECT id,year_name,is_active FROM ep_academic_years ORDER BY year_name DESC")
    return cursor.fetchall()
def get_active_academic_year(cursor):
    cursor.execute("SELECT id,year_name FROM ep_academic_years WHERE is_active=1 LIMIT 1")
    row=cursor.fetchone()
    if row: return row
    cursor.execute("SELECT id,year_name FROM ep_academic_years ORDER BY id DESC LIMIT 1")
    return cursor.fetchone()
def get_faculty_semesters(cursor, faculty_id, academic_year_id=None):
    if academic_year_id is None:
        ay=get_active_academic_year(cursor); academic_year_id=ay["id"] if ay else 0
    cursor.execute("""
        SELECT DISTINCT a.semester_no, cs.year_no, cs.course_type, cs.section
        FROM ep_faculty_assignments a
        LEFT JOIN ep_class_semesters cs ON cs.class_id=a.class_id AND cs.academic_year_id=a.academic_year_id
        WHERE a.faculty_id=%s AND a.academic_year_id=%s
        ORDER BY a.semester_no
    """, (faculty_id,academic_year_id))
    return cursor.fetchall()
def _ensure_faculty_student_semester_mappings(cursor, faculty_id, academic_year_id):
    """Self-heal student semester rows for classes actually assigned to a faculty.
    Older/legacy student records keep their current class in students.class_id.
    Some databases created before the semester layer can therefore have valid
    faculty assignments but no matching ep_student_semesters rows, which makes
    mark-entry pages look empty.  This helper creates only the missing rows for
    the faculty's real active-year assignments and never assigns a student to a
    different class.
    """
    if not faculty_id or not academic_year_id:
        return 0
    cursor.execute("""
        SELECT DISTINCT a.class_id, a.semester_no,
               COALESCE(cs.year_no, (a.semester_no + 1) DIV 2) AS year_no,
               COALESCE(NULLIF(a.section,''), NULLIF(cs.section,''), 'A') AS section
        FROM ep_faculty_assignments a
        LEFT JOIN ep_class_semesters cs
          ON cs.class_id=a.class_id
         AND cs.academic_year_id=a.academic_year_id
         AND cs.semester_no=a.semester_no
        WHERE a.faculty_id=%s AND a.academic_year_id=%s
    """, (faculty_id, academic_year_id))
    assignments = cursor.fetchall()
    changed = 0
    for a in assignments:
        cursor.execute("""
            INSERT INTO ep_student_semesters
                (student_id,academic_year_id,semester_no,year_no,class_id,section,status)
            SELECT s.id,%s,%s,%s,%s,%s,'Current'
            FROM students s
            WHERE s.class_id=%s
            ON DUPLICATE KEY UPDATE
                year_no=VALUES(year_no),
                class_id=VALUES(class_id),
                section=VALUES(section),
                status='Current'
        """, (academic_year_id, a['semester_no'], a['year_no'], a['class_id'],
              a.get('section') or 'A', a['class_id']))
        changed += max(int(getattr(cursor, 'rowcount', 0) or 0), 0)
    return changed
def get_faculty_subjects_for_semester(cursor, faculty_id, semester_no, academic_year_id):
    cursor.execute("""
        SELECT a.subject_id,s.subject_code,s.subject_name,a.class_id,c.class_name,a.section,
               a.semester_no, a.academic_year_id
        FROM ep_faculty_assignments a
        INNER JOIN subjects s ON s.id=a.subject_id
        INNER JOIN classes c ON c.id=a.class_id
        WHERE a.faculty_id=%s AND a.semester_no=%s AND a.academic_year_id=%s
        ORDER BY s.subject_name
    """, (faculty_id,semester_no,academic_year_id))
    return cursor.fetchall()
def get_selected_faculty_semester(cursor, faculty_id):
    """Return ONLY the current active academic year and one of its assigned semesters.
    A faculty request cannot switch to a historical academic year by query-string
    manipulation. This is the central access rule used by all faculty mark pages.
    """
    ay = get_active_academic_year(cursor)
    if not ay:
        return None, None, []
    active_year_id = int(ay["id"])
    sem = request.args.get("semester", type=int)
    if sem is None:
        sem = request.form.get("semester", type=int)
    available = get_faculty_semesters(cursor, faculty_id, active_year_id)
    valid = [int(x["semester_no"]) for x in available]
    if sem not in valid:
        sem = valid[0] if valid else None
    return ay, sem, available
@app.route("/admin/academic-years")
def admin_academic_years():
    """Display-only Academic Year page. The active year is fully automatic."""
    if not admin_required():
        return redirect(url_for("login"))
    conn=get_db_connection(); cur=conn.cursor(dictionary=True)
    try:
        auto_rollover_academic_year(cur)
        sync_current_active_semesters(cur)
        conn.commit()
        active = get_active_academic_year(cur)
        return render_template("admin/academic_years.html", active=active)
    finally:
        cur.close(); conn.close()
@app.route("/admin/semester-setup")
def admin_semester_setup():
    if not admin_required(): return redirect(url_for("login"))
    conn=get_db_connection(); cur=conn.cursor(dictionary=True)
    try:
        years=get_academic_years(cur); active=get_active_academic_year(cur)
        cur.execute("""
            SELECT cs.*,c.class_name,d.department_name,d.department_code
            FROM ep_class_semesters cs
            INNER JOIN classes c ON c.id=cs.class_id
            INNER JOIN departments d ON d.id=c.department_id
            ORDER BY d.department_name,cs.year_no,cs.semester_no,c.class_name
        """)
        mappings=cur.fetchall()
        cur.execute("SELECT c.id,c.class_name,d.department_name FROM classes c INNER JOIN departments d ON d.id=c.department_id ORDER BY d.department_name,c.class_name")
        classes=cur.fetchall()
        return render_template("admin/semester_setup.html",years=years,active=active,mappings=mappings,classes=classes)
    finally: cur.close(); conn.close()
@app.route("/admin/semester-setup/save", methods=["POST"])
def admin_semester_setup_save():
    if not admin_required(): return redirect(url_for("login"))
    conn=get_db_connection(); cur=conn.cursor(dictionary=True)
    try:
        ay=int(request.form.get("academic_year_id")); class_id=int(request.form.get("class_id")); sem=int(request.form.get("semester_no")); section='A'
        cur.execute("SELECT department_id FROM classes WHERE id=%s",(class_id,)); row=cur.fetchone()
        if not row: raise ValueError("Class not found.")
        ctype=_course_type_for_department(cur,row["department_id"]); maxsem=6 if ctype=="UG" else 4
        if sem<1 or sem>maxsem: raise ValueError(f"{ctype} supports semesters 1-{maxsem}.")
        year_no=(sem+1)//2
        cur.execute("""
            INSERT INTO ep_class_semesters(class_id,academic_year_id,course_type,year_no,semester_no,section)
            VALUES(%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE academic_year_id=VALUES(academic_year_id),course_type=VALUES(course_type),year_no=VALUES(year_no),semester_no=VALUES(semester_no),section=VALUES(section)
        """,(class_id,ay,ctype,year_no,sem,section))
        conn.commit(); flash("Class semester mapping saved.","success")
    except Exception as e:
        conn.rollback(); flash(str(e),"danger")
    finally: cur.close(); conn.close()
    return redirect(url_for("admin_semester_setup"))
@app.route("/admin/faculty-assignments", methods=["GET", "POST"])
@app.route("/admin/faculty-subject-assignments", methods=["GET", "POST"])
@app.route("/faculty-subject-assignments", methods=["GET", "POST"])
def admin_faculty_assignments():
    """Admin Faculty Subject Assignment manager.
    One faculty is shown once on the main page.  Individual assignment rows
    remain normalized in ep_faculty_assignments so the Faculty Dashboard can
    later use the exact Faculty + Subject + Class + Academic Year + Semester
    relationship.
    """
    if not admin_required():
        return redirect(url_for("login"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        auto_rollover_academic_year(cur)
        sync_class_semesters(cur)
        sync_current_active_semesters(cur)
        conn.commit()
        if request.method == "POST":
            try:
                faculty_id = int(request.form.get("faculty_id"))
                academic_year_id = int(request.form.get("academic_year_id"))
                semester_no = int(request.form.get("semester_no"))
                class_id = int(request.form.get("class_id"))
                subject_ids = request.form.getlist("subject_ids")
                if not subject_ids and request.form.get("subject_id"):
                    subject_ids = [request.form.get("subject_id")]
                subject_ids = list(dict.fromkeys(int(x) for x in subject_ids if str(x).isdigit()))
                if not subject_ids:
                    raise ValueError("Please select at least one subject.")
                _validate_assignment_context(cur, faculty_id, class_id, academic_year_id, semester_no)
                added, duplicates, repeat_conflicts = _insert_faculty_subject_assignments(
                    cur, faculty_id, class_id, academic_year_id, semester_no, subject_ids
                )
                conn.commit()
                if repeat_conflicts:
                    flash("Subject repeat blocked for the same student batch: " + ", ".join(repeat_conflicts) + ". A subject can be assigned only once across Semesters 1-6 for that cohort.", "warning")
                if added and duplicates:
                    flash(f"Subjects assigned successfully. {duplicates} duplicate assignment(s) skipped.", "warning")
                elif added and not repeat_conflicts:
                    flash("Subjects assigned successfully.", "success")
                elif not added and duplicates and not repeat_conflicts:
                    flash("This subject is already assigned for the selected Class, Academic Year and Semester.", "warning")
            except (TypeError, ValueError) as e:
                conn.rollback()
                flash(str(e), "danger")
            except mysql.connector.Error as e:
                conn.rollback()
                flash(_friendly_db_error(e, "Unable to save assignment."), "danger")
        cur.execute("SELECT id,faculty_name FROM faculty ORDER BY faculty_name")
        faculty_rows = cur.fetchall()
        cur.execute("SELECT id,department_name,department_code FROM departments ORDER BY department_name")
        departments = cur.fetchall()
        cur.execute("SELECT department_id,course_type,start_year,end_year,course_duration FROM ep_department_batches")
        department_batches = cur.fetchall()
        cur.execute("SELECT id,class_name,department_id FROM classes ORDER BY department_id,class_name,id")
        classes = cur.fetchall()
        cur.execute("SELECT id,subject_code,subject_name,department_id,class_id FROM subjects ORDER BY subject_name,subject_code")
        subjects = cur.fetchall()
        years = get_academic_years(cur)
        cur.execute("""
            SELECT cs.class_id,cs.academic_year_id,cs.semester_no,cs.year_no,
                   cs.course_type,cs.section,c.class_name,c.department_id
            FROM ep_class_semesters cs
            INNER JOIN classes c ON c.id=cs.class_id
            ORDER BY cs.academic_year_id DESC,cs.class_id,cs.semester_no
        """)
        class_semester_rows = cur.fetchall()
        cur.execute("""
            SELECT a.id,a.faculty_id,f.faculty_name,
                   c.id AS class_id,c.class_name,c.department_id,
                   d.department_name,
                   s.id AS subject_id,s.subject_code,s.subject_name,
                   a.academic_year_id,ay.year_name,a.semester_no,a.section,
                   cs.year_no,cs.course_type
            FROM ep_faculty_assignments a
            INNER JOIN faculty f ON f.id=a.faculty_id
            INNER JOIN classes c ON c.id=a.class_id
            LEFT JOIN departments d ON d.id=c.department_id
            INNER JOIN subjects s ON s.id=a.subject_id
            INNER JOIN ep_academic_years ay ON ay.id=a.academic_year_id
            LEFT JOIN ep_class_semesters cs
              ON cs.class_id=a.class_id
             AND cs.academic_year_id=a.academic_year_id
             AND cs.semester_no=a.semester_no
            ORDER BY f.faculty_name,ay.year_name DESC,a.semester_no,c.class_name,s.subject_name
        """)
        assignment_rows = cur.fetchall()
        assignments_by_faculty = {}
        for row in assignment_rows:
            item = dict(row)
            year_no = item.get("year_no")
            if not year_no:
                year_no = _infer_year_from_class(item.get("class_name"))
            item["year_label"] = f"{_roman_year(year_no)} Year" if year_no else "—"
            item["semester_label"] = f"Semester {item.get('semester_no')}"
            item["subject_label"] = f"{item.get('subject_code')} - {item.get('subject_name')}"
            assignments_by_faculty.setdefault(str(item["faculty_id"]), []).append(item)
        grouped = []
        for f in faculty_rows:
            rows = assignments_by_faculty.get(str(f["id"]), [])
            grouped.append({
                "id": f["id"],
                "faculty_name": f["faculty_name"],
                "assignment_count": len(rows),
            })
        return render_template(
            "admin/assignments.html",
            faculties=grouped,
            departments=departments,
            department_batches=department_batches,
            classes=classes,
            subjects=subjects,
            years=years,
            active_academic_year=get_active_academic_year(cur),
            class_semester_rows=class_semester_rows,
            assignments_by_faculty=assignments_by_faculty,
        )
    finally:
        cur.close()
        conn.close()
def _roman_year(number):
    return {1: "I", 2: "II", 3: "III", 4: "IV"}.get(int(number), str(number))
def _friendly_db_error(error, fallback):
    if getattr(error, "errno", None) == 1062:
        return "This subject is already assigned for the selected Class, Academic Year and Semester."
    return fallback
def _validate_assignment_context(cur, faculty_id, class_id, academic_year_id, semester_no):
    """Validate the exact assignment context and create the class/year mapping
    on demand when it is missing.
    Classes such as ``1st Year`` are reusable across academic years.  The old
    implementation only allowed academic years that happened to have a legacy
    ``ep_class_semesters`` row, which caused the Academic Year dropdown to show
    incomplete/stale values (for example only 2026-2027 and 2023-2024).
    """
    cur.execute("SELECT id FROM faculty WHERE id=%s LIMIT 1", (faculty_id,))
    if not cur.fetchone():
        raise ValueError("Selected faculty was not found.")
    cur.execute("SELECT id,department_id,class_name FROM classes WHERE id=%s LIMIT 1", (class_id,))
    cls = cur.fetchone()
    if not cls:
        raise ValueError("Selected class was not found.")
    enforce_active_assignment_year(cur, academic_year_id)
    cur.execute("SELECT id,year_name FROM ep_academic_years WHERE id=%s LIMIT 1", (academic_year_id,))
    selected_academic_year = cur.fetchone()
    if not selected_academic_year:
        raise ValueError("Selected academic year was not found.")
    year_no = _infer_year_from_class(cls.get("class_name"))
    if not year_no:
        raise ValueError("Unable to determine the selected Class / Year.")
    cur.execute("SELECT course_type,start_year,end_year FROM ep_department_batches WHERE department_id=%s LIMIT 1", (cls["department_id"],))
    batch = cur.fetchone()
    if not batch:
        raise ValueError("Academic structure is not configured for the selected department.")
    course_type = str(batch.get("course_type") or "UG").upper()
    max_years = 3 if course_type == "UG" else 2
    if int(year_no) < 1 or int(year_no) > max_years:
        raise ValueError("The selected Class / Year is not valid for this department.")
    allowed_semesters = ((int(year_no) * 2) - 1, int(year_no) * 2)
    if int(semester_no) not in allowed_semesters:
        raise ValueError("The selected semester does not belong to the selected Class / Year.")
    for sem in allowed_semesters:
        cur.execute("""
            INSERT INTO ep_class_semesters
                (class_id,academic_year_id,course_type,year_no,semester_no,section)
            VALUES(%s,%s,%s,%s,%s,'A')
            ON DUPLICATE KEY UPDATE
                course_type=VALUES(course_type), year_no=VALUES(year_no)
        """, (class_id, academic_year_id, course_type, year_no, sem))
    cur.execute("""
        SELECT id,year_no,section FROM ep_class_semesters
        WHERE class_id=%s AND academic_year_id=%s AND semester_no=%s LIMIT 1
    """, (class_id, academic_year_id, semester_no))
    mapping = cur.fetchone()
    if not mapping:
        raise ValueError("Unable to create the selected class, academic year and semester context.")
    return mapping
def _cohort_start_year(cur, class_id, academic_year_id, semester_no):
    """Return the student cohort start year for a class/semester context.
    Example: a UG Year-1 Semester-1/2 in 2026-27 belongs to the 2026 cohort;
    Year-2 Semester-3/4 in 2027-28 belongs to the same 2026 cohort.
    This lets us prevent the same subject from being repeated across later
    semesters for the same batch while still allowing the next batch to study
    the subject normally.
    """
    cur.execute("""
        SELECT c.department_id, c.class_name, ay.year_name, cs.year_no
        FROM classes c
        LEFT JOIN ep_class_semesters cs
          ON cs.class_id=c.id AND cs.academic_year_id=%s AND cs.semester_no=%s
        INNER JOIN ep_academic_years ay ON ay.id=%s
        WHERE c.id=%s LIMIT 1
    """, (academic_year_id, semester_no, academic_year_id, class_id))
    row = cur.fetchone()
    if not row:
        return None, None
    try:
        start = int(str(row.get('year_name') if isinstance(row, dict) else row[2]).split('-')[0])
        year_no = int((row.get('year_no') if isinstance(row, dict) else row[3]) or _infer_year_from_class(row.get('class_name') if isinstance(row, dict) else row[1]) or 1)
        return (row.get('department_id') if isinstance(row, dict) else row[0]), start - year_no + 1
    except Exception:
        return (row.get('department_id') if isinstance(row, dict) else row[0]), None
def _subject_repeat_conflict(cur, subject_id, class_id, academic_year_id, semester_no, exclude_assignment_id=None):
    """Find an earlier/later semester using the same subject for the same cohort.
    The rule is cohort-aware: a subject taken by AI 2026-27 students in
    Semester 1 cannot be assigned again to that same cohort in Semesters 2-6,
    but a new AI batch may use the same subject.
    """
    department_id, cohort_start = _cohort_start_year(cur, class_id, academic_year_id, semester_no)
    if department_id is None or cohort_start is None:
        return None
    cur.execute("""
        SELECT a.id, a.semester_no, ay.year_name, c.class_name, s.subject_code, s.subject_name
        FROM ep_faculty_assignments a
        INNER JOIN classes c ON c.id=a.class_id AND c.department_id=%s
        INNER JOIN ep_academic_years ay ON ay.id=a.academic_year_id
        INNER JOIN subjects s ON s.id=a.subject_id
        LEFT JOIN ep_class_semesters cs
          ON cs.class_id=a.class_id
         AND cs.academic_year_id=a.academic_year_id
         AND cs.semester_no=a.semester_no
        WHERE a.subject_id=%s
        ORDER BY ay.year_name, a.semester_no
    """, (department_id, subject_id))
    rows = cur.fetchall()
    for row in rows:
        if exclude_assignment_id is not None and int(row.get("id") or 0) == int(exclude_assignment_id):
            continue
        ay_name = row.get('year_name') if isinstance(row, dict) else row[2]
        class_name = row.get('class_name') if isinstance(row, dict) else row[3]
        try:
            row_start = int(str(ay_name).split('-')[0])
            row_year_no = _infer_year_from_class(class_name) or 1
            row_cohort = row_start - row_year_no + 1
        except Exception:
            continue
        if row_cohort == cohort_start:
            return row
    return None
def _insert_faculty_subject_assignments(cur, faculty_id, class_id, academic_year_id, semester_no, subject_ids):
    added = 0
    duplicates = 0
    repeat_conflicts = []
    for subject_id in subject_ids:
        cur.execute("SELECT id FROM subjects WHERE id=%s LIMIT 1", (subject_id,))
        subject = cur.fetchone()
        if not subject:
            continue
        cur.execute("""
            SELECT id FROM ep_faculty_assignments
            WHERE academic_year_id=%s AND semester_no=%s
              AND class_id=%s AND subject_id=%s LIMIT 1
        """, (academic_year_id, semester_no, class_id, subject_id))
        if cur.fetchone():
            duplicates += 1
            continue
        conflict = _subject_repeat_conflict(cur, subject_id, class_id, academic_year_id, semester_no)
        if conflict:
            code = subject.get('subject_code') if isinstance(subject, dict) else None
            name = subject.get('subject_name') if isinstance(subject, dict) else None
            label = code or name or str(subject_id)
            repeat_conflicts.append(label)
            continue
        cur.execute("""
            INSERT INTO ep_faculty_assignments
              (faculty_id,academic_year_id,semester_no,class_id,subject_id,section)
            VALUES(%s,%s,%s,%s,%s,'A')
        """, (faculty_id, academic_year_id, semester_no, class_id, subject_id))
        _ensure_class_subject_mapping(cur, subject_id, class_id, academic_year_id, semester_no)
        added += 1
    return added, duplicates, repeat_conflicts
def _ensure_class_subject_mapping(cur, subject_id, class_id, academic_year_id, semester_no):
    """Connect the master subject to the selected class/context without
    changing Subject Master fields.  Safe for legacy databases."""
    try:
        cur.execute("SELECT 1 FROM class_subjects WHERE class_id=%s AND subject_id=%s LIMIT 1", (class_id, subject_id))
        if not cur.fetchone():
            cur.execute("INSERT INTO class_subjects(class_id,subject_id) VALUES(%s,%s)", (class_id, subject_id))
    except mysql.connector.Error:
        pass
    try:
        cur.execute("""
            SELECT id FROM ep_subject_semesters
            WHERE subject_id=%s AND class_id=%s AND academic_year_id=%s AND semester_no=%s LIMIT 1
        """, (subject_id, class_id, academic_year_id, semester_no))
        if not cur.fetchone():
            cur.execute("""
                INSERT INTO ep_subject_semesters(subject_id,class_id,academic_year_id,semester_no)
                VALUES(%s,%s,%s,%s)
            """, (subject_id, class_id, academic_year_id, semester_no))
    except mysql.connector.Error as e:
        if getattr(e, "errno", None) != 1062:
            raise
@app.route("/admin/faculty-assignments/bulk-add", methods=["POST"])
def admin_faculty_assignments_bulk_add():
    if not admin_required():
        return redirect(url_for("login"))
    conn = get_db_connection(); cur = conn.cursor(dictionary=True)
    try:
        faculty_id = int(request.form.get("faculty_id"))
        class_id = int(request.form.get("class_id"))
        academic_year_id = int(request.form.get("academic_year_id"))
        semester_no = int(request.form.get("semester_no"))
        subject_ids = list(dict.fromkeys(int(x) for x in request.form.getlist("subject_ids") if str(x).isdigit()))
        if not subject_ids:
            raise ValueError("Please select at least one subject.")
        _validate_assignment_context(cur, faculty_id, class_id, academic_year_id, semester_no)
        added, duplicates, repeat_conflicts = _insert_faculty_subject_assignments(cur, faculty_id, class_id, academic_year_id, semester_no, subject_ids)
        conn.commit()
        if repeat_conflicts:
            flash("Subject repeat blocked for the same student batch: " + ", ".join(repeat_conflicts) + ". A subject can be assigned only once across Semesters 1-6 for that cohort.", "warning")
        if added and duplicates:
            flash(f"Subjects assigned successfully. {duplicates} already assigned and skipped.", "warning")
        elif added and not repeat_conflicts:
            flash("Subjects assigned successfully.", "success")
        elif not added and duplicates and not repeat_conflicts:
            flash("This subject is already assigned for the selected Class, Academic Year and Semester.", "warning")
    except (TypeError, ValueError) as e:
        conn.rollback(); flash(str(e), "danger")
    except mysql.connector.Error as e:
        conn.rollback(); flash(_friendly_db_error(e, "Unable to assign subjects."), "danger")
    finally:
        cur.close(); conn.close()
    return redirect(url_for("admin_faculty_assignments"))
@app.route("/admin/faculty-assignments/bulk-delete", methods=["POST"], endpoint="admin_faculty_assignments_bulk_delete")
def admin_faculty_assignments_bulk_delete():
    if not admin_required():
        return redirect(url_for("login"))
    conn = get_db_connection(); cur = conn.cursor(dictionary=True)
    try:
        ids = list(dict.fromkeys(int(x) for x in request.form.getlist("assignment_ids") if str(x).isdigit()))
        if not ids:
            raise ValueError("Please select at least one assignment to remove.")
        placeholders = ",".join(["%s"] * len(ids))
        cur.execute(f"SELECT id,subject_id,class_id,academic_year_id,semester_no FROM ep_faculty_assignments WHERE id IN ({placeholders})", tuple(ids))
        rows = cur.fetchall()
        if not rows:
            raise ValueError("No valid assignments were selected.")
        cur.execute(f"DELETE FROM ep_faculty_assignments WHERE id IN ({placeholders})", tuple(ids))
        for row in rows:
            cur.execute("""
                SELECT COUNT(*) AS n FROM ep_faculty_assignments
                WHERE subject_id=%s AND class_id=%s AND academic_year_id=%s AND semester_no=%s
            """, (row["subject_id"],row["class_id"],row["academic_year_id"],row["semester_no"]))
            if int(cur.fetchone()["n"] or 0) == 0:
                try:
                    cur.execute("DELETE FROM ep_subject_semesters WHERE subject_id=%s AND class_id=%s AND academic_year_id=%s AND semester_no=%s", (row["subject_id"],row["class_id"],row["academic_year_id"],row["semester_no"]))
                except mysql.connector.Error:
                    pass
        conn.commit()
        flash(f"{len(rows)} subject assignment(s) removed successfully.", "success")
    except (TypeError, ValueError) as e:
        conn.rollback(); flash(str(e), "danger")
    except mysql.connector.Error as e:
        conn.rollback(); flash("Unable to remove assignments.", "danger")
    finally:
        cur.close(); conn.close()
    return redirect(url_for("admin_faculty_assignments"))
@app.route("/admin/faculty-assignments/edit/<int:assignment_id>", methods=["POST"])
def admin_faculty_assignment_edit(assignment_id):
    if not admin_required():
        return redirect(url_for("login"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        faculty_id = int(request.form.get("faculty_id"))
        academic_year_id = int(request.form.get("academic_year_id"))
        semester_no = int(request.form.get("semester_no"))
        class_id = int(request.form.get("class_id"))
        subject_id = int(request.form.get("subject_id"))
        cur.execute("SELECT department_id FROM faculty WHERE id=%s", (faculty_id,))
        f = cur.fetchone()
        cur.execute("SELECT department_id FROM classes WHERE id=%s", (class_id,))
        c = cur.fetchone()
        if not f or not c:
            raise ValueError("Faculty or class not found.")
        cur.execute("SELECT 1 FROM ep_class_semesters WHERE class_id=%s AND academic_year_id=%s AND semester_no=%s",
                    (class_id, academic_year_id, semester_no))
        if not cur.fetchone():
            raise ValueError("The selected class is not mapped to this semester.")
        cur.execute("SELECT 1 FROM subjects WHERE id=%s", (subject_id,))
        if not cur.fetchone():
            raise ValueError("Selected subject does not exist.")
        cur.execute("""
            SELECT id FROM ep_faculty_assignments
            WHERE faculty_id=%s AND academic_year_id=%s AND semester_no=%s
              AND class_id=%s AND subject_id=%s AND id<>%s
        """, (faculty_id, academic_year_id, semester_no, class_id, subject_id, assignment_id))
        if cur.fetchone():
            raise ValueError("This faculty-subject assignment already exists.")
        conflict = _subject_repeat_conflict(cur, subject_id, class_id, academic_year_id, semester_no, exclude_assignment_id=assignment_id)
        if conflict:
            raise ValueError("This subject is already used in another semester for the same student batch. A subject cannot be repeated across Semesters 1-6 for that cohort.")
        cur.execute("""
            UPDATE ep_faculty_assignments
            SET faculty_id=%s, academic_year_id=%s, semester_no=%s,
                class_id=%s, subject_id=%s
            WHERE id=%s
        """, (faculty_id, academic_year_id, semester_no, class_id, subject_id, assignment_id))
        try:
            cur.execute("SELECT 1 FROM class_subjects WHERE class_id=%s AND subject_id=%s LIMIT 1", (class_id, subject_id))
            if not cur.fetchone():
                cur.execute("INSERT INTO class_subjects(class_id,subject_id) VALUES(%s,%s)", (class_id, subject_id))
        except mysql.connector.Error:
            pass
        try:
            cur.execute("""
                INSERT INTO ep_subject_semesters(subject_id,class_id,academic_year_id,semester_no)
                VALUES(%s,%s,%s,%s)
            """, (subject_id, class_id, academic_year_id, semester_no))
        except mysql.connector.Error as e:
            if getattr(e, 'errno', None) != 1062:
                raise
        conn.commit()
        flash("Faculty subject assignment updated successfully.", "success")
    except (TypeError, ValueError) as e:
        conn.rollback()
        flash(str(e), "danger")
    except mysql.connector.Error as e:
        conn.rollback()
        flash(f"Unable to update assignment: {e}", "danger")
    finally:
        cur.close()
        conn.close()
    return redirect(url_for("admin_faculty_assignments"))
@app.route("/admin/faculty-assignments/delete/<int:assignment_id>", methods=["POST"])
def admin_faculty_assignment_delete(assignment_id):
    if not admin_required():
        return redirect(url_for("login"))
    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM ep_faculty_assignments WHERE id=%s", (assignment_id,))
        conn.commit()
        flash("Assignment deleted successfully.", "success")
    except mysql.connector.Error as e:
        conn.rollback()
        flash(f"Unable to delete assignment: {e}", "danger")
    finally:
        cur.close()
        conn.close()
    return redirect(url_for("admin_faculty_assignments"))
def admin_required():
    return "user_id" in session and session.get("role") == "admin"
def student_logged_in():
    """Boolean helper for student-session checks.
    Kept separate from the @student_required decorator below so route
    authorization cannot be accidentally overwritten.
    """
    return "user_id" in session and session.get("role") == "student"
@app.route("/")
def index():
    return render_template("index.html")
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if not username or not password:
            flash("Please enter username and password.", "danger")
            return render_template("login.html")
        try:
            db = get_db_connection()
            cursor = db.cursor(dictionary=True)
            cursor.execute(
                "SELECT * FROM users WHERE LOWER(username) = LOWER(%s) LIMIT 1",
                (username,)
            )
            user = cursor.fetchone()
        except Exception:
            flash("Database connection failed. Check DB_HOST, DB_NAME, DB_USER, and DB_PASSWORD.", "danger")
            return render_template("login.html"), 503
        finally:
            if "cursor" in locals():
                cursor.close()
            if "db" in locals():
                db.close()
        if user:
            stored_password = user["password"]
            password_valid = False
            try:
                password_valid = check_password_hash(
                    stored_password,
                    password
                )
            except Exception:
                password_valid = False
            if (
                username == "admin"
                and stored_password == "admin123"
                and password == "admin123"
            ):
                password_valid = True
            if password_valid:
                if stored_password == "admin123":
                    new_password = generate_password_hash(password)
                    db = get_db_connection()
                    cursor = db.cursor()
                    cursor.execute(
                        """
                        UPDATE users
                        SET password = %s
                        WHERE id = %s
                        """,
                        (new_password, user["id"])
                    )
                    db.commit()
                    cursor.close()
                    db.close()
                session.clear()
                session["user_id"] = user["id"]
                session["username"] = user["username"]
                session["role"] = user["role"]
                if user["role"] == "admin":
                    return redirect(url_for("admin_dashboard"))
                if user["role"] == "faculty":
                    return redirect(url_for("faculty_dashboard"))
                if user["role"] == "student":
                    return redirect(url_for("student_dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html")
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))
@app.route("/admin/dashboard")
def admin_dashboard():
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT COUNT(*) AS total FROM departments")
    departments_count = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(*) AS total FROM classes")
    classes_count = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(*) AS total FROM faculty")
    faculty_count = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(*) AS total FROM students")
    students_count = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(*) AS total FROM subjects")
    subjects_count = cursor.fetchone()["total"]
    cursor.execute("""
        SELECT visibility
        FROM result_settings
        WHERE id = 1
    """)
    setting = cursor.fetchone()
    result_visibility = (
        bool(setting["visibility"])
        if setting
        else False
    )
    cursor.close()
    db.close()
    return render_template(
        "admin/dashboard.html",
        departments_count=departments_count,
        classes_count=classes_count,
        faculty_count=faculty_count,
        students_count=students_count,
        subjects_count=subjects_count,
        result_visibility=result_visibility
    )
@app.route("/admin/profile", methods=["GET", "POST"])
def admin_profile():
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection()
    cursor = db.cursor(dictionary=True)
    if request.method == "POST":
        username = request.form.get(
            "username",
            ""
        ).strip()
        if not username:
            flash(
                "Username is required.",
                "danger"
            )
        else:
            try:
                cursor.execute("""
                    UPDATE users
                    SET username = %s
                    WHERE id = %s
                """, (
                    username,
                    session["user_id"]
                ))
                db.commit()
                session["username"] = username
                flash(
                    "Profile updated successfully.",
                    "success"
                )
            except mysql.connector.Error:
                db.rollback()
                flash(
                    "Username already exists.",
                    "danger"
                )
    cursor.execute("""
        SELECT id, username, role
        FROM users
        WHERE id = %s
    """, (
        session["user_id"],
    ))
    admin = cursor.fetchone()
    cursor.close()
    db.close()
    return render_template(
        "admin/profile.html",
        admin=admin
    )
@app.route("/admin/departments")
def departments():
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection(); cursor = db.cursor(dictionary=True)
    try:
        cursor.execute("""
            SELECT d.id, d.department_name, d.department_code, d.hod_name,
                   b.course_type, b.start_year, b.end_year, b.course_duration,
                   COUNT(DISTINCT c.id) AS class_count
            FROM departments d
            LEFT JOIN classes c ON d.id = c.department_id
            LEFT JOIN ep_department_batches b ON b.department_id = d.id
            GROUP BY d.id, d.department_name, d.department_code, d.hod_name,
                     b.course_type, b.start_year, b.end_year, b.course_duration
            ORDER BY d.id DESC
        """)
        rows = cursor.fetchall()
        return render_template("admin/departments.html", departments=rows)
    finally:
        cursor.close(); db.close()
@app.route("/admin/departments/add", methods=["GET", "POST"])
def add_department():
    if not admin_required():
        return redirect(url_for("login"))
    if request.method == "GET":
        return render_template("admin/department_form.html", department=None)
    department_name = request.form.get("department_name", "").strip()
    department_code = request.form.get("department_code", "").strip().upper()
    start_year = request.form.get("start_year", "").strip()
    hod_name = request.form.get("hod_name", "").strip()
    if not department_name or not department_code or not start_year or not hod_name:
        flash("Department/Course Name, Department Code, Start Year and HOD Name are required.", "danger")
        return render_template("admin/department_form.html", department={
            "department_name": department_name, "department_code": department_code,
            "start_year": start_year, "hod_name": hod_name
        })
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9_-]{1,19}", department_code):
        flash("Department Code must be 2-20 characters using only letters, numbers, underscore or hyphen.", "danger")
        return render_template("admin/department_form.html", department={
            "department_name": department_name, "department_code": department_code,
            "start_year": start_year, "hod_name": hod_name
        })
    course_type = _infer_course_type(department_name)
    db = get_db_connection(); cursor = db.cursor()
    try:
        cursor.execute("SELECT id FROM departments WHERE UPPER(department_code)=UPPER(%s) LIMIT 1", (department_code,))
        if cursor.fetchone():
            raise ValueError("Department Code already exists. Please enter a unique code.")
        cursor.execute("INSERT INTO departments(department_name,department_code,hod_name) VALUES(%s,%s,%s)",
                       (department_name, department_code, hod_name))
        department_id = cursor.lastrowid
        start, end, academic_batch = save_department_batch(
            cursor, department_id, course_type, start_year
        )
        years = 2 if course_type == "PG" else 3
        for class_name in [f"{n}st Year" if n == 1 else f"{n}nd Year" if n == 2 else f"{n}rd Year" for n in range(1, years + 1)]:
            cursor.execute("INSERT INTO classes(department_id,class_name) VALUES(%s,%s)", (department_id, class_name))
        sync_class_semesters(cursor)
        db.commit()
        flash(f"{department_name} added successfully with {years} valid year classes.", "success")
        return redirect(url_for("departments"))
    except (mysql.connector.Error, ValueError) as e:
        db.rollback()
        flash(f"Unable to add department: {e}", "danger")
        return render_template("admin/department_form.html", department={
            "department_name": department_name, "department_code": department_code,
            "start_year": start_year, "hod_name": hod_name,
            "course_type": course_type
        })
    finally:
        cursor.close(); db.close()
@app.route("/admin/departments/edit/<int:department_id>", methods=["GET", "POST"])
def edit_department(department_id):
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection(); cursor = db.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM departments WHERE id=%s", (department_id,))
        department = cursor.fetchone()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("departments"))
        cursor.execute("""
            SELECT course_type,start_year,end_year,course_duration
            FROM ep_department_batches
            WHERE department_id=%s LIMIT 1
        """, (department_id,))
        batch = cursor.fetchone() or {}
        department["course_type"] = batch.get("course_type", _infer_course_type(department.get("department_name", "")))
        department["start_year"] = batch.get("start_year", "")
        department["end_year"] = batch.get("end_year", "")
        department["academic_batch"] = batch.get("course_duration", "")
        if request.method == "POST":
            name = request.form.get("department_name", "").strip()
            department_code = request.form.get("department_code", "").strip().upper()
            start_year = request.form.get("start_year", "").strip()
            hod_name = request.form.get("hod_name", "").strip()
            if not name or not department_code or not start_year or not hod_name:
                raise ValueError("Department/Course Name, Department Code, Start Year and HOD Name are required.")
            if not re.fullmatch(r"[A-Z0-9][A-Z0-9_-]{1,19}", department_code):
                raise ValueError("Department Code must be 2-20 characters using only letters, numbers, underscore or hyphen.")
            cursor.execute("SELECT id FROM departments WHERE UPPER(department_code)=UPPER(%s) AND id<>%s LIMIT 1", (department_code, department_id))
            if cursor.fetchone():
                raise ValueError("Department Code already exists. Please enter a unique code.")
            course_type = _infer_course_type(name)
            cursor.execute(
                "UPDATE departments SET department_name=%s, department_code=%s, hod_name=%s WHERE id=%s",
                (name, department_code, hod_name, department_id)
            )
            start, end, academic_batch = save_department_batch(
                cursor, department_id, course_type, start_year
            )
            max_years = 2 if course_type == "PG" else 3
            cursor.execute("SELECT id,class_name FROM classes WHERE department_id=%s", (department_id,))
            existing = cursor.fetchall()
            existing_years = {_infer_year_from_class(r["class_name"]) for r in existing}
            for year_no in range(1, max_years + 1):
                if year_no not in existing_years:
                    suffix = "st" if year_no == 1 else "nd" if year_no == 2 else "rd"
                    cursor.execute("INSERT INTO classes(department_id,class_name) VALUES(%s,%s)", (department_id, f"{year_no}{suffix} Year"))
            sync_class_semesters(cursor)
            db.commit()
            flash("Department updated successfully.", "success")
            return redirect(url_for("departments"))
        return render_template("admin/department_form.html", department=department)
    except (mysql.connector.Error, ValueError) as e:
        db.rollback()
        department = department or {}
        if request.method == "POST":
            department["department_name"] = request.form.get("department_name", "")
            department["department_code"] = request.form.get("department_code", "").upper()
            department["start_year"] = request.form.get("start_year", "")
            department["hod_name"] = request.form.get("hod_name", "")
            department["course_type"] = _infer_course_type(department["department_name"])
        flash(f"Unable to update department: {e}", "danger")
        return render_template("admin/department_form.html", department=department)
    finally:
        cursor.close(); db.close()
@app.route("/admin/departments/delete/<int:department_id>", methods=["POST"])
def delete_department(department_id):
    if not admin_required(): return redirect(url_for("login"))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute("SELECT id FROM students WHERE department_id=%s LIMIT 1",(department_id,))
        if cur.fetchone(): raise ValueError("This department has students. Move/delete the students first before deleting the department.")
        cur.execute("SELECT id FROM faculty WHERE department_id=%s LIMIT 1",(department_id,))
        if cur.fetchone(): raise ValueError("This department has faculty. Move/delete the faculty first before deleting the department.")
        cur.execute("SELECT id FROM classes WHERE department_id=%s",(department_id,)); class_ids=[r['id'] for r in cur.fetchall()]
        if class_ids:
            ph=','.join(['%s']*len(class_ids)); cur.execute(f"SELECT id FROM subjects WHERE class_id IN ({ph})",tuple(class_ids)); subject_ids=[r['id'] for r in cur.fetchall()]
            if subject_ids:
                sph=','.join(['%s']*len(subject_ids)); cur.execute(f"DELETE FROM ep_faculty_assignments WHERE subject_id IN ({sph})",tuple(subject_ids)); cur.execute(f"DELETE FROM ep_subject_semesters WHERE subject_id IN ({sph})",tuple(subject_ids)); cur.execute(f"DELETE FROM class_subjects WHERE subject_id IN ({sph})",tuple(subject_ids)); cur.execute(f"DELETE FROM subjects WHERE id IN ({sph})",tuple(subject_ids))
            cur.execute(f"DELETE FROM ep_class_semesters WHERE class_id IN ({ph})",tuple(class_ids)); cur.execute(f"DELETE FROM classes WHERE id IN ({ph})",tuple(class_ids))
        cur.execute("DELETE FROM ep_department_batches WHERE department_id=%s",(department_id,)); cur.execute("DELETE FROM departments WHERE id=%s",(department_id,)); db.commit(); flash("Department and its generated academic structure were deleted successfully.","success")
    except (ValueError,mysql.connector.Error) as e:
        db.rollback(); flash(f"Unable to delete department: {e}","danger")
    finally: cur.close(); db.close()
    return redirect(url_for("departments"))
@app.route("/admin/classes")
def classes():
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SHOW COLUMNS FROM classes")
    class_columns = {str(row.get('Field', '')).lower() for row in cursor.fetchall()}
    faculty_select = 'f.faculty_name' if 'faculty_id' in class_columns else 'NULL AS faculty_name'
    faculty_join = 'LEFT JOIN faculty f ON c.faculty_id = f.id' if 'faculty_id' in class_columns else ''
    cursor.execute(f"""
        SELECT
            c.id,
            c.class_name,
            c.department_id,
            d.department_name,
            d.department_code,
            {faculty_select},
            COALESCE(sc.subject_count, 0) AS subject_count,
            COALESCE(sm.semester_count, 0) AS semester_count
        FROM classes c
        INNER JOIN departments d ON c.department_id = d.id
        {faculty_join}
        LEFT JOIN (SELECT class_id, COUNT(*) AS subject_count FROM subjects GROUP BY class_id) sc ON sc.class_id=c.id
        LEFT JOIN (SELECT class_id, COUNT(*) AS semester_count FROM ep_class_semesters GROUP BY class_id) sm ON sm.class_id=c.id
        ORDER BY
            d.department_name,
            c.id
    """)
    class_list = cursor.fetchall()
    cursor.close()
    db.close()
    return render_template(
        "admin/classes.html",
        classes=class_list
    )
@app.route('/admin/faculty')
def faculty():
    if not admin_required():
        return redirect(url_for('login'))
    db = get_db_connection()
    cur = db.cursor(dictionary=True)
    try:
        cur.execute("SELECT f.user_id, f.faculty_name, u.username FROM faculty f INNER JOIN users u ON u.id=f.user_id WHERE u.role='faculty'")
        for account in cur.fetchall():
            desired = _faculty_username(account.get('faculty_name'))
            current = str(account.get('username') or '').strip()
            if not desired or desired == current:
                continue
            cur.execute('SELECT id FROM users WHERE LOWER(username)=LOWER(%s) AND id<>%s LIMIT 1', (desired, account['user_id']))
            if not cur.fetchone():
                cur.execute('UPDATE users SET username=%s WHERE id=%s', (desired, account['user_id']))
        db.commit()
        cur.execute('''SELECT f.id, f.faculty_name, f.email, f.phone, f.department_id,
                              d.department_name
                       FROM faculty f
                       LEFT JOIN departments d ON d.id=f.department_id
                       ORDER BY f.id DESC''')
        rows = cur.fetchall()
        return render_template('admin/faculties.html', faculty=rows)
    finally:
        cur.close()
        db.close()
@app.route('/admin/faculty/view/<int:faculty_id>')
def view_faculty(faculty_id):
    if not admin_required():
        return redirect(url_for('login'))
    db = get_db_connection()
    cur = db.cursor(dictionary=True)
    try:
        cur.execute('''SELECT f.id, f.faculty_name, f.email, f.phone, f.department_id,
                              d.department_name, fd.gender, fd.dob, fd.address
                       FROM faculty f
                       LEFT JOIN departments d ON d.id=f.department_id
                       LEFT JOIN ep_faculty_details fd ON fd.faculty_id=f.id
                       WHERE f.id=%s''', (faculty_id,))
        row = cur.fetchone()
        if not row:
            flash('Faculty not found.', 'danger')
            return redirect(url_for('faculty'))
        return render_template('admin/faculty_profile.html', faculty=row)
    finally:
        cur.close()
        db.close()
def _faculty_departments(cur):
    cur.execute('SELECT id, department_name, department_code FROM departments ORDER BY department_name')
    return cur.fetchall()
def _faculty_username(name):
    """Create a faculty username from the name without spaces or punctuation.
    Example: "Dr. Suresh Kumar" -> "drsureshkumar".
    The faculty DOB/password rule is unchanged.
    """
    return re.sub(r'[^a-z0-9]', '', str(name or '').strip().lower())
def _dob_password(dob_value):
    """Return the faculty login password in DD/MM/YYYY format."""
    value = str(dob_value or '').strip()
    if not value:
        return ''
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d', '%d-%m-%Y'):
        try:
            return datetime.strptime(value, fmt).strftime('%d/%m/%Y')
        except ValueError:
            continue
    return value
def _dob_db_value(dob_value):
    """Normalize a faculty DOB to MySQL DATE format YYYY-MM-DD."""
    value = str(dob_value or '').strip()
    if not value:
        return None
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d', '%d-%m-%Y'):
        try:
            return datetime.strptime(value, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None
def _normalize_excel_header(value):
    return re.sub(r'[^a-z0-9]+', '_', str(value or '').strip().lower()).strip('_')
def _department_key(value):
    """Normalize department names for reliable bulk-upload matching."""
    words = re.findall(r'[a-z0-9]+', str(value or '').lower())
    return ' '.join(word[:-1] if len(word) > 3 and word.endswith('s') else word for word in words)
def _faculty_excel_aliases():
    return {
        'faculty_name': {'faculty_name', 'name', 'faculty'},
        'email': {'email', 'faculty_email'},
        'phone': {'phone', 'mobile', 'mobile_number', 'phone_number'},
        'department': {'department', 'department_name', 'course'},
        'dob': {'dob', 'date_of_birth', 'birth_date'},
    }
def _read_faculty_excel(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    records = []
    if ext == '.csv':
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as fh:
            reader = csv.DictReader(fh)
            if not reader.fieldnames:
                raise ValueError('The CSV file has no header row.')
            headers = {_normalize_excel_header(h): h for h in reader.fieldnames}
            aliases = _faculty_excel_aliases()
            mapped = {}
            for field, names in aliases.items():
                hit = next((headers[n] for n in names if n in headers), None)
                if hit:
                    mapped[field] = hit
            missing = [x.replace('_', ' ').title() for x in ('faculty_name','email','phone','department') if x not in mapped]
            if missing:
                raise ValueError('Missing Excel columns: ' + ', '.join(missing))
            for row_no, row in enumerate(reader, start=2):
                values = {field: str(row.get(col, '') or '').strip() for field, col in mapped.items()}
                if not any(values.values()):
                    continue
                values['_row_no'] = row_no
                records.append(values)
    elif ext == '.xlsx':
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_cells:
            wb.close()
            raise ValueError('The Excel file has no header row.')
        headers = {_normalize_excel_header(h): i for i, h in enumerate(header_cells) if h is not None}
        aliases = _faculty_excel_aliases()
        mapped = {}
        for field, names in aliases.items():
            hit = next((headers[n] for n in names if n in headers), None)
            if hit is not None:
                mapped[field] = hit
        missing = [x.replace('_', ' ').title() for x in ('faculty_name','email','phone','department') if x not in mapped]
        if missing:
            wb.close()
            raise ValueError('Missing Excel columns: ' + ', '.join(missing))
        for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            values = {}
            for field, idx in mapped.items():
                value = row[idx] if idx < len(row) else ''
                if field == 'dob' and value is not None and hasattr(value, 'strftime'):
                    value = value.strftime('%Y-%m-%d')
                values[field] = str(value or '').strip()
            if not any(values.values()):
                continue
            values['_row_no'] = row_no
            records.append(values)
        wb.close()
    else:
        raise ValueError('Only .xlsx and .csv files are supported.')
    return records
def _validate_faculty_bulk(conn, rows):
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute('SELECT id, department_name, department_code FROM departments')
        departments = cur.fetchall()
        by_name = {_department_key(d['department_name']): d for d in departments}
        by_code = {str(d['department_code']).strip().lower(): d for d in departments if d.get('department_code')}
        cur.execute('SELECT LOWER(email) AS email, phone FROM faculty')
        existing_faculty = cur.fetchall()
        db_emails = {str(x['email']).strip().lower() for x in existing_faculty if x.get('email')}
        db_phones = {str(x['phone']).strip() for x in existing_faculty if x.get('phone')}
        cur.execute('SELECT LOWER(username) AS username FROM users')
        db_usernames = {str(x['username']).strip().lower() for x in cur.fetchall() if x.get('username')}
        seen_emails, seen_phones, seen_usernames = set(), set(), set()
        out = []
        email_re = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')
        for index, row in enumerate(rows, start=1):
            errors = []
            def clean(value):
                return str(value or '').strip()
            name = clean(row.get('faculty_name'))
            email = clean(row.get('email'))
            phone = clean(row.get('phone'))
            dept_value = clean(row.get('department'))
            dob_raw = clean(row.get('dob'))
            username = _faculty_username(name)
            password = _dob_password(dob_raw)
            if not name: errors.append('Faculty Name is required')
            if not email: errors.append('Email is required')
            elif not email_re.match(email): errors.append('Invalid email format')
            if not phone: errors.append('Phone is required')
            elif not phone.isdigit() or len(phone) != 10: errors.append('Phone must contain exactly 10 digits')
            if not dept_value: errors.append('Department is required')
            if not dob_raw: errors.append('Date of Birth is required')
            elif not password: errors.append('Date of Birth must be in DD/MM/YYYY or YYYY-MM-DD format')
            dept = by_name.get(_department_key(dept_value)) or by_code.get(dept_value.lower())
            if dept_value and not dept:
                errors.append(f'Department "{dept_value}" does not exist')
            email_key = email.lower()
            phone_key = phone
            username_key = username.lower()
            if username and username_key in db_usernames: errors.append('Faculty name/username already exists')
            if username and username_key in seen_usernames: errors.append('Duplicate faculty name/username in this file')
            if email and email_key in db_emails: errors.append('Email already exists')
            if email and email_key in seen_emails: errors.append('Duplicate email in this file')
            if phone and phone_key in db_phones: errors.append('Phone already exists')
            if phone and phone_key in seen_phones: errors.append('Duplicate phone in this file')
            valid = not errors
            if valid:
                if email: seen_emails.add(email_key)
                if phone: seen_phones.add(phone_key)
                if username: seen_usernames.add(username_key)
            out.append({
                'sno': index,
                'row_no': row.get('_row_no', index + 1),
                'faculty_name': name,
                'email': email,
                'phone': phone,
                'department': dept['department_name'] if dept else dept_value,
                'department_id': dept['id'] if dept else None,
                'username': username,
                'password': password,
                'dob': _dob_db_value(dob_raw),
                'dob_display': dob_raw or None,
                'valid': valid,
                'status': 'Valid' if valid else 'Invalid',
                'errors': errors
            })
        return out
    finally:
        cur.close()
def _bulk_upload_dir():
    path = os.path.join(tempfile.gettempdir(), 'e_progress_card_faculty_bulk')
    os.makedirs(path, exist_ok=True)
    return path
@app.route('/admin/faculties/template')
def download_faculty_template():
    if not admin_required(): return redirect(url_for('login'))
    wb = Workbook(); ws = wb.active; ws.title = 'Faculty Data'
    headers = ['Faculty Name','Email','Phone','Department','DOB']
    ws.append(headers)
    ws.append(['Faculty One','faculty1@example.com','9876543210','BCA','12/04/1989'])
    ws.append(['Faculty Two','faculty2@example.com','9876543211','MCA','25/07/1990'])
    for cell in ws[1]: cell.font = cell.font.copy(bold=True)
    ws.freeze_panes = 'A2'
    for i, width in enumerate([24,30,16,24,16], start=1): ws.column_dimensions[chr(64+i)].width = width
    output = BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, as_attachment=True, download_name='faculty_upload_template.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/admin/faculties/add', methods=['GET', 'POST'])
def add_faculty():
    if not admin_required():
        return redirect(url_for('login'))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        departments = _faculty_departments(cur)
        if request.method == 'POST':
            name = request.form.get('faculty_name', '').strip()
            email = request.form.get('email', '').strip()
            phone = request.form.get('phone', '').strip()
            dept = request.form.get('department_id', '').strip()
            dob = request.form.get('dob', '').strip()
            username = _faculty_username(name)
            password = _dob_password(dob)
            if not all([name, email, phone, dept, dob]):
                raise ValueError('Faculty Name, Email, Phone, Department and Date of Birth are required.')
            if not password:
                raise ValueError('Enter a valid Date of Birth.')
            if not re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', email):
                raise ValueError('Enter a valid email address.')
            if not phone.isdigit() or len(phone) != 10:
                raise ValueError('Phone number must contain exactly 10 digits.')
            cur.execute('SELECT id FROM users WHERE LOWER(username)=LOWER(%s)', (username,))
            if cur.fetchone():
                raise ValueError('A faculty with the same generated username already exists.')
            cur.execute('SELECT id FROM faculty WHERE LOWER(email)=LOWER(%s)', (email,))
            if cur.fetchone():
                raise ValueError('Faculty email already exists.')
            cur.execute('SELECT id FROM faculty WHERE phone=%s', (phone,))
            if cur.fetchone():
                raise ValueError('Faculty phone number already exists.')
            cur.execute('SELECT id FROM departments WHERE id=%s', (dept,))
            if not cur.fetchone():
                raise ValueError('Selected department does not exist.')
            cur.execute(
                "INSERT INTO users(username,password,role) VALUES(%s,%s,'faculty')",
                (username, generate_password_hash(password))
            )
            uid = cur.lastrowid
            cur.execute(
                "SELECT IS_NULLABLE FROM information_schema.COLUMNS "
                "WHERE TABLE_SCHEMA=%s AND TABLE_NAME='faculty' AND COLUMN_NAME='class_id' LIMIT 1",
                (config.DB_NAME,)
            )
            class_meta = cur.fetchone()
            if class_meta and str(class_meta.get('IS_NULLABLE', 'YES')).upper() == 'NO':
                cur.execute('SELECT id FROM classes WHERE department_id=%s ORDER BY id LIMIT 1', (dept,))
                legacy_class = cur.fetchone()
                if not legacy_class:
                    raise ValueError('Create the department classes before adding faculty.')
                cur.execute(
                    'INSERT INTO faculty(user_id,faculty_name,email,phone,department_id,class_id) '
                    'VALUES(%s,%s,%s,%s,%s,%s)',
                    (uid, name, email, phone, dept, legacy_class['id'])
                )
            else:
                try:
                    cur.execute(
                        'INSERT INTO faculty(user_id,faculty_name,email,phone,department_id,class_id) '
                        'VALUES(%s,%s,%s,%s,%s,NULL)',
                        (uid, name, email, phone, dept)
                    )
                except mysql.connector.Error:
                    cur.execute(
                        'INSERT INTO faculty(user_id,faculty_name,email,phone,department_id) '
                        'VALUES(%s,%s,%s,%s,%s)',
                        (uid, name, email, phone, dept)
                    )
            fid = cur.lastrowid
            cur.execute(
                '''INSERT INTO ep_faculty_details(faculty_id,dob)
                   VALUES(%s,%s)
                   ON DUPLICATE KEY UPDATE dob=VALUES(dob)''',
                (fid, _dob_db_value(dob))
            )
            conn.commit()
            flash('Faculty added successfully. Username is the faculty name and password is the DOB in DD/MM/YYYY format.', 'success')
            return redirect(url_for('faculty'))
        return render_template('admin/faculty_form.html', faculty=None, departments=departments, classes=[])
    except (ValueError, mysql.connector.Error) as e:
        conn.rollback()
        flash(f'Unable to add faculty: {e}', 'danger')
        return render_template('admin/faculty_form.html', faculty=None, departments=departments, classes=[])
    finally:
        cur.close()
        conn.close()
@app.route('/admin/faculties/bulk-upload', methods=['POST'])
def faculty_bulk_upload():
    if not admin_required(): return redirect(url_for('login'))
    upload=request.files.get('faculty_excel')
    if not upload or not upload.filename:
        flash('Please select an Excel or CSV file.','danger'); return redirect(url_for('add_faculty'))
    ext=os.path.splitext(upload.filename)[1].lower()
    if ext not in {'.xlsx','.csv'}:
        flash('Only .xlsx and .csv files are supported.','danger'); return redirect(url_for('add_faculty'))
    token=uuid.uuid4().hex; path=os.path.join(_bulk_upload_dir(),token+ext)
    try:
        upload.save(path); rows=_read_faculty_excel(path)
        if not rows: raise ValueError('The uploaded file contains no faculty records.')
        conn=get_db_connection()
        try: preview=_validate_faculty_bulk(conn,rows)
        finally: conn.close()
        preview_path=os.path.join(_bulk_upload_dir(),token+'.json')
        with open(preview_path,'w',encoding='utf-8') as fh: json.dump(preview,fh,ensure_ascii=False)
        session['faculty_bulk_token']=token
        return render_template('admin/faculty_bulk_preview.html',rows=preview,token=token)
    except Exception as e:
        try:
            if os.path.exists(path): os.remove(path)
        except OSError: pass
        flash(f'Unable to read faculty file: {e}','danger'); return redirect(url_for('add_faculty'))
@app.route('/admin/faculties/bulk-confirm', methods=['POST'])
def faculty_bulk_confirm():
    if not admin_required(): return redirect(url_for('login'))
    token=request.form.get('token','').strip()
    if not token or token != session.get('faculty_bulk_token') or not re.fullmatch(r'[a-f0-9]{32}',token):
        flash('Bulk upload session expired. Please upload the Excel file again.','danger'); return redirect(url_for('add_faculty'))
    preview_path=os.path.join(_bulk_upload_dir(),token+'.json')
    try:
        with open(preview_path,'r',encoding='utf-8') as fh: preview=json.load(fh)
    except Exception:
        flash('Bulk upload preview expired. Please upload the Excel file again.','danger'); return redirect(url_for('add_faculty'))
    conn=get_db_connection(); success=0; failed=[]
    try:
        rows=[]
        for r in preview:
            rows.append({'_row_no':r.get('row_no'),'faculty_name':r.get('faculty_name',''),'email':r.get('email',''),'phone':r.get('phone',''),
                         'department':r.get('department',''),'dob':r.get('dob')})
        fresh=_validate_faculty_bulk(conn,rows)
        for row in fresh:
            if not row['valid']:
                failed.append({'name':row['faculty_name'] or f"Row {row['row_no']}",'errors':row['errors']}); continue
            cur=conn.cursor(dictionary=True)
            try:
                cur.execute("INSERT INTO users(username,password,role) VALUES(%s,%s,'faculty')",(row['username'],generate_password_hash(row['password']))); uid=cur.lastrowid
                cur.execute("SELECT IS_NULLABLE FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=%s AND TABLE_NAME='faculty' AND COLUMN_NAME='class_id' LIMIT 1",(config.DB_NAME,)); class_meta=cur.fetchone()
                if class_meta and str(class_meta.get('IS_NULLABLE','YES')).upper()=='NO':
                    cur.execute('SELECT id FROM classes WHERE department_id=%s ORDER BY id LIMIT 1',(row['department_id'],)); legacy_class=cur.fetchone()
                    if not legacy_class: raise ValueError('Create department classes before adding faculty.')
                    cur.execute('INSERT INTO faculty(user_id,faculty_name,email,phone,department_id,class_id) VALUES(%s,%s,%s,%s,%s,%s)',(uid,row['faculty_name'],row['email'],row['phone'],row['department_id'],legacy_class['id']))
                else:
                    try: cur.execute('INSERT INTO faculty(user_id,faculty_name,email,phone,department_id,class_id) VALUES(%s,%s,%s,%s,%s,NULL)',(uid,row['faculty_name'],row['email'],row['phone'],row['department_id']))
                    except mysql.connector.Error: cur.execute('INSERT INTO faculty(user_id,faculty_name,email,phone,department_id) VALUES(%s,%s,%s,%s,%s)',(uid,row['faculty_name'],row['email'],row['phone'],row['department_id']))
                fid=cur.lastrowid
                cur.execute('''INSERT INTO ep_faculty_details(faculty_id,dob) VALUES(%s,%s)
                                 ON DUPLICATE KEY UPDATE dob=VALUES(dob)''',
                            (fid,row.get('dob') or None))
                conn.commit(); success+=1
            except (ValueError,mysql.connector.Error) as e:
                conn.rollback(); failed.append({'name':row['faculty_name'] or f"Row {row['row_no']}",'errors':[str(e)]})
            finally: cur.close()
        flash(f'Faculty added successfully: {success} added, {len(failed)} failed.','success' if not failed else 'warning')
    except Exception as e:
        conn.rollback(); flash(f'Bulk faculty import failed: {e}','danger')
    finally:
        conn.close()
        for suffix in ('.json','.xlsx','.csv'):
            path=os.path.join(_bulk_upload_dir(),token+suffix)
            try:
                if os.path.exists(path): os.remove(path)
            except OSError: pass
        session.pop('faculty_bulk_token',None)
    if failed: session['faculty_bulk_failures']=failed[:200]
    else: session.pop('faculty_bulk_failures',None)
    return redirect(url_for('faculty'))
@app.route('/admin/faculty/edit/<int:faculty_id>', methods=['GET', 'POST'])
def edit_faculty(faculty_id):
    if not admin_required():
        return redirect(url_for('login'))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(
            '''SELECT f.*,d.department_name,fd.dob,u.username
               FROM faculty f
               LEFT JOIN departments d ON d.id=f.department_id
               LEFT JOIN ep_faculty_details fd ON fd.faculty_id=f.id
               LEFT JOIN users u ON u.id=f.user_id
               WHERE f.id=%s''',
            (faculty_id,)
        )
        faculty_data = cur.fetchone()
        if not faculty_data:
            flash('Faculty not found.', 'danger')
            return redirect(url_for('faculty'))
        departments = _faculty_departments(cur)
        if request.method == 'POST':
            name = request.form.get('faculty_name', '').strip()
            email = request.form.get('email', '').strip()
            phone = request.form.get('phone', '').strip()
            dept = request.form.get('department_id', '').strip()
            dob = request.form.get('dob', '').strip()
            username = _faculty_username(name)
            password = _dob_password(dob)
            if not all([name, email, phone, dept, dob]):
                raise ValueError('Faculty Name, Email, Phone, Department and Date of Birth are required.')
            if not password:
                raise ValueError('Enter a valid Date of Birth.')
            if not re.match(r'^[^\s@]+@[^\s@]+\.[^\s@]+$', email):
                raise ValueError('Enter a valid email address.')
            if not phone.isdigit() or len(phone) != 10:
                raise ValueError('Phone number must contain exactly 10 digits.')
            cur.execute(
                'SELECT id FROM users WHERE LOWER(username)=LOWER(%s) AND id<>%s',
                (username, faculty_data['user_id'])
            )
            if cur.fetchone():
                raise ValueError('A faculty with the same generated username already exists.')
            cur.execute(
                'SELECT id FROM faculty WHERE LOWER(email)=LOWER(%s) AND id<>%s',
                (email, faculty_id)
            )
            if cur.fetchone():
                raise ValueError('Faculty email already exists.')
            cur.execute(
                'SELECT id FROM faculty WHERE phone=%s AND id<>%s',
                (phone, faculty_id)
            )
            if cur.fetchone():
                raise ValueError('Faculty phone number already exists.')
            cur.execute('SELECT id FROM departments WHERE id=%s', (dept,))
            if not cur.fetchone():
                raise ValueError('Selected department does not exist.')
            cur.execute(
                'UPDATE faculty SET faculty_name=%s,email=%s,phone=%s,department_id=%s WHERE id=%s',
                (name, email, phone, dept, faculty_id)
            )
            cur.execute(
                'UPDATE users SET username=%s,password=%s WHERE id=%s',
                (username, generate_password_hash(password), faculty_data['user_id'])
            )
            cur.execute(
                '''INSERT INTO ep_faculty_details(faculty_id,dob)
                   VALUES(%s,%s)
                   ON DUPLICATE KEY UPDATE dob=VALUES(dob)''',
                (faculty_id, _dob_db_value(dob))
            )
            conn.commit()
            flash('Faculty updated successfully. Username is the faculty name and password is the DOB in DD/MM/YYYY format.', 'success')
            return redirect(url_for('faculty'))
        return render_template(
            'admin/faculty_form.html',
            faculty=faculty_data,
            departments=departments,
            classes=[],
            assignments=[],
            subjects=[],
            years=[],
            class_semester_rows=[],
            edit_mode=True
        )
    except (ValueError, mysql.connector.Error) as e:
        conn.rollback()
        flash(f'Unable to update faculty: {e}', 'danger')
        return redirect(url_for('edit_faculty', faculty_id=faculty_id))
    finally:
        cur.close()
        conn.close()
@app.route('/admin/faculty/delete/<int:faculty_id>', methods=['POST'])
def delete_faculty(faculty_id):
    if not admin_required():
        return redirect(url_for('login'))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(
            'SELECT id, user_id, faculty_name FROM faculty WHERE id=%s LIMIT 1',
            (faculty_id,)
        )
        row = cur.fetchone()
        if not row:
            flash('Faculty not found.', 'danger')
            return redirect(url_for('faculty'))
        def table_has_column(table_name, column_name):
            try:
                cur.execute(f"SHOW COLUMNS FROM `{table_name}`")
                columns = {str(r.get('Field', '')).lower() for r in cur.fetchall()}
                return column_name.lower() in columns
            except mysql.connector.Error as table_error:
                if getattr(table_error, 'errno', None) == 1146:
                    return False
                raise
        for table_name in (
            'ep_question_papers',
            'ep_faculty_assignments',
            'ep_faculty_details',
        ):
            if table_has_column(table_name, 'faculty_id'):
                cur.execute(
                    f'DELETE FROM `{table_name}` WHERE `faculty_id`=%s',
                    (faculty_id,)
                )
        if table_has_column('classes', 'faculty_id'):
            cur.execute(
                'UPDATE classes SET faculty_id=NULL WHERE faculty_id=%s',
                (faculty_id,)
            )
        user_id = row.get('user_id')
        cur.execute('DELETE FROM `faculty` WHERE `id`=%s', (faculty_id,))
        if user_id and table_has_column('users', 'id'):
            cur.execute('DELETE FROM `users` WHERE `id`=%s', (user_id,))
        conn.commit()
        flash(
            f"Faculty '{row.get('faculty_name') or faculty_id}' deleted successfully.",
            'success'
        )
    except (mysql.connector.Error, ValueError) as e:
        conn.rollback()
        flash(f'Unable to delete faculty: {e}', 'danger')
    finally:
        cur.close()
        conn.close()
    return redirect(url_for('faculty'))
def _student_academic_rows(cur):
    """Return the department, class, and batch options used by student pages."""
    cur.execute("SELECT id, department_name, department_code FROM departments ORDER BY department_name")
    departments = cur.fetchall()
    cur.execute("SELECT id, department_id, class_name FROM classes ORDER BY department_id, id")
    classes = cur.fetchall()
    cur.execute("""
        SELECT department_id, course_type, start_year, end_year, course_duration
        FROM ep_department_batches
        ORDER BY department_id
    """)
    batches = cur.fetchall()
    return departments, classes, batches
def _academic_context(batch, class_name=None, academic_start=None):
    """Build the display context for a department batch and academic year."""
    start_year = int(batch.get('start_year') or 0)
    end_year = int(batch.get('end_year') or start_year)
    raw_duration = batch.get('course_duration')
    try:
        duration = int(raw_duration)
    except (TypeError, ValueError):
        duration = max(end_year - start_year, 1)
        if isinstance(raw_duration, str):
            years = re.findall(r'\d{4}', raw_duration)
            if len(years) >= 2:
                duration = max(int(years[-1]) - int(years[0]), 1)
    if academic_start is None:
        year_no = _infer_year_from_class(class_name)
        academic_start = start_year + year_no - 1
    else:
        year_no = max(1, int(academic_start) - start_year + 1)
    return {
        'academic_year': f'{academic_start}-{int(academic_start) + 1}',
        'year_no': year_no,
        'year_label': f'{year_no}{"st" if year_no == 1 else "nd" if year_no == 2 else "rd" if year_no == 3 else "th"} Year',
        'batch': f'{start_year}-{end_year}',
        'course_duration': duration,
    }
@app.route('/admin/students')
def students():
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        departments,classes,batches=_student_academic_rows(cur)
        dept_id=request.args.get('department_id','').strip()
        academic_year=request.args.get('academic_year','').strip()
        search=request.args.get('search','').strip()
        students=[]; selected_context=None
        if dept_id:
            batch=next((b for b in batches if str(b['department_id'])==dept_id),None)
            if batch:
                start=int(batch['start_year']); end=int(batch['end_year'])
                years=[f'{y}-{y+1}' for y in range(start,end)]
                if academic_year in years:
                    ay_start=int(academic_year.split('-')[0]); ctx=_academic_context(batch,academic_start=ay_start); selected_context=ctx
                    year_no=ctx['year_no']
                    class_row=next((c for c in classes if str(c['department_id'])==dept_id and _infer_year_from_class(c['class_name'])==year_no),None)
                    if class_row:
                        q="""SELECT s.id,s.register_number,s.student_name,s.email,s.phone,d.department_name,c.class_name,b.course_duration AS academic_batch FROM students s JOIN departments d ON d.id=s.department_id JOIN classes c ON c.id=s.class_id LEFT JOIN ep_department_batches b ON b.department_id=s.department_id WHERE s.department_id=%s AND s.class_id=%s"""
                        params=[dept_id,class_row['id']]
                        if search:
                            q+=' AND (s.register_number LIKE %s OR s.student_name LIKE %s OR s.email LIKE %s OR s.phone LIKE %s)'; like='%'+search+'%'; params.extend([like]*4)
                        q+=' ORDER BY s.student_name ASC'; cur.execute(q,tuple(params)); students=cur.fetchall()
        return render_template('admin/students.html',students=students,departments=departments,classes=classes,batches=batches,selected_department=dept_id,selected_academic_year=academic_year,search=search,selected_context=selected_context)
    finally: cur.close(); db.close()
def _student_login_password_from_dob(dob_value):
    """Return the student's login password in DD/MM/YYYY format."""
    if dob_value is None or str(dob_value).strip() == "":
        raise ValueError("Date of Birth is required for student login.")
    value = dob_value
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    raise ValueError("Invalid Date of Birth. Use DD/MM/YYYY format.")
def _student_form_data(cur):
    return _student_academic_rows(cur)
def _student_bulk_dir():
    path = os.path.join(tempfile.gettempdir(), 'e_progress_card_student_bulk')
    os.makedirs(path, exist_ok=True)
    return path
def _student_excel_value(value):
    if value is None:
        return ''
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value).strip()
def _read_student_excel(path):
    aliases = {
        'register_number': {'register_number', 'register_no', 'register', 'roll_no', 'roll_number'},
        'student_name': {'student_name', 'name', 'student'},
        'email': {'email', 'student_email'},
        'phone': {'phone', 'mobile', 'mobile_number', 'phone_number'},
        'dob': {'dob', 'date_of_birth', 'birth_date'},
    }
    ext = os.path.splitext(path)[1].lower()
    records = []
    if ext == '.csv':
        with open(path, 'r', encoding='utf-8-sig', newline='') as fh:
            reader = csv.DictReader(fh)
            if not reader.fieldnames:
                raise ValueError('The CSV file has no header row.')
            headers = {_normalize_excel_header(h): h for h in reader.fieldnames}
            mapped = {field: next((headers[name] for name in names if name in headers), None) for field, names in aliases.items()}
            missing = [field.replace('_', ' ').title() for field in aliases if not mapped[field]]
            if missing:
                raise ValueError('Missing Excel columns: ' + ', '.join(sorted(missing)))
            for row_no, row in enumerate(reader, start=2):
                values = {field: _student_excel_value(row.get(column)) for field, column in mapped.items()}
                if any(values.values()):
                    values['_row_no'] = row_no
                    records.append(values)
    elif ext == '.xlsx':
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if not header:
                raise ValueError('The Excel file has no header row.')
            headers = {_normalize_excel_header(value): index for index, value in enumerate(header) if value is not None}
            mapped = {field: next((headers[name] for name in names if name in headers), None) for field, names in aliases.items()}
            missing = [field.replace('_', ' ').title() for field in aliases if mapped[field] is None]
            if missing:
                raise ValueError('Missing Excel columns: ' + ', '.join(sorted(missing)))
            for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                values = {field: _student_excel_value(row[index] if index < len(row) else '') for field, index in mapped.items()}
                if any(values.values()):
                    values['_row_no'] = row_no
                    records.append(values)
        finally:
            wb.close()
    else:
        raise ValueError('Only .xlsx and .csv files are supported.')
    return records
def _validate_student_bulk(conn, rows, department_id, class_id):
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute('SELECT id, department_name FROM departments WHERE id=%s', (department_id,))
        department = cur.fetchone()
        cur.execute('SELECT id, department_id, class_name FROM classes WHERE id=%s AND department_id=%s', (class_id, department_id))
        selected_class = cur.fetchone()
        if not department or not selected_class:
            raise ValueError('Selected class does not belong to the selected department.')
        cur.execute('SELECT * FROM ep_department_batches WHERE department_id=%s LIMIT 1', (department_id,))
        batch = cur.fetchone()
        if not batch:
            raise ValueError('No academic batch is configured for the selected department.')
        context = _academic_context(batch, selected_class['class_name'])
        cur.execute('SELECT register_number, email, phone FROM students')
        existing = cur.fetchall()
        existing_registers = {str(row['register_number']).strip().upper() for row in existing if row.get('register_number')}
        existing_emails = {str(row['email']).strip().lower() for row in existing if row.get('email')}
        existing_phones = {str(row['phone']).strip() for row in existing if row.get('phone')}
        seen_registers, seen_emails, seen_phones = set(), set(), set()
        output = []
        for index, row in enumerate(rows, start=1):
            register_number = _student_excel_value(row.get('register_number')).upper()
            student_name = _student_excel_value(row.get('student_name'))
            email = _student_excel_value(row.get('email')).lower()
            phone = _student_excel_value(row.get('phone'))
            dob = _student_excel_value(row.get('dob'))
            errors = []
            if not register_number:
                errors.append('Register Number is required')
            elif register_number in existing_registers or register_number in seen_registers:
                errors.append('Register Number already exists or is duplicated')
            if not student_name:
                errors.append('Student Name is required')
            if not email or not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email):
                errors.append('Valid Email is required')
            elif email in existing_emails or email in seen_emails:
                errors.append('Email already exists or is duplicated')
            if not phone or not phone.isdigit() or len(phone) != 10:
                errors.append('Phone number must contain exactly 10 digits')
            elif phone in existing_phones or phone in seen_phones:
                errors.append('Phone already exists or is duplicated')
            try:
                dob = _student_login_password_from_dob(dob)
            except ValueError as exc:
                errors.append(str(exc))
            valid = not errors
            if valid:
                seen_registers.add(register_number)
                seen_emails.add(email)
                seen_phones.add(phone)
            output.append({
                'sno': index, 'row_no': row.get('_row_no', index + 1),
                'register_number': register_number, 'student_name': student_name,
                'email': email, 'phone': phone, 'dob': dob,
                'valid': valid, 'status': 'Valid' if valid else 'Invalid', 'errors': errors,
            })
        return output, department, selected_class, batch, context
    finally:
        cur.close()
@app.route('/admin/students/add',methods=['GET','POST'])
def add_student():
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        departments,classes,batches=_student_form_data(cur)
        if request.method=='POST':
            name=request.form.get('student_name','').strip(); email=request.form.get('email','').strip(); phone=request.form.get('phone','').strip(); dept=request.form.get('department_id','').strip(); cls=request.form.get('class_id','').strip(); gender=request.form.get('gender','').strip(); dob=request.form.get('dob','').strip() or None; address=request.form.get('address','').strip()
            if not all([name,email,phone,dept,cls,dob]): raise ValueError('All required student fields must be filled, including Date of Birth.')
            login_password=_student_login_password_from_dob(dob)
            if not phone.isdigit() or len(phone)!=10: raise ValueError('Phone number must contain exactly 10 digits.')
            cur.execute('SELECT id FROM students WHERE LOWER(email)=LOWER(%s) OR phone=%s',(email,phone));
            if cur.fetchone(): raise ValueError('Student email or phone already exists.')
            cur.execute('SELECT id,class_name FROM classes WHERE id=%s AND department_id=%s',(cls,dept)); c=cur.fetchone()
            if not c: raise ValueError('Selected class does not belong to the selected department.')
            cur.execute('SELECT department_code FROM departments WHERE id=%s',(dept,)); d=cur.fetchone(); code=(d['department_code'] or '').upper()
            ycode=get_year_code(c['class_name'])
            if not ycode: raise ValueError('Invalid class/year.')
            prefix=f'{code}{ycode}'; cur.execute('SELECT register_number FROM students WHERE register_number LIKE %s',(prefix+'%',)); highest=0
            for r in cur.fetchall():
                m=re.fullmatch(re.escape(prefix)+r'(\d+)',(r['register_number'] or '').upper())
                if m: highest=max(highest,int(m.group(1)))
            reg=f'{prefix}{highest+1:02d}'
            username=reg
            cur.execute("INSERT INTO users(username,password,role) VALUES(%s,%s,'student')",(username,generate_password_hash(login_password))); uid=cur.lastrowid
            cur.execute('INSERT INTO students(user_id,register_number,student_name,email,phone,department_id,class_id) VALUES(%s,%s,%s,%s,%s,%s,%s)',(uid,reg,name,email,phone,dept,cls)); sid=cur.lastrowid
            cur.execute('INSERT INTO ep_student_details(student_id,gender,dob,address) VALUES(%s,%s,%s,%s)',(sid,gender,dob,address))
            db.commit(); flash(f'Student added successfully. Register Number: {reg}','success'); return redirect(url_for('students'))
        return render_template('admin/student_form.html',departments=departments,classes=classes,batches=batches,student=None)
    except (ValueError,mysql.connector.Error) as e:
        db.rollback(); flash(f'Unable to add student: {e}','danger'); return render_template('admin/student_form.html',departments=departments,classes=classes,batches=batches,student=None)
    finally: cur.close(); db.close()
@app.route('/admin/students/template')
def download_student_template():
    if not admin_required(): return redirect(url_for('login'))
    wb=Workbook(); ws=wb.active; ws.title='Student Data'
    ws.append(['Register Number','Student Name','Email','Phone','DOB'])
    ws.append(['26BCA001','Student One','student1@example.com','9876543210','12/04/2007'])
    ws.append(['26BCA002','Student Two','student2@example.com','9876543211','15/08/2007'])
    for cell in ws[1]: cell.font=cell.font.copy(bold=True)
    ws.freeze_panes='A2'
    for i,w in enumerate([22,28,32,16,16],1): ws.column_dimensions[chr(64+i)].width=w
    out=BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name='student_upload_template.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/admin/students/bulk-upload',methods=['POST'])
def student_bulk_upload():
    if not admin_required(): return redirect(url_for('login'))
    dept=request.form.get('department_id','').strip(); cls=request.form.get('class_id','').strip(); upload=request.files.get('student_excel')
    if not dept or not cls: flash('Select Department and Class / Year before uploading students.','danger'); return redirect(url_for('add_student'))
    if not upload or not upload.filename: flash('Please select an Excel or CSV file.','danger'); return redirect(url_for('add_student'))
    ext=os.path.splitext(upload.filename)[1].lower()
    if ext not in {'.xlsx','.csv'}: flash('Only .xlsx and .csv files are supported.','danger'); return redirect(url_for('add_student'))
    token=uuid.uuid4().hex; path=os.path.join(_student_bulk_dir(),token+ext)
    try:
        upload.save(path); rows=_read_student_excel(path)
        if not rows: raise ValueError('The uploaded file contains no student records.')
        conn=get_db_connection()
        try: preview,dept_row,class_row,batch,context=_validate_student_bulk(conn,rows,int(dept),int(cls))
        finally: conn.close()
        preview_path=os.path.join(_student_bulk_dir(),token+'.json')
        with open(preview_path,'w',encoding='utf-8') as fh: json.dump({'rows':preview,'department_id':int(dept),'class_id':int(cls),'department_name':dept_row['department_name'],'class_name':class_row['class_name'],'batch':context['batch'],'academic_year':context['academic_year'],'year_label':context['year_label']},fh,ensure_ascii=False)
        session['student_bulk_token']=token
        return render_template('admin/student_bulk_preview.html',rows=preview,token=token,department_name=dept_row['department_name'],class_name=class_row['class_name'],academic_batch=context['batch'],academic_year=context['academic_year'],year_label=context['year_label'])
    except Exception as e:
        try: os.remove(path)
        except OSError: pass
        flash(f'Unable to read student file: {e}','danger'); return redirect(url_for('add_student'))
@app.route('/admin/students/bulk-confirm',methods=['POST'])
def student_bulk_confirm():
    if not admin_required(): return redirect(url_for('login'))
    token=request.form.get('token','').strip()
    if not token or token!=session.get('student_bulk_token') or not re.fullmatch(r'[a-f0-9]{32}',token):
        flash('Student bulk upload session expired. Please upload the Excel file again.','danger'); return redirect(url_for('add_student'))
    preview_path=os.path.join(_student_bulk_dir(),token+'.json')
    try:
        with open(preview_path,'r',encoding='utf-8') as fh: data=json.load(fh)
    except Exception:
        flash('Student bulk preview expired. Please upload the Excel file again.','danger'); return redirect(url_for('add_student'))
    conn=get_db_connection(); success=0; failed=[]
    try:
        cur=conn.cursor(dictionary=True)
        rows=data['rows']; dept=int(data['department_id']); cls=int(data['class_id'])
        fresh,_,_,_,_= _validate_student_bulk(conn,[{'register_number':r.get('register_number'),'student_name':r.get('student_name'),'email':r.get('email'),'phone':r.get('phone'),'dob':r.get('dob'),'_row_no':r.get('row_no')} for r in rows],dept,cls)
        for row in fresh:
            if not row['valid']:
                failed.append({'name':row['student_name'] or f"Row {row['row_no']}",'errors':row['errors']}); continue
            try:
                username=row['register_number']
                login_password=_student_login_password_from_dob(row['dob'])
                cur.execute("INSERT INTO users(username,password,role) VALUES(%s,%s,'student')",(username,generate_password_hash(login_password))); uid=cur.lastrowid
                cur.execute('INSERT INTO students(user_id,register_number,student_name,email,phone,department_id,class_id) VALUES(%s,%s,%s,%s,%s,%s,%s)',(uid,row['register_number'],row['student_name'],row['email'],row['phone'],dept,cls)); sid=cur.lastrowid
                cur.execute('INSERT INTO ep_student_details(student_id,gender,dob,address) VALUES(%s,NULL,%s,NULL)',(sid,row['dob']))
                conn.commit(); success+=1
            except (ValueError,mysql.connector.Error) as e:
                conn.rollback(); failed.append({'name':row['student_name'] or f"Row {row['row_no']}",'errors':[str(e)]})
        cur.close()
        flash(f'Students added successfully: {success} added, {len(failed)} skipped.','success' if not failed else 'warning')
    except Exception as e:
        conn.rollback(); flash(f'Student bulk import failed: {e}','danger')
    finally:
        conn.close()
        for suffix in ('.json','.xlsx','.csv'):
            try: os.remove(os.path.join(_student_bulk_dir(),token+suffix))
            except OSError: pass
        session.pop('student_bulk_token',None)
    if failed: session['student_bulk_failures']=failed[:200]
    else: session.pop('student_bulk_failures',None)
    return redirect(url_for('students'))
@app.route('/admin/students/view/<int:student_id>')
def view_student(student_id):
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute("""SELECT s.*,d.department_name,d.department_code,c.class_name,b.course_type,b.start_year,b.end_year,b.course_duration AS academic_batch,u.username,sd.gender,sd.dob,sd.address FROM students s LEFT JOIN departments d ON d.id=s.department_id LEFT JOIN classes c ON c.id=s.class_id LEFT JOIN ep_department_batches b ON b.department_id=s.department_id LEFT JOIN users u ON u.id=s.user_id LEFT JOIN ep_student_details sd ON sd.student_id=s.id WHERE s.id=%s LIMIT 1""",(student_id,))
        student=cur.fetchone()
        if not student: flash('Student not found.','danger'); return redirect(url_for('students'))
        return render_template('admin/student_view.html',student=student)
    finally: cur.close(); db.close()
@app.route('/admin/students/edit/<int:student_id>',methods=['GET','POST'])
def edit_student(student_id):
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute("""SELECT s.*,u.username,sd.gender,sd.dob,sd.address,b.course_duration AS academic_batch FROM students s LEFT JOIN users u ON u.id=s.user_id LEFT JOIN ep_student_details sd ON sd.student_id=s.id LEFT JOIN ep_department_batches b ON b.department_id=s.department_id WHERE s.id=%s""",(student_id,)); st=cur.fetchone()
        if not st: flash('Student not found.','danger'); return redirect(url_for('students'))
        departments,classes,batches=_student_form_data(cur)
        if request.method=='POST':
            name=request.form.get('student_name','').strip(); email=request.form.get('email','').strip(); phone=request.form.get('phone','').strip(); dept=request.form.get('department_id','').strip(); cls=request.form.get('class_id','').strip(); gender=request.form.get('gender','').strip(); dob=request.form.get('dob','').strip() or None; address=request.form.get('address','').strip()
            if not all([name,email,phone,dept,cls,dob]): raise ValueError('All required student fields must be filled, including Date of Birth.')
            login_password=_student_login_password_from_dob(dob)
            if not phone.isdigit() or len(phone)!=10: raise ValueError('Phone number must contain exactly 10 digits.')
            cur.execute('SELECT id,class_name FROM classes WHERE id=%s AND department_id=%s',(cls,dept)); selected_class=cur.fetchone()
            if not selected_class: raise ValueError('Selected class does not belong to the selected department.')
            cur.execute('SELECT id FROM students WHERE (LOWER(email)=LOWER(%s) OR phone=%s) AND id<>%s',(email,phone,student_id))
            if cur.fetchone(): raise ValueError('Student email or phone already exists.')
            cur.execute('UPDATE students SET student_name=%s,email=%s,phone=%s,department_id=%s,class_id=%s WHERE id=%s',(name,email,phone,dept,cls,student_id)); cur.execute('UPDATE users SET username=%s,password=%s WHERE id=%s',(st['register_number'],generate_password_hash(login_password),st['user_id']))
            cur.execute("""INSERT INTO ep_student_details(student_id,gender,dob,address) VALUES(%s,%s,%s,%s) ON DUPLICATE KEY UPDATE gender=VALUES(gender),dob=VALUES(dob),address=VALUES(address)""",(student_id,gender,dob,address)); db.commit(); flash('Student updated successfully.','success'); return redirect(url_for('students'))
        return render_template('admin/student_form.html',departments=departments,classes=classes,batches=batches,student=st,edit_mode=True)
    except (ValueError,mysql.connector.Error) as e:
        db.rollback(); flash(f'Unable to update student: {e}','danger'); return redirect(url_for('edit_student',student_id=student_id))
    finally: cur.close(); db.close()
@app.route('/admin/students/delete/<int:student_id>',methods=['POST'])
def delete_student(student_id):
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute('SELECT user_id FROM students WHERE id=%s',(student_id,)); st=cur.fetchone()
        if not st: flash('Student not found.','danger'); return redirect(url_for('students'))
        cur.execute('DELETE FROM marks WHERE student_id=%s',(student_id,)); cur.execute('DELETE FROM ep_student_semesters WHERE student_id=%s',(student_id,)); cur.execute('DELETE FROM ep_student_details WHERE student_id=%s',(student_id,)); cur.execute('DELETE FROM students WHERE id=%s',(student_id,))
        if st.get('user_id'): cur.execute('DELETE FROM users WHERE id=%s',(st['user_id'],))
        db.commit(); flash('Student deleted successfully.','success')
    except mysql.connector.Error as e: db.rollback(); flash(f'Unable to delete student: {e}','danger')
    finally: cur.close(); db.close()
    return redirect(url_for('students'))
@app.route("/admin/subjects")
def subjects():
    if not admin_required(): return redirect(url_for("login"))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute("""
            SELECT s.id,s.subject_code,s.subject_name,
                   COUNT(DISTINCT a.id) AS assignment_count
            FROM subjects s
            LEFT JOIN ep_faculty_assignments a ON a.subject_id=s.id
            GROUP BY s.id,s.subject_code,s.subject_name
            ORDER BY s.subject_name,s.subject_code
        """)
        rows=cur.fetchall()
        return render_template('admin/subjects.html',subjects=rows)
    finally: cur.close(); db.close()
@app.route('/admin/subjects/add', methods=['GET','POST'])
def add_subject():
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        if request.method=='POST':
            code=request.form.get('subject_code','').strip().upper()
            name=request.form.get('subject_name','').strip()
            if not code or not name:
                raise ValueError('Subject code and subject name are required.')
            cur.execute('SELECT id FROM subjects WHERE subject_code=%s LIMIT 1',(code,))
            if cur.fetchone(): raise ValueError('Subject code already exists.')
            cur.execute('INSERT INTO subjects(subject_code,subject_name,department_id,class_id) VALUES(%s,%s,NULL,NULL)',(code,name))
            db.commit(); flash('Subject added to Subject Master successfully. Assign it later from Faculty Subject.','success'); return redirect(url_for('subjects'))
        return render_template('admin/subject_form.html',subject=None,edit_mode=False)
    except (ValueError,mysql.connector.Error) as e:
        db.rollback(); flash(f'Unable to add subject: {e}','danger'); return render_template('admin/subject_form.html',subject=None,edit_mode=False)
    finally: cur.close(); db.close()
@app.route('/admin/subjects/edit/<int:subject_id>', methods=['GET','POST'])
def edit_subject(subject_id):
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor(dictionary=True)
    try:
        cur.execute('SELECT id,subject_code,subject_name FROM subjects WHERE id=%s',(subject_id,)); subject=cur.fetchone()
        if not subject: flash('Subject not found.','danger'); return redirect(url_for('subjects'))
        if request.method=='POST':
            code=request.form.get('subject_code','').strip().upper()
            name=request.form.get('subject_name','').strip()
            if not code or not name: raise ValueError('Subject code and subject name are required.')
            cur.execute('SELECT id FROM subjects WHERE subject_code=%s AND id<>%s',(code,subject_id))
            if cur.fetchone(): raise ValueError('Subject code already exists.')
            cur.execute('UPDATE subjects SET subject_code=%s,subject_name=%s WHERE id=%s',(code,name,subject_id))
            db.commit(); flash('Subject updated successfully. Assignments were kept unchanged.','success'); return redirect(url_for('subjects'))
        return render_template('admin/subject_form.html',subject=subject,edit_mode=True)
    except (ValueError,mysql.connector.Error) as e:
        db.rollback(); flash(f'Unable to update subject: {e}','danger'); return render_template('admin/subject_form.html',subject=subject,edit_mode=True)
    finally: cur.close(); db.close()
def _subject_bulk_dir():
    path=os.path.join(tempfile.gettempdir(),'e_progress_card_subject_bulk')
    os.makedirs(path,exist_ok=True)
    return path
def _read_subject_excel(path):
    """Read a subject master file containing only Subject Code and Subject Name."""
    ext=os.path.splitext(path)[1].lower()
    aliases={
        'subject_code': {'subject_code','subjectcode','subject_code_no','code','paper_code'},
        'subject_name': {'subject_name','subjectname','name','subject','paper_name'},
    }
    records=[]
    if ext=='.csv':
        with open(path,'r',encoding='utf-8-sig',newline='') as fh:
            reader=csv.DictReader(fh)
            if not reader.fieldnames: raise ValueError('The CSV file has no header row.')
            normalized={_normalize_excel_header(h):h for h in reader.fieldnames}
            mapped={field:next((normalized[a] for a in names if a in normalized),None) for field,names in aliases.items()}
            missing=[x.replace('_',' ').title() for x in aliases if not mapped[x]]
            if missing: raise ValueError('Missing Excel columns: '+', '.join(sorted(missing)))
            for row_no,row in enumerate(reader,start=2):
                vals={f:_student_excel_value(row.get(col)) for f,col in mapped.items()}
                if not any(vals.values()): continue
                vals['_row_no']=row_no; records.append(vals)
    elif ext=='.xlsx':
        wb=load_workbook(path,read_only=True,data_only=True); ws=wb.active
        header=next(ws.iter_rows(min_row=1,max_row=1,values_only=True),None)
        if not header:
            wb.close(); raise ValueError('The Excel file has no header row.')
        normalized={_normalize_excel_header(h):i for i,h in enumerate(header) if h is not None}
        mapped={field:next((normalized[a] for a in names if a in normalized),None) for field,names in aliases.items()}
        missing=[x.replace('_',' ').title() for x in aliases if mapped[x] is None]
        if missing:
            wb.close(); raise ValueError('Missing Excel columns: '+', '.join(sorted(missing)))
        for row_no,row in enumerate(ws.iter_rows(min_row=2,values_only=True),start=2):
            vals={f:_student_excel_value(row[col_idx] if col_idx < len(row) else '') for f,col_idx in mapped.items()}
            if not any(vals.values()): continue
            vals['_row_no']=row_no; records.append(vals)
        wb.close()
    else:
        raise ValueError('Only .xlsx and .csv files are supported.')
    return records
def _validate_subject_bulk(conn,rows):
    """Validate subject master rows; assignment details are deliberately excluded."""
    cur=conn.cursor(dictionary=True)
    try:
        cur.execute('SELECT id,subject_code FROM subjects')
        existing_codes={str(r['subject_code']).strip().upper():r['id'] for r in cur.fetchall() if r.get('subject_code')}
        seen_codes=set(); out=[]
        for index,row in enumerate(rows,start=1):
            code=_student_excel_value(row.get('subject_code')).upper()
            name=_student_excel_value(row.get('subject_name'))
            errors=[]
            if not code: errors.append('Subject Code is required')
            elif not re.fullmatch(r'[A-Z0-9][A-Z0-9._-]{1,29}',code): errors.append('Invalid Subject Code')
            if not name: errors.append('Subject Name is required')
            if code and code in existing_codes: errors.append('Subject Code already exists')
            if code and code in seen_codes: errors.append('Duplicate Subject Code in this file')
            valid=not errors
            if valid: seen_codes.add(code)
            out.append({
                'sno':index,'row_no':row.get('_row_no',index+1),
                'subject_code':code,'subject_name':name,
                'valid':valid,'status':'Valid' if valid else 'Invalid','errors':errors
            })
        return out
    finally:
        cur.close()
@app.route('/admin/subjects/template')
def download_subject_template():
    if not admin_required(): return redirect(url_for('login'))
    wb=Workbook(); ws=wb.active; ws.title='Subject Data'
    ws.append(['Subject Code','Subject Name'])
    ws.append(['MCA101','Programming in Python'])
    ws.append(['MCA102','Database Management Systems'])
    ws.append(['MCA103','Data Structures'])
    ws.append(['MCA104','Artificial Intelligence'])
    for cell in ws[1]: cell.font=cell.font.copy(bold=True)
    ws.freeze_panes='A2'
    ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=42
    out=BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name='subject_master_upload_template.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/admin/subjects/bulk-upload',methods=['POST'])
def subject_bulk_upload():
    if not admin_required(): return redirect(url_for('login'))
    upload=request.files.get('subject_excel')
    if not upload or not upload.filename:
        flash('Please select an Excel or CSV file.','danger'); return redirect(url_for('subjects'))
    ext=os.path.splitext(upload.filename)[1].lower()
    if ext not in {'.xlsx','.csv'}:
        flash('Only .xlsx and .csv files are supported.','danger'); return redirect(url_for('subjects'))
    token=uuid.uuid4().hex; path=os.path.join(_subject_bulk_dir(),token+ext)
    try:
        upload.save(path); rows=_read_subject_excel(path)
        if not rows: raise ValueError('The uploaded file contains no subject records.')
        conn=get_db_connection()
        try: preview=_validate_subject_bulk(conn,rows)
        finally: conn.close()
        preview_path=os.path.join(_subject_bulk_dir(),token+'.json')
        with open(preview_path,'w',encoding='utf-8') as fh: json.dump({'rows':preview},fh,ensure_ascii=False)
        session['subject_bulk_token']=token
        return render_template('admin/subject_bulk_preview.html',rows=preview,token=token)
    except Exception as e:
        try: os.remove(path)
        except OSError: pass
        try: os.remove(os.path.join(_subject_bulk_dir(),token+'.json'))
        except OSError: pass
        flash(f'Unable to read subject file: {e}','danger'); return redirect(url_for('subjects'))
@app.route('/admin/subjects/bulk-confirm',methods=['POST'])
def subject_bulk_confirm():
    if not admin_required(): return redirect(url_for('login'))
    token=request.form.get('token','').strip()
    if not token or token!=session.get('subject_bulk_token') or not re.fullmatch(r'[a-f0-9]{32}',token):
        flash('Subject bulk upload session expired. Please upload the Excel file again.','danger'); return redirect(url_for('subjects'))
    preview_path=os.path.join(_subject_bulk_dir(),token+'.json')
    try:
        with open(preview_path,'r',encoding='utf-8') as fh: data=json.load(fh)
    except Exception:
        flash('Subject bulk preview expired. Please upload the Excel file again.','danger'); return redirect(url_for('subjects'))
    conn=get_db_connection(); success=0; failed=[]
    try:
        cur=conn.cursor(dictionary=True)
        rows=data['rows']
        fresh=_validate_subject_bulk(conn,rows)
        for row in fresh:
            if not row['valid']:
                failed.append({'name':row['subject_name'] or f"Row {row['row_no']}",'errors':row['errors']}); continue
            try:
                cur.execute('SELECT id FROM subjects WHERE subject_code=%s LIMIT 1',(row['subject_code'],))
                if cur.fetchone():
                    failed.append({'name':row['subject_name'],'errors':['Subject Code already exists']}); continue
                cur.execute('INSERT INTO subjects(subject_code,subject_name,department_id,class_id) VALUES(%s,%s,NULL,NULL)',(row['subject_code'],row['subject_name']))
                conn.commit(); success+=1
            except mysql.connector.Error as e:
                conn.rollback(); failed.append({'name':row['subject_name'] or f"Row {row['row_no']}",'errors':[str(e)]})
        cur.close()
        flash(f'Subjects added successfully: {success} added, {len(failed)} skipped.','success' if not failed else 'warning')
    except Exception as e:
        conn.rollback(); flash(f'Subject bulk import failed: {e}','danger')
    finally:
        conn.close()
        for suffix in ('.json','.xlsx','.csv'):
            try: os.remove(os.path.join(_subject_bulk_dir(),token+suffix))
            except OSError: pass
        session.pop('subject_bulk_token',None)
    if failed: session['subject_bulk_failures']=failed[:200]
    else: session.pop('subject_bulk_failures',None)
    return redirect(url_for('subjects'))
@app.route('/admin/subjects/delete/<int:subject_id>', methods=['POST'])
def delete_subject(subject_id):
    if not admin_required(): return redirect(url_for('login'))
    db=get_db_connection(); cur=db.cursor()
    try:
        cur.execute('DELETE FROM ep_faculty_assignments WHERE subject_id=%s',(subject_id,)); cur.execute('DELETE FROM ep_subject_semesters WHERE subject_id=%s',(subject_id,)); cur.execute('DELETE FROM class_subjects WHERE subject_id=%s',(subject_id,)); cur.execute('DELETE FROM subjects WHERE id=%s',(subject_id,)); db.commit(); flash('Subject deleted successfully.','success')
    except mysql.connector.Error as e: db.rollback(); flash(f'Unable to delete subject: {e}','danger')
    finally: cur.close(); db.close()
    return redirect(url_for('subjects'))
@app.route("/admin/assignments")
def assignments():
    """Compatibility route for the Admin Assignments menu.
    The Admin UI now uses explicit Faculty + Class + Subject +
    Academic Year + Semester assignments.  Keep this original endpoint
    so existing links do not break.
    """
    return redirect(url_for("admin_faculty_assignments"))
@app.route(
    "/admin/results",
    methods=["GET", "POST"]
)
def results():
    if not admin_required():
        return redirect(url_for("login"))
    db = get_db_connection()
    cursor = db.cursor(dictionary=True)
    if request.method == "POST":
        visibility = request.form.get(
            "visibility"
        )
        value = (
            True
            if visibility == "on"
            else False
        )
        cursor.execute("""
            UPDATE result_settings
            SET visibility = %s
            WHERE id = 1
        """, (
            value,
        ))
        db.commit()
        flash(
            "Result visibility updated successfully.",
            "success"
        )
    cursor.execute("""
        SELECT visibility
        FROM result_settings
        WHERE id = 1
    """)
    setting = cursor.fetchone()
    visibility = (
        bool(setting["visibility"])
        if setting
        else False
    )
    cursor.close()
    db.close()
    return render_template(
        "admin/results.html",
        visibility=visibility
    )
from functools import wraps
from flask import jsonify
def faculty_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Please login first.", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "faculty":
            flash("You are not authorized to access this page.", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function
def get_faculty():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(
            """
            SELECT
                f.id AS faculty_id,
                f.user_id,
                f.faculty_name,
                f.email,
                f.phone,
                f.department_id,
                f.class_id,
                d.department_name,
                d.department_code,
                c.class_name,
                u.username,
                u.role
            FROM faculty f
            JOIN users u ON f.user_id = u.id
            LEFT JOIN departments d ON f.department_id = d.id
            LEFT JOIN classes c ON f.class_id = c.id
            WHERE f.user_id = %s
            LIMIT 1
            """,
            (session.get("user_id"),)
        )
        return cursor.fetchone()
    finally:
        cursor.close()
        conn.close()
def ensure_assignment_marks_schema(cur):
    """Create the assignment marks table on-demand for old databases.
    Some installations skipped the startup migration, so Assignment Mark Entry
    must never depend on ep_assignment_marks already existing.
    """
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ep_assignment_marks (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_id INT NOT NULL,
            subject_id INT NOT NULL,
            academic_year_id INT NOT NULL,
            semester_no INT NOT NULL,
            marks DECIMAL(6,2) NOT NULL DEFAULT 0,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY uq_ep_assignment_mark
                (student_id,subject_id,academic_year_id,semester_no)
        )
    """)
    try:
        for col in ("ca1", "ca2", "assignment", "external", "internal",
                    "total", "grade", "result", "appreciation"):
            cur.execute("""
                SELECT COLUMN_TYPE, IS_NULLABLE
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA=DATABASE()
                  AND TABLE_NAME='marks' AND COLUMN_NAME=%s
            """, (col,))
            meta=cur.fetchone()
            if meta:
                if isinstance(meta, dict):
                    col_type=meta.get("COLUMN_TYPE")
                    nullable=meta.get("IS_NULLABLE")
                else:
                    col_type, nullable=meta[0], meta[1]
                if str(nullable).upper()=="NO" and col_type:
                    cur.execute(f"ALTER TABLE marks MODIFY `{col}` {col_type} NULL")
    except mysql.connector.Error:
        pass
def faculty_page_data():
    faculty = get_faculty()
    if not faculty:
        return None, None
    letter = (faculty.get("faculty_name") or "F").strip()[:1].upper()
    return faculty, letter
def get_assigned_subjects(cursor, class_id):
    cursor.execute(
        """
        SELECT
            s.id AS subject_id,
            s.subject_code,
            s.subject_name
        FROM subjects s
        INNER JOIN class_subjects cs ON cs.subject_id = s.id
        WHERE cs.class_id = %s
        ORDER BY s.subject_name
        """,
        (class_id,)
    )
    return cursor.fetchall()
def verify_student_access(cursor, faculty, student_id):
    """Authorize a student through the active faculty assignment.
    ep_student_semesters is preferred, with a safe legacy-class fallback for
    older rows that have not yet been migrated.
    """
    ay, sem, _ = get_selected_faculty_semester(cursor, faculty["faculty_id"])
    if ay and sem is not None:
        cursor.execute("""
            SELECT 1
            FROM students s
            INNER JOIN ep_faculty_assignments a
              ON a.faculty_id=%s
             AND a.academic_year_id=%s
             AND a.semester_no=%s
            LEFT JOIN ep_student_semesters es
              ON es.student_id=s.id
             AND es.academic_year_id=a.academic_year_id
             AND es.semester_no=a.semester_no
             AND es.class_id=a.class_id
            WHERE s.id=%s
              AND (es.id IS NOT NULL OR s.class_id=a.class_id)
            LIMIT 1
        """, (faculty["faculty_id"], ay["id"], sem, student_id))
        return bool(cursor.fetchone())
    if faculty.get("class_id"):
        cursor.execute("SELECT id FROM students WHERE id=%s AND class_id=%s LIMIT 1", (student_id, faculty.get("class_id")))
        return bool(cursor.fetchone())
    return False
def verify_subject_access(cursor, faculty, subject_id):
    ay, sem, _ = get_selected_faculty_semester(cursor, faculty["faculty_id"])
    if ay and sem is not None:
        cursor.execute("SELECT 1 FROM ep_faculty_assignments WHERE faculty_id=%s AND academic_year_id=%s AND semester_no=%s AND subject_id=%s LIMIT 1", (faculty["faculty_id"], ay["id"], sem, subject_id))
        return bool(cursor.fetchone())
    cursor.execute("SELECT s.id FROM subjects s INNER JOIN class_subjects cs ON cs.subject_id=s.id WHERE s.id=%s AND cs.class_id=%s LIMIT 1", (subject_id, faculty.get("class_id")))
    return bool(cursor.fetchone())
def verify_mark_access(cursor, faculty, student_id, subject_id):
    """Authorize exact faculty/student/subject using active-year assignment data."""
    if not faculty:
        return False
    ay, sem, _ = get_selected_faculty_semester(cursor, faculty["faculty_id"])
    if ay and sem is not None:
        cursor.execute("""
            SELECT 1
            FROM students s
            INNER JOIN ep_faculty_assignments a
              ON a.faculty_id=%s
             AND a.subject_id=%s
             AND a.academic_year_id=%s
             AND a.semester_no=%s
            LEFT JOIN ep_student_semesters es
              ON es.student_id=s.id
             AND es.academic_year_id=a.academic_year_id
             AND es.semester_no=a.semester_no
             AND es.class_id=a.class_id
            WHERE s.id=%s
              AND (es.id IS NOT NULL OR s.class_id=a.class_id)
            LIMIT 1
        """, (faculty["faculty_id"], subject_id, ay["id"], sem, student_id))
        return bool(cursor.fetchone())
    return bool(verify_student_access(cursor, faculty, student_id) and verify_subject_access(cursor, faculty, subject_id))
def get_year_code(class_name):
    text = (class_name or "").strip().lower()
    if "1st" in text or "first" in text:
        return "I"
    if "2nd" in text or "second" in text:
        return "II"
    if "3rd" in text or "third" in text:
        return "III"
    return None
def generate_faculty_register_number(cursor, faculty):
    department_code = (faculty.get("department_code") or "").strip().upper()
    year_code = get_year_code(faculty.get("class_name"))
    if not department_code or not year_code:
        raise ValueError("Unable to generate Register Number for the assigned class.")
    register_prefix = f"{department_code}{year_code}"
    cursor.execute(
        """
        SELECT register_number
        FROM students
        WHERE register_number LIKE %s
        """,
        (f"{register_prefix}%",)
    )
    highest_number = 0
    for row in cursor.fetchall():
        value = (row.get("register_number") or "").strip().upper()
        match = re.fullmatch(rf"{re.escape(register_prefix)}(\d+)", value)
        if match:
            highest_number = max(highest_number, int(match.group(1)))
    return f"{register_prefix}{highest_number + 1:02d}"
def calculate_internal(ca1, ca2, assignment):
    if not 0 <= ca1 <= 75:
        raise ValueError("CIA 1 must be between 0 and 75.")
    if not 0 <= ca2 <= 75:
        raise ValueError("CIA 2 must be between 0 and 75.")
    if not 0 <= assignment <= 5:
        raise ValueError("Assignment must be between 0 and 5.")
    ca1_converted = (ca1 / 75) * 10
    ca2_converted = (ca2 / 75) * 10
    return round(ca1_converted + ca2_converted + assignment, 2)
def calculate_marks(ca1, ca2, assignment, external):
    internal = calculate_internal(ca1, ca2, assignment)
    if not 0 <= external <= 75:
        raise ValueError("External must be between 0 and 75.")
    total = round(internal + external, 2)
    if total >= 90:
        grade, appreciation = "A+", "Excellent"
    elif total >= 80:
        grade, appreciation = "A", "Very Good"
    elif total >= 70:
        grade, appreciation = "B+", "Good"
    elif total >= 60:
        grade, appreciation = "B", "Satisfactory"
    elif total >= 30:
        grade, appreciation = "C", "Needs Improvement"
    else:
        grade, appreciation = "F", "Fail"
    result = "PASS" if total >= 30 else "FAIL"
    return internal, total, grade, result, appreciation
@app.route("/faculty/dashboard")
@faculty_required
def faculty_dashboard():
    """Simple faculty dashboard: exactly three main cards + mark status strip.
    All academic/mark access is restricted to the active academic year.
    """
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("logout"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cur)
        classes = []
        status = {
            "internal_completed": 0, "internal_pending": 0,
            "external_completed": 0, "external_pending": 0,
            "assignment_completed": 0, "assignment_pending": 0,
            "total_entries": 0
        }
        if ay:
            try:
                _ensure_faculty_student_semester_mappings(cur, faculty["faculty_id"], ay["id"])
                conn.commit()
            except mysql.connector.Error:
                conn.rollback()
            cur.execute("""
                SELECT DISTINCT
                       a.class_id,
                       c.class_name,
                       d.department_name,
                       d.department_code
                FROM ep_faculty_assignments a
                INNER JOIN classes c ON c.id = a.class_id
                INNER JOIN departments d ON d.id = c.department_id
                WHERE a.faculty_id=%s AND a.academic_year_id=%s
                ORDER BY d.department_name, c.class_name
            """, (faculty["faculty_id"], ay["id"]))
            classes = cur.fetchall()
            cur.execute("""
                SELECT
                    COUNT(*) AS total_entries,
                    SUM(CASE WHEN m.ca1 IS NOT NULL
                               AND m.ca2 IS NOT NULL
                               AND m.assignment IS NOT NULL THEN 1 ELSE 0 END) AS internal_completed,
                    SUM(CASE WHEN m.external IS NOT NULL THEN 1 ELSE 0 END) AS external_completed,
                    SUM(CASE WHEN m.assignment IS NOT NULL THEN 1 ELSE 0 END) AS assignment_completed
                FROM (
                    SELECT DISTINCT
                        a.class_id, a.subject_id, a.semester_no,
                        es.student_id
                    FROM ep_faculty_assignments a
                    INNER JOIN ep_student_semesters es
                        ON es.class_id=a.class_id
                       AND es.academic_year_id=a.academic_year_id
                       AND es.semester_no=a.semester_no
                    WHERE a.faculty_id=%s
                      AND a.academic_year_id=%s
                ) x
                LEFT JOIN (
                    SELECT m1.*
                    FROM marks m1
                    INNER JOIN (
                        SELECT student_id,subject_id,MAX(id) AS id
                        FROM marks
                        GROUP BY student_id,subject_id
                    ) lm ON lm.id=m1.id
                ) m
                  ON m.student_id=x.student_id
                 AND m.subject_id=x.subject_id
            """, (faculty["faculty_id"], ay["id"]))
            r = cur.fetchone() or {}
            total = int(r.get("total_entries") or 0)
            ic = int(r.get("internal_completed") or 0)
            ec = int(r.get("external_completed") or 0)
            ac = int(r.get("assignment_completed") or 0)
            status.update({
                "total_entries": total,
                "internal_completed": ic,
                "internal_pending": max(total-ic, 0),
                "external_completed": ec,
                "external_pending": max(total-ec, 0),
                "assignment_completed": ac,
                "assignment_pending": max(total-ac, 0),
            })
        total_students = 0
        if ay:
            cur.execute("""
                SELECT COUNT(DISTINCT es.student_id) AS cnt
                FROM ep_faculty_assignments a
                INNER JOIN ep_student_semesters es
                  ON es.class_id=a.class_id
                 AND es.academic_year_id=a.academic_year_id
                 AND es.semester_no=a.semester_no
                WHERE a.faculty_id=%s AND a.academic_year_id=%s
            """, (faculty["faculty_id"], ay["id"]))
            total_students = int((cur.fetchone() or {}).get("cnt") or 0)
        cur.execute("""
            SELECT COUNT(DISTINCT CONCAT(a.subject_id,'-',a.semester_no)) AS cnt
            FROM ep_faculty_assignments a
            WHERE a.faculty_id=%s AND a.academic_year_id=%s
        """, (faculty["faculty_id"], ay["id"] if ay else 0))
        total_subjects = int((cur.fetchone() or {}).get("cnt") or 0)
        return render_template(
            "faculty/dashboard.html",
            faculty=faculty,
            profile_letter=profile_letter,
            academic_year=ay,
            classes=classes,
            assigned_class_count=len(classes),
            subject_count=total_subjects,
            student_count=total_students,
            mark_status=status
        )
    finally:
        cur.close()
        conn.close()
@app.route("/faculty/classes")
@faculty_required
def faculty_classes():
    """Show every class assigned to this faculty in the active academic year."""
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cur)
        rows = []
        if ay:
            cur.execute("""
                SELECT
                    a.class_id,
                    c.class_name,
                    d.department_name,
                    d.department_code,
                    COUNT(DISTINCT CONCAT(a.subject_id,'-',a.semester_no)) AS subject_count,
                    COUNT(DISTINCT es.student_id) AS student_count,
                    GROUP_CONCAT(DISTINCT a.semester_no ORDER BY a.semester_no SEPARATOR ',') AS semesters
                FROM ep_faculty_assignments a
                INNER JOIN classes c ON c.id=a.class_id
                INNER JOIN departments d ON d.id=c.department_id
                LEFT JOIN ep_student_semesters es
                    ON es.class_id=a.class_id
                   AND es.academic_year_id=a.academic_year_id
                   AND es.semester_no=a.semester_no
                WHERE a.faculty_id=%s AND a.academic_year_id=%s
                GROUP BY a.class_id,c.class_name,d.department_name,d.department_code
                ORDER BY d.department_name,c.class_name
            """, (faculty["faculty_id"], ay["id"]))
            rows = cur.fetchall()
        return render_template(
            "faculty/my_classes.html",
            faculty=faculty,
            profile_letter=profile_letter,
            classes=rows,
            academic_year=ay
        )
    finally:
        cur.close()
        conn.close()
@app.route("/faculty/classes/<int:class_id>")
@faculty_required
def faculty_class_subjects(class_id):
    """List subjects assigned to the selected faculty/class in the active year."""
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cur)
        cls = None
        subjects = []
        if ay:
            cur.execute("""
                SELECT c.id,c.class_name,d.department_name,d.department_code
                FROM classes c
                INNER JOIN departments d ON d.id=c.department_id
                WHERE c.id=%s
                  AND EXISTS (
                      SELECT 1 FROM ep_faculty_assignments a
                      WHERE a.faculty_id=%s AND a.class_id=c.id
                        AND a.academic_year_id=%s
                  )
            """, (class_id, faculty["faculty_id"], ay["id"]))
            cls = cur.fetchone()
            if cls:
                cur.execute("""
                    SELECT DISTINCT
                        a.subject_id,
                        s.subject_code,
                        s.subject_name,
                        a.semester_no,
                        a.section,
                        cs.year_no
                    FROM ep_faculty_assignments a
                    INNER JOIN subjects s ON s.id=a.subject_id
                    LEFT JOIN ep_class_semesters cs
                      ON cs.class_id=a.class_id
                     AND cs.academic_year_id=a.academic_year_id
                     AND cs.semester_no=a.semester_no
                    WHERE a.faculty_id=%s
                      AND a.class_id=%s
                      AND a.academic_year_id=%s
                    ORDER BY a.semester_no,s.subject_name
                """, (faculty["faculty_id"], class_id, ay["id"]))
                subjects = cur.fetchall()
            else:
                flash("This class is not assigned to you for the active academic year.", "danger")
                return redirect(url_for("faculty_classes"))
        return render_template(
            "faculty/class_subjects.html",
            faculty=faculty,
            profile_letter=profile_letter,
            subjects=subjects,
            class_info=cls,
            academic_year=ay
        )
    finally:
        cur.close()
        conn.close()
@app.route("/faculty/subjects")
@faculty_required
def faculty_subjects():
    """Tabular subject details for all active-year faculty assignments."""
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cur)
        subjects = []
        if ay:
            cur.execute("""
                SELECT DISTINCT
                    a.subject_id,
                    s.subject_name,
                    s.subject_code,
                    d.department_name,
                    d.department_code,
                    c.class_name,
                    cs.year_no,
                    a.semester_no,
                    a.section,
                    COALESCE(db.course_duration, '—') AS academic_batch
                FROM ep_faculty_assignments a
                INNER JOIN subjects s ON s.id=a.subject_id
                INNER JOIN classes c ON c.id=a.class_id
                INNER JOIN departments d ON d.id=c.department_id
                LEFT JOIN ep_class_semesters cs
                  ON cs.class_id=a.class_id
                 AND cs.academic_year_id=a.academic_year_id
                 AND cs.semester_no=a.semester_no
                LEFT JOIN ep_department_batches db
                  ON db.department_id=d.id
                WHERE a.faculty_id=%s
                  AND a.academic_year_id=%s
                ORDER BY s.subject_name,c.class_name,a.semester_no
            """, (faculty["faculty_id"], ay["id"]))
            subjects = cur.fetchall()
        return render_template(
            "faculty/my_subjects.html",
            faculty=faculty,
            profile_letter=profile_letter,
            subjects=subjects,
            academic_year=ay
        )
    finally:
        cur.close()
        conn.close()
def _assessment_completion_status(cur, faculty, row, academic_year_id):
    """Return student-wise completion status for each assessment button.
    An assessment is marked complete only when every student mapped to the
    assigned class/semester has a complete set of marks for that assessment.
    Zero is a valid entered mark; missing/null values are treated as pending.
    """
    class_id = row["class_id"]
    subject_id = row["subject_id"]
    semester_no = row["semester_no"]
    cur.execute("""
        SELECT DISTINCT s.id AS student_id
        FROM ep_student_semesters es
        INNER JOIN students s ON s.id=es.student_id
        INNER JOIN ep_faculty_assignments a
          ON a.class_id=es.class_id
         AND a.academic_year_id=es.academic_year_id
         AND a.semester_no=es.semester_no
         AND a.faculty_id=%s
         AND a.subject_id=%s
        WHERE es.academic_year_id=%s AND es.semester_no=%s AND es.class_id=%s
        ORDER BY s.id
    """, (faculty["faculty_id"], subject_id, academic_year_id, semester_no, class_id))
    student_ids = [int(r["student_id"]) for r in cur.fetchall()]
    total_students = len(student_ids)
    if total_students == 0:
        return {
            "complete": False, "entered": 0, "total": 0,
            "label": "No students", "icon": "bi-dash-circle", "state": "empty"
        }
    def result(entered):
        complete = entered == total_students
        return {
            "complete": complete, "entered": entered, "total": total_students,
            "label": "All students" if complete else f"{entered}/{total_students} entered",
            "icon": "bi-check-circle-fill" if complete else "bi-hourglass-split",
            "state": "complete" if complete else "pending"
        }
    cur.execute("""
        SELECT COUNT(DISTINCT student_id) AS entered
        FROM ep_assignment_marks
        WHERE subject_id=%s AND academic_year_id=%s AND semester_no=%s
          AND student_id IN (%s)
    """ % ("%s", "%s", "%s", ",".join(["%s"] * total_students)),
        (subject_id, academic_year_id, semester_no, *student_ids))
    assignment_entered = int(cur.fetchone()["entered"] or 0)
    statuses = {}
    for cia in ("cia1", "cia2"):
        details = _load_cia_detail_marks(
            faculty["faculty_id"], subject_id, cia, academic_year_id, semester_no
        )
        entered = 0
        for sid in student_ids:
            detail = details.get(str(sid), {})
            try:
                section_a = detail.get("section_a", [])
                section_b = detail.get("section_b", [])
                section_c = detail.get("section_c", [])
                a_ok = len(section_a) == 10 and all(x not in (None, "") and 0 <= float(x) <= 2 for x in section_a)
                b_ok = len(section_b) == 5 and all(
                    isinstance(x, dict) and str(x.get("choice", "")).upper() in ("A", "B")
                    and x.get("mark") not in (None, "") and 0 <= float(x.get("mark")) <= 5
                    for x in section_b
                )
                c_values = [x for x in section_c if x not in (None, "")]
                c_ok = len(section_c) == 5 and len(c_values) == 3 and all(0 <= float(x) <= 10 for x in c_values)
                if a_ok and b_ok and c_ok:
                    entered += 1
            except (TypeError, ValueError):
                pass
        statuses[cia] = result(entered)
    cur.execute("""
        SELECT student_id, question_no, obtained_marks, choice
        FROM ep_question_marks
        WHERE subject_id=%s AND academic_year_id=%s AND semester_no=%s
          AND exam_type='external' AND student_id IN (%s)
    """ % ("%s", "%s", "%s", ",".join(["%s"] * total_students)),
        (subject_id, academic_year_id, semester_no, *student_ids))
    ext = {}
    for r in cur.fetchall():
        ext.setdefault(int(r["student_id"]), {})[str(r["question_no"])] = r
    external_entered = 0
    for sid in student_ids:
        q = ext.get(sid, {})
        a_ok = all(q.get(str(i), {}).get("obtained_marks") is not None for i in range(1, 11))
        b_ok = all(
            q.get(str(i), {}).get("obtained_marks") is not None
            and str(q.get(str(i), {}).get("choice") or "").upper() in ("A", "B")
            for i in range(11, 16)
        )
        c_answered = sum(q.get(str(i), {}).get("obtained_marks") is not None for i in range(16, 21))
        if a_ok and b_ok and c_answered == 3:
            external_entered += 1
    statuses["assignment"] = result(assignment_entered)
    statuses["external"] = result(external_entered)
    return statuses
@app.route("/faculty/marks-entry")
@faculty_required
def faculty_marks_entry():
    """Simple mark-entry hub: list every active-year faculty assignment in one table."""
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cur = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cur)
        assignments = []
        if ay:
            _ensure_faculty_student_semester_mappings(cur, faculty["faculty_id"], ay["id"])
            conn.commit()
            cur.execute("""
                SELECT DISTINCT
                    a.class_id, a.subject_id, a.semester_no, a.section,
                    c.class_name,
                    d.department_name, d.department_code,
                    s.subject_name, s.subject_code,
                    COALESCE(cs.year_no, CEIL(a.semester_no / 2)) AS year_no
                FROM ep_faculty_assignments a
                INNER JOIN classes c ON c.id=a.class_id
                INNER JOIN departments d ON d.id=c.department_id
                INNER JOIN subjects s ON s.id=a.subject_id
                LEFT JOIN ep_class_semesters cs
                  ON cs.class_id=a.class_id
                 AND cs.academic_year_id=a.academic_year_id
                 AND cs.semester_no=a.semester_no
                WHERE a.faculty_id=%s
                  AND a.academic_year_id=%s
                ORDER BY d.department_name,c.class_name,a.semester_no,s.subject_name
            """, (faculty["faculty_id"], ay["id"]))
            assignments = cur.fetchall()
            for row in assignments:
                row["assessment_status"] = _assessment_completion_status(cur, faculty, row, ay["id"])
        return render_template(
            "faculty/marks_entry.html",
            faculty=faculty,
            profile_letter=profile_letter,
            academic_year=ay,
            assignments=assignments
        )
    finally:
        cur.close()
        conn.close()
@app.route("/faculty/assignment", methods=["GET","POST"])
@faculty_required
def faculty_assignment():
    faculty, profile_letter=faculty_page_data()
    if not faculty:
        return redirect(url_for("faculty_dashboard"))
    subject_id=request.values.get("subject_id",type=int)
    class_id=request.values.get("class_id",type=int)
    conn=get_db_connection()
    cur=conn.cursor(dictionary=True)
    try:
        ensure_assignment_marks_schema(cur)
        conn.commit()
        ay, sem, semester_options=get_selected_faculty_semester(cur,faculty["faculty_id"])
        if not (ay and sem is not None and subject_id and class_id):
            flash("Select a class and subject from My Classes.","warning")
            return redirect(url_for("faculty_classes"))
        _ensure_faculty_student_semester_mappings(cur, faculty["faculty_id"], ay["id"])
        conn.commit()
        cur.execute("""
            SELECT 1 FROM ep_faculty_assignments
            WHERE faculty_id=%s AND class_id=%s AND subject_id=%s
              AND academic_year_id=%s AND semester_no=%s
        """, (faculty["faculty_id"],class_id,subject_id,ay["id"],sem))
        if not cur.fetchone():
            flash("You cannot access this assignment.","danger")
            return redirect(url_for("faculty_classes"))
        if request.method=="POST":
            form_errors = []
            for key,val in request.form.items():
                if not key.startswith("mark_"):
                    continue
                try:
                    sid=int(key[5:])
                    raw=(val or "").strip()
                    mark=None if raw=="" else float(raw)
                except (TypeError, ValueError):
                    form_errors.append("Enter valid assignment marks between 0 and 5.")
                    continue
                if mark is not None and (mark < 0 or mark > 5):
                    form_errors.append("Assignment marks must be between 0 and 5.")
                    continue
                if mark is None:
                    cur.execute("""
                        DELETE FROM ep_assignment_marks
                        WHERE student_id=%s AND subject_id=%s
                          AND academic_year_id=%s AND semester_no=%s
                    """, (sid,subject_id,ay["id"],sem))
                else:
                    cur.execute("""
                        INSERT INTO ep_assignment_marks
                        (student_id,subject_id,academic_year_id,semester_no,marks)
                        VALUES(%s,%s,%s,%s,%s)
                        ON DUPLICATE KEY UPDATE marks=VALUES(marks)
                    """, (sid,subject_id,ay["id"],sem,mark))
                cur.execute("""
                    SELECT id,ca1,ca2,external
                    FROM marks
                    WHERE student_id=%s AND subject_id=%s
                    ORDER BY id DESC LIMIT 1 FOR UPDATE
                """, (sid,subject_id))
                m=cur.fetchone()
                ca1=m["ca1"] if m else None
                ca2=m["ca2"] if m else None
                external=m["external"] if m else None
                internal=total=grade=result=appreciation=None
                if mark is not None and all(v is not None for v in (ca1,ca2,external)):
                    internal,total,grade,result,appreciation=calculate_marks(
                        float(ca1),float(ca2),float(mark),float(external)
                    )
                if m:
                    cur.execute("""
                        UPDATE marks
                        SET assignment=%s,internal=%s,total=%s,grade=%s,
                            result=%s,appreciation=%s
                        WHERE id=%s
                    """, (mark,internal,total,grade,result,appreciation,m["id"]))
                elif mark is not None:
                    cur.execute("""
                        INSERT INTO marks
                        (student_id,subject_id,ca1,ca2,assignment,internal,
                         external,total,grade,result,appreciation)
                        VALUES(%s,%s,NULL,NULL,%s,%s,NULL,%s,%s,%s,%s)
                    """, (sid,subject_id,mark,internal,total,grade,result,appreciation))
            conn.commit()
            flash("Assignment marks saved successfully.","success")
            return redirect(url_for("faculty_assignment", subject_id=subject_id, class_id=class_id))
        cur.execute("""
            SELECT s.id,s.register_number,s.student_name,
                   am.marks AS marks
            FROM students s
            LEFT JOIN ep_student_semesters es
              ON es.student_id=s.id
             AND es.academic_year_id=%s
             AND es.semester_no=%s
             AND es.class_id=%s
            LEFT JOIN ep_assignment_marks am
              ON am.student_id=s.id
             AND am.subject_id=%s
             AND am.academic_year_id=%s
             AND am.semester_no=%s
            WHERE es.student_id IS NOT NULL OR s.class_id=%s
            ORDER BY s.student_name
        """, (ay["id"],sem,class_id,subject_id,ay["id"],sem,class_id))
        students=cur.fetchall()
        cur.execute("SELECT subject_code,subject_name FROM subjects WHERE id=%s",(subject_id,))
        subject=cur.fetchone()
        cur.execute("SELECT class_name FROM classes WHERE id=%s",(class_id,))
        cls=cur.fetchone()
        return render_template(
            "faculty/assignment_marks.html",
            faculty=faculty,profile_letter=profile_letter,
            students=students,subject=subject,class_info=cls,
            subject_id=subject_id,class_id=class_id,
            academic_year=ay,semester_no=sem
        )
    finally:
        cur.close()
        conn.close()
@app.route("/faculty/profile")
@faculty_required
def faculty_profile():
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cursor)
        subjects = []
        if ay:
            cursor.execute("""
                SELECT DISTINCT
                    s.subject_code,s.subject_name,
                    c.class_name,a.semester_no,a.section
                FROM ep_faculty_assignments a
                INNER JOIN subjects s ON s.id=a.subject_id
                INNER JOIN classes c ON c.id=a.class_id
                WHERE a.faculty_id=%s AND a.academic_year_id=%s
                ORDER BY c.class_name,a.semester_no,s.subject_name
            """, (faculty["faculty_id"], ay["id"]))
            subjects = cursor.fetchall()
        cursor.execute("""
            SELECT gender,dob,address
            FROM ep_faculty_details
            WHERE faculty_id=%s
            LIMIT 1
        """, (faculty["faculty_id"],))
        details = cursor.fetchone() or {}
        cursor.execute("""
            SELECT DISTINCT c.class_name
            FROM ep_faculty_assignments a
            INNER JOIN classes c ON c.id=a.class_id
            WHERE a.faculty_id=%s AND a.academic_year_id=%s
            ORDER BY c.class_name
        """, (faculty["faculty_id"], ay["id"] if ay else 0))
        assigned_classes = cursor.fetchall()
        return render_template(
            "faculty/profile.html",
            faculty=faculty,
            profile_letter=profile_letter,
            subjects=subjects,
            details=details,
            assigned_classes=assigned_classes,
            academic_year=ay
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/students")
@faculty_required
def faculty_students():
    """Show students belonging to classes assigned to the faculty in the active year."""
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_dashboard"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cursor)
        students = []
        if ay:
            _ensure_faculty_student_semester_mappings(cursor, faculty["faculty_id"], ay["id"])
            conn.commit()
            cursor.execute("""
                SELECT DISTINCT
                    s.id,s.user_id,s.register_number,s.student_name,
                    s.email,s.phone,s.department_id,es.class_id,
                    d.department_name,c.class_name,u.username
                FROM ep_student_semesters es
                INNER JOIN students s ON s.id=es.student_id
                INNER JOIN classes c ON c.id=es.class_id
                INNER JOIN departments d ON d.id=c.department_id
                LEFT JOIN users u ON s.user_id=u.id
                INNER JOIN ep_faculty_assignments a
                  ON a.class_id=es.class_id
                 AND a.academic_year_id=es.academic_year_id
                 AND a.semester_no=es.semester_no
                 AND a.faculty_id=%s
                WHERE es.academic_year_id=%s
                ORDER BY c.class_name,s.student_name
            """, (faculty["faculty_id"],ay["id"]))
            students = cursor.fetchall()
        return render_template(
            "faculty/students.html",
            faculty=faculty,
            profile_letter=profile_letter,
            students=students,
            academic_year=ay
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/students/view/<int:student_id>")
@faculty_required
def faculty_view_student(student_id):
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_students"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_student_access(cursor, faculty, student_id):
            student = None
        else:
            cursor.execute(
                """
                SELECT
                    s.id,s.user_id,s.register_number,s.student_name,s.email,s.phone,
                    s.department_id,s.class_id,d.department_name,c.class_name,u.username
                FROM students s
                INNER JOIN departments d ON s.department_id=d.id
                INNER JOIN classes c ON s.class_id=c.id
                LEFT JOIN users u ON s.user_id=u.id
                WHERE s.id=%s
                LIMIT 1
                """,
                (student_id,)
            )
            student = cursor.fetchone()
    finally:
        cursor.close()
        conn.close()
    if not student:
        flash("You cannot access this student.", "danger")
        return redirect(url_for("faculty_students"))
    return render_template(
        "faculty/view_student.html",
        faculty=faculty,
        profile_letter=profile_letter,
        student=student
    )
@app.route("/faculty/students/add", methods=["GET", "POST"])
@faculty_required
def faculty_add_student():
    faculty, profile_letter = faculty_page_data()
    if not faculty or not faculty.get("class_id"):
        flash("You do not have an assigned class.", "danger")
        return redirect(url_for("faculty_dashboard"))
    conn=get_db_connection(); cursor=conn.cursor(dictionary=True)
    try:
        if request.method == "GET":
            register_number=generate_faculty_register_number(cursor, faculty)
            cursor.execute("""SELECT id,register_number,student_name FROM students
                              WHERE department_id=%s AND class_id=%s ORDER BY student_name""",
                           (faculty["department_id"], faculty["class_id"]))
            existing_students=cursor.fetchall()
            return render_template("faculty/add_student.html", faculty=faculty,
                                   profile_letter=profile_letter, register_number=register_number,
                                   students=existing_students)
        student_name=request.form.get("student_name","").strip()
        email=request.form.get("email","").strip()
        phone=request.form.get("phone","").strip()
        dob=request.form.get("dob","").strip()
        if not student_name or not dob:
            flash("Student name and Date of Birth are required.","danger")
            return redirect(url_for("faculty_add_student"))
        try:
            login_password=_student_login_password_from_dob(dob)
        except ValueError as e:
            flash(str(e),"danger")
            return redirect(url_for("faculty_add_student"))
        if phone and (not phone.isdigit() or len(phone)!=10):
            flash("Phone number must contain exactly 10 digits.","danger")
            return redirect(url_for("faculty_add_student"))
        if email:
            cursor.execute("SELECT id FROM students WHERE LOWER(email)=LOWER(%s) LIMIT 1",(email,))
            if cursor.fetchone():
                flash("Email address is already registered.","danger"); return redirect(url_for("faculty_add_student"))
            cursor.execute("SELECT id FROM faculty WHERE LOWER(email)=LOWER(%s) LIMIT 1",(email,))
            if cursor.fetchone():
                flash("Email address is already registered.","danger"); return redirect(url_for("faculty_add_student"))
        if phone:
            cursor.execute("SELECT id FROM students WHERE phone=%s LIMIT 1",(phone,))
            if cursor.fetchone():
                flash("Phone number is already registered.","danger"); return redirect(url_for("faculty_add_student"))
            cursor.execute("SELECT id FROM faculty WHERE phone=%s LIMIT 1",(phone,))
            if cursor.fetchone():
                flash("Phone number is already registered.","danger"); return redirect(url_for("faculty_add_student"))
        register_number=generate_faculty_register_number(cursor, faculty)
        cursor.execute("SELECT id FROM students WHERE register_number=%s LIMIT 1",(register_number,))
        if cursor.fetchone():
            flash("Register number collision. Please try again.","danger"); return redirect(url_for("faculty_add_student"))
        cursor.execute("INSERT INTO users (username,password,role) VALUES (%s,%s,'student')",
                       (register_number,generate_password_hash(login_password)))
        user_id=cursor.lastrowid
        cursor.execute("""INSERT INTO students
                         (user_id,register_number,student_name,email,phone,department_id,class_id)
                         VALUES (%s,%s,%s,%s,%s,%s,%s)""",
                       (user_id,register_number,student_name,email,phone,faculty["department_id"],faculty["class_id"]))
        student_id=cursor.lastrowid
        cursor.execute("INSERT INTO ep_student_details(student_id,gender,dob,address) VALUES(%s,NULL,%s,NULL)",(student_id,dob))
        conn.commit()
        flash(f"Student added successfully. Register Number: {register_number}","success")
        return redirect(url_for("faculty_students"))
    except mysql.connector.Error as e:
        conn.rollback(); flash(f"Unable to add student: {e}","danger"); return redirect(url_for("faculty_add_student"))
    except Exception as e:
        conn.rollback(); flash(f"Unable to add student: {e}","danger"); return redirect(url_for("faculty_add_student"))
    finally:
        cursor.close(); conn.close()
@app.route("/faculty/students/edit/<int:student_id>", methods=["GET", "POST"])
@faculty_required
def faculty_edit_student(student_id):
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_students"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_student_access(cursor, faculty, student_id):
            student = None
        else:
            cursor.execute(
                """
                SELECT
                    s.id, s.user_id, s.register_number, s.student_name,
                    s.email, s.phone, s.department_id, s.class_id, u.username, sd.dob
                FROM students s
                LEFT JOIN users u ON s.user_id = u.id
                LEFT JOIN ep_student_details sd ON sd.student_id = s.id
                WHERE s.id = %s
                LIMIT 1
                """,
                (student_id,)
            )
            student = cursor.fetchone()
        if not student:
            flash("You cannot edit this student.", "danger")
            return redirect(url_for("faculty_students"))
        if request.method == "GET":
            return render_template(
                "faculty/edit_student.html",
                faculty=faculty,
                profile_letter=profile_letter,
                student=student
            )
        student_name = request.form.get("student_name", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        dob = request.form.get("dob", "").strip()
        if not student_name or not dob:
            flash("Student name and Date of Birth are required.", "danger")
            return redirect(url_for("faculty_edit_student", student_id=student_id))
        try:
            login_password=_student_login_password_from_dob(dob)
        except ValueError as e:
            flash(str(e), "danger")
            return redirect(url_for("faculty_edit_student", student_id=student_id))
        if not student_name:
            flash("Student name is required.", "danger")
            return redirect(url_for("faculty_edit_student", student_id=student_id))
        if phone and (not phone.isdigit() or len(phone) != 10):
            flash("Phone number must contain exactly 10 digits.", "danger")
            return redirect(url_for("faculty_edit_student", student_id=student_id))
        cursor.execute(
            """
            UPDATE students
            SET student_name=%s, email=%s, phone=%s
            WHERE id=%s
            """,
            (student_name, email, phone, student_id)
        )
        cursor.execute("UPDATE users SET username=%s, password=%s WHERE id=%s",
                       (student["register_number"], generate_password_hash(login_password), student["user_id"]))
        cursor.execute("""INSERT INTO ep_student_details(student_id,gender,dob,address)
                         VALUES(%s,NULL,%s,NULL)
                         ON DUPLICATE KEY UPDATE dob=VALUES(dob)""", (student_id,dob))
        conn.commit()
        flash("Student updated successfully.", "success")
        return redirect(url_for("faculty_students"))
    except Exception as e:
        conn.rollback()
        flash(f"Unable to update student: {str(e)}", "danger")
        return redirect(url_for("faculty_students"))
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/students/delete/<int:student_id>", methods=["POST"])
@app.route("/faculty/delete-student/<int:student_id>", methods=["POST"])
@faculty_required
def faculty_delete_student(student_id):
    faculty = get_faculty()
    if not faculty:
        return redirect(url_for("faculty_students"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_student_access(cursor, faculty, student_id):
            flash("You cannot delete this student.", "danger")
            return redirect(url_for("faculty_students"))
        cursor.execute("SELECT user_id FROM students WHERE id = %s LIMIT 1", (student_id,))
        row = cursor.fetchone()
        cursor.execute("DELETE FROM ep_question_marks WHERE student_id = %s", (student_id,))
        cursor.execute("DELETE FROM ep_assignment_marks WHERE student_id = %s", (student_id,))
        cursor.execute("DELETE FROM marks WHERE student_id = %s", (student_id,))
        cursor.execute("DELETE FROM ep_student_semesters WHERE student_id = %s", (student_id,))
        cursor.execute("DELETE FROM ep_student_details WHERE student_id = %s", (student_id,))
        cursor.execute("DELETE FROM students WHERE id = %s", (student_id,))
        if row and row.get("user_id"):
            cursor.execute("DELETE FROM users WHERE id = %s", (row["user_id"],))
        conn.commit()
        flash("Student deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Unable to delete student: {str(e)}", "danger")
    finally:
        cursor.close()
        conn.close()
    return redirect(url_for("faculty_students"))
def _faculty_subject_context():
    """Return faculty subjects for the active year and the selected class/semester.
    Semester choices are class-specific, preventing a semester from another
    assigned class from producing an empty subject/student list.
    """
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        return faculty, profile_letter, [], None, None, None, None
    selected_subject_id = request.args.get("subject_id", type=int)
    selected_class_id = request.args.get("class_id", type=int)
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        ay = get_active_academic_year(cursor)
        if not ay:
            return faculty, profile_letter, [], None, None, None, None
        cursor.execute("""
            SELECT DISTINCT class_id
            FROM ep_faculty_assignments
            WHERE faculty_id=%s AND academic_year_id=%s
            ORDER BY class_id
        """, (faculty["faculty_id"], ay["id"]))
        valid_classes = [int(r["class_id"]) for r in cursor.fetchall()]
        if selected_class_id not in valid_classes:
            selected_class_id = valid_classes[0] if len(valid_classes) == 1 else None
        q = """
            SELECT DISTINCT a.semester_no, cs.year_no, cs.course_type,
                   COALESCE(NULLIF(a.section,''), cs.section, 'A') AS section
            FROM ep_faculty_assignments a
            LEFT JOIN ep_class_semesters cs
              ON cs.class_id=a.class_id
             AND cs.academic_year_id=a.academic_year_id
             AND cs.semester_no=a.semester_no
            WHERE a.faculty_id=%s AND a.academic_year_id=%s
        """
        params=[faculty["faculty_id"], ay["id"]]
        if selected_class_id:
            q += " AND a.class_id=%s"
            params.append(selected_class_id)
        q += " ORDER BY a.semester_no"
        cursor.execute(q, tuple(params))
        available = cursor.fetchall()
        valid_semesters = [int(x["semester_no"]) for x in available]
        sem = request.args.get("semester", type=int)
        if sem is None:
            sem = request.form.get("semester", type=int)
        if sem not in valid_semesters:
            sem = valid_semesters[0] if valid_semesters else None
        subjects=[]
        if sem is not None:
            query = """
                SELECT a.subject_id,s.subject_code,s.subject_name,
                       a.class_id,c.class_name,a.section,
                       a.semester_no,a.academic_year_id
                FROM ep_faculty_assignments a
                INNER JOIN subjects s ON s.id=a.subject_id
                INNER JOIN classes c ON c.id=a.class_id
                WHERE a.faculty_id=%s
                  AND a.academic_year_id=%s
                  AND a.semester_no=%s
            """
            params=[faculty["faculty_id"], ay["id"], sem]
            if selected_class_id:
                query += " AND a.class_id=%s"
                params.append(selected_class_id)
            query += " ORDER BY c.class_name,s.subject_name"
            cursor.execute(query, tuple(params))
            subjects=cursor.fetchall()
        subject=None
        if selected_subject_id:
            subject=next((x for x in subjects if int(x["subject_id"])==int(selected_subject_id)), None)
            if subject is None:
                selected_subject_id=None
        return faculty,profile_letter,subjects,subject,selected_subject_id,ay,sem
    finally:
        cursor.close()
        conn.close()
def _get_mark_students(cursor, faculty, subject_id, semester_no=None, academic_year_id=None, class_id=None):
    """Fetch exactly the students mapped to the selected class/year/semester."""
    if academic_year_id and semester_no is not None:
        query = """
            SELECT DISTINCT s.id AS student_id,s.register_number,s.student_name,
                   m.id AS mark_id,m.ca1,m.ca2,m.assignment,m.internal,
                   m.external,m.total,m.grade,m.result,m.appreciation
            FROM ep_student_semesters es
            INNER JOIN students s ON s.id=es.student_id
            INNER JOIN ep_faculty_assignments a
              ON a.class_id=es.class_id
             AND a.academic_year_id=es.academic_year_id
             AND a.semester_no=es.semester_no
             AND a.faculty_id=%s
             AND a.subject_id=%s
            LEFT JOIN (
                SELECT m1.* FROM marks m1
                INNER JOIN (
                    SELECT student_id,subject_id,MAX(id) AS id
                    FROM marks GROUP BY student_id,subject_id
                ) lm ON lm.id=m1.id
            ) m ON m.student_id=s.id AND m.subject_id=%s
            WHERE es.academic_year_id=%s AND es.semester_no=%s
        """
        params=[faculty["faculty_id"], subject_id, subject_id, academic_year_id, semester_no]
        if class_id:
            query += " AND es.class_id=%s"
            params.append(class_id)
        query += " ORDER BY s.student_name ASC"
        cursor.execute(query, tuple(params))
    else:
        cursor.execute("""SELECT s.id AS student_id,s.register_number,s.student_name,
                      m.id AS mark_id,m.ca1,m.ca2,m.assignment,m.internal,
                      m.external,m.total,m.grade,m.result,m.appreciation
               FROM students s LEFT JOIN marks m ON m.student_id=s.id AND m.subject_id=%s
               WHERE s.class_id=%s ORDER BY s.student_name ASC""", (subject_id, faculty.get("class_id")))
    return cursor.fetchall()
def _save_partial_mark(student_id, subject_id, field, value):
    """Save only CIA1, CIA2, Assignment or External without requiring all marks."""
    allowed = {"ca1", "ca2", "assignment", "external"}
    if field not in allowed:
        raise ValueError("Invalid mark field.")
    faculty = get_faculty()
    if not faculty:
        raise PermissionError("Faculty profile not found.")
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError("Enter a valid mark.")
    limits = {
        "ca1": 75,
        "ca2": 75,
        "assignment": 5,
        "external": 75
    }
    if value < 0 or value > limits[field]:
        raise ValueError(
            f"{field.upper()} must be between 0 and {limits[field]}."
        )
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_mark_access(cursor, faculty, student_id, subject_id):
            raise PermissionError("You cannot edit these marks.")
        cursor.execute(
            """
            SELECT id, ca1, ca2, assignment, external
            FROM marks
            WHERE student_id = %s
              AND subject_id = %s
            ORDER BY id DESC
            LIMIT 1
            FOR UPDATE
            """,
            (student_id, subject_id)
        )
        mark = cursor.fetchone()
        values = {
            "ca1": mark["ca1"] if mark else None,
            "ca2": mark["ca2"] if mark else None,
            "assignment": mark["assignment"] if mark else None,
            "external": mark["external"] if mark else None
        }
        values[field] = value
        internal = total = grade = result = appreciation = None
        if all(values[x] is not None for x in ("ca1", "ca2", "assignment", "external")):
            internal, total, grade, result, appreciation = calculate_marks(
                float(values["ca1"]),
                float(values["ca2"]),
                float(values["assignment"]),
                float(values["external"])
            )
        if mark:
            cursor.execute(
                """
                UPDATE marks
                SET ca1=%s, ca2=%s, assignment=%s,
                    internal=%s, external=%s, total=%s,
                    grade=%s, result=%s, appreciation=%s
                WHERE id=%s
                """,
                (
                    values["ca1"], values["ca2"], values["assignment"],
                    internal, values["external"], total,
                    grade, result, appreciation, mark["id"]
                )
            )
        else:
            cursor.execute(
                """
                INSERT INTO marks
                (student_id, subject_id, ca1, ca2, assignment, internal,
                 external, total, grade, result, appreciation)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (
                    student_id, subject_id,
                    values["ca1"], values["ca2"], values["assignment"],
                    internal, values["external"], total,
                    grade, result, appreciation
                )
            )
        conn.commit()
        return {
            "ca1": values["ca1"],
            "ca2": values["ca2"],
            "assignment": values["assignment"],
            "external": values["external"],
            "ca1_converted": round(float(values["ca1"]) / 75 * 10, 2)
                if values["ca1"] is not None else None,
            "ca2_converted": round(float(values["ca2"]) / 75 * 10, 2)
                if values["ca2"] is not None else None,
            "internal": internal,
            "total": total,
            "grade": grade,
            "result": result
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
def _cia_detail_json_path(faculty_id, subject_id, cia, academic_year_id=None, semester_no=None):
    """Path for question-wise CIA marks, isolated by academic year + semester."""
    return os.path.join(
        QUESTION_PAPER_DIR,
        _question_paper_file_key(faculty_id, subject_id, cia, academic_year_id, semester_no) + "_marks.json"
    )
def _load_cia_detail_marks(faculty_id, subject_id, cia, academic_year_id=None, semester_no=None):
    path = _cia_detail_json_path(faculty_id, subject_id, cia, academic_year_id, semester_no)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}
def _save_cia_detail_marks(faculty_id, subject_id, cia, data, academic_year_id=None, semester_no=None):
    path = _cia_detail_json_path(faculty_id, subject_id, cia, academic_year_id, semester_no)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
def _cia_total_from_questions(detail):
    """Return raw CIA total /75 from 10×2 + 5×5 + exactly 3 of 5×10."""
    if not isinstance(detail, dict):
        raise ValueError("Invalid CIA mark data.")
    section_a = detail.get("section_a", [])
    section_b = detail.get("section_b", [])
    section_c = detail.get("section_c", [])
    if len(section_a) != 10 or len(section_b) != 5 or len(section_c) != 5:
        raise ValueError("CIA must contain 10 two-mark, 5 five-mark A/B choice pairs and 5 ten-mark questions.")
    def values(items, maximum):
        out = []
        for x in items:
            if x in (None, ""):
                out.append(None)
                continue
            try:
                n = float(x)
            except (TypeError, ValueError):
                raise ValueError("Enter valid question-wise marks.")
            if n < 0 or n > maximum:
                raise ValueError(f"Each question mark must be between 0 and {maximum}.")
            out.append(round(n, 2))
        return out
    a = values(section_a, 2)
    b_total = 0.0
    normalized_b = []
    for pair in section_b:
        if isinstance(pair, dict):
            choice = str(pair.get("choice", "")).upper().strip()
            raw_mark = pair.get("mark")
        else:
            choice = "A"
            raw_mark = pair
        if choice not in ("A", "B"):
            raise ValueError("For every Section B question, select either A or B.")
        if raw_mark in (None, ""):
            raise ValueError("Enter one mark for every Section B A/B choice.")
        try:
            n = float(raw_mark)
        except (TypeError, ValueError):
            raise ValueError("Enter valid Section B marks.")
        if n < 0 or n > 5:
            raise ValueError("Each Section B answer has a maximum of 5 marks.")
        normalized_b.append({"choice": choice, "mark": round(n, 2)})
        b_total += n
    c = values(section_c, 10)
    if any(x is None for x in a):
        raise ValueError("Enter all Section A marks.")
    if sum(x is not None for x in c) != 3:
        raise ValueError("Section C requires exactly any 3 of the 5 questions.")
    return round(sum(a) + b_total + sum(x for x in c if x is not None), 2)
def _extract_question_schema(paper):
    """Normalize saved OCR question-paper sections into question entry rows."""
    if not isinstance(paper,dict): return []
    out=[]
    section_names=["A","B","C","D","E","F"]
    for si,section in enumerate(paper.get("sections",[]) or []):
        questions=section.get("questions",[]) if isinstance(section,dict) else []
        for qi,q in enumerate(questions):
            if not isinstance(q,dict): continue
            text=str(q.get("question") or "").strip()
            number=str(q.get("number") or (qi+1))
            raw=q.get("marks",q.get("max_marks",q.get("mark")))
            max_marks=None
            if raw not in (None,""):
                try: max_marks=float(raw)
                except (TypeError,ValueError): max_marks=None
            if max_marks is None:
                m=re.search(r"(?:max(?:imum)?\s*)?(?:mark|marks)\s*[:\-]?\s*(\d+(?:\.\d+)?)",text,re.I)
                if m: max_marks=float(m.group(1))
            if max_marks is None:
                max_marks={0:2,1:5,2:10}.get(si,0)
            out.append({"number":number,"question":text,"max_marks":max_marks,"section":section_names[si] if si<len(section_names) else str(si+1),"co":q.get("co","") ,"k":q.get("k","")})
    if not out:
        for i in range(1,11): out.append({"number":str(i),"question":"","max_marks":2,"section":"A","co":"","k":""})
        for i in range(11,16): out.append({"number":str(i),"question":"","max_marks":5,"section":"B","co":"","k":""})
        for i in range(16,21): out.append({"number":str(i),"question":"","max_marks":10,"section":"C","co":"","k":""})
    return out
def _load_saved_question_paper(faculty_id,subject_id,cia,academic_year_id,semester_no):
    path=_question_paper_json_path(faculty_id,subject_id,cia,academic_year_id,semester_no)
    if os.path.exists(path):
        try:
            with open(path,"r",encoding="utf-8") as f: return json.load(f)
        except (OSError,json.JSONDecodeError): pass
    return None
def _render_assessment_page(template_name, field, title, subtitle):
    faculty, profile_letter, subjects, subject, selected_subject_id, academic_year, semester_no = _faculty_subject_context()
    selected_class_id = request.args.get("class_id", type=int)
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_dashboard"))
    if selected_class_id is None and subject and subject.get("class_id"):
        selected_class_id = int(subject["class_id"])
    students = []
    detail_marks = {}
    saved_detail_students = set()
    external_details = {}
    available = []
    if faculty and academic_year:
        conn2=get_db_connection(); cur2=conn2.cursor(dictionary=True)
        try:
            q="""SELECT DISTINCT a.semester_no,cs.year_no,cs.course_type,
                       COALESCE(NULLIF(a.section,''),cs.section,'A') AS section
                 FROM ep_faculty_assignments a
                 LEFT JOIN ep_class_semesters cs
                   ON cs.class_id=a.class_id AND cs.academic_year_id=a.academic_year_id
                  AND cs.semester_no=a.semester_no
                 WHERE a.faculty_id=%s AND a.academic_year_id=%s"""
            params=[faculty["faculty_id"],academic_year["id"]]
            if selected_class_id:
                q += " AND a.class_id=%s"
                params.append(selected_class_id)
            q += " ORDER BY a.semester_no"
            cur2.execute(q,tuple(params)); available=cur2.fetchall()
        finally: cur2.close(); conn2.close()
    if selected_subject_id:
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        try:
            if academic_year:
                _ensure_faculty_student_semester_mappings(cursor, faculty["faculty_id"], academic_year["id"])
                conn.commit()
            students = _get_mark_students(cursor, faculty, selected_subject_id, semester_no, academic_year["id"] if academic_year else None, selected_class_id)
        finally:
            cursor.close()
            conn.close()
        if field in ("ca1", "ca2"):
            cia = "cia1" if field == "ca1" else "cia2"
            detail_marks = _load_cia_detail_marks(faculty["faculty_id"], selected_subject_id, cia, academic_year["id"] if academic_year else None, semester_no)
            for student_id, detail in detail_marks.items():
                try:
                    _cia_total_from_questions(detail)
                    saved_detail_students.add(str(student_id))
                except (TypeError, ValueError):
                    pass
        paper_type = "cia1" if field == "ca1" else ("cia2" if field == "ca2" else "external")
        saved_paper = _load_saved_question_paper(faculty["faculty_id"], selected_subject_id, paper_type, academic_year["id"] if academic_year else None, semester_no)
        if field == "external" and academic_year and semester_no is not None:
            conn3=get_db_connection(); cur3=conn3.cursor(dictionary=True)
            try:
                cur3.execute("SELECT student_id,question_no,question_text,max_marks,obtained_marks,choice FROM ep_question_marks WHERE subject_id=%s AND academic_year_id=%s AND semester_no=%s AND exam_type='external'",(selected_subject_id,academic_year["id"],semester_no))
                for q in cur3.fetchall():
                    external_details.setdefault(str(q["student_id"]),{})[str(q["question_no"])] = q
            finally:
                cur3.close(); conn3.close()
    else:
        saved_paper = None
    question_schema = _extract_question_schema(saved_paper)
    return render_template(
        template_name,
        faculty=faculty,
        profile_letter=profile_letter,
        subjects=subjects,
        students=students,
        subject=subject,
        selected_subject_id=selected_subject_id,
        mark_field=field,
        page_title=title,
        page_subtitle=subtitle,
        detail_marks=detail_marks,
        saved_detail_students=saved_detail_students,
        academic_year=academic_year,
        semester_no=semester_no,
        semester_options=available, question_schema=question_schema, saved_paper=saved_paper, external_details=external_details, selected_class_id=selected_class_id
    )
@app.route("/faculty/cia1")
@faculty_required
def faculty_cia1():
    return _render_assessment_page(
        "faculty/cia_marks.html",
        "ca1",
        "CIA 1 Mark Entry",
        "Enter CIA 1 marks out of 75. The system automatically converts them to 10 marks."
    )
@app.route("/faculty/cia2")
@faculty_required
def faculty_cia2():
    return _render_assessment_page(
        "faculty/cia_marks.html",
        "ca2",
        "CIA 2 Mark Entry",
        "Enter CIA 2 marks out of 75. The system automatically converts them to 10 marks."
    )
@app.route("/faculty/external")
@faculty_required
def faculty_external():
    return _render_assessment_page(
        "faculty/external_marks.html",
        "external",
        "External Mark Entry",
        "Enter External marks question-wise. CIA 1, CIA 2 and Assignment are fetched/calculated automatically."
    )
@app.route("/faculty/marks/save-partial", methods=["POST"])
@faculty_required
def faculty_save_partial_mark():
    try:
        student_id = int(request.form.get("student_id", "0"))
        subject_id = int(request.form.get("subject_id", "0"))
        field = request.form.get("field", "").strip()
        value = request.form.get("value", "").strip()
        result = _save_partial_mark(student_id, subject_id, field, value)
        return jsonify({"success": True, "message": "Mark saved successfully.", **result})
    except PermissionError as e:
        return jsonify({"success": False, "message": str(e)}), 403
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except mysql.connector.Error as e:
        return jsonify({"success": False, "message": "Database error: " + str(e)}), 500
    except Exception as e:
        return jsonify({"success": False, "message": "Unable to save mark: " + str(e)}), 500
@app.route("/faculty/external/save-question-marks", methods=["POST"])
@faculty_required
def faculty_save_external_question_marks():
    """Save External using the same 75-mark A/B/C pattern as CIA."""
    try:
        student_id = int(request.form.get("student_id", "0"))
        subject_id = int(request.form.get("subject_id", "0"))
        payload = json.loads(request.form.get("question_marks", "{}"))
        total = _cia_total_from_questions(payload)
        faculty = get_faculty()
        conn = get_db_connection()
        cur = conn.cursor(dictionary=True)
        try:
            if not verify_mark_access(cur, faculty, student_id, subject_id):
                raise PermissionError("You cannot edit these marks.")
            ay, sem, _ = get_selected_faculty_semester(cur, faculty["faculty_id"])
            if not ay or sem is None:
                raise ValueError("No active academic-year semester is available.")
            section_a = payload["section_a"]
            section_b = payload["section_b"]
            section_c = payload["section_c"]
            for i, mark in enumerate(section_a, 1):
                cur.execute("""
                    INSERT INTO ep_question_marks
                    (student_id,subject_id,academic_year_id,semester_no,exam_type,
                     question_no,question_text,max_marks,obtained_marks,choice)
                    VALUES(%s,%s,%s,%s,'external',%s,%s,2,%s,NULL)
                    ON DUPLICATE KEY UPDATE obtained_marks=VALUES(obtained_marks),
                    max_marks=VALUES(max_marks),choice=NULL
                """, (student_id,subject_id,ay["id"],sem,str(i),"Section A",float(mark)))
            for i, pair in enumerate(section_b, 1):
                choice = str(pair.get("choice","")).upper()
                mark = float(pair.get("mark"))
                qno = str(10+i)
                cur.execute("""
                    INSERT INTO ep_question_marks
                    (student_id,subject_id,academic_year_id,semester_no,exam_type,
                     question_no,question_text,max_marks,obtained_marks,choice)
                    VALUES(%s,%s,%s,%s,'external',%s,%s,5,%s,%s)
                    ON DUPLICATE KEY UPDATE obtained_marks=VALUES(obtained_marks),
                    max_marks=VALUES(max_marks),choice=VALUES(choice)
                """, (student_id,subject_id,ay["id"],sem,qno,"Section B",mark,choice))
            for i, mark in enumerate(section_c, 1):
                qno = str(15+i)
                value = None if mark in (None,"") else float(mark)
                cur.execute("""
                    INSERT INTO ep_question_marks
                    (student_id,subject_id,academic_year_id,semester_no,exam_type,
                     question_no,question_text,max_marks,obtained_marks,choice)
                    VALUES(%s,%s,%s,%s,'external',%s,%s,10,%s,NULL)
                    ON DUPLICATE KEY UPDATE obtained_marks=VALUES(obtained_marks),
                    max_marks=VALUES(max_marks),choice=NULL
                """, (student_id,subject_id,ay["id"],sem,qno,"Section C",value))
            cur.execute("""
                SELECT id,ca1,ca2,assignment
                FROM marks
                WHERE student_id=%s AND subject_id=%s
                ORDER BY id DESC LIMIT 1 FOR UPDATE
            """, (student_id,subject_id))
            m = cur.fetchone()
            ca1 = m["ca1"] if m else None
            ca2 = m["ca2"] if m else None
            assignment = m["assignment"] if m else None
            internal = total_final = grade = result = appreciation = None
            if all(v is not None for v in (ca1,ca2,assignment)):
                internal = calculate_internal(float(ca1),float(ca2),float(assignment))
                total_final = round(internal + total, 2)
                if total_final >= 90:
                    grade,appreciation="A+","Excellent"
                elif total_final >= 80:
                    grade,appreciation="A","Very Good"
                elif total_final >= 70:
                    grade,appreciation="B+","Good"
                elif total_final >= 60:
                    grade,appreciation="B","Satisfactory"
                elif total_final >= 30:
                    grade,appreciation="C","Needs Improvement"
                else:
                    grade,appreciation="F","Fail"
                result = "PASS" if total_final >= 30 else "FAIL"
            if m:
                cur.execute("""
                    UPDATE marks SET external=%s,internal=%s,total=%s,
                    grade=%s,result=%s,appreciation=%s WHERE id=%s
                """, (total,internal,total_final,grade,result,appreciation,m["id"]))
            else:
                cur.execute("""
                    INSERT INTO marks
                    (student_id,subject_id,ca1,ca2,assignment,internal,external,
                     total,grade,result,appreciation)
                    VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (student_id,subject_id,ca1,ca2,assignment,internal,total,
                      total_final,grade,result,appreciation))
            conn.commit()
            return jsonify({
                "success":True,
                "message":"External marks saved successfully.",
                "external":round(total,2),
                "internal":internal,
                "total":total_final,
                "grade":grade,
                "result":result
            })
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()
    except PermissionError as e:
        return jsonify({"success":False,"message":str(e)}),403
    except (ValueError,TypeError,json.JSONDecodeError) as e:
        return jsonify({"success":False,"message":str(e)}),400
    except Exception as e:
        return jsonify({"success":False,"message":"Unable to save external marks: "+str(e)}),500
@app.route("/faculty/marks")
@faculty_required
def faculty_marks():
    return redirect(url_for("faculty_cia1", subject_id=request.args.get("subject_id", type=int)))
@app.route("/faculty/marks/submit", methods=["POST"])
@faculty_required
def faculty_marks_submit():
    """Legacy complete-row save retained for existing View Marks edit buttons."""
    faculty = get_faculty()
    if not faculty:
        return jsonify({"success": False, "message": "Faculty profile not found."}), 403
    try:
        student_id = int(request.form.get("student_id", "0"))
        subject_id = int(request.form.get("subject_id", "0"))
        posted_class_id = request.form.get("class_id", type=int)
        ca1 = float(request.form.get("ca1", ""))
        ca2 = float(request.form.get("ca2", ""))
        assignment = float(request.form.get("assignment", ""))
        external = float(request.form.get("external", ""))
        internal, total, grade, result, appreciation = calculate_marks(
            ca1, ca2, assignment, external
        )
    except (TypeError, ValueError) as e:
        return jsonify({"success": False, "message": str(e) or "Enter all marks correctly."}), 400
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_mark_access(cursor, faculty, student_id, subject_id):
            return jsonify({"success": False, "message": "You cannot submit these marks."}), 403
        if posted_class_id is not None:
            ay_ctx, sem_ctx, _ = get_selected_faculty_semester(cursor, faculty["faculty_id"])
            cursor.execute("""
                SELECT 1 FROM ep_faculty_assignments
                WHERE faculty_id=%s AND class_id=%s AND subject_id=%s
                  AND academic_year_id=%s AND semester_no=%s
                LIMIT 1
            """, (faculty["faculty_id"],posted_class_id,subject_id,
                  ay_ctx["id"] if ay_ctx else 0,sem_ctx or 0))
            if not cursor.fetchone():
                return jsonify({"success": False, "message": "Class/subject is not assigned to you for the active academic year."}), 403
        cursor.execute(
            """
            SELECT id FROM marks
            WHERE student_id=%s AND subject_id=%s
            ORDER BY id DESC LIMIT 1 FOR UPDATE
            """,
            (student_id, subject_id)
        )
        mark = cursor.fetchone()
        if mark:
            cursor.execute(
                """
                UPDATE marks
                SET ca1=%s, ca2=%s, assignment=%s, internal=%s,
                    external=%s, total=%s, grade=%s, result=%s, appreciation=%s
                WHERE id=%s
                """,
                (ca1, ca2, assignment, internal, external, total,
                 grade, result, appreciation, mark["id"])
            )
        else:
            cursor.execute(
                """
                INSERT INTO marks
                (student_id, subject_id, ca1, ca2, assignment, internal,
                 external, total, grade, result, appreciation)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (student_id, subject_id, ca1, ca2, assignment, internal,
                 external, total, grade, result, appreciation)
            )
        conn.commit()
        return jsonify({
            "success": True,
            "message": "Marks saved successfully.",
            "ca1": ca1,
            "ca2": ca2,
            "assignment": assignment,
            "internal": internal,
            "external": external,
            "total": total,
            "percentage": total,
            "grade": grade,
            "result": result,
            "appreciation": appreciation
        })
    except mysql.connector.Error as e:
        conn.rollback()
        return jsonify({"success": False, "message": "Database error: " + str(e)}), 500
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/marks/edit", methods=["POST"])
@faculty_required
def faculty_marks_edit():
    faculty = get_faculty()
    if not faculty:
        return jsonify({"success": False, "message": "Faculty profile not found."}), 403
    try:
        student_id = int(request.form.get("student_id", "0"))
        subject_id = int(request.form.get("subject_id", "0"))
    except (TypeError, ValueError):
        return jsonify({"success": False, "message": "Invalid request."}), 400
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not verify_mark_access(cursor, faculty, student_id, subject_id):
            return jsonify({"success": False, "message": "You cannot edit these marks."}), 403
        return jsonify({"success": True, "message": "Marks are ready for editing."})
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/cia/save-question-marks", methods=["POST"])
@faculty_required
def faculty_save_cia_question_marks():
    """Save question-wise CIA marks and sync the /75 total to marks.ca1/ca2."""
    try:
        student_id = int(request.form.get("student_id", "0"))
        subject_id = int(request.form.get("subject_id", "0"))
        cia = request.form.get("cia", "").lower()
        if cia not in ("cia1", "cia2", "external"):
            raise ValueError("Invalid examination selection.")
        payload = request.form.get("question_marks", "")
        detail = json.loads(payload)
        total_raw = _cia_total_from_questions(detail)
        normalized_b = []
        for pair in detail.get("section_b", []):
            if isinstance(pair, dict):
                normalized_b.append({"choice": str(pair.get("choice", "A")).upper(), "mark": round(float(pair.get("mark", 0)), 2)})
            else:
                normalized_b.append({"choice": "A", "mark": round(float(pair), 2)})
        detail["section_b"] = normalized_b
        faculty = get_faculty()
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        try:
            if not verify_mark_access(cursor, faculty, student_id, subject_id):
                raise PermissionError("You cannot edit these marks.")
            ay_ctx, sem_ctx, _ = get_selected_faculty_semester(cursor, faculty["faculty_id"])
            all_details = _load_cia_detail_marks(
                faculty["faculty_id"], subject_id, cia, ay_ctx["id"] if ay_ctx else None, sem_ctx
            )
            all_details[str(student_id)] = detail
            _save_cia_detail_marks(
                faculty["faculty_id"], subject_id, cia, all_details, ay_ctx["id"] if ay_ctx else None, sem_ctx
            )
            if ay_ctx and sem_ctx is not None:
                qrows=[]
                for idx,val in enumerate(detail.get("section_a",[]),1): qrows.append((str(idx),2,val,None))
                for idx,pair in enumerate(detail.get("section_b",[]),11):
                    if isinstance(pair,dict): qrows.append((str(idx),5,pair.get("mark"),pair.get("choice")))
                    else: qrows.append((str(idx),5,pair,None))
                for idx,val in enumerate(detail.get("section_c",[]),16): qrows.append((str(idx),10,val,None))
                for qno,qmax,qmark,qchoice in qrows:
                    if qmark in (None,""): continue
                    cursor.execute("""INSERT INTO ep_question_marks(student_id,subject_id,academic_year_id,semester_no,exam_type,question_no,question_text,max_marks,obtained_marks,choice)
                        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        ON DUPLICATE KEY UPDATE max_marks=VALUES(max_marks),obtained_marks=VALUES(obtained_marks),choice=VALUES(choice)
                    """,(student_id,subject_id,ay_ctx["id"],sem_ctx,cia,qno,None,qmax,float(qmark),qchoice))
            field = "ca1" if cia == "cia1" else "ca2"
            cursor.execute(
                """
                SELECT id, ca1, ca2, assignment, external
                FROM marks
                WHERE student_id=%s AND subject_id=%s
                ORDER BY id DESC LIMIT 1 FOR UPDATE
                """,
                (student_id, subject_id)
            )
            mark = cursor.fetchone()
            ca1 = total_raw if field == "ca1" else (mark["ca1"] if mark else None)
            ca2 = total_raw if field == "ca2" else (mark["ca2"] if mark else None)
            assignment = mark["assignment"] if mark else None
            external = mark["external"] if mark else None
            internal = total = grade = result = appreciation = None
            if all(v is not None for v in (ca1, ca2, assignment, external)):
                internal, total, grade, result, appreciation = calculate_marks(
                    float(ca1), float(ca2), float(assignment), float(external)
                )
            if mark:
                cursor.execute(
                    """
                    UPDATE marks
                    SET ca1=%s, ca2=%s, assignment=%s, internal=%s,
                        external=%s, total=%s, grade=%s, result=%s, appreciation=%s
                    WHERE id=%s
                    """,
                    (ca1, ca2, assignment, internal, external, total,
                     grade, result, appreciation, mark["id"])
                )
            else:
                cursor.execute(
                    """
                    INSERT INTO marks
                    (student_id, subject_id, ca1, ca2, assignment, internal,
                     external, total, grade, result, appreciation)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """,
                    (student_id, subject_id, ca1, ca2, assignment, internal,
                     external, total, grade, result, appreciation)
                )
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cursor.close()
            conn.close()
        return jsonify({
            "success": True,
            "message": f"{cia.upper()} question-wise marks saved.",
            "raw": total_raw,
            "converted": round(total_raw / 75 * 10, 2)
        })
    except PermissionError as e:
        return jsonify({"success": False, "message": str(e)}), 403
    except (ValueError, TypeError, json.JSONDecodeError) as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except mysql.connector.Error as e:
        return jsonify({"success": False, "message": "Database error: " + str(e)}), 500
    except Exception as e:
        return jsonify({"success": False, "message": "Unable to save CIA marks: " + str(e)}), 500
QUESTION_PAPER_DIR = os.path.join(
    app.root_path, "static", "uploads", "question_papers"
)
os.makedirs(QUESTION_PAPER_DIR, exist_ok=True)
def _question_paper_file_key(faculty_id, subject_id, cia, academic_year_id=None, semester_no=None):
    suffix = f"_ay_{academic_year_id}_sem_{semester_no}" if academic_year_id and semester_no else ""
    return f"faculty_{faculty_id}_subject_{subject_id}_{cia}{suffix}"
def _question_paper_json_path(faculty_id, subject_id, cia, academic_year_id=None, semester_no=None):
    return os.path.join(
        QUESTION_PAPER_DIR,
        _question_paper_file_key(faculty_id, subject_id, cia, academic_year_id, semester_no) + ".json"
    )
def _question_paper_access(cursor, faculty, subject_id):
    if not faculty:
        return False
    return verify_subject_access(cursor, faculty, subject_id)
def _normalize_question_paper_payload(payload):
    """Return saved QP data in one canonical 3-section shape for view/edit.
    This accepts current saves, older saves, nested paper_data and JSON-string
    variants so the View Extracted QP page never renders blank merely because
    a previous version used a slightly different payload structure.
    """
    if payload is None:
        return None
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except Exception:
            return None
    if not isinstance(payload, dict):
        return None
    for key in ("paper_data", "data", "payload"):
        nested = payload.get(key)
        if isinstance(nested, str):
            try: nested = json.loads(nested)
            except Exception: nested = None
        if isinstance(nested, dict) and (nested.get("sections") or nested.get("questions")):
            merged = dict(payload)
            merged.update(nested)
            payload = merged
            break
    raw_sections = payload.get("sections")
    if isinstance(raw_sections, str):
        try: raw_sections = json.loads(raw_sections)
        except Exception: raw_sections = None
    canonical = [{"questions": []}, {"questions": []}, {"questions": []}]
    def section_index(name, number):
        text = str(name or "").upper()
        if "SECTION" in text or text in ("A", "B", "C"):
            if "A" in text and "B" not in text: return 0
            if "B" in text: return 1
            if "C" in text: return 2
        try:
            n = int(re.search(r"\d+", str(number or ""))[0])
            return 0 if n <= 10 else 1 if n <= 15 else 2
        except Exception:
            return 0
    def add_question(item, fallback_section=None):
        if not isinstance(item, dict): return
        number = item.get("number", item.get("qno", item.get("question_no", "")))
        question = item.get("question", item.get("text", item.get("question_text", "")))
        co = item.get("co", item.get("co_level", item.get("coLevel", "")))
        k = item.get("k", item.get("k_level", item.get("kLevel", "")))
        if not any([number, question, co, k]): return
        idx = section_index(item.get("section", item.get("title", fallback_section)), number)
        canonical[idx]["questions"].append({
            "number": str(number or ""),
            "question": str(question or ""),
            "co": str(co or ""),
            "k": str(k or "")
        })
    if isinstance(raw_sections, dict):
        raw_sections = [raw_sections.get(k, raw_sections.get(k.lower(), [])) for k in ("A", "B", "C")]
    if isinstance(raw_sections, list):
        if raw_sections and all(isinstance(x, dict) and not isinstance(x.get("questions"), list) for x in raw_sections):
            for q in raw_sections: add_question(q)
        else:
            for i, sec in enumerate(raw_sections):
                if isinstance(sec, str):
                    try: sec = json.loads(sec)
                    except Exception: continue
                if isinstance(sec, dict):
                    rows = sec.get("questions", sec.get("rows", sec.get("items", [])))
                    if isinstance(rows, str):
                        try: rows = json.loads(rows)
                        except Exception: rows = []
                    if isinstance(rows, list):
                        for q in rows: add_question(q, sec.get("title", ["A","B","C"][min(i,2)]))
    if not any(x["questions"] for x in canonical) and isinstance(payload.get("questions"), list):
        for q in payload["questions"]: add_question(q)
    payload = dict(payload)
    payload["sections"] = canonical
    return payload
def _question_paper_has_questions(payload):
    """True only when the payload contains at least one real extracted question."""
    try:
        for sec in (payload or {}).get("sections", []):
            for q in (sec or {}).get("questions", []):
                if str((q or {}).get("question", "")).strip():
                    return True
    except Exception:
        pass
    return False
def _normalize_saved_question_paper(payload):
    """Extra defensive normalizer used by both SAVE and VIEW.
    Older versions stored section rows in several shapes.  This routine walks
    nested dict/list/string JSON values and recovers question rows instead of
    returning an empty table in View Extracted QP.
    """
    base = _normalize_question_paper_payload(payload)
    if base and _question_paper_has_questions(base):
        return base
    if isinstance(payload, str):
        try: payload = json.loads(payload)
        except Exception: return base
    if not isinstance(payload, (dict, list)):
        return base
    canonical = [{"questions": []}, {"questions": []}, {"questions": []}]
    seen=set()
    def add(item, fallback=None):
        if not isinstance(item, dict): return
        num=item.get("number", item.get("qno", item.get("question_no", item.get("Q.No", ""))))
        text=item.get("question", item.get("text", item.get("question_text", item.get("Questions", ""))))
        co=item.get("co", item.get("co_level", item.get("coLevel", item.get("CO Level", ""))))
        k=item.get("k", item.get("k_level", item.get("kLevel", item.get("K Level", ""))))
        if not any(str(x or "").strip() for x in (num,text,co,k)): return
        try:
            m=re.search(r"\d+", str(num or "")); n=int(m.group()) if m else 0
        except Exception: n=0
        sec=str(item.get("section", fallback or "")).upper()
        idx=0 if n and n<=10 else 1 if n and n<=15 else 2 if n else (0 if 'A' in sec else 1 if 'B' in sec else 2 if 'C' in sec else 0)
        row={"number":str(num or ""),"question":str(text or "").strip(),"co":str(co or "").strip(),"k":str(k or "").strip()}
        key=(idx,row['number'],row['question'],row['co'],row['k'])
        if key not in seen:
            seen.add(key); canonical[idx]['questions'].append(row)
    def walk(obj, fallback=None, depth=0):
        if depth>8: return
        if isinstance(obj, str):
            t=obj.strip()
            if t[:1] in '[{':
                try: walk(json.loads(t), fallback, depth+1)
                except Exception: pass
            return
        if isinstance(obj, list):
            for x in obj: walk(x, fallback, depth+1)
            return
        if not isinstance(obj, dict): return
        keys={str(k).lower() for k in obj.keys()}
        if keys & {'question','question_text','text','questions'} and not isinstance(obj.get('questions'), list):
            add(obj, fallback)
        for key,val in obj.items():
            low=str(key).lower()
            fb = 'A' if low in ('a','section_a','sectiona') else 'B' if low in ('b','section_b','sectionb') else 'C' if low in ('c','section_c','sectionc') else fallback
            if low in ('sections','questions','rows','items','paper_data','data','payload','section_a','section_b','section_c','a','b','c') or isinstance(val,(dict,list)):
                walk(val, fb, depth+1)
    walk(payload)
    if not any(sec['questions'] for sec in canonical):
        return base
    for sec in canonical:
        sec['questions'].sort(key=lambda q: int(re.search(r'\d+',q['number']).group()) if re.search(r'\d+',q['number']) else 999)
    meta=dict(payload) if isinstance(payload,dict) else {}
    meta['sections']=canonical
    return meta
def _load_question_paper_saved_data(faculty_id, subject_id, cia, academic_year_id=None, semester_no=None):
    """Load the newest non-empty extracted QP without depending on one filename.
    JSON is checked first because it is the exact browser payload.  Database rows
    are then checked from newest to oldest.  Every candidate is normalized and
    rejected unless it contains real question text, preventing blank rows in
    View Extracted QP.
    """
    import glob
    paths=[]
    if academic_year_id is not None and semester_no is not None:
        paths.append(_question_paper_json_path(faculty_id, subject_id, cia, academic_year_id, semester_no))
    paths.append(_question_paper_json_path(faculty_id, subject_id, cia))
    paths.extend(sorted(glob.glob(os.path.join(QUESTION_PAPER_DIR, f"faculty_{faculty_id}_subject_{subject_id}_{cia}*.json")), key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True))
    seen=set()
    for path in paths:
        if not path or path in seen or not os.path.exists(path):
            continue
        seen.add(path)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                candidate=_normalize_saved_question_paper(json.load(f))
            if candidate and _question_paper_has_questions(candidate):
                return candidate
        except Exception:
            continue
    conn=None; cur=None
    try:
        conn=get_db_connection(); cur=conn.cursor(dictionary=True)
        cur.execute("""SELECT paper_json FROM ep_question_papers
                       WHERE faculty_id=%s AND subject_id=%s AND cia=%s
                       ORDER BY created_at DESC, id DESC""", (faculty_id, subject_id, cia))
        for row in cur.fetchall() or []:
            candidate=_normalize_saved_question_paper(row.get('paper_json'))
            if candidate and _question_paper_has_questions(candidate):
                return candidate
    except Exception:
        try:
            if cur: cur.close()
            if conn: conn.close()
            conn=get_db_connection(); cur=conn.cursor(dictionary=True)
            cur.execute("SELECT paper_json FROM ep_question_papers WHERE faculty_id=%s AND subject_id=%s AND cia=%s", (faculty_id, subject_id, cia))
            for row in cur.fetchall() or []:
                candidate=_normalize_saved_question_paper(row.get('paper_json'))
                if candidate and _question_paper_has_questions(candidate):
                    return candidate
        except Exception:
            pass
    finally:
        try:
            if cur: cur.close()
        except Exception: pass
        try:
            if conn: conn.close()
        except Exception: pass
    return None
@app.route("/faculty/question-paper/data")
@faculty_required
def faculty_question_paper_data():
    """AJAX source for View Extracted QP. Never returns the original PDF."""
    faculty=get_faculty()
    if not faculty:
        return jsonify({"success":False,"message":"Faculty profile not found."}), 404
    subject_id=request.args.get('subject_id', type=int)
    requested_semester=request.args.get('semester', type=int)
    cia=request.args.get('cia','cia1').lower()
    if cia not in ('cia1','cia2','external'):
        return jsonify({"success":False,"message":"Invalid examination."}), 400
    if not subject_id:
        return jsonify({"success":False,"message":"Subject is required."}), 400
    conn=get_db_connection(); cur=conn.cursor(dictionary=True)
    try:
        if not _question_paper_access(cur, faculty, subject_id):
            return jsonify({"success":False,"message":"You cannot access this subject."}), 403
        ay, sem, _ = get_selected_faculty_semester(cur, faculty['faculty_id'])
        if requested_semester is not None:
            try:
                assigned = get_faculty_semesters(cur, faculty['faculty_id'], ay['id'] if ay else None)
                valid = {int(x.get('semester_no')) for x in (assigned or []) if x.get('semester_no') is not None}
                if requested_semester in valid:
                    sem = requested_semester
            except Exception:
                pass
    finally:
        cur.close(); conn.close()
    saved=_load_question_paper_saved_data(faculty['faculty_id'], subject_id, cia, ay['id'] if ay else None, sem)
    if not saved:
        return jsonify({"success":False,"message":"No saved extracted questions found."}), 404
    return jsonify({"success":True,"paper":saved})
@app.route("/faculty/question-paper")
@faculty_required
def faculty_question_paper():
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_dashboard"))
    selected_subject_id = request.args.get("subject_id", type=int)
    view_mode = request.args.get("view", "0") == "1"
    cia = request.args.get("cia", "cia1").lower()
    if cia not in ("cia1", "cia2", "external"):
        cia = "cia1"
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        academic_year, semester_no, semester_options = get_selected_faculty_semester(cursor, faculty["faculty_id"])
        subjects = get_faculty_subjects_for_semester(cursor,faculty["faculty_id"],semester_no,academic_year["id"]) if academic_year and semester_no is not None else get_assigned_subjects(cursor, faculty["class_id"])
        subject = None
        if selected_subject_id and _question_paper_access(cursor, faculty, selected_subject_id):
            subject = next(
                (x for x in subjects if x["subject_id"] == selected_subject_id),
                None
            )
        elif selected_subject_id:
            flash("You cannot access this subject.", "danger")
            selected_subject_id = None
    finally:
        cursor.close()
        conn.close()
    saved = None
    if selected_subject_id:
        saved = _load_question_paper_saved_data(
            faculty["faculty_id"], selected_subject_id, cia,
            academic_year["id"] if academic_year else None, semester_no
        )
    return render_template(
        "faculty/question_paper.html",
        faculty=faculty,
        profile_letter=profile_letter,
        subjects=subjects,
        subject=subject,
        selected_subject_id=selected_subject_id,
        cia=cia,
        saved=saved,
        academic_year=academic_year,
        semester_no=semester_no,
        semester_options=semester_options,
        view_mode=view_mode
    )
@app.route("/faculty/question-paper/save", methods=["POST"])
@faculty_required
def faculty_question_paper_save():
    faculty = get_faculty()
    if not faculty:
        return jsonify({"success": False, "message": "Faculty profile not found."}), 403
    try:
        subject_id = int(request.form.get("subject_id", "0"))
        cia = request.form.get("cia", "cia1").lower()
        if cia not in ("cia1", "cia2", "external"):
            raise ValueError("Invalid examination selection.")
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        try:
            if not _question_paper_access(cursor, faculty, subject_id):
                return jsonify({"success": False, "message": "You cannot access this subject."}), 403
        finally:
            cursor.close()
            conn.close()
        data_text = request.form.get("paper_data", "{}")
        data = json.loads(data_text)
        data = _normalize_saved_question_paper(data) or data
        if not _question_paper_has_questions(data):
            raise ValueError("No extracted questions to save. Please scan the question paper and verify the question rows before saving.")
        data["cia"] = cia
        data["subject_id"] = subject_id
        conn_sem=get_db_connection(); cur_sem=conn_sem.cursor(dictionary=True)
        try:
            ay, sem, _ = get_selected_faculty_semester(cur_sem, faculty["faculty_id"])
            data["academic_year_id"] = ay["id"] if ay else None
            data["academic_year"] = ay["year_name"] if ay else data.get("academic_year", "")
            data["semester_no"] = sem
            cur_sem.execute("""SELECT c.class_name, cs.year_no, d.department_name, s.subject_code, s.subject_name
                               FROM ep_faculty_assignments a
                               INNER JOIN classes c ON c.id=a.class_id
                               INNER JOIN departments d ON d.id=c.department_id
                               INNER JOIN subjects s ON s.id=a.subject_id
                               LEFT JOIN ep_class_semesters cs
                                 ON cs.class_id=a.class_id
                                AND cs.academic_year_id=a.academic_year_id
                                AND cs.semester_no=a.semester_no
                               WHERE a.faculty_id=%s AND a.subject_id=%s
                                 AND a.academic_year_id=%s AND a.semester_no=%s
                               LIMIT 1""",(faculty["faculty_id"],subject_id,ay["id"] if ay else 0,sem or 0))
            meta=cur_sem.fetchone()
        finally:
            cur_sem.close(); conn_sem.close()
        if not meta: raise ValueError("The selected subject is not assigned for this semester.")
        data["department"]=meta["department_name"]
        data["subject_code"]=meta["subject_code"]
        data["subject_name"]=meta["subject_name"]
        data["semester"]=str(sem or "")
        data["year"]=str(meta.get("year_no") or ((int(sem)+1)//2 if sem else ""))
        data["duration"]="3 Hours"
        data["entered_date"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["exam_name"]={"cia1":"CIA 1","cia2":"CIA 2","external":"External"}[cia]
        data["saved_at"] = datetime.now().isoformat(timespec="seconds")
        upload = request.files.get("question_paper")
        if upload and upload.filename:
            safe = secure_filename(upload.filename)
            ext = os.path.splitext(safe)[1].lower()
            if ext not in {".png", ".jpg", ".jpeg", ".webp", ".pdf", ".doc", ".docx"}:
                return jsonify({"success": False, "message": "Upload PNG, JPG, WEBP, PDF, DOC or DOCX only."}), 400
            filename = _question_paper_file_key(faculty["faculty_id"], subject_id, cia, data.get("academic_year_id"), data.get("semester_no")) + ext
            upload.save(os.path.join(QUESTION_PAPER_DIR, filename))
            data["original_file"] = url_for(
                "static",
                filename=f"uploads/question_papers/{filename}"
            )
        with open(
            _question_paper_json_path(faculty["faculty_id"], subject_id, cia, data.get("academic_year_id"), data.get("semester_no")),
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        conn2=get_db_connection(); cur2=conn2.cursor()
        try:
            cur2.execute("""INSERT INTO ep_question_papers(faculty_id,subject_id,academic_year_id,semester_no,cia,file_path,paper_json) VALUES(%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE file_path=VALUES(file_path),paper_json=VALUES(paper_json),created_at=CURRENT_TIMESTAMP""",(faculty["faculty_id"],subject_id,data.get("academic_year_id") or 0,data.get("semester_no") or 0,cia,data.get("original_file"),json.dumps(data,ensure_ascii=False)))
            conn2.commit()
        finally:
            cur2.close(); conn2.close()
        return jsonify({"success": True, "message": f"{cia.upper()} question paper saved successfully.", "view_url": url_for("faculty_question_paper", subject_id=subject_id, cia=cia, semester=data.get("semester_no"), view=1)})
    except (ValueError, json.JSONDecodeError) as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except Exception as e:
        return jsonify({"success": False, "message": "Unable to save question paper: " + str(e)}), 500
@app.route("/faculty/view-marks")
@faculty_required
def faculty_view_marks():
    """View/edit marks using Department -> Class -> Semester -> Subject.
    Only the faculty's assigned department and active academic year are available.
    """
    faculty, profile_letter = faculty_page_data()
    if not faculty:
        flash("Faculty profile not found.", "danger")
        return redirect(url_for("faculty_dashboard"))
    selected_department_id = request.args.get("department_id", type=int)
    selected_class_id = request.args.get("class_id", type=int)
    selected_subject_id = request.args.get("subject_id", type=int)
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        academic_year = get_active_academic_year(cursor)
        departments = []
        classes = []
        subjects = []
        students = []
        semester_no = request.args.get("semester", type=int)
        semester_options = []
        if academic_year:
            cursor.execute("""
                SELECT DISTINCT
                    d.id, d.department_name, d.department_code
                FROM ep_faculty_assignments a
                INNER JOIN classes c ON c.id = a.class_id
                INNER JOIN departments d ON d.id = c.department_id
                WHERE a.faculty_id=%s
                  AND a.academic_year_id=%s
                ORDER BY d.department_name
            """, (faculty["faculty_id"], academic_year["id"]))
            departments = cursor.fetchall()
            valid_departments = {int(x["id"]) for x in departments}
            if selected_department_id not in valid_departments:
                selected_department_id = None
            if selected_department_id:
                cursor.execute("""
                    SELECT DISTINCT
                        c.id AS class_id,
                        c.class_name
                    FROM ep_faculty_assignments a
                    INNER JOIN classes c ON c.id=a.class_id
                    WHERE a.faculty_id=%s
                      AND a.academic_year_id=%s
                      AND c.department_id=%s
                    ORDER BY c.class_name
                """, (
                    faculty["faculty_id"], academic_year["id"],
                    selected_department_id
                ))
                classes = cursor.fetchall()
                valid_classes = {int(x["class_id"]) for x in classes}
                if selected_class_id not in valid_classes:
                    selected_class_id = None
            if selected_department_id and selected_class_id:
                cursor.execute("""
                    SELECT DISTINCT
                        a.semester_no,
                        COALESCE(cs.year_no, (a.semester_no+1) DIV 2) AS year_no
                    FROM ep_faculty_assignments a
                    INNER JOIN classes c ON c.id=a.class_id
                    LEFT JOIN ep_class_semesters cs
                      ON cs.class_id=a.class_id
                     AND cs.academic_year_id=a.academic_year_id
                     AND cs.semester_no=a.semester_no
                    WHERE a.faculty_id=%s
                      AND a.class_id=%s
                      AND a.academic_year_id=%s
                      AND c.department_id=%s
                    ORDER BY a.semester_no
                """, (
                    faculty["faculty_id"], selected_class_id,
                    academic_year["id"], selected_department_id
                ))
                semester_options = cursor.fetchall()
                valid_semesters = {int(x["semester_no"]) for x in semester_options}
                if semester_no not in valid_semesters:
                    semester_no = None
                if semester_no is not None:
                    cursor.execute("""
                        SELECT DISTINCT
                            a.subject_id,
                            s.subject_code,
                            s.subject_name,
                            a.semester_no
                        FROM ep_faculty_assignments a
                        INNER JOIN subjects s ON s.id=a.subject_id
                        INNER JOIN classes c ON c.id=a.class_id
                        WHERE a.faculty_id=%s
                          AND a.class_id=%s
                          AND a.academic_year_id=%s
                          AND a.semester_no=%s
                          AND c.department_id=%s
                        ORDER BY s.subject_name
                    """, (
                        faculty["faculty_id"], selected_class_id,
                        academic_year["id"], semester_no,
                        selected_department_id
                    ))
                    subjects = cursor.fetchall()
                    valid_subjects = {int(x["subject_id"]) for x in subjects}
                    if selected_subject_id not in valid_subjects:
                        selected_subject_id = None
            if selected_department_id and selected_class_id and selected_subject_id and semester_no is not None:
                cursor.execute("""
                    SELECT
                        s.id AS student_id,
                        s.register_number,
                        s.student_name,
                        m.id AS mark_id,
                        m.ca1,m.ca2,m.assignment,m.internal,
                        m.external,m.total,m.grade,m.result,m.appreciation
                    FROM students s
                    INNER JOIN ep_student_semesters es
                      ON es.student_id=s.id
                     AND es.class_id=%s
                     AND es.academic_year_id=%s
                     AND es.semester_no=%s
                    INNER JOIN ep_faculty_assignments a
                      ON a.faculty_id=%s
                     AND a.class_id=%s
                     AND a.subject_id=%s
                     AND a.academic_year_id=%s
                     AND a.semester_no=%s
                    LEFT JOIN (
                        SELECT m1.*
                        FROM marks m1
                        INNER JOIN (
                            SELECT student_id,subject_id,MAX(id) AS id
                            FROM marks
                            GROUP BY student_id,subject_id
                        ) lm ON lm.id=m1.id
                    ) m
                      ON m.student_id=s.id AND m.subject_id=%s
                    WHERE s.department_id=%s
                    ORDER BY s.student_name
                """, (
                    selected_class_id, academic_year["id"], semester_no,
                    faculty["faculty_id"], selected_class_id,
                    selected_subject_id, academic_year["id"], semester_no,
                    selected_subject_id, selected_department_id
                ))
                students = cursor.fetchall()
        return render_template(
            "faculty/view_marks.html",
            faculty=faculty,
            profile_letter=profile_letter,
            departments=departments,
            classes=classes,
            subjects=subjects,
            students=students,
            academic_year=academic_year,
            semester_no=semester_no,
            semester_options=semester_options,
            selected_department_id=selected_department_id,
            selected_class_id=selected_class_id,
            selected_subject_id=selected_subject_id
        )
    finally:
        cursor.close()
        conn.close()
from functools import wraps
def student_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Please login first.", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "student":
            flash("You are not authorized to access this page.", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function
def get_student():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""
            SELECT
                s.id AS student_id,
                s.user_id,
                s.register_number,
                s.student_name,
                s.email,
                s.phone,
                s.department_id,
                s.class_id,
                d.department_name,
                d.department_code,
                c.class_name,
                u.username
            FROM students s
            LEFT JOIN departments d
                ON s.department_id = d.id
            LEFT JOIN classes c
                ON s.class_id = c.id
            LEFT JOIN users u
                ON s.user_id = u.id
            WHERE s.user_id = %s
        """, (session.get("user_id"),))
        student = cursor.fetchone()
        if student:
            try:
                cursor.execute("SELECT course_type,start_year,end_year,course_duration FROM ep_department_batches WHERE department_id=%s LIMIT 1", (student["department_id"],))
                batch = cursor.fetchone()
            except mysql.connector.Error:
                batch = None
            if batch:
                context = _academic_context(batch, student.get("class_name"))
                student["academic_batch"] = context["batch"]
                student["current_academic_year"] = context["academic_year"]
                student["current_year"] = _year_label_for_class(student.get("class_name"))
            else:
                student["academic_batch"] = None
                student["current_academic_year"] = None
                student["current_year"] = _year_label_for_class(student.get("class_name"))
        return student
    finally:
        cursor.close()
        conn.close()
@app.route("/student/dashboard")
@student_required
def student_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        student = get_student()
        if not student:
            flash("Student profile not found.", "danger")
            return redirect(url_for("login"))
        cursor.execute("""
            SELECT COUNT(DISTINCT semester_no) AS semester_count
            FROM ep_student_semesters
            WHERE student_id=%s
        """, (student["student_id"],))
        semester_count = int((cursor.fetchone() or {}).get("semester_count") or 0)
        return render_template(
            "student/dashboard.html",
            student=student,
            semester_count=semester_count
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/student/profile")
@student_required
def student_profile():
    student = get_student()
    if not student:
        flash(
            "Student profile not found.",
            "danger"
        )
        return redirect(
            url_for("student_dashboard")
        )
    return render_template(
        "student/profile.html",
        student=student
    )
@app.route("/student/subjects")
@student_required
def student_subjects():
    student = get_student()
    if not student:
        flash(
            "Student profile not found.",
            "danger"
        )
        return redirect(
            url_for("student_dashboard")
        )
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""
            SELECT
                s.id AS subject_id,
                s.subject_code,
                s.subject_name
            FROM subjects s
            INNER JOIN class_subjects cs
                ON s.id = cs.subject_id
            WHERE cs.class_id = %s
            ORDER BY s.subject_name
        """, (
            student["class_id"],
        ))
        subjects = cursor.fetchall()
        return render_template(
            "student/subjects.html",
            student=student,
            subjects=subjects
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/student/marks")
@student_required
def student_marks():
    student = get_student()
    if not student:
        flash("Student profile not found.", "danger")
        return redirect(url_for("student_dashboard"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""
            SELECT DISTINCT
                es.semester_no, es.year_no, es.academic_year_id, es.class_id,
                es.section, es.status, ay.year_name
            FROM ep_student_semesters es
            INNER JOIN ep_academic_years ay ON ay.id=es.academic_year_id
                        WHERE es.student_id=%s
                            AND NOT EXISTS (
                                    SELECT 1
                                    FROM ep_student_semesters newer
                                    WHERE newer.student_id=es.student_id
                                        AND newer.semester_no=es.semester_no
                                        AND newer.academic_year_id>es.academic_year_id
                            )
            ORDER BY es.academic_year_id DESC, es.semester_no
        """, (student["student_id"],))
        student_semesters = cursor.fetchall()
        selected_semester = request.args.get("semester", type=int)
        active_year = get_active_academic_year(cursor)
        active_maps = [
            x for x in student_semesters
            if active_year and int(x["academic_year_id"]) == int(active_year["id"])
        ]
        valid = sorted({int(x["semester_no"]) for x in active_maps or student_semesters})
        if selected_semester not in valid:
            selected_semester = None
            if active_maps:
                cursor.execute("""
                    SELECT ess.semester_no, COUNT(m.id) AS mark_count
                    FROM ep_subject_semesters ess
                    LEFT JOIN marks m
                      ON m.subject_id=ess.subject_id AND m.student_id=%s
                    WHERE ess.class_id=%s AND ess.academic_year_id=%s
                    GROUP BY ess.semester_no
                    ORDER BY mark_count DESC, ess.semester_no DESC
                    LIMIT 1
                """, (student["student_id"], active_maps[0]["class_id"], active_year["id"]))
                marked_semester = cursor.fetchone()
                selected_semester = int(marked_semester["semester_no"]) if marked_semester else int(active_maps[-1]["semester_no"])
            elif valid:
                selected_semester = valid[-1]
        selected_map = None
        if selected_semester is not None:
            selected_map = next(
                (x for x in student_semesters if int(x["semester_no"]) == selected_semester),
                None
            )
        marks = []
        if selected_map:
            cursor.execute("""
                SELECT
                    s.id AS subject_id,
                    s.subject_code,
                    s.subject_name,
                    m.ca1,
                    m.ca2,
                    m.assignment,
                    m.internal,
                    m.external,
                    m.total,
                    m.grade,
                    m.result,
                    m.appreciation
                FROM ep_subject_semesters ess
                INNER JOIN subjects s ON s.id=ess.subject_id
                LEFT JOIN marks m
                  ON m.subject_id=s.id
                 AND m.student_id=%s
                WHERE ess.class_id=%s
                  AND ess.academic_year_id=%s
                  AND ess.semester_no=%s
                ORDER BY s.subject_name
            """, (
                student["student_id"], selected_map["class_id"],
                selected_map["academic_year_id"], selected_semester
            ))
            marks = cursor.fetchall()
        completed = [float(r["total"]) for r in marks if r.get("total") is not None]
        average_percentage = round(sum(completed) / len(completed), 2) if completed else None
        if average_percentage is None:
            appreciation = "Marks not available yet"
        elif average_percentage >= 90:
            appreciation = "Outstanding"
        elif average_percentage >= 80:
            appreciation = "Excellent"
        elif average_percentage >= 70:
            appreciation = "Very Good"
        elif average_percentage >= 60:
            appreciation = "Good"
        elif average_percentage >= 50:
            appreciation = "Satisfactory"
        else:
            appreciation = "Improve"
        for row in marks:
            total = row.get("total")
            if total is None:
                row["performance"] = "pending"
            elif float(total) >= 75:
                row["performance"] = "strong"
            elif float(total) >= 50:
                row["performance"] = "moderate"
            else:
                row["performance"] = "weak"
        return render_template(
            "student/marks.html",
            student=student,
            marks=marks,
            student_semesters=student_semesters,
            selected_semester=selected_semester,
            selected_map=selected_map,
            average_percentage=average_percentage,
            appreciation=appreciation,
            selected_academic_year=(selected_map.get("year_name") if selected_map else None)
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/student/results")
@student_required
def student_results():
    semester = request.args.get("semester", type=int)
    return redirect(url_for("student_marks", semester=semester) if semester else url_for("student_marks"))
def build_progress_card_pdf(student, results):
    buffer=BytesIO()
    doc=SimpleDocTemplate(buffer,pagesize=landscape(A4),rightMargin=28,leftMargin=28,topMargin=28,bottomMargin=28)
    styles=getSampleStyleSheet()
    title_style=ParagraphStyle("EPTitle",parent=styles["Title"],fontSize=20,leading=23,alignment=TA_CENTER,textColor=colors.HexColor("#123b78"),spaceAfter=6)
    meta_style=ParagraphStyle("EPMeta",parent=styles["Normal"],fontSize=10.5,leading=14,alignment=TA_CENTER,textColor=colors.HexColor("#52667a"))
    cell_style=ParagraphStyle("EPCell",parent=styles["Normal"],fontSize=10,leading=12,alignment=TA_CENTER)
    header_style=ParagraphStyle("EPHead",parent=cell_style,fontSize=9.5,leading=11,textColor=colors.white)
    elems=[Paragraph("E-PROGRESS CARD",title_style),Paragraph(f"<b>{student.get('student_name','')}</b> &nbsp; | &nbsp; Register No: <b>{student.get('register_number','')}</b> &nbsp; | &nbsp; {student.get('department_name','')} &nbsp; | &nbsp; {student.get('class_name','')}",meta_style),Spacer(1,16)]
    headers=["S.No","Subject Code","Subject","Internal /25","External /75","Total /100","Grade","Result"]
    data=[[Paragraph(x,header_style) for x in headers]]
    for i,r in enumerate(results,1):
        vals=[str(i),r.get("subject_code") or "-",r.get("subject_name") or "-",r.get("internal") if r.get("internal") is not None else "-",r.get("external") if r.get("external") is not None else "-",r.get("total") if r.get("total") is not None else "-",r.get("grade") or "-",r.get("result") or "-"]
        data.append([Paragraph(str(v),cell_style) for v in vals])
    if len(data)==1: data.append([Paragraph(x,cell_style) for x in ["-","-","No marks available","-","-","-","-","-"]])
    table=Table(data,colWidths=[42,90,245,90,90,90,60,65],repeatRows=1)
    table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1f416d")),("GRID",(0,0),(-1,-1),0.55,colors.HexColor("#d4dde7")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f7f9fc")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    completed_totals = [float(r.get("total")) for r in results if r.get("total") is not None]
    average_percentage = round(sum(completed_totals) / len(completed_totals), 2) if completed_totals else 0.0
    if not completed_totals:
        appreciation = "No marks available"
    elif average_percentage >= 90:
        appreciation = "Excellent"
    elif average_percentage >= 80:
        appreciation = "Very Good"
    elif average_percentage >= 70:
        appreciation = "Good"
    elif average_percentage >= 60:
        appreciation = "Satisfactory"
    elif average_percentage >= 30:
        appreciation = "Needs Improvement"
    else:
        appreciation = "Fail"
    summary_label = ParagraphStyle("EPSummaryLabel", parent=styles["Normal"], fontSize=12, leading=15, alignment=TA_CENTER, textColor=colors.HexColor("#123b78"), spaceBefore=12)
    summary_table = Table([[
        Paragraph(f"<b>Average Percentage: {average_percentage:.2f}%</b>", summary_label),
        Paragraph(f"<b>Appreciation: {appreciation}</b>", summary_label)
    ]], colWidths=[295,295])
    summary_table.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.7, colors.HexColor("#d4dde7")),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.HexColor("#d4dde7")),
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f7f9fc")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,-1), 9),
    ]))
    elems.append(table)
    elems.append(Spacer(1, 8))
    elems.append(summary_table)
    doc.build(elems); buffer.seek(0); return buffer
def fetch_student_results_for_pdf(cursor, student_id):
    cursor.execute("""SELECT s.id,s.register_number,s.student_name,s.email,s.phone,d.department_name,d.department_code,c.class_name,s.class_id FROM students s LEFT JOIN departments d ON d.id=s.department_id LEFT JOIN classes c ON c.id=s.class_id WHERE s.id=%s LIMIT 1""",(student_id,))
    student=cursor.fetchone()
    if not student: return None,[]
    cursor.execute("""SELECT sub.subject_code,sub.subject_name,m.internal,m.external,m.total,m.grade,m.result FROM marks m INNER JOIN subjects sub ON sub.id=m.subject_id WHERE m.student_id=%s ORDER BY sub.subject_name""",(student_id,))
    return student,cursor.fetchall()
@app.route("/student/progress-card/pdf")
@student_required
def student_progress_card_pdf():
    semester = request.args.get("semester", type=int)
    if not semester:
        flash("Please select a semester before generating the PDF.", "warning")
        return redirect(url_for("student_marks"))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        student = get_student()
        if not student:
            flash("Student profile not found.", "danger")
            return redirect(url_for("student_dashboard"))
        cursor.execute("""
            SELECT es.academic_year_id, es.class_id, ay.year_name
            FROM ep_student_semesters es
            INNER JOIN ep_academic_years ay ON ay.id=es.academic_year_id
            WHERE es.student_id=%s AND es.semester_no=%s
            ORDER BY es.academic_year_id DESC
            LIMIT 1
        """, (student["student_id"], semester))
        sem = cursor.fetchone()
        if not sem:
            flash("Selected semester is not available.", "warning")
            return redirect(url_for("student_marks"))
        cursor.execute("""
            SELECT
                sub.subject_code,
                sub.subject_name,
                m.internal,
                m.external,
                m.total,
                m.grade,
                m.result
            FROM ep_subject_semesters ess
            INNER JOIN subjects sub ON sub.id=ess.subject_id
            LEFT JOIN marks m
              ON m.subject_id=sub.id
             AND m.student_id=%s
            WHERE ess.class_id=%s
              AND ess.academic_year_id=%s
              AND ess.semester_no=%s
            ORDER BY sub.subject_name
        """, (student["student_id"], sem["class_id"], sem["academic_year_id"], semester))
        results = cursor.fetchall()
        pdf = build_progress_card_pdf(student, results)
        return send_file(
            pdf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{student['register_number']}_Semester_{semester}_Progress_Card.pdf"
        )
    finally:
        cursor.close()
        conn.close()
@app.route("/faculty/students/<int:student_id>/progress-card/pdf")
@faculty_required
def faculty_progress_card_pdf(student_id):
    faculty=get_faculty()
    if not faculty: flash("Faculty profile not found.","danger"); return redirect(url_for("faculty_dashboard"))
    conn=get_db_connection(); cursor=conn.cursor(dictionary=True)
    try:
        if not verify_student_access(cursor,faculty,student_id): flash("You cannot access this student.","danger"); return redirect(url_for("faculty_students"))
        student,results=fetch_student_results_for_pdf(cursor,student_id); pdf=build_progress_card_pdf(student,results)
        return send_file(pdf,mimetype="application/pdf",as_attachment=True,download_name=f"{student['register_number']}_Progress_Card.pdf")
    finally: cursor.close(); conn.close()
@app.route(
    "/student/change-password",
    methods=["GET", "POST"]
)
@student_required
def student_change_password():
    student = get_student()
    if not student:
        flash(
            "Student profile not found.",
            "danger"
        )
        return redirect(
            url_for("student_dashboard")
        )
    if request.method == "POST":
        current_password = request.form.get(
            "current_password",
            ""
        )
        new_password = request.form.get(
            "new_password",
            ""
        )
        confirm_password = request.form.get(
            "confirm_password",
            ""
        )
        if not current_password:
            flash(
                "Enter your current password.",
                "danger"
            )
            return render_template(
                "student/change_password.html",
                student=student
            )
        if not new_password:
            flash(
                "Enter a new password.",
                "danger"
            )
            return render_template(
                "student/change_password.html",
                student=student
            )
        if len(new_password) < 6:
            flash(
                "New password must contain at least 6 characters.",
                "danger"
            )
            return render_template(
                "student/change_password.html",
                student=student
            )
        if new_password != confirm_password:
            flash(
                "New passwords do not match.",
                "danger"
            )
            return render_template(
                "student/change_password.html",
                student=student
            )
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        try:
            cursor.execute("""
                SELECT password
                FROM users
                WHERE id = %s
                AND role = 'student'
            """, (
                session["user_id"],
            ))
            user = cursor.fetchone()
            if not user:
                flash(
                    "Student account not found.",
                    "danger"
                )
                return render_template(
                    "student/change_password.html",
                    student=student
                )
            if not check_password_hash(
                user["password"],
                current_password
            ):
                flash(
                    "Current password is incorrect.",
                    "danger"
                )
                return render_template(
                    "student/change_password.html",
                    student=student
                )
            password_hash = generate_password_hash(
                new_password
            )
            cursor.execute("""
                UPDATE users
                SET password = %s
                WHERE id = %s
                AND role = 'student'
            """, (
                password_hash,
                session["user_id"]
            ))
            conn.commit()
            flash(
                "Password changed successfully.",
                "success"
            )
            return redirect(
                url_for("student_dashboard")
            )
        except mysql.connector.Error:
            conn.rollback()
            flash(
                "Unable to change password.",
                "danger"
            )
            return render_template(
                "student/change_password.html",
                student=student
            )
        finally:
            cursor.close()
            conn.close()
    return render_template(
        "student/change_password.html",
        student=student
    )
if __name__ == "__main__":
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
    )
