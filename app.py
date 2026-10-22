#!/usr/bin/env python3
"""
Syllabus Parser Web App
Upload Excel syllabus files and generate marks distribution tables.
"""

from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import os
import io
import sqlite3
import hashlib
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
# On Vercel the code directory is read-only, so uploads and the SQLite
# database live in /tmp (reset on each cold start).
_DATA_DIR = '/tmp/syllabus_parser' if os.environ.get('VERCEL') else '.'
app.config['UPLOAD_FOLDER'] = os.path.join(_DATA_DIR, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.config['DATABASE'] = os.path.join(_DATA_DIR, 'syllabus_data.db')

import re
import pdfplumber

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_file_hash(file_data):
    """Calculate MD5 hash of file content for duplicate detection."""
    if isinstance(file_data, bytes):
        return hashlib.md5(file_data).hexdigest()
    elif hasattr(file_data, 'read'):
        content = file_data.read()
        file_data.seek(0)  # Reset file pointer
        return hashlib.md5(content).hexdigest()
    return None

# ============== DATABASE INITIALIZATION ==============
def get_db():
    """Get database connection."""
    db = sqlite3.connect(app.config['DATABASE'])
    db.row_factory = sqlite3.Row
    return db

def init_db():
    """Initialize database with required tables."""
    db = get_db()
    db.executescript('''
        -- Persistent student master table (survives resets)
        CREATE TABLE IF NOT EXISTS student_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            prn TEXT UNIQUE NOT NULL,
            name TEXT,
            gender TEXT,
            category TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        -- Upload history for duplicate detection
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            step INTEGER NOT NULL,
            filename TEXT,
            file_hash TEXT,
            semester TEXT,
            session TEXT,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(file_hash, step)
        );
        
        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            semester TEXT,
            program TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        CREATE TABLE IF NOT EXISTS courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER REFERENCES classes(id),
            code TEXT NOT NULL,
            title TEXT,
            credits INTEGER DEFAULT 0,
            max_marks INTEGER DEFAULT 100,
            min_marks INTEGER DEFAULT 40,
            th_max INTEGER DEFAULT 0,
            th_min INTEGER DEFAULT 0,
            pr_max INTEGER DEFAULT 0,
            pr_min INTEGER DEFAULT 0,
            ce_max INTEGER DEFAULT 0,
            ce_min INTEGER DEFAULT 0,
            vv_max INTEGER DEFAULT 0,
            vv_min INTEGER DEFAULT 0,
            UNIQUE(class_id, code)
        );
        
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_id INTEGER REFERENCES classes(id),
            usn TEXT NOT NULL,
            roll_no TEXT,
            name TEXT,
            section TEXT,
            gender TEXT,
            category TEXT,
            semester TEXT,
            session TEXT,
            sgpa REAL,
            UNIQUE(class_id, usn)
        );
        
        CREATE TABLE IF NOT EXISTS student_courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER REFERENCES students(id),
            course_id INTEGER REFERENCES courses(id),
            UNIQUE(student_id, course_id)
        );
        
        CREATE TABLE IF NOT EXISTS marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER REFERENCES students(id),
            course_id INTEGER REFERENCES courses(id),
            ise_th REAL DEFAULT 0,
            ise_pr_tu REAL DEFAULT 0,
            mse REAL DEFAULT 0,
            ese REAL DEFAULT 0,
            total REAL DEFAULT 0,
            grade TEXT,
            grade_points REAL DEFAULT 0,
            credits_earned REAL DEFAULT 0,
            UNIQUE(student_id, course_id)
        );
    ''')
    db.commit()
    
    # Migration: Add new columns to existing tables if they don't exist
    # This handles upgrading existing databases
    migration_queries = [
        "ALTER TABLE students ADD COLUMN semester TEXT",
        "ALTER TABLE students ADD COLUMN session TEXT",
        "ALTER TABLE students ADD COLUMN sgpa REAL"
    ]
    for query in migration_queries:
        try:
            db.execute(query)
            db.commit()
        except sqlite3.OperationalError as e:
            # Column already exists, ignore
            if 'duplicate column name' not in str(e).lower():
                pass
    
    db.close()

# Initialize database on startup
with app.app_context():
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    init_db()

def parse_pdf_syllabus(file_path):
    """Parse uploaded PDF file and extract syllabus data."""
    courses = []
    generic_courses = {}  # Map generic code pattern -> course dict
    seen_codes = set()  # Prevent duplicates
    
    with pdfplumber.open(file_path) as pdf:
        # Phase 1: Parse Scheme Tables (detect format automatically)
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 3:
                    continue
                
                # Find header row and detect table format
                header_row_idx = -1
                col_indices = {'code': 0, 'name': 1, 'credits': -1}
                
                for row_idx, row in enumerate(table[:5]):
                    row_str = ' '.join([str(c).replace('\n', ' ') for c in row if c]).lower()
                    
                    # Check if this is a header row with course info
                    if 'course code' in row_str or 'course name' in row_str:
                        header_row_idx = row_idx
                        
                        # Find column indices by header text
                        for col_idx, cell in enumerate(row):
                            cell_str = str(cell).replace('\n', ' ').lower() if cell else ''
                            if 'course code' in cell_str or cell_str == 'course code':
                                col_indices['code'] = col_idx
                            elif 'course name' in cell_str:
                                col_indices['name'] = col_idx
                            elif 'credit' in cell_str:
                                col_indices['credits'] = col_idx
                        break
                    
                    # Alternative: BE format with "Examination Marks" 
                    elif 'examination marks' in row_str:
                        header_row_idx = row_idx
                        col_indices = {'code': 0, 'name': 1, 'credits': -1}  # BE format
                        break
                
                if header_row_idx == -1:
                    continue
                
                # Process data rows after header
                # Skip sub-header rows (like TH/PR or ISE1/MSE)
                start_row = header_row_idx + 1
                while start_row < len(table):
                    first_cell = str(table[start_row][0]) if table[start_row][0] else ''
                    # Skip if first cell is None/empty or is a sub-header
                    if not first_cell.strip() or first_cell.lower() in ['none', 'th', 'pr', '']:
                        start_row += 1
                    else:
                        break
                
                for row in table[start_row:]:
                    if not row or len(row) < 2:
                        continue
                    
                    # Get code
                    code = str(row[col_indices['code']]).replace('\n', '').strip() if row[col_indices['code']] else ''
                    if not code or len(code) < 3 or code.lower() in ['none', 'total', 'remark']:
                        continue
                    
                    # Skip duplicate codes
                    if code in seen_codes:
                        continue
                    seen_codes.add(code)
                    
                    # Get name from detected column
                    name_col = col_indices['name']
                    title = str(row[name_col]).replace('\n', ' ').strip() if name_col < len(row) and row[name_col] else ''
                    
                    # Skip if title is empty or looks like an abbreviation only (e.g., "PCPEC")
                    if not title or (len(title) <= 6 and title.isupper()):
                        # Try other columns for a better title
                        for alt_col in range(1, min(5, len(row))):
                            alt_val = str(row[alt_col]).replace('\n', ' ').strip() if row[alt_col] else ''
                            if alt_val and len(alt_val) > 6 and not alt_val.isupper():
                                title = alt_val
                                break
                    
                    if 'total' in title.lower() or 'remark' in title.lower():
                        continue
                    
                    # Parse Credits
                    credits = 3  # Default
                    if col_indices['credits'] >= 0 and col_indices['credits'] < len(row):
                        try:
                            credits = int(float(row[col_indices['credits']])) if row[col_indices['credits']] else 3
                        except:
                            credits = 3
                    elif len(row) > 0:
                        # Try last column
                        try:
                            credits = int(float(row[-1])) if str(row[-1]).replace('.', '', 1).isdigit() else 3
                        except:
                            credits = 3
                    
                    # Determine course type for marks
                    max_marks = 100; min_marks = 40; th_max = 60; th_min = 24
                    pr_max = 0; pr_min = 0; ce_max = 40; ce_min = 16; vv_max = 0; vv_min = 0
                    
                    title_upper = title.upper()
                    code_upper = code.upper()
                    
                    is_lab = 'LAB' in title_upper or ('L' in code_upper[3:5] if len(code_upper) >= 5 else False)
                    is_project = 'PROJECT' in title_upper or ('P' in code_upper[3:5] if len(code_upper) >= 5 else False)
                    
                    if is_project:
                        max_marks = 150; min_marks = 60; th_max = 50; th_min = 20
                        pr_max = 100; pr_min = 40; ce_max = 0; ce_min = 0; credits = max(credits, 6)
                    elif is_lab:
                        max_marks = 50; min_marks = 20; th_max = 0; th_min = 0
                        pr_max = 50; pr_min = 20; ce_max = 0; ce_min = 0; credits = max(credits, 1)
                    
                    course_obj = {
                        'code': code, 'title': title if title else code, 'credits': credits,
                        'max': max_marks, 'min': min_marks,
                        'th_max': th_max, 'th_min': th_min, 'pr_max': pr_max, 'pr_min': pr_min,
                        'ce_max': ce_max, 'ce_min': ce_min, 'vv_max': vv_max, 'vv_min': vv_min
                    }
                    
                    # If Generic (ends in X or 'Optional' in title), store for mapping
                    if code.upper().endswith('X') or 'OPTIONAL' in title.upper():
                        base_key = code.upper().replace(' ', '').replace('X', '')
                        generic_courses[base_key] = course_obj
                    else:
                        courses.append(course_obj)
                        # Don't break - continue to find more scheme tables (Sem 8, etc.)

        # Phase 2: Parse Elective Mapping Tables
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                flat_text = ' '.join([str(c) for row in table for c in row if c]).lower()
                if 'optional' in flat_text and 'course' in flat_text:
                    # Scan rows for specific codes
                    for row in table:
                        for cell in row:
                            if not cell: continue
                            cell_text = str(cell).replace('\n', ' ')
                            # Regex to find "CODE: Title" patterns
                            # Matches e.g. "CSDO7011: NLP" or "CSDO7011 : NLP"
                            matches = re.findall(r'([A-Z]{3,5}\s?[0-9]{3,4}[A-Z]?)\s*:\s*([^:]+?)(?=[A-Z]{3,5}\s?[0-9]{3,4}|$)', cell_text)
                            
                            for m_code, m_title in matches:
                                m_code = m_code.strip()
                                m_title = m_title.strip()
                                
                                # Find matching parent
                                m_base = m_code.replace(' ', '')[:7] # Match first 7 chars e.g. CSDO701
                                parent = None
                                
                                # Try exact match or prefix match against generic keys
                                for gen_key, gen_obj in generic_courses.items():
                                    if gen_key in m_base.upper():
                                        parent = gen_obj
                                        break
                                
                                if parent:
                                    # Create specific course inheriting parent's stats
                                    new_course = parent.copy()
                                    new_course['code'] = m_code
                                    new_course['title'] = m_title
                                    # Avoid duplicates
                                    if not any(c['code'] == m_code for c in courses):
                                        courses.append(new_course)

    # If no specific electives found, fallback to adding the generics
    # (Optional: user might prefer seeing generics if specifics fail)
    # For now, let's ADD the generics if they weren't expanded? 
    # Or just return mixed list. The Reference Excel lists Specifics.
    
    # Sort by code
    courses.sort(key=lambda x: x['code'])
    return courses

def parse_syllabus(file_path):
    """Parse uploaded Excel file and extract syllabus data."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    # Try to find a syllabus sheet
    ws = None
    target_sheets = ['BE Syllabus', 'Syllabus', 'Scheme']
    for name in target_sheets:
        for sheet_name in wb.sheetnames:
            if name.lower() in sheet_name.lower():
                ws = wb[sheet_name]
                break
        if ws:
            break
    
    # If no specific sheet found, search for header in all sheets
    if not ws:
        for sheet_name in wb.sheetnames:
            temp_ws = wb[sheet_name]
            # Check first 10 rows for "Course Code"
            found = False
            for row in range(1, min(15, temp_ws.max_row + 1)):
                cell_val = str(temp_ws.cell(row=row, column=1).value or '').lower()
                if 'course' in cell_val and 'code' in cell_val:
                    ws = temp_ws
                    found = True
                    break
            if found:
                break
    
    # Fallback to active sheet
    if not ws:
        ws = wb.active
    
    courses = []
    
    # Try to find header row again in the selected sheet
    header_row = 1
    for row in range(1, min(15, ws.max_row + 1)):
        cell_val = str(ws.cell(row=row, column=1).value or '').lower()
        if 'course' in cell_val and 'code' in cell_val:
            header_row = row
            break
    
    # Parse courses starting from row after header
    for row in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=row, column=1).value
        title = ws.cell(row=row, column=2).value
        
        if not code or code == 'Total' or 'total' in str(code).lower():
            continue
        
        # Try to get credits from column 3 or extract from course code
        credits = ws.cell(row=row, column=3).value
        if not credits:
            # Estimate credits: Labs=1, Theory=3, Project=6
            if 'L' in str(code).upper() and 'LAB' in str(title or '').upper():
                credits = 1
            elif 'P' in str(code).upper() and 'PROJECT' in str(title or '').upper():
                credits = 6
            else:
                credits = 3
        
        # Determine course type and marks
        title_upper = str(title or '').upper()
        code_upper = str(code).upper()
        
        # Default values (Theory)
        max_marks = 100
        min_marks = 40
        th_max, th_min = 60, 24
        pr_max, pr_min = 0, 0
        ce_max, ce_min = 40, 16
        vv_max, vv_min = 0, 0
        
        # Check patterns
        is_lab = False
        is_project = False
        
        # Pattern 1: 'L' in the 4th position (e.g., CSDL8011 vs CSDC8011)
        if len(code_upper) >= 4 and code_upper[3] == 'L':
            is_lab = True
        # Pattern 2: 'Lab' in title
        elif 'LAB' in title_upper:
            is_lab = True
        # Pattern 3: CSL prefix (e.g., CSL801)
        elif code_upper.startswith('CSL'):
            is_lab = True
            
        # Project detection
        if 'MAJOR PROJECT' in title_upper or 'MINI PROJECT' in title_upper or (len(code_upper) >= 4 and code_upper[3] == 'P'):
            is_project = True
            
        if is_project:
            # Project: 150 marks, TH=50, PR=100
            max_marks = 150
            min_marks = 60
            th_max, th_min = 50, 20
            pr_max, pr_min = 100, 40
            ce_max, ce_min = 0, 0
            vv_max, vv_min = 0, 0
            credits = 6 # Force 6 credits for project
        elif is_lab:
            # Lab course: 50 marks, all practical
            max_marks = 50
            min_marks = 20
            th_max, th_min = 0, 0
            pr_max, pr_min = 50, 20
            ce_max, ce_min = 0, 0
            vv_max, vv_min = 0, 0
            credits = 1 # Force 1 credit for lab
        else:
            # Theory (default)
            pass
        
        courses.append({
            'code': code,
            'title': title,
            'credits': credits,
            'max': max_marks,
            'min': min_marks,
            'th_max': th_max,
            'th_min': th_min,
            'pr_max': pr_max,
            'pr_min': pr_min,
            'ce_max': ce_max,
            'ce_min': ce_min,
            'vv_max': vv_max,
            'vv_min': vv_min
        })
    
    return courses

def generate_excel(filtered_courses, all_courses=None):
    """Generate Excel file with marks distribution table (2 sheets)."""
    wb = openpyxl.Workbook()
    
    # Headers
    headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN', 
               'TH_MAX', 'TH_MIN', 'PR_MAX', 'PR_MIN', 'CE_MAX', 'CE_MIN', 
               'VV_MAX', 'VV_MIN']
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    def write_sheet(ws, courses, title):
        ws.title = title
        
        # Write headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
        
        # Write data
        for row_idx, course in enumerate(courses, 2):
            data = [
                course.get('code', ''), course.get('title', ''), course.get('credits', 0),
                course.get('max', 0), course.get('min', 0),
                course.get('th_max', 0), course.get('th_min', 0),
                course.get('pr_max', 0), course.get('pr_min', 0),
                course.get('ce_max', 0), course.get('ce_min', 0),
                course.get('vv_max', 0), course.get('vv_min', 0)
            ]
            for col, value in enumerate(data, 1):
                cell = ws.cell(row=row_idx, column=col, value=value)
                cell.border = thin_border
                if col > 2:
                    cell.alignment = Alignment(horizontal='center')
        
        # Add total row
        if courses:
            total_row = len(courses) + 2
            total_cell = ws.cell(row=total_row, column=2, value="Total")
            total_cell.font = Font(bold=True)
            credits_cell = ws.cell(row=total_row, column=3, value=sum(c.get('credits', 0) for c in courses))
            credits_cell.font = Font(bold=True)
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 40
        for col in 'CDEFGHIJKLM':
            ws.column_dimensions[col].width = 10
    
    # Sheet 1: Selected/Filtered courses
    ws1 = wb.active
    write_sheet(ws1, filtered_courses, "Selected Courses")
    
    # Sheet 2: All courses (if provided)
    if all_courses:
        ws2 = wb.create_sheet()
        write_sheet(ws2, all_courses, "All Syllabus Courses")
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    print(f"DEBUG: Received file: {file.filename}")
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        print(f"DEBUG: Saved to {filepath}")
        
        try:
            if filename.lower().endswith('.pdf'):
                print("DEBUG: Parsing as PDF...")
                courses = parse_pdf_syllabus(filepath)
            else:
                print("DEBUG: Parsing as Excel...")
                courses = parse_syllabus(filepath)
            
            print(f"DEBUG: Parsed {len(courses)} courses")
            
            # Save to database
            if courses:
                try:
                    db = get_db()
                    cursor = db.cursor()
                    
                    # Create default class/batch for now (TODO: make dynamic)
                    class_name = "Upload Batch " + courses[0]['code'][:3] if courses else "Default"
                    cursor.execute('INSERT OR IGNORE INTO classes (name) VALUES (?)', (class_name,))
                    cursor.execute('SELECT id FROM classes WHERE name = ?', (class_name,))
                    class_id = cursor.fetchone()[0]
                    
                    for course in courses:
                        cursor.execute('''
                            INSERT OR REPLACE INTO courses 
                            (class_id, code, title, credits, max_marks, th_max, th_min, pr_max, pr_min, ce_max, ce_min)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            class_id,
                            course['code'],
                            course['title'],
                            course.get('credits', 3),
                            course.get('max', 100),
                            course.get('th_max', 60),
                            course.get('th_min', 24),
                            course.get('pr_max', 0),
                            course.get('pr_min', 0),
                            course.get('ce_max', 40),
                            course.get('ce_min', 16)
                        ))
                    
                    db.commit()
                    db.close()
                    print("DEBUG: Saved courses to database")
                except Exception as db_err:
                    print(f"DEBUG: Database save warning: {db_err}")
                    # Don't fail the upload if DB save fails
            
            os.remove(filepath)  # Clean up
            return jsonify({'courses': courses})
        except Exception as e:
            print(f"DEBUG: ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls'}), 400

@app.route('/get-courses', methods=['GET'])
def get_courses():
    """Get courses from database for the latest valid syllabus class."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get the class_id that has the MOST valid syllabus data 
        # Valid course codes start with 3+ letters (e.g., CSDC8011, LLC012, AEC12CE021)
        cursor.execute('''
            SELECT class_id, COUNT(*) as valid_count FROM courses 
            WHERE code GLOB '[A-Za-z][A-Za-z][A-Za-z]*'
            GROUP BY class_id
            ORDER BY valid_count DESC LIMIT 1
        ''')
        class_row = cursor.fetchone()
        
        if not class_row:
            # Fallback: get the latest class_id
            cursor.execute('SELECT id FROM classes ORDER BY id DESC LIMIT 1')
            class_row = cursor.fetchone()
        
        if not class_row:
            db.close()
            return jsonify({'courses': []})
        
        class_id = class_row[0]
        
        cursor.execute('''
            SELECT code, title, credits, max_marks, min_marks, 
                   th_max, th_min, pr_max, pr_min, ce_max, ce_min, vv_max, vv_min
            FROM courses
            WHERE class_id = ?
            ORDER BY code
        ''', (class_id,))
        rows = cursor.fetchall()
        db.close()
        
        courses = []
        for row in rows:
            courses.append({
                'code': row[0],
                'title': row[1],
                'credits': row[2],
                'max': row[3],
                'min': row[4],
                'th_max': row[5],
                'th_min': row[6],
                'pr_max': row[7],
                'pr_min': row[8],
                'ce_max': row[9],
                'ce_min': row[10],
                'vv_max': row[11],
                'vv_min': row[12]
            })
        
        return jsonify({'courses': courses})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get-students', methods=['GET'])
def get_students():
    """Get all students from database to display on page load."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        cursor.execute('''
            SELECT usn, roll_no, name, section, gender, category
            FROM students
            ORDER BY roll_no, usn
        ''')
        rows = cursor.fetchall()
        db.close()
        
        students = []
        for row in rows:
            students.append({
                'usn': row[0],
                'roll_no': row[1],
                'name': row[2],
                'section': row[3],
                'gender': row[4],
                'category': row[5]
            })
        
        return jsonify({'students': students})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download', methods=['POST'])
def download_excel():
    data = request.json
    filtered_courses = data.get('filtered_courses', data.get('courses', []))
    all_courses = data.get('all_courses', None)
    
    if not filtered_courses:
        return jsonify({'error': 'No courses selected to export'}), 400
    
    try:
        output = generate_excel(filtered_courses, all_courses)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='BE_Syllabus_Generated.xlsx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/upload-students', methods=['POST'])
def upload_students():
    """Parse student enrollment Excel file."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    filename = secure_filename(file.filename)
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    try:
        students = parse_student_list(filepath)
        
        # Save to database
        if students:
            try:
                db = get_db()
                cursor = db.cursor()
                
                # Get the last added class_id (likely from Step 1)
                # Or find/create a default one
                cursor.execute('SELECT id FROM classes ORDER BY id DESC LIMIT 1')
                class_row = cursor.fetchone()
                
                if class_row:
                    class_id = class_row[0]
                else:
                    cursor.execute("INSERT OR IGNORE INTO classes (name) VALUES ('Default Batch')")
                    class_id = cursor.lastrowid
                
                for student in students:
                    cursor.execute('''
                        INSERT OR REPLACE INTO students 
                        (class_id, usn, roll_no, name, section, gender, category)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        class_id,
                        student.get('usn', ''),
                        student.get('roll_no', ''),
                        student.get('name', ''),
                        student.get('section', ''),
                        student.get('gender', ''),
                        student.get('category', '')
                    ))
                    
                    # Also save student_courses enrollment if courses provided
                    student_id = cursor.lastrowid
                    if not student_id: # If replace/update didn't return id
                         cursor.execute('SELECT id FROM students WHERE usn = ?', (student.get('usn', ''),))
                         student_id = cursor.fetchone()[0]
                    
                    for course_code in student.get('courses', []):
                         cursor.execute('SELECT id FROM courses WHERE code = ? AND class_id = ?', 
                                       (course_code, class_id))
                         course_row = cursor.fetchone()
                         if course_row:
                             cursor.execute('INSERT OR IGNORE INTO student_courses (student_id, course_id) VALUES (?, ?)',
                                           (student_id, course_row[0]))
                
                db.commit()
                db.close()
                print("DEBUG: Saved students to database")
            except Exception as db_err:
                print(f"DEBUG: Database save warning: {db_err}")
                import traceback
                traceback.print_exc()

        os.remove(filepath)
        return jsonify({'students': students, 'count': len(students)})
    except Exception as e:
        import traceback
        traceback.print_exc()
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': str(e)}), 500

def parse_student_list(file_path):
    """Parse student enrollment file and extract USN, Roll No, Name, and opted courses."""
    students = []
    
    wb = openpyxl.load_workbook(file_path)
    
    # Find the sheet with student data
    target_sheet = None
    for sheet_name in wb.sheetnames:
        if 'subject' in sheet_name.lower() or 'list' in sheet_name.lower() or 'student' in sheet_name.lower():
            target_sheet = wb[sheet_name]
            break
    
    if not target_sheet:
        target_sheet = wb.active  # Use first sheet as fallback
    
    ws = target_sheet
    
    # Find header row
    header_row = 1
    usn_col = roll_col = name_col = None
    
    for row in range(1, min(10, ws.max_row + 1)):
        for col in range(1, min(10, ws.max_column + 1)):
            cell_val = str(ws.cell(row=row, column=col).value or '').strip().upper()
            if cell_val in ['USN', 'PRN', 'ENROLL', 'ENROLLMENT']:
                usn_col = col
                header_row = row
            elif cell_val in ['ROLLNO', 'ROLL NO', 'ROLL', 'ROLL_NO']:
                roll_col = col
            elif cell_val == 'NAME':
                name_col = col
        
        if usn_col and name_col:
            break
    
    # Default column positions if not found
    if not usn_col: usn_col = 1
    if not roll_col: roll_col = 2
    if not name_col: name_col = 3
    
    # Parse student rows (start after header, skip empty rows)
    for row in range(header_row + 1, ws.max_row + 1):
        usn = ws.cell(row=row, column=usn_col).value
        roll_no = ws.cell(row=row, column=roll_col).value
        name = ws.cell(row=row, column=name_col).value
        
        # Skip empty rows
        if not usn and not name:
            continue
        
        # Get course codes from remaining columns
        courses = []
        for col in range(name_col + 1, ws.max_column + 1):
            course_code = ws.cell(row=row, column=col).value
            if course_code and str(course_code).strip():
                course_code_str = str(course_code).strip()
                # Only add if it looks like a course code (alphanumeric, 4+ chars)
                if len(course_code_str) >= 4 and any(c.isalpha() for c in course_code_str):
                    courses.append(course_code_str)
        
        students.append({
            'usn': str(usn).strip() if usn else '',
            'roll_no': str(roll_no).strip() if roll_no else '',
            'name': str(name).strip() if name else '',
            'courses': courses
        })
    
    return students

@app.route('/download-students', methods=['POST'])
def download_students():
    """Generate and download students list Excel."""
    data = request.json
    students = data.get('students', [])
    
    if not students:
        return jsonify({'error': 'No student data to export'}), 400
    
    try:
        output = generate_students_excel(students)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Student_List.xlsx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def generate_students_excel(students):
    """Generate Excel with students list - matching Step 1 format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Students List"
    
    # Use same blue color as Step 1
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    headers = ['USN', 'Roll No', 'Name', 'Courses Opted']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    # Data rows
    for row_idx, student in enumerate(students, 2):
        ws.cell(row=row_idx, column=1, value=student.get('usn', '')).border = thin_border
        ws.cell(row=row_idx, column=2, value=student.get('roll_no', '')).border = thin_border
        ws.cell(row=row_idx, column=3, value=student.get('name', '')).border = thin_border
        ws.cell(row=row_idx, column=4, value=', '.join(student.get('courses', []))).border = thin_border
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 60
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/download-combined', methods=['POST'])
def download_combined():
    """Generate and download combined syllabus + students Excel."""
    data = request.json
    courses = data.get('courses', [])
    students = data.get('students', [])
    
    if not courses and not students:
        return jsonify({'error': 'No data to export'}), 400
    
    try:
        output = generate_combined_excel(courses, students)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Syllabus_Students_Combined.xlsx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def generate_combined_excel(courses, students):
    """Generate combined Excel with syllabus and students - matching Step 1 format."""
    wb = openpyxl.Workbook()
    
    # Use same blue color as Step 1
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Sheet 1: Syllabus/Courses (matching Step 1 format exactly)
    ws1 = wb.active
    ws1.title = "Syllabus Courses"
    
    course_headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN', 
                      'TH_MAX', 'TH_MIN', 'PR_MAX', 'PR_MIN', 'CE_MAX', 'CE_MIN', 
                      'VV_MAX', 'VV_MIN']
    
    # Write headers
    for col, header in enumerate(course_headers, 1):
        cell = ws1.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    # Write course data
    for row_idx, course in enumerate(courses, 2):
        data = [
            course.get('code', ''), course.get('title', ''), course.get('credits', 0),
            course.get('max', 0), course.get('min', 0),
            course.get('th_max', 0), course.get('th_min', 0),
            course.get('pr_max', 0), course.get('pr_min', 0),
            course.get('ce_max', 0), course.get('ce_min', 0),
            course.get('vv_max', 0), course.get('vv_min', 0)
        ]
        for col, value in enumerate(data, 1):
            cell = ws1.cell(row=row_idx, column=col, value=value)
            cell.border = thin_border
            if col > 2:
                cell.alignment = Alignment(horizontal='center')
    
    # Add total row
    if courses:
        total_row = len(courses) + 2
        total_cell = ws1.cell(row=total_row, column=2, value="Total")
        total_cell.font = Font(bold=True)
        credits_cell = ws1.cell(row=total_row, column=3, value=sum(c.get('credits', 0) for c in courses))
        credits_cell.font = Font(bold=True)
    
    # Set column widths for Sheet 1
    ws1.column_dimensions['A'].width = 15
    ws1.column_dimensions['B'].width = 40
    for col in 'CDEFGHIJKLM':
        ws1.column_dimensions[col].width = 10
    
    # Sheet 2: Students List
    ws2 = wb.create_sheet("Students List")
    
    student_headers = ['USN', 'Roll No', 'Name', 'Courses Opted']
    for col, header in enumerate(student_headers, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for row_idx, student in enumerate(students, 2):
        ws2.cell(row=row_idx, column=1, value=student.get('usn', '')).border = thin_border
        ws2.cell(row=row_idx, column=2, value=student.get('roll_no', '')).border = thin_border
        ws2.cell(row=row_idx, column=3, value=student.get('name', '')).border = thin_border
        ws2.cell(row=row_idx, column=4, value=', '.join(student.get('courses', []))).border = thin_border
    
    # Set column widths for Sheet 2
    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['B'].width = 10
    ws2.column_dimensions['C'].width = 30
    ws2.column_dimensions['D'].width = 60
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_matrix_excel(courses, students):
    """Generate Excel with student-course matrix."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Student-Course Matrix"
    
    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="6C5CE7", end_color="6C5CE7", fill_type="solid")
    
    # Build header row
    headers = ['USN', 'Roll No', 'Name'] + [c['code'] for c in courses] + ['Total Credits']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    
    # Course code to credits map
    course_credits = {c['code']: c.get('credits', 0) for c in courses}
    
    # Add student rows
    for row_idx, student in enumerate(students, 2):
        ws.cell(row=row_idx, column=1, value=student.get('usn', ''))
        ws.cell(row=row_idx, column=2, value=student.get('roll_no', ''))
        ws.cell(row=row_idx, column=3, value=student.get('name', ''))
        
        student_courses = set(student.get('courses', []))
        total_credits = 0
        
        for col_offset, course in enumerate(courses):
            col = 4 + col_offset
            if course['code'] in student_courses:
                ws.cell(row=row_idx, column=col, value=course.get('credits', 0))
                total_credits += course.get('credits', 0)
            else:
                ws.cell(row=row_idx, column=col, value='-')
        
        # Total credits column
        ws.cell(row=row_idx, column=len(courses) + 4, value=total_credits)
    
    # Create second sheet with course marks structure
    ws2 = wb.create_sheet("Course Marks Structure")
    marks_headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN', 'TH_MAX', 'TH_MIN', 'PR_MAX', 'PR_MIN', 'CE_MAX', 'CE_MIN']
    for col, header in enumerate(marks_headers, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    for row_idx, course in enumerate(courses, 2):
        ws2.cell(row=row_idx, column=1, value=course.get('code', ''))
        ws2.cell(row=row_idx, column=2, value=course.get('title', ''))
        ws2.cell(row=row_idx, column=3, value=course.get('credits', 0))
        ws2.cell(row=row_idx, column=4, value=course.get('max', 100))
        ws2.cell(row=row_idx, column=5, value=course.get('min', 40))
        ws2.cell(row=row_idx, column=6, value=course.get('th_max', 60))
        ws2.cell(row=row_idx, column=7, value=course.get('th_min', 24))
        ws2.cell(row=row_idx, column=8, value=course.get('pr_max', 0))
        ws2.cell(row=row_idx, column=9, value=course.get('pr_min', 0))
        ws2.cell(row=row_idx, column=10, value=course.get('ce_max', 40))
        ws2.cell(row=row_idx, column=11, value=course.get('ce_min', 16))
    
    # Adjust column widths
    for ws_sheet in [ws, ws2]:
        for col in ws_sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws_sheet.column_dimensions[column].width = min(max_length + 2, 25)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============== STEP 3: BE RESULTS ROUTES ==============

@app.route('/get-sheets', methods=['POST'])
def get_sheets():
    """Get list of sheets from an Excel file."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        sheets = wb.sheetnames
        wb.close()
        return jsonify({'sheets': sheets})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/upload-results', methods=['POST'])
def upload_results():
    """Parse BE Results Excel sheet and return structured data."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    sheet_name = request.form.get('sheet', '')
    
    if file.filename == '' or not sheet_name:
        return jsonify({'error': 'File and sheet name required'}), 400
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        
        result = parse_be_results(ws)
        wb.close()
        
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def parse_be_results(ws):
    """Parse BE Results sheet format.
    
    Extracts:
    - SEM and SESSION from header rows (1-7)
    - Course codes from row 8
    - Course names from row 9
    - Student data from row 12+
    - SGPA for each student (usually after last course column)
    """
    courses = []
    students = []
    marks = []
    
    # Extract SEM and SESSION from header rows (rows 1-7)
    semester = ''
    session = ''
    
    for row in range(1, 8):
        for col in range(1, min(20, ws.max_column + 1)):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val and isinstance(cell_val, str):
                cell_str = cell_val.strip().upper()
                # Look for semester info (e.g., "SEM 8", "SEMESTER 8", "8TH SEM")
                import re
                sem_match = re.search(r'(?:SEM(?:ESTER)?\.?\s*[-:]?\s*)?(\d+)(?:ST|ND|RD|TH)?\s*SEM', cell_str)
                if not sem_match:
                    sem_match = re.search(r'SEM(?:ESTER)?\.?\s*[-:]?\s*(\d+)', cell_str)
                if sem_match:
                    semester = 'Sem' + sem_match.group(1)
                
                # Look for session/year info (e.g., "2024-25", "2024-2025", "AY 2024-25")
                session_match = re.search(r'(\d{4})[-/](\d{2,4})', cell_str)
                if session_match:
                    year1 = session_match.group(1)
                    year2 = session_match.group(2)
                    if len(year2) == 2:
                        session = f"{year1}-{year2}"
                    else:
                        session = f"{year1}-{year2[-2:]}"
    
    # Row 8 has course codes starting from column H (8)
    # Row 9 has course names
    # Row 11 has exam type headers (ISE0TH, ISE0PR/ISE0TU, MSE, ESE, TOTAL, GR, GP, CRE) - 8 columns per subject
    # Row 12+ has student data
    
    # Find all course codes in row 8 (they appear every 8 columns starting from H)
    course_columns = {}  # course_code -> start_column
    last_course_col = 8  # Track last course column to find SGPA
    
    for col in range(8, ws.max_column + 1):  # Start from column H (8)
        cell_val = ws.cell(row=8, column=col).value
        if cell_val and isinstance(cell_val, str) and len(cell_val.strip()) >= 4:
            code = str(cell_val).strip()
            # Check if it looks like a course code (has letters and numbers)
            if any(c.isalpha() for c in code) and any(c.isdigit() for c in code):
                course_name = ws.cell(row=9, column=col).value or ''
                course_columns[code] = col
                last_course_col = max(last_course_col, col + 7)  # Each course has 8 columns
                courses.append({
                    'code': code,
                    'name': str(course_name).strip() if course_name else code,
                    'start_col': col
                })
    
    # Find SGPA column - usually after all course columns
    # Look for "SGPA", "S.G.P.A", or similar in row 11 or nearby
    sgpa_col = None
    for col in range(last_course_col + 1, min(last_course_col + 15, ws.max_column + 1)):
        for check_row in [8, 9, 10, 11]:
            cell_val = ws.cell(row=check_row, column=col).value
            if cell_val and isinstance(cell_val, str):
                cell_upper = cell_val.strip().upper().replace('.', '').replace(' ', '')
                if 'SGPA' in cell_upper or cell_upper == 'SGPA':
                    sgpa_col = col
                    break
        if sgpa_col:
            break
    
    # Parse student data starting from row 12
    for row in range(12, ws.max_row + 1):
        # Column A: SlNo (skip)
        # Column B: USN
        # Column C: RollNo
        # Column D: Name
        # Column E: Section
        # Column F: Gender
        # Column G: Category
        
        usn = ws.cell(row=row, column=2).value
        if not usn:
            continue
            
        usn = str(usn).strip()
        roll_no = ws.cell(row=row, column=3).value
        name = ws.cell(row=row, column=4).value
        section = ws.cell(row=row, column=5).value
        gender = ws.cell(row=row, column=6).value
        category = ws.cell(row=row, column=7).value
        
        # Get SGPA if column found
        sgpa = None
        if sgpa_col:
            sgpa_val = ws.cell(row=row, column=sgpa_col).value
            if sgpa_val:
                try:
                    sgpa = float(sgpa_val)
                except (ValueError, TypeError):
                    sgpa = None
        
        student = {
            'usn': usn,
            'roll_no': str(roll_no).strip() if roll_no else '',
            'name': str(name).strip() if name else '',
            'section': str(section).strip() if section else '',
            'gender': str(gender).strip() if gender else '',
            'category': str(category).strip() if category else '',
            'sgpa': sgpa,
            'semester': semester,
            'session': session
        }
        students.append(student)
        
        # Parse marks for each course
        for course in courses:
            start_col = course['start_col']
            
            # Columns relative to start: ISE0TH(0), ISE0PR/TU(1), MSE(2), ESE(3), TOTAL(4), GR(5), GP(6), CRE(7)
            ise_th = ws.cell(row=row, column=start_col).value
            ise_pr_tu = ws.cell(row=row, column=start_col + 1).value
            mse = ws.cell(row=row, column=start_col + 2).value
            ese = ws.cell(row=row, column=start_col + 3).value
            total = ws.cell(row=row, column=start_col + 4).value
            grade = ws.cell(row=row, column=start_col + 5).value
            gp = ws.cell(row=row, column=start_col + 6).value
            cre = ws.cell(row=row, column=start_col + 7).value
            
            # Only add mark entry if there's some data
            if any([ise_th, ise_pr_tu, mse, ese, total]):
                marks.append({
                    'usn': usn,
                    'course_code': course['code'],
                    'ise_th': float(ise_th) if ise_th and str(ise_th).replace('.', '').isdigit() else 0,
                    'ise_pr_tu': float(ise_pr_tu) if ise_pr_tu and str(ise_pr_tu).replace('.', '').isdigit() else 0,
                    'mse': float(mse) if mse and str(mse).replace('.', '').isdigit() else 0,
                    'ese': float(ese) if ese and str(ese).replace('.', '').isdigit() else 0,
                    'total': float(total) if total and str(total).replace('.', '').isdigit() else 0,
                    'grade': str(grade).strip() if grade else '',
                    'grade_points': float(gp) if gp and str(gp).replace('.', '').isdigit() else 0,
                    'credits_earned': float(cre) if cre and str(cre).replace('.', '').isdigit() else 0
                })
    
    return {
        'courses': courses,
        'students': students,
        'marks': marks,
        'marks_count': len(marks),
        'semester': semester,
        'session': session,
        'sgpa_column_found': sgpa_col is not None
    }

@app.route('/save-results', methods=['POST'])
def save_results():
    """Save parsed results to the database."""
    data = request.json
    
    courses = data.get('courses', [])
    students = data.get('students', [])
    marks_data = data.get('marks', [])
    syllabus_courses = data.get('syllabus_courses', [])
    student_list = data.get('student_list', [])
    
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Create or get class (use first course prefix as identifier)
        class_name = "BE Results " + (courses[0]['code'][:3] if courses else "Unknown")
        cursor.execute('INSERT OR IGNORE INTO classes (name) VALUES (?)', (class_name,))
        cursor.execute('SELECT id FROM classes WHERE name = ?', (class_name,))
        class_id = cursor.fetchone()[0]
        
        courses_saved = 0
        students_saved = 0
        marks_saved = 0
        
        # Save courses (merge syllabus info if available)
        syllabus_map = {c.get('code', ''): c for c in syllabus_courses}
        for course in courses:
            code = course['code']
            syllabus_info = syllabus_map.get(code, {})
            
            cursor.execute('''
                INSERT OR REPLACE INTO courses 
                (class_id, code, title, credits, max_marks, th_max, th_min, pr_max, pr_min, ce_max, ce_min)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                class_id,
                code,
                course.get('name') or syllabus_info.get('title', ''),
                syllabus_info.get('credits', 0),
                syllabus_info.get('max_marks', 100),
                syllabus_info.get('th_max', 0),
                syllabus_info.get('th_min', 0),
                syllabus_info.get('pr_max', 0),
                syllabus_info.get('pr_min', 0),
                syllabus_info.get('ce_max', 0),
                syllabus_info.get('ce_min', 0)
            ))
            courses_saved += 1
        
        # Save students
        for student in students:
            cursor.execute('''
                INSERT OR REPLACE INTO students 
                (class_id, usn, roll_no, name, section, gender, category, semester, session, sgpa)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                class_id,
                student['usn'],
                student.get('roll_no', ''),
                student.get('name', ''),
                student.get('section', ''),
                student.get('gender', ''),
                student.get('category', ''),
                student.get('semester', ''),
                student.get('session', ''),
                student.get('sgpa')
            ))
            students_saved += 1
        
        # Save marks
        for mark in marks_data:
            # Get student_id
            cursor.execute('SELECT id FROM students WHERE usn = ? AND class_id = ?', 
                          (mark['usn'], class_id))
            student_row = cursor.fetchone()
            if not student_row:
                continue
            student_id = student_row[0]
            
            # Get course_id
            cursor.execute('SELECT id FROM courses WHERE code = ? AND class_id = ?', 
                          (mark['course_code'], class_id))
            course_row = cursor.fetchone()
            if not course_row:
                continue
            course_id = course_row[0]
            
            cursor.execute('''
                INSERT OR REPLACE INTO marks 
                (student_id, course_id, ise_th, ise_pr_tu, mse, ese, total, grade, grade_points, credits_earned)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                student_id,
                course_id,
                mark.get('ise_th', 0),
                mark.get('ise_pr_tu', 0),
                mark.get('mse', 0),
                mark.get('ese', 0),
                mark.get('total', 0),
                mark.get('grade', ''),
                mark.get('grade_points', 0),
                mark.get('credits_earned', 0)
            ))
            marks_saved += 1
        
        db.commit()
        db.close()
        
        return jsonify({
            'success': True,
            'class_id': class_id,
            'courses_saved': courses_saved,
            'students_saved': students_saved,
            'marks_saved': marks_saved
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download-results', methods=['POST'])
def download_results():
    """Generate and download processed results Excel."""
    data = request.json
    courses = data.get('courses', [])
    students = data.get('students', [])
    marks_data = data.get('marks', [])
    
    try:
        output = generate_results_excel(courses, students, marks_data)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='BE_Results_Processed.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def generate_results_excel(courses, students, marks_data):
    """Generate processed results Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    headers = ['USN', 'Roll No', 'Name', 'Section']
    for course in courses:
        headers.append(f"{course['code']}\n(Total)")
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    
    # Data rows
    marks_map = {}
    for mark in marks_data:
        key = (mark['usn'], mark['course_code'])
        marks_map[key] = mark
    
    for row_idx, student in enumerate(students, 2):
        ws.cell(row=row_idx, column=1, value=student['usn']).border = thin_border
        ws.cell(row=row_idx, column=2, value=student.get('roll_no', '')).border = thin_border
        ws.cell(row=row_idx, column=3, value=student.get('name', '')).border = thin_border
        ws.cell(row=row_idx, column=4, value=student.get('section', '')).border = thin_border
        
        for col_offset, course in enumerate(courses):
            mark = marks_map.get((student['usn'], course['code']))
            total = mark['total'] if mark else ''
            cell = ws.cell(row=row_idx, column=5 + col_offset, value=total)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 12
    for col in range(5, 5 + len(courses)):
        ws.column_dimensions[get_column_letter(col)].width = 12
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============== RESET AND SAVE ROUTES ==============

@app.route('/reset-all-data', methods=['POST'])
def reset_all_data():
    """Reset current session data but preserve student_master records.
    
    Clears: classes, courses, students, marks
    Preserves: student_master (persistent student records), uploads (file history)
    """
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Delete current session data (respecting foreign keys)
        # NOTE: student_master and uploads are NOT cleared - they are persistent
        cursor.execute('DELETE FROM marks')
        cursor.execute('DELETE FROM students')
        cursor.execute('DELETE FROM courses')
        cursor.execute('DELETE FROM classes')
        
        db.commit()
        db.close()
        
        return jsonify({
            'success': True, 
            'message': 'Session data reset. Student master records preserved.'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/save-step1', methods=['POST'])
def save_step1():
    """Save Step 1 courses to database."""
    data = request.json
    courses = data.get('courses', [])
    
    if not courses:
        return jsonify({'success': False, 'error': 'No courses provided'}), 400
    
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Create a class for this session
        class_name = "Syllabus " + datetime.now().strftime("%Y-%m-%d %H:%M")
        cursor.execute('INSERT INTO classes (name) VALUES (?)', (class_name,))
        class_id = cursor.lastrowid
        
        courses_saved = 0
        for course in courses:
            cursor.execute('''
                INSERT OR REPLACE INTO courses 
                (class_id, code, title, credits, max_marks, min_marks, th_max, th_min, pr_max, pr_min, ce_max, ce_min, vv_max, vv_min)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                class_id,
                course.get('code', ''),
                course.get('title', ''),
                course.get('credits', 0),
                course.get('max', 100),
                course.get('min', 40),
                course.get('th_max', 0),
                course.get('th_min', 0),
                course.get('pr_max', 0),
                course.get('pr_min', 0),
                course.get('ce_max', 0),
                course.get('ce_min', 0),
                course.get('vv_max', 0),
                course.get('vv_min', 0)
            ))
            courses_saved += 1
        
        db.commit()
        db.close()
        
        return jsonify({'success': True, 'courses_saved': courses_saved, 'class_id': class_id})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/save-step2', methods=['POST'])
def save_step2():
    """Save Step 2 students to database."""
    data = request.json
    students = data.get('students', [])
    
    if not students:
        return jsonify({'success': False, 'error': 'No students provided'}), 400
    
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get the latest class_id
        cursor.execute('SELECT id FROM classes ORDER BY id DESC LIMIT 1')
        class_row = cursor.fetchone()
        
        if not class_row:
            # Create a default class if none exists
            cursor.execute('INSERT INTO classes (name) VALUES (?)', ('Students Import',))
            class_id = cursor.lastrowid
        else:
            class_id = class_row[0]
        
        students_saved = 0
        for student in students:
            cursor.execute('''
                INSERT OR REPLACE INTO students 
                (class_id, usn, roll_no, name, section, gender, category)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                class_id,
                student.get('usn', ''),
                student.get('roll_no', ''),
                student.get('name', ''),
                student.get('section', ''),
                student.get('gender', ''),
                student.get('category', '')
            ))
            students_saved += 1
        
        db.commit()
        db.close()
        
        return jsonify({'success': True, 'students_saved': students_saved, 'class_id': class_id})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ============== PROGRESSIVE DOWNLOAD ROUTES ==============

@app.route('/check-preconditions', methods=['GET'])
def check_preconditions():
    """Check if all preconditions for Step 4 are met."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Count courses
        cursor.execute('SELECT COUNT(*) FROM courses')
        courses_count = cursor.fetchone()[0]
        
        # Count students
        cursor.execute('SELECT COUNT(*) FROM students')
        students_count = cursor.fetchone()[0]
        
        # Count marks
        cursor.execute('SELECT COUNT(*) FROM marks')
        marks_count = cursor.fetchone()[0]
        
        db.close()
        
        return jsonify({
            'courses': courses_count,
            'students': students_count,
            'marks': marks_count,
            'ready': courses_count > 0 and students_count > 0 and marks_count > 0
        })
    except Exception as e:
        return jsonify({'error': str(e), 'courses': 0, 'students': 0, 'marks': 0, 'ready': False})

@app.route('/download-step1', methods=['GET'])
def download_step1():
    """Download Step 1 only (Syllabus/Courses from database)."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get courses from database
        cursor.execute('''
            SELECT code, title, credits, max_marks, min_marks, 
                   th_max, th_min, pr_max, pr_min, ce_max, ce_min, vv_max, vv_min
            FROM courses
        ''')
        courses = cursor.fetchall()
        db.close()
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Courses"
        
        headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN', 
                   'TH_MAX', 'TH_MIN', 'PR_MAX', 'PR_MIN', 'CE_MAX', 'CE_MIN', 'VV_MAX', 'VV_MIN']
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        
        for row_idx, course in enumerate(courses, 2):
            for col_idx, value in enumerate(course, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Step1_Syllabus.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download-step2', methods=['GET'])
def download_step2():
    """Download Step 1+2 combined (Syllabus + Students from database)."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get courses
        cursor.execute('''
            SELECT code, title, credits, max_marks, min_marks, 
                   th_max, th_min, pr_max, pr_min, ce_max, ce_min, vv_max, vv_min
            FROM courses
        ''')
        courses = cursor.fetchall()
        
        # Get students
        cursor.execute('''
            SELECT usn, roll_no, name, section, gender, category
            FROM students
        ''')
        students = cursor.fetchall()
        
        db.close()
        
        wb = openpyxl.Workbook()
        
        # Sheet 1: Courses
        ws1 = wb.active
        ws1.title = "Courses"
        headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN', 
                   'TH_MAX', 'TH_MIN', 'PR_MAX', 'PR_MIN', 'CE_MAX', 'CE_MIN', 'VV_MAX', 'VV_MIN']
        for col, header in enumerate(headers, 1):
            ws1.cell(row=1, column=col, value=header)
        for row_idx, course in enumerate(courses, 2):
            for col_idx, value in enumerate(course, 1):
                ws1.cell(row=row_idx, column=col_idx, value=value)
        
        # Sheet 2: Students
        ws2 = wb.create_sheet("Students")
        student_headers = ['USN', 'Roll No', 'Name', 'Section', 'Gender', 'Category']
        for col, header in enumerate(student_headers, 1):
            ws2.cell(row=1, column=col, value=header)
        for row_idx, student in enumerate(students, 2):
            for col_idx, value in enumerate(student, 1):
                ws2.cell(row=row_idx, column=col_idx, value=value)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Step1_2_Syllabus_Students.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download-step3', methods=['GET'])
def download_step3():
    """Download Step 1+2+3 combined (Syllabus + Students + Results from database)."""
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get courses
        cursor.execute('''
            SELECT code, title, credits, max_marks, min_marks
            FROM courses
        ''')
        courses = cursor.fetchall()
        
        # Get students
        cursor.execute('''
            SELECT usn, roll_no, name, section, gender, category
            FROM students
        ''')
        students = cursor.fetchall()
        
        # Get marks with student and course info using proper foreign keys
        cursor.execute('''
            SELECT s.usn, s.name, c.code, c.title, m.ise_th, m.ise_pr_tu, m.mse, m.ese, m.total, m.grade, m.grade_points, m.credits_earned
            FROM marks m
            JOIN students s ON m.student_id = s.id
            JOIN courses c ON m.course_id = c.id
        ''')
        all_marks = cursor.fetchall()
        
        # Get unique course codes with titles
        cursor.execute('SELECT DISTINCT c.code, c.title FROM marks m JOIN courses c ON m.course_id = c.id')
        course_info = cursor.fetchall()
        course_codes = [row[0] for row in course_info]
        course_titles = {row[0]: row[1] for row in course_info}
        
        db.close()
        
        wb = openpyxl.Workbook()
        
        # Sheet 1: Courses
        ws1 = wb.active
        ws1.title = "Courses"
        headers = ['Course Code', 'Course Title', 'Credits', 'MAX', 'MIN']
        for col, header in enumerate(headers, 1):
            ws1.cell(row=1, column=col, value=header)
        for row_idx, course in enumerate(courses, 2):
            for col_idx, value in enumerate(course, 1):
                ws1.cell(row=row_idx, column=col_idx, value=value)
        
        # Sheet 2: Students
        ws2 = wb.create_sheet("Students")
        student_headers = ['USN', 'Roll No', 'Name', 'Section', 'Gender', 'Category']
        for col, header in enumerate(student_headers, 1):
            ws2.cell(row=1, column=col, value=header)
        for row_idx, student in enumerate(students, 2):
            for col_idx, value in enumerate(student, 1):
                ws2.cell(row=row_idx, column=col_idx, value=value)
        
        # Sheet 3: Results - Use course TITLES instead of codes
        ws3 = wb.create_sheet("Results")
        result_headers = ['PRN', 'Name'] + [course_titles.get(code, code) for code in course_codes]
        for col, header in enumerate(result_headers, 1):
            ws3.cell(row=1, column=col, value=header)
        
        # Group marks by student
        student_marks = {}
        for row in all_marks:
            usn, name, course_code, course_title, ise_th, ise_pr_tu, mse, ese, total, grade, gp, cre = row
            if usn not in student_marks:
                student_marks[usn] = {'name': name, 'marks': {}}
            student_marks[usn]['marks'][course_code] = total
        
        for row_idx, (usn, data) in enumerate(student_marks.items(), 2):
            ws3.cell(row=row_idx, column=1, value=usn)
            ws3.cell(row=row_idx, column=2, value=data['name'])
            for col_idx, code in enumerate(course_codes, 3):
                ws3.cell(row=row_idx, column=col_idx, value=data['marks'].get(code, ''))
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Step1_2_3_Combined.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ============== STEP 4: Final Result Sheet ==============

# Grade point mapping
GRADE_POINTS = {
    'O': 10, 'A+': 9, 'A': 8, 'B+': 7, 'B': 6, 
    'C': 5, 'P': 4, 'F': 0, 'AB': 0, '': 0
}

@app.route('/generate-result-sheet', methods=['POST'])
def generate_result_sheet():
    """Generate the final comprehensive result sheet combining all data."""
    data = request.json
    semester = data.get('semester', 'Sem8')
    session = data.get('session', '2025-05')
    syllabus_courses = data.get('syllabus_courses', [])
    students_data = data.get('students', [])
    results_data = data.get('results', {})
    
    try:
        db = get_db()
        cursor = db.cursor()
        
        # Get all courses (for reference)
        cursor.execute('''
            SELECT DISTINCT id, code, title, credits, max_marks, 
                   th_max, th_min, pr_max, pr_min, ce_max, ce_min
            FROM courses
            ORDER BY code
        ''')
        courses = [dict(row) for row in cursor.fetchall()]
        
        # Get all students who have marks data
        cursor.execute('''
            SELECT DISTINCT s.id, s.usn, s.roll_no, s.name, s.section, s.gender, s.category,
                   s.semester, s.session, s.sgpa
            FROM students s
            INNER JOIN marks m ON m.student_id = s.id
            ORDER BY s.usn
        ''')
        students_db = [dict(row) for row in cursor.fetchall()]
        
        # Build result for each student
        result_students = []
        
        for student in students_db:
            student_id = student['id']
            student_usn = student['usn']
            
            # Get marks for this student using proper foreign keys
            cursor.execute('''
                SELECT m.id, m.student_id, m.course_id, m.ise_th, m.ise_pr_tu, m.mse, m.ese, 
                       m.total, m.grade, m.grade_points, m.credits_earned,
                       c.code, c.title, c.credits, c.max_marks, c.th_max, c.th_min, 
                       c.pr_max, c.pr_min, c.ce_max, c.ce_min
                FROM marks m
                JOIN courses c ON m.course_id = c.id
                WHERE m.student_id = ?
                ORDER BY c.code
            ''', (student_id,))
            marks_rows = [dict(row) for row in cursor.fetchall()]
            
            # Build student record
            # Use semester/session from database if available, otherwise use params
            db_semester = student.get('semester') or semester
            db_session = student.get('session') or session
            db_sgpa = student.get('sgpa')  # Will calculate if not available
            
            student_record = {
                'roll_no': student['roll_no'],  # Roll number as primary identifier
                'prn': student['usn'],
                'name': student['name'],
                'section': student['section'],
                'gender': student['gender'],
                'category': student['category'],
                'semester': db_semester,
                'session': db_session,
                'sgpa_from_db': db_sgpa  # SGPA from database (if extracted from Step 3)
            }
            
            # Calculate totals
            tot_mrks = 0
            tot_max = 0
            tot_credits = 0
            tot_credit_points = 0
            all_passed = True
            
            # Assign subjects to SUB1, SUB2, etc. - Use course TITLE (name) instead of code
            for i, mark in enumerate(marks_rows[:15], 1):
                prefix = f'sub{i}'
                # Store title as the main subject identifier, code is secondary
                student_record[prefix] = mark['title'] or mark['code']  # Course NAME instead of code
                student_record[f'{prefix}nm'] = mark['title']
                student_record[f'{prefix}code'] = mark['code']  # Keep code for reference
                student_record[f'{prefix}max'] = mark['max_marks']
                student_record[f'{prefix}min'] = int(mark['max_marks'] * 0.4) if mark['max_marks'] else 0
                student_record[f'{prefix}_th_max'] = mark['th_max']
                student_record[f'{prefix}_th_min'] = mark['th_min']
                student_record[f'{prefix}_pr_max'] = mark['pr_max']
                student_record[f'{prefix}_pr_min'] = mark['pr_min']
                student_record[f'{prefix}_ce_max'] = mark['ce_max']
                student_record[f'{prefix}_ce_min'] = mark['ce_min']
                # SUB_TH_MRKS = MSE + ESE (Theory marks from mid-sem + end-sem exams)
                mse_val = mark.get('mse', 0) or 0
                ese_val = mark.get('ese', 0) or 0
                student_record[f'{prefix}_th_mrks'] = mse_val + ese_val
                
                # SUB_PR_MRKS = ISE-PR/ISE-TU (Practical marks)
                student_record[f'{prefix}_pr_mrks'] = mark.get('ise_pr_tu', 0)
                
                # SUB_CE_MRKS = ISE-TH (Internal/Continuous Evaluation marks)
                student_record[f'{prefix}_ce_mrks'] = mark.get('ise_th', 0)
                
                # SUB_VV_MRKS = Only for courses with "Project" in name (not "Project Management")
                course_title = (mark['title'] or '').lower()
                if 'project' in course_title and 'management' not in course_title:
                    student_record[f'{prefix}_vv_mrks'] = ese_val  # Use ESE for project viva
                else:
                    student_record[f'{prefix}_vv_mrks'] = 0
                
                # SUB_TOT = Directly from TOTAL column
                student_record[f'{prefix}_tot'] = mark.get('total', 0)
                
                # SUB_GRADE = From GR column
                student_record[f'{prefix}_grade'] = mark.get('grade', '')
                
                # SUB_GRADE_POINTS = From GP column
                student_record[f'{prefix}_grade_points'] = mark.get('grade_points', 0)
                
                # SUB_CREDIT = From CRE column
                student_record[f'{prefix}_credit'] = mark.get('credits', 0)
                
                # SUB_STATUS = FAIL if GR="F", else PASS
                if mark.get('grade', '') == 'F':
                    all_passed = False
                    student_record[f'{prefix}_status'] = 'FAIL'
                else:
                    student_record[f'{prefix}_status'] = 'PASS'
                
                # Accumulate totals
                tot_mrks += mark['total'] or 0
                tot_max += mark['max_marks'] or 0
                credits = mark['credits'] or 0
                gp = mark['grade_points'] or GRADE_POINTS.get(mark['grade'], 0)
                tot_credits += credits
                tot_credit_points += credits * gp
            
            student_record['tot_mrks'] = tot_mrks
            student_record['tot_max'] = tot_max
            student_record['tot_credit'] = tot_credits
            
            # Use SGPA from database if available, otherwise calculate it
            calculated_sgpa = round(tot_credit_points / tot_credits, 2) if tot_credits > 0 else 0
            if db_sgpa is not None:
                student_record['sgpa'] = db_sgpa
                # Warn if there's a significant difference between calculated and extracted
                if abs(db_sgpa - calculated_sgpa) > 0.1:
                    student_record['sgpa_warning'] = f'Mismatch: calculated={calculated_sgpa}, extracted={db_sgpa}'
            else:
                student_record['sgpa'] = calculated_sgpa
            
            student_record['result'] = 'PASS' if all_passed else 'FAIL'
            
            result_students.append(student_record)
        
        db.close()
        
        # Calculate extraction statistics
        # Each student has 15 possible subjects with 19 fields each = 285 possible fields per student
        # Plus 12 base fields per student
        total_fields_per_student = 12 + (15 * 19)  # 297 fields per student
        total_possible = len(result_students) * total_fields_per_student
        
        # Count extracted (non-empty) vs missing
        extracted_count = 0
        missing_count = 0
        for student in result_students:
            for key, value in student.items():
                if value is not None and value != '' and value != 0:
                    extracted_count += 1
                else:
                    missing_count += 1
        
        return jsonify({
            'students': result_students,
            'courses': courses,
            'semester': semester,
            'session': session,
            'extraction_stats': {
                'extracted': extracted_count,
                'missing': missing_count,
                'total': total_possible,
                'fill_rate': round((extracted_count / (extracted_count + missing_count)) * 100, 1) if (extracted_count + missing_count) > 0 else 0
            }
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download-result-sheet', methods=['POST'])
def download_result_sheet():
    """Generate and download the final comprehensive result sheet Excel."""
    data = request.json
    result_data = data.get('data', {})
    semester = data.get('semester', 'Sem8')
    session = data.get('session', '2025-05')
    
    try:
        output = generate_final_result_excel(result_data, semester, session)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Result_Sheet_{semester}_{session.replace("-", "_")}.xlsx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def generate_final_result_excel(result_data, semester, session):
    """Generate the comprehensive result sheet Excel file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Result_sheet"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E86AB", end_color="2E86AB", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Build headers
    headers = [
        'PRN', 'Name', 'REGN_NO', 'GENDER', 'CASTE', 'SEM', 'SESSION',
        'TOT_MAX', 'TOT_MRKS', 'TOT_CREDIT', 'SGPA', 'RESULT'
    ]
    
    # Add subject headers (SUB1 to SUB15)
    for i in range(1, 16):
        prefix = f'SUB{i}'
        headers.extend([
            f'{prefix}NM', f'{prefix}', f'{prefix}MAX', f'{prefix}MIN',
            f'{prefix}_TH_MAX', f'{prefix}_TH_MIN', f'{prefix}_PR_MAX', f'{prefix}_PR_MIN',
            f'{prefix}_CE_MAX', f'{prefix}_CE_MIN',
            f'{prefix}_TH_MRKS', f'{prefix}_PR_MRKS', f'{prefix}_CE_MRKS', f'{prefix}_VV_MRKS',
            f'{prefix}_TOT', f'{prefix}_STATUS', f'{prefix}_GRADE', f'{prefix}_GRADE_POINTS',
            f'{prefix}_CREDIT'
        ])
    
    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    
    # Write data rows
    students = result_data.get('students', [])
    for row_idx, student in enumerate(students, 2):
        # Basic columns
        ws.cell(row=row_idx, column=1, value=student.get('prn', '')).border = thin_border
        ws.cell(row=row_idx, column=2, value=student.get('name', '')).border = thin_border
        ws.cell(row=row_idx, column=3, value=student.get('prn', '')).border = thin_border
        ws.cell(row=row_idx, column=4, value=student.get('gender', '')).border = thin_border
        ws.cell(row=row_idx, column=5, value=student.get('category', '')).border = thin_border
        ws.cell(row=row_idx, column=6, value=semester).border = thin_border
        ws.cell(row=row_idx, column=7, value=session).border = thin_border
        ws.cell(row=row_idx, column=8, value=student.get('tot_max', 0)).border = thin_border
        ws.cell(row=row_idx, column=9, value=student.get('tot_mrks', 0)).border = thin_border
        ws.cell(row=row_idx, column=10, value=student.get('tot_credit', 0)).border = thin_border
        ws.cell(row=row_idx, column=11, value=student.get('sgpa', 0)).border = thin_border
        ws.cell(row=row_idx, column=12, value=student.get('result', '')).border = thin_border
        
        # Subject columns
        col_offset = 13
        for i in range(1, 16):
            prefix = f'sub{i}'
            values = [
                student.get(f'{prefix}nm', ''),
                student.get(prefix, ''),
                student.get(f'{prefix}max', ''),
                student.get(f'{prefix}min', ''),
                student.get(f'{prefix}_th_max', ''),
                student.get(f'{prefix}_th_min', ''),
                student.get(f'{prefix}_pr_max', ''),
                student.get(f'{prefix}_pr_min', ''),
                student.get(f'{prefix}_ce_max', ''),
                student.get(f'{prefix}_ce_min', ''),
                student.get(f'{prefix}_th_mrks', ''),
                student.get(f'{prefix}_pr_mrks', ''),
                student.get(f'{prefix}_ce_mrks', ''),
                student.get(f'{prefix}_vv_mrks', ''),
                student.get(f'{prefix}_tot', ''),
                student.get(f'{prefix}_status', ''),
                student.get(f'{prefix}_grade', ''),
                student.get(f'{prefix}_grade_points', ''),
                student.get(f'{prefix}_credit', '')
            ]
            for j, val in enumerate(values):
                cell = ws.cell(row=row_idx, column=col_offset + j, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center')
            col_offset += 19
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 30
    for col in range(3, 13):
        ws.column_dimensions[get_column_letter(col)].width = 12
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    print("🎓 Syllabus Parser running at http://localhost:5501")
    app.run(debug=True, port=5501)
