
[web_app.py](https://github.com/user-attachments/files/29495718/web_app.py)
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, make_response
import sqlite3
from datetime import datetime, timedelta
import csv
import io
import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT

app = Flask(__name__)
app.secret_key = 'web-app-secret-key-2026'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True

DATABASE = 'attendance.db'
ACCESS_CODE = 'alusman@123'
LICENSE_FILE = '.license'
EXPIRY_DAYS = 60
ACTIVATION_KEY = 'azeem@taizco.com'

# ============ LICENSE FUNCTIONS ============

def get_license_data():
    """Read license file"""
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
                return data
        except:
            return None
    return None

def is_license_valid():
    """Check if license is valid"""
    data = get_license_data()
    if not data:
        return False
    
    try:
        expiry = datetime.strptime(data['expiry'], '%Y-%m-%d %H:%M:%S')
        now = datetime.now()
        if now <= expiry:
            return True
        else:
            return False
    except:
        return False

# ============ LICENSE CHECK DECORATOR ============

def license_required(f):
    """Decorator to check license before accessing routes"""
    def decorated_function(*args, **kwargs):
        if not is_license_valid():
            return redirect(url_for('license_expired'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# ============ DATABASE FUNCTIONS ============

def get_db():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

def get_user_name(conn, user_id, floor):
    """Get user name from database"""
    cursor = conn.cursor()
    cursor.execute('''
        SELECT name FROM users 
        WHERE user_id = ? AND device_floor = ?
    ''', (user_id, floor))
    result = cursor.fetchone()
    return result['name'] if result else f"User {user_id}"

# ============ HELPER: GROUP PUNCHES BY DAY ============

def group_punches_by_day(logs):
    """
    Group punches by user, date, and floor.
    Returns dict with key: user_id|date|floor
    """
    grouped = {}
    
    for log in logs:
        key = f"{log['user_id']}|{log['date_only']}|{log['device_floor']}"
        
        if key not in grouped:
            grouped[key] = {
                'user_id': log['user_id'],
                'user_name': log['user_name'] or 'Unknown',
                'device_floor': log['device_floor'],
                'date': log['date_only'],
                'day': log['day_name'],
                'in_punches': [],
                'out_punches': []
            }
        
        if log['punch_type'] == 0:
            grouped[key]['in_punches'].append(log['time'])
        else:
            grouped[key]['out_punches'].append(log['time'])
    
    return grouped

def calculate_remark_in(time_str, day_name):
    """Calculate IN remark based on time"""
    if not time_str or time_str == '--':
        return 'Missing In'
    
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        
        if day_name == 'Saturday':
            start = datetime.strptime('10:00', '%H:%M').time()
            grace = datetime.strptime('10:30', '%H:%M').time()
        else:
            start = datetime.strptime('09:15', '%H:%M').time()
            grace = datetime.strptime('09:30', '%H:%M').time()
        
        check_time = time_obj.time()
        
        if check_time <= start:
            return 'On Time'
        elif check_time <= grace:
            return 'Grace'
        else:
            return 'Late'
    except:
        return 'On Time'

def calculate_remark_out(time_str, day_name):
    """Calculate OUT remark based on time"""
    if not time_str or time_str == '--':
        return 'Missing Out'
    
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        
        if day_name == 'Saturday':
            early = datetime.strptime('14:00', '%H:%M').time()
            overtime = datetime.strptime('16:00', '%H:%M').time()
        else:
            early = datetime.strptime('17:00', '%H:%M').time()
            overtime = datetime.strptime('18:00', '%H:%M').time()
        
        check_time = time_obj.time()
        
        if check_time < early:
            return 'Early Dept'
        elif check_time > overtime:
            return 'Overtime'
        else:
            return 'On Time'
    except:
        return 'On Time'

def calculate_status(in_time, out_time, day_name):
    """Calculate combined status based on IN and OUT presence"""
    in_present = in_time is not None and in_time != '--'
    out_present = out_time is not None and out_time != '--'
    
    # Sunday handling
    if day_name == 'Sunday':
        return 'Weekly Off'
    
    if in_present and out_present:
        in_remark = calculate_remark_in(in_time, day_name)
        out_remark = calculate_remark_out(out_time, day_name)
        return f"Present / {in_remark} / {out_remark}"
    elif in_present and not out_present:
        return f"Present / Missing Out"
    elif not in_present and out_present:
        return f"Present / Missing In"
    else:
        return "Leave / Absent(Document your Truancy) / Missing In-Out"

def get_display_time(time_str):
    """Convert 24-hour time to 12-hour AM/PM format"""
    if not time_str or time_str == '--':
        return '--'
    
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        return time_obj.strftime('%I:%M %p').lstrip('0')
    except:
        return time_str

def get_sunday_display():
    """Return the Sunday merged display text"""
    return '***************SUNDAY*****************'

# ============ ROUTES ============

@app.route('/license-expired')
def license_expired():
    """Show license expired page"""
    return render_template('license_expired.html'), 403

@app.route('/activate', methods=['GET', 'POST'])
def activate():
    """Hidden activation page"""
    if request.method == 'POST':
        key = request.form.get('key', '').strip()
        
        if key == ACTIVATION_KEY:
            data = {
                'expiry': (datetime.now() + timedelta(days=EXPIRY_DAYS)).strftime('%Y-%m-%d %H:%M:%S'),
                'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(LICENSE_FILE, 'w') as f:
                json.dump(data, f)
            flash('License activated successfully!', 'success')
            return redirect(url_for('login'))
        else:
            flash('Invalid activation key', 'danger')
            return render_template('activate.html', error='Invalid activation key')
    
    return render_template('activate.html', error=None)

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login page"""
    if not is_license_valid():
        return redirect(url_for('license_expired'))
    
    if 'access_code' in session:
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        access_code = request.form.get('access_code', '').strip()
        
        if access_code == ACCESS_CODE:
            session['access_code'] = access_code
            session.permanent = False
            flash('Login successful!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid access code', 'danger')
            return render_template('login.html', error='Invalid access code')
    
    return render_template('login.html', error=None)

@app.route('/logout')
def logout():
    """Logout and clear session"""
    session.clear()
    flash('Logged out successfully', 'info')
    return redirect(url_for('login'))

# ============ MAIN INDEX ============

@app.route('/')
@license_required
def index():
    """Main attendance view page - 10 column format"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    # Get filter parameters
    page = request.args.get('page', 1, type=int)
    per_page = 30
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Build query for logs
    query = '''
        SELECT 
            a.user_id,
            a.device_floor,
            DATE(a.timestamp) as date_only,
            strftime('%w', a.timestamp) as weekday,
            strftime('%H:%M', a.timestamp) as time,
            a.punch_type,
            u.name as user_name
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE 1=1
    '''
    count_query = 'SELECT COUNT(*) as total FROM attendance_logs WHERE 1=1'
    params = []
    
    if date_from:
        query += ' AND DATE(a.timestamp) >= ?'
        count_query += ' AND DATE(timestamp) >= ?'
        params.append(date_from)
    
    if date_to:
        query += ' AND DATE(a.timestamp) <= ?'
        count_query += ' AND DATE(timestamp) <= ?'
        params.append(date_to)
    
    if selected_floor:
        query += ' AND a.device_floor = ?'
        count_query += ' AND device_floor = ?'
        params.append(selected_floor)
    
    if search_user:
        query += ' AND (a.user_id LIKE ? OR u.name LIKE ?)'
        count_query += ' AND (user_id LIKE ?)'
        search_pattern = f'%{search_user}%'
        params.extend([search_pattern, search_pattern])
    
    # Get total count (for pagination)
    cursor.execute(count_query, params)
    total_count = cursor.fetchone()['total']
    
    # Get logs with pagination
    query += ' ORDER BY a.timestamp DESC LIMIT ? OFFSET ?'
    offset = (page - 1) * per_page
    cursor.execute(query, params + [per_page, offset])
    raw_logs = cursor.fetchall()
    
    # Process logs - add day names
    weekday_map = {
        '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
        '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
    }
    
    processed_logs = []
    for log in raw_logs:
        processed_logs.append({
            'user_id': log['user_id'],
            'user_name': log['user_name'] or 'Unknown',
            'device_floor': log['device_floor'],
            'date_only': log['date_only'],
            'day_name': weekday_map.get(log['weekday'], 'Unknown'),
            'time': log['time'],
            'punch_type': log['punch_type']
        })
    
    # Group punches by user, date, floor
    grouped = group_punches_by_day(processed_logs)
    
    # Build final table data
    table_data = []
    for key, data in grouped.items():
        user_id = data['user_id']
        user_name = data['user_name']
        device_floor = data['device_floor']
        date_str = data['date']
        day_name = data['day']
        
        # Format date
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        date_formatted = date_obj.strftime('%d-%b-%Y')
        
        # Get earliest IN and latest OUT
        in_times = data['in_punches']
        out_times = data['out_punches']
        
        in_time = min(in_times) if in_times else None
        out_time = max(out_times) if out_times else None
        
        # Sunday handling - merged display
        if day_name == 'Sunday':
            table_data.append({
                'user_id': user_id,
                'user_name': user_name,
                'device_floor': device_floor,
                'date': date_formatted,
                'day': day_name,
                'chk_in': get_sunday_display(),
                'remark_in': '',
                'chk_out': '',
                'remark_out': '',
                'status': 'Weekly Off',
                'is_sunday': True
            })
        else:
            # Working day - normal display
            chk_in = get_display_time(in_time) if in_time else '--'
            chk_out = get_display_time(out_time) if out_time else '--'
            
            remark_in = calculate_remark_in(in_time, day_name) if in_time else 'Missing In'
            remark_out = calculate_remark_out(out_time, day_name) if out_time else 'Missing Out'
            
            status = calculate_status(in_time, out_time, day_name)
            
            table_data.append({
                'user_id': user_id,
                'user_name': user_name,
                'device_floor': device_floor,
                'date': date_formatted,
                'day': day_name,
                'chk_in': chk_in,
                'remark_in': remark_in,
                'chk_out': chk_out,
                'remark_out': remark_out,
                'status': status,
                'is_sunday': False
            })
    
    # Sort by date descending
    table_data.sort(key=lambda x: datetime.strptime(x['date'], '%d-%b-%Y'), reverse=True)
    
    total_pages = (total_count + per_page - 1) // per_page if total_count > 0 else 1
    
    conn.close()
    
    return render_template('index.html', 
                         logs=table_data,
                         page=page,
                         total_pages=total_pages,
                         total_count=total_count,
                         date_from=date_from,
                         date_to=date_to,
                         selected_floor=selected_floor,
                         search_user=search_user)

# ============ USERS ============

@app.route('/users')
@license_required
def users():
    """View all users"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    search_query = request.args.get('search', '')
    selected_floor = request.args.get('floor', '')
    selected_status = request.args.get('status', '')
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT u.*, 
               (SELECT COUNT(*) FROM attendance_logs a WHERE a.user_id = u.user_id AND a.device_floor = u.device_floor) as log_count
        FROM users u
        WHERE 1=1
    '''
    params = []
    
    if search_query:
        query += ' AND (u.user_id LIKE ? OR u.name LIKE ?)'
        search_pattern = f'%{search_query}%'
        params.extend([search_pattern, search_pattern])
    
    if selected_floor:
        query += ' AND u.device_floor = ?'
        params.append(selected_floor)
    
    if selected_status:
        query += ' AND u.is_active = ?'
        params.append(selected_status)
    
    query += ' ORDER BY u.name'
    cursor.execute(query, params)
    users = cursor.fetchall()
    
    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']
    
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    active_count = cursor.fetchone()['total']
    
    inactive_count = total_users - active_count
    
    conn.close()
    
    return render_template('users.html',
                         users=users,
                         search_query=search_query,
                         selected_floor=selected_floor,
                         selected_status=selected_status,
                         active_count=active_count,
                         inactive_count=inactive_count)

# ============ REPORTS ============

@app.route('/reports')
@license_required
def reports():
    """Reports and statistics page"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Total counts - only ACTIVE users
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    total_users = cursor.fetchone()['total']
    
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 1')
    active_users = cursor.fetchone()['total']
    
    cursor.execute('SELECT COUNT(*) as total FROM users WHERE is_active = 0')
    inactive_users = cursor.fetchone()['total'] or 0
    
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    
    # Date range
    cursor.execute('SELECT MIN(timestamp) as min_date, MAX(timestamp) as max_date FROM attendance_logs')
    result = cursor.fetchone()
    date_from = result['min_date'] if result['min_date'] else 'N/A'
    date_to = result['max_date'] if result['max_date'] else 'N/A'
    
    if date_from != 'N/A' and date_to != 'N/A':
        days_covered = (datetime.strptime(date_to, '%Y-%m-%d %H:%M:%S') - 
                       datetime.strptime(date_from, '%Y-%m-%d %H:%M:%S')).days + 1
    else:
        days_covered = 0
    
    # Users by floor - only ACTIVE users
    cursor.execute('SELECT device_floor, COUNT(*) as count FROM users WHERE is_active = 1 GROUP BY device_floor')
    floor_users = {row['device_floor']: row['count'] for row in cursor.fetchall()}
    
    # Logs by floor
    cursor.execute('SELECT device_floor, COUNT(*) as count FROM attendance_logs GROUP BY device_floor')
    floor_logs = {row['device_floor']: row['count'] for row in cursor.fetchall()}
    
    # Daily stats (last 7 days)
    daily_stats = []
    for i in range(7):
        date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
        cursor.execute('''
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN punch_type = 0 THEN 1 ELSE 0 END) as in_count,
                SUM(CASE WHEN punch_type = 1 THEN 1 ELSE 0 END) as out_count
            FROM attendance_logs
            WHERE DATE(timestamp) = ?
        ''', (date,))
        row = cursor.fetchone()
        daily_stats.append({
            'date': date,
            'total': row['total'] or 0,
            'in_count': row['in_count'] or 0,
            'out_count': row['out_count'] or 0
        })
    
    # Get employees for activity report dropdown - ONLY ACTIVE USERS
    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()
    
    conn.close()
    
    return render_template('reports.html',
                         active_tab='summary',
                         total_users=total_users,
                         active_users=active_users,
                         inactive_users=inactive_users,
                         total_logs=total_logs,
                         date_from=date_from,
                         date_to=date_to,
                         days_covered=days_covered,
                         floor_users=floor_users,
                         floor_logs=floor_logs,
                         daily_stats=daily_stats,
                         employees=employees,
                         selected_employee='',
                         date_from_activity='',
                         date_to_activity='',
                         activity_logs=[])

# ============ EMPLOYEE ACTIVITY REPORT ============

@app.route('/activity')
@license_required
def activity():
    """Employee Activity Report - shows all punches without grouping"""
    if 'access_code' not in session:
        return redirect(url_for('login'))

    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

    conn = get_db()
    cursor = conn.cursor()

    # Get all employees for dropdown - ONLY ACTIVE USERS
    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()

    activity_logs = []
    selected_employee_name = ''

    if selected_employee:
        parts = selected_employee.split('|')
        if len(parts) == 2:
            user_id = parts[0]
            device_floor = parts[1]

            cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
            user = cursor.fetchone()
            if user:
                selected_employee_name = user['name']

            query = '''
                SELECT
                    user_id,
                    device_floor,
                    DATE(timestamp) as date,
                    strftime('%w', timestamp) as weekday,
                    strftime('%H:%M', timestamp) as time,
                    punch_type,
                    strftime('%Y-%m-%d', timestamp) as date_only
                FROM attendance_logs
                WHERE user_id = ?
                    AND device_floor = ?
                    AND DATE(timestamp) >= ?
                    AND DATE(timestamp) <= ?
                ORDER BY timestamp ASC
            '''
            cursor.execute(query, (user_id, device_floor, date_from, date_to))
            punches = cursor.fetchall()

            weekday_map = {
                '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
                '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
            }

            for punch in punches:
                day_name = weekday_map.get(punch['weekday'], 'Unknown')
                date_obj = datetime.strptime(punch['date_only'], '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')

                time_formatted = punch['time']
                try:
                    time_obj = datetime.strptime(time_formatted, '%H:%M')
                    time_formatted = time_obj.strftime('%I:%M %p').lstrip('0')
                except:
                    pass

                activity_logs.append({
                    'user_id': punch['user_id'],
                    'device_floor': punch['device_floor'],
                    'date': date_formatted,
                    'day': day_name,
                    'user_name': selected_employee_name,
                    'time': time_formatted,
                    'punch_type': punch['punch_type']
                })

    conn.close()

    return render_template('activity.html',
                         employees=employees,
                         selected_employee=selected_employee,
                         date_from_activity=date_from,
                         date_to_activity=date_to,
                         activity_logs=activity_logs)@app.route('/export/excel')
@license_required
def export_excel():
    """Export to Excel (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    
    # Headers - 10 columns
    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
    ws.append(headers)
    
    # Style headers
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Add data
    for row in data:
        ws.append([
            row['user_id'],
            row['device_floor'],
            row['date'],
            row['day'],
            row['user_name'],
            row['chk_in'],
            row['remark_in'],
            row['chk_out'],
            row['remark_out'],
            row['status']
        ])
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 35)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='attendance.xlsx'
    )

@app.route('/export/csv')
@license_required
def export_csv():
    """Export to CSV (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    output = io.StringIO()
    writer = csv.writer(output)
    
    writer.writerow(['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status'])
    
    for row in data:
        writer.writerow([
            row['user_id'],
            row['device_floor'],
            row['date'],
            row['day'],
            row['user_name'],
            row['chk_in'],
            row['remark_in'],
            row['chk_out'],
            row['remark_out'],
            row['status']
        ])
    
    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='attendance.csv'
    )

@app.route('/export/pdf')
@license_required
def export_pdf():
    """Export to PDF (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=20
    )
    
    content = []
    
    title = Paragraph(f"🅐∞🅐 USMAN ENTERPRISE - Attendance Report", title_style)
    content.append(title)
    content.append(Spacer(1, 10))
    
    table_data = []
    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
    table_data.append(headers)
    
    for row in data:
        table_data.append([
            str(row['user_id']),
            str(row['device_floor']),
            str(row['date']),
            str(row['day']),
            str(row['user_name']),
            str(row['chk_in']),
            str(row['remark_in']),
            str(row['chk_out']),
            str(row['remark_out']),
            str(row['status'])
        ])
    
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
    ]))
    
    content.append(table)
    doc.build(content)
    buffer.seek(0)
    
    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name='attendance.pdf'
    )

@app.route('/export/sheets')
@license_required
def export_sheets():
    """Export to Excel with multiple sheets (10-column format)"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    selected_floor = request.args.get('floor', '')
    search_user = request.args.get('user_id', '')
    
    data = get_export_data_grouped(date_from, date_to, selected_floor, search_user)
    
    # Group data by user
    user_data = {}
    for row in data:
        user_key = f"{row['user_id']}_{row['user_name']}"
        if user_key not in user_data:
            user_data[user_key] = {
                'user_id': row['user_id'],
                'user_name': row['user_name'],
                'records': []
            }
        user_data[user_key]['records'].append(row)
    
    wb = Workbook()
    
    # Create SUMMARY sheet
    summary_ws = wb.active
    summary_ws.title = "SUMMARY"
    
    summary_data = [
        ['USMAN ENTERPRISE - Attendance Summary'],
        [''],
        ['Period:', date_from if date_from else 'N/A', 'to', date_to if date_to else 'N/A'],
        ['Total Users:', len(user_data)],
        ['Total Logs:', len(data)],
        [''],
        ['Users by Floor:'],
    ]
    
    floor_counts = {}
    for row in data:
        floor = row['device_floor']
        floor_counts[floor] = floor_counts.get(floor, 0) + 1
    
    for floor, count in floor_counts.items():
        summary_data.append([floor, count])
    
    summary_data.append([''])
    summary_data.append(['All Rights Reserved ® 2026'])
    
    for row_idx, row_data in enumerate(summary_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = summary_ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = Font(bold=True, size=14)
            elif row_idx == len(summary_data):
                cell.font = Font(italic=True, size=10)
    
    summary_ws.column_dimensions['A'].width = 25
    summary_ws.column_dimensions['B'].width = 15
    summary_ws.column_dimensions['C'].width = 15
    summary_ws.column_dimensions['D'].width = 15
    
    # Remove default sheet if we have user data
    if user_data:
        for user_key, user_info in user_data.items():
            sheet_name = f"{user_info['user_id']}_{user_info['user_name']}"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            
            ws = wb.create_sheet(title=sheet_name)
            
            headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Remarks In', 'Chk Out', 'Remarks Out', 'Status']
            ws.append(headers)
            
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
            
            for row in user_info['records']:
                ws.append([
                    row['user_id'],
                    row['device_floor'],
                    row['date'],
                    row['day'],
                    row['user_name'],
                    row['chk_in'],
                    row['remark_in'],
                    row['chk_out'],
                    row['remark_out'],
                    row['status']
                ])
            
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 35)
                ws.column_dimensions[column].width = adjusted_width
    
    if not user_data:
        ws = wb.active
        ws.title = "No Data"
        ws.append(['No attendance records found for the selected filters'])
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='attendance_sheets.xlsx'
    )

# ============ EMPLOYEE ACTIVITY EXPORT ============

@app.route('/activity/export')
@license_required
def export_activity():
    """Export Employee Activity Report to Excel"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    
    if not selected_employee or not date_from or not date_to:
        flash('Please select employee and date range', 'error')
        return redirect(url_for('activity'))
    
    parts = selected_employee.split('|')
    if len(parts) != 2:
        flash('Invalid employee selection', 'error')
        return redirect(url_for('activity'))
    
    user_id = parts[0]
    device_floor = parts[1]
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
    user = cursor.fetchone()
    employee_name = user['name'] if user else 'Unknown'
    
    query = '''
        SELECT 
            user_id,
            device_floor,
            DATE(timestamp) as date,
            strftime('%w', timestamp) as weekday,
            strftime('%H:%M', timestamp) as time,
            punch_type,
            strftime('%Y-%m-%d', timestamp) as date_only
        FROM attendance_logs
        WHERE user_id = ? 
            AND device_floor = ?
            AND DATE(timestamp) >= ?
            AND DATE(timestamp) <= ?
        ORDER BY timestamp ASC
    '''
    cursor.execute(query, (user_id, device_floor, date_from, date_to))
    punches = cursor.fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Activity"
    
    headers = ['ID', 'Floor', 'Date', 'Day', 'Name', 'Chk In', 'Chk Out']
    ws.append(headers)
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1a237e", end_color="1a237e", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    weekday_map = {
        '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
        '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
    }
    
    for punch in punches:
        day_name = weekday_map.get(punch['weekday'], 'Unknown')
        date_obj = datetime.strptime(punch['date_only'], '%Y-%m-%d')
        date_formatted = date_obj.strftime('%d-%b-%Y')
        
        time_formatted = punch['time']
        try:
            time_obj = datetime.strptime(time_formatted, '%H:%M')
            time_formatted = time_obj.strftime('%I:%M %p').lstrip('0')
        except:
            pass
        
        chk_in = time_formatted if punch['punch_type'] == 0 else '--'
        chk_out = time_formatted if punch['punch_type'] == 1 else '--'
        
        ws.append([
            punch['user_id'],
            punch['device_floor'],
            date_formatted,
            day_name,
            employee_name,
            chk_in,
            chk_out
        ])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    file_name = f"employee_activity_{employee_name}_{date_from}_to_{date_to}.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=file_name
    )

# ============ CONTEXT PROCESSOR ============

@app.context_processor
def utility_processor():
    """Make variables available to all templates"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']
    conn.close()
    return {'total_logs': total_logs, 'total_users': total_users}

# ============ MAIN ============

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)




    [fix_recursion.py](https://github.com/user-attachments/files/29495731/fix_recursion.py)
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Fix RecursionError in web_app.py
This script creates activity.html and updates web_app.py
"""

import os
import re

print("=" * 60)
print("FIX RECURSIONERROR IN WEB_APP.PY")
print("=" * 60)

# ============================================================
# STEP 1: CREATE activity.html
# ============================================================

print("\n[1] Creating templates/activity.html...")

activity_html = '''{% extends "layout.html" %}

{% block title %}Employee Activity Report{% endblock %}

{% block content %}
<div class="card-modern">
    <div class="card-header-modern">
        <i class="bi bi-clock-history"></i> Employee Activity Report
        <span class="badge bg-secondary float-end">Detailed punch audit</span>
    </div>
    <div class="card-body-modern">

        <div class="alert alert-info alert-modern mb-4">
            <i class="bi bi-info-circle"></i>
            <strong>Shows every individual punch</strong> - No grouping, no remarks, just raw punch data.
            Each IN and OUT punch appears as a separate row.
        </div>

        <form method="GET" action="{{ url_for('activity') }}" class="filter-row">
            <div class="filter-item">
                <label class="form-label fw-semibold"><i class="bi bi-person"></i> Employee</label>
                <select name="employee" class="form-select form-select-modern" required>
                    <option value="">-- Select Employee --</option>
                    {% for emp in employees %}
                        <option value="{{ emp.user_id }}|{{ emp.device_floor }}"
                                {% if selected_employee == emp.user_id|string + '|' + emp.device_floor %}selected{% endif %}>
                            {{ emp.user_id }} - {{ emp.name }} ({{ emp.device_floor }})
                        </option>
                    {% endfor %}
                </select>
            </div>
            <div class="filter-item">
                <label class="form-label fw-semibold"><i class="bi bi-calendar"></i> From</label>
                <input type="date" name="date_from" class="form-control form-control-modern"
                       value="{{ date_from_activity or '' }}" required>
            </div>
            <div class="filter-item">
                <label class="form-label fw-semibold"><i class="bi bi-calendar"></i> To</label>
                <input type="date" name="date_to" class="form-control form-control-modern"
                       value="{{ date_to_activity or '' }}" required>
            </div>
            <div class="filter-item-btn">
                <button type="submit" class="btn btn-primary-custom">
                    <i class="bi bi-search"></i> View
                </button>
            </div>
            <div class="filter-item-btn">
                <a href="{{ url_for('activity') }}" class="btn btn-secondary-custom">
                    <i class="bi bi-arrow-counterclockwise"></i> Reset
                </a>
            </div>
        </form>

        {% if activity_logs %}
        <div class="table-responsive">
            <table class="table table-custom">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Floor</th>
                        <th>Date</th>
                        <th>Day</th>
                        <th>Name</th>
                        <th>Chk In</th>
                        <th>Chk Out</th>
                    </tr>
                </thead>
                <tbody>
                    {% for log in activity_logs %}
                    <tr>
                        <td><strong>{{ log.user_id }}</strong></td>
                        <td>
                            {% if log.device_floor == '2nd_floor' %}
                                <span class="badge-floor-2nd">2nd</span>
                            {% else %}
                                <span class="badge-floor-3rd">3rd</span>
                            {% endif %}
                        </td>
                        <td>{{ log.date }}</td>
                        <td>{{ log.day }}</td>
                        <td>{{ log.user_name }}</td>
                        <td>
                            {% if log.punch_type == 0 %}
                                <span class="badge-in">{{ log.time }}</span>
                            {% else %}
                                <span class="text-muted">--</span>
                            {% endif %}
                        </td>
                        <td>
                            {% if log.punch_type == 1 %}
                                <span class="badge-out">{{ log.time }}</span>
                            {% else %}
                                <span class="text-muted">--</span>
                            {% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>

        <div class="mt-3 d-flex justify-content-between align-items-center">
            <div>
                <a href="{{ url_for('export_activity', employee=selected_employee, date_from=date_from_activity, date_to=date_to_activity) }}"
                   class="btn btn-excel">
                    <i class="bi bi-file-earmark-excel"></i> Download Excel
                </a>
                <span class="text-muted ms-3" style="font-size: 0.85rem;">
                    <i class="bi bi-list-ul"></i> {{ activity_logs|length }} punches found
                </span>
            </div>
        </div>

        {% else %}
            {% if selected_employee %}
            <div class="alert alert-info alert-modern">
                <i class="bi bi-info-circle"></i> No punches found for the selected employee and date range.
            </div>
            {% else %}
            <div class="alert alert-info alert-modern">
                <i class="bi bi-info-circle"></i> Select an employee and date range to view activity.
            </div>
            {% endif %}
        {% endif %}

    </div>
</div>
{% endblock %}'''

# Write activity.html with UTF-8 encoding
os.makedirs('templates', exist_ok=True)
with open('templates/activity.html', 'w', encoding='utf-8') as f:
    f.write(activity_html)

print("   ✅ templates/activity.html created")

# ============================================================
# STEP 2: FIX web_app.py - Replace activity() function
# ============================================================

print("\n[2] Fixing web_app.py - updating activity() function...")

web_app_path = 'web_app.py'

if not os.path.exists(web_app_path):
    print("   ❌ web_app.py not found!")
    exit(1)

# Read with UTF-8 encoding
with open(web_app_path, 'r', encoding='utf-8') as f:
    content = f.read()

# New activity function (clean version)
new_activity = """@app.route('/activity')
@license_required
def activity():
    \"\"\"Employee Activity Report - shows all punches without grouping\"\"\"
    if 'access_code' not in session:
        return redirect(url_for('login'))

    selected_employee = request.args.get('employee', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    if not date_from and not date_to:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

    conn = get_db()
    cursor = conn.cursor()

    # Get all employees for dropdown - ONLY ACTIVE USERS
    cursor.execute('SELECT user_id, name, device_floor FROM users WHERE is_active = 1 ORDER BY name')
    employees = cursor.fetchall()

    activity_logs = []
    selected_employee_name = ''

    if selected_employee:
        parts = selected_employee.split('|')
        if len(parts) == 2:
            user_id = parts[0]
            device_floor = parts[1]

            cursor.execute('SELECT name FROM users WHERE user_id = ? AND device_floor = ?', (user_id, device_floor))
            user = cursor.fetchone()
            if user:
                selected_employee_name = user['name']

            query = '''
                SELECT
                    user_id,
                    device_floor,
                    DATE(timestamp) as date,
                    strftime('%w', timestamp) as weekday,
                    strftime('%H:%M', timestamp) as time,
                    punch_type,
                    strftime('%Y-%m-%d', timestamp) as date_only
                FROM attendance_logs
                WHERE user_id = ?
                    AND device_floor = ?
                    AND DATE(timestamp) >= ?
                    AND DATE(timestamp) <= ?
                ORDER BY timestamp ASC
            '''
            cursor.execute(query, (user_id, device_floor, date_from, date_to))
            punches = cursor.fetchall()

            weekday_map = {
                '0': 'Sunday', '1': 'Monday', '2': 'Tuesday', '3': 'Wednesday',
                '4': 'Thursday', '5': 'Friday', '6': 'Saturday'
            }

            for punch in punches:
                day_name = weekday_map.get(punch['weekday'], 'Unknown')
                date_obj = datetime.strptime(punch['date_only'], '%Y-%m-%d')
                date_formatted = date_obj.strftime('%d-%b-%Y')

                time_formatted = punch['time']
                try:
                    time_obj = datetime.strptime(time_formatted, '%H:%M')
                    time_formatted = time_obj.strftime('%I:%M %p').lstrip('0')
                except:
                    pass

                activity_logs.append({
                    'user_id': punch['user_id'],
                    'device_floor': punch['device_floor'],
                    'date': date_formatted,
                    'day': day_name,
                    'user_name': selected_employee_name,
                    'time': time_formatted,
                    'punch_type': punch['punch_type']
                })

    conn.close()

    return render_template('activity.html',
                         employees=employees,
                         selected_employee=selected_employee,
                         date_from_activity=date_from,
                         date_to_activity=date_to,
                         activity_logs=activity_logs)"""

# Find and replace the activity function
pattern = r'@app\.route\(\'/activity\'\)\s*@license_required\s*def activity\(\):.*?(?=@app\.route|\Z)'

if re.search(pattern, content, re.DOTALL):
    content = re.sub(pattern, new_activity, content, flags=re.DOTALL)
    print("   ✅ activity() function updated")
else:
    print("   ⚠️ Could not find activity() function - checking manually...")
    # Try to find and replace manually
    if 'def activity():' in content:
        lines = content.split('\n')
        new_lines = []
        in_activity = False
        activity_found = False

        for line in lines:
            if 'def activity():' in line and not activity_found:
                in_activity = True
                activity_found = True
                # Add the new function
                new_lines.extend(new_activity.split('\n'))
                continue

            if in_activity:
                # Check if we're out of the function
                if line.strip() and not line.startswith(('    ', '\t')) and line.strip():
                    in_activity = False
                    new_lines.append(line)
                continue

            new_lines.append(line)

        content = '\n'.join(new_lines)
        print("   ✅ activity() function updated (manual method)")
    else:
        print("   ❌ Could not find activity() function in web_app.py")

# Write updated web_app.py with UTF-8 encoding
with open(web_app_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("   ✅ web_app.py updated")

# ============================================================
# STEP 3: Update reports.html - Remove activity tab
# ============================================================

print("\n[3] Updating reports.html - removing activity tab...")

reports_path = 'templates/reports.html'

if os.path.exists(reports_path):
    with open(reports_path, 'r', encoding='utf-8') as f:
        reports_content = f.read()

    # Remove activity tab content
    pattern_tab = r'<!-- ==================== EMPLOYEE ACTIVITY TAB ==================== -->.*?<!-- ==================== END EMPLOYEE ACTIVITY TAB ==================== -->'

    if re.search(pattern_tab, reports_content, re.DOTALL):
        reports_content = re.sub(pattern_tab, '', reports_content, flags=re.DOTALL)
        print("   ✅ Removed activity tab from reports.html")
    else:
        # Try alternative pattern
        pattern_tab2 = r'<div class="tab-pane fade[^>]*id="activity-tab".*?</div>\s*</div>'
        if re.search(pattern_tab2, reports_content, re.DOTALL):
            reports_content = re.sub(pattern_tab2, '', reports_content, flags=re.DOTALL)
            print("   ✅ Removed activity tab from reports.html (alternative)")

    # Update the nav link for activity
    nav_pattern = r'<a class="nav-link[^>]*href="#activity-tab"[^>]*>.*?</a>'
    nav_replacement = '<a class="nav-link" href="{{ url_for(\'activity\') }}" style="color: #6c757d; font-weight: 600; padding: 10px 24px; border: none;"><i class="bi bi-clock-history"></i> Employee Activity</a>'

    if re.search(nav_pattern, reports_content, re.DOTALL):
        reports_content = re.sub(nav_pattern, nav_replacement, reports_content, flags=re.DOTALL)
        print("   ✅ Updated activity nav link")

    with open(reports_path, 'w', encoding='utf-8') as f:
        f.write(reports_content)
else:
    print("   ⚠️ templates/reports.html not found - skipping")

# ============================================================
# SUMMARY
# ============================================================

print("\n" + "=" * 60)
print("FIX COMPLETED!")
print("=" * 60)
print("\n✅ Created: templates/activity.html")
print("✅ Updated: web_app.py (activity() function)")
print("✅ Updated: templates/reports.html (removed activity tab)")
print("\n" + "=" * 60)
print("Now restart your Flask app:")
print("   python web_app.py")
print("=" * 60)





















from flask import Flask, render_template, request, redirect, url_for, flash, session
import sqlite3
from datetime import datetime, timedelta
import os
import json

app = Flask(__name__)
app.secret_key = 'edit-app-secret-key-2026'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True

DATABASE = 'attendance.db'
ACCESS_CODE = 'alusman@123'
LICENSE_FILE = '.license'
EXPIRY_DAYS = 60
ACTIVATION_KEY = 'azeem@taizco.com'

# ============ LICENSE FUNCTIONS ============

def get_license_data():
    """Read license file"""
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
                return data
        except:
            return None
    return None

def is_license_valid():
    """Check if license is valid"""
    data = get_license_data()
    if not data:
        return False
    
    try:
        expiry = datetime.strptime(data['expiry'], '%Y-%m-%d %H:%M:%S')
        now = datetime.now()
        if now <= expiry:
            return True
        else:
            return False
    except:
        return False

# ============ LICENSE CHECK DECORATOR ============

def license_required(f):
    """Decorator to check license before accessing routes"""
    def decorated_function(*args, **kwargs):
        if not is_license_valid():
            return redirect(url_for('license_expired'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# ============ DATABASE FUNCTIONS ============

def get_db():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

# ============ ROUTES ============

@app.route('/license-expired')
def license_expired():
    """Show license expired page"""
    return render_template('license_expired.html'), 403

@app.route('/activate', methods=['GET', 'POST'])
def activate():
    """Hidden activation page"""
    if request.method == 'POST':
        key = request.form.get('key', '').strip()
        
        if key == ACTIVATION_KEY:
            # Activation successful - reset expiry
            data = {
                'expiry': (datetime.now() + timedelta(days=EXPIRY_DAYS)).strftime('%Y-%m-%d %H:%M:%S'),
                'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(LICENSE_FILE, 'w') as f:
                json.dump(data, f)
            flash('License activated successfully!', 'success')
            return redirect(url_for('login'))
        else:
            flash('Invalid activation key', 'danger')
            return render_template('edit_activate.html', error='Invalid activation key')
    
    return render_template('edit_activate.html', error=None)

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login page"""
    # Check license first
    if not is_license_valid():
        return redirect(url_for('license_expired'))
    
    # If already logged in, redirect to edit list
    if 'access_code' in session:
        return redirect(url_for('edit_list'))
    
    if request.method == 'POST':
        access_code = request.form.get('access_code', '').strip()
        
        if access_code == ACCESS_CODE:
            session['access_code'] = access_code
            session.permanent = False
            flash('Login successful!', 'success')
            return redirect(url_for('edit_list'))
        else:
            flash('Invalid access code', 'danger')
            return render_template('edit_login.html', error='Invalid access code')
    
    return render_template('edit_login.html', error=None)

@app.route('/logout')
def logout():
    """Logout and clear session"""
    session.clear()
    flash('Logged out successfully', 'info')
    return redirect(url_for('login'))

@app.route('/edit')
@license_required
def edit_list():
    """List all records that can be edited"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    page = request.args.get('page', 1, type=int)
    per_page = 20
    search = request.args.get('search', '')
    floor = request.args.get('floor', '')
    show_only_auto = request.args.get('auto_only', 'false') == 'true'
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT a.*, u.name as user_name
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE 1=1
    '''
    params = []
    
    if search:
        query += ' AND (a.user_id LIKE ? OR u.name LIKE ?)'
        search_pattern = f'%{search}%'
        params.extend([search_pattern, search_pattern])
    
    if floor:
        query += ' AND a.device_floor = ?'
        params.append(floor)
    
    if show_only_auto:
        query += ' AND a.is_manually_edited = 0'
    
    # Get total count
    count_query = query.replace('SELECT a.*, u.name as user_name', 'SELECT COUNT(*) as total')
    cursor.execute(count_query, params)
    total_count = cursor.fetchone()['total']
    
    # Get paginated results
    query += ' ORDER BY a.timestamp DESC LIMIT ? OFFSET ?'
    offset = (page - 1) * per_page
    cursor.execute(query, params + [per_page, offset])
    logs = cursor.fetchall()
    
    total_pages = (total_count + per_page - 1) // per_page if total_count > 0 else 1
    
    conn.close()
    
    return render_template('edit_list.html',
                         logs=logs,
                         page=page,
                         total_pages=total_pages,
                         total_count=total_count,
                         search=search,
                         floor=floor,
                         show_only_auto=show_only_auto)

@app.route('/edit/<int:log_id>', methods=['GET', 'POST'])
@license_required
def edit_single(log_id):
    """Edit a single attendance record"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get the log record with lock
    cursor.execute('BEGIN IMMEDIATE')
    
    cursor.execute('''
        SELECT a.*, u.name as user_name
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE a.id = ?
    ''', (log_id,))
    log = cursor.fetchone()
    
    if not log:
        conn.rollback()
        conn.close()
        flash('Record not found', 'error')
        return redirect(url_for('edit_list'))
    
    # Check if record is manually edited
    is_locked = log['is_manually_edited'] == 1
    
    if request.method == 'POST':
        if is_locked:
            conn.rollback()
            conn.close()
            flash('This record is locked (manually edited) and cannot be modified', 'error')
            return redirect(url_for('edit_list'))
        
        try:
            # Get form data with validation
            remark_in = request.form.get('remark_in', '').strip()
            remark_out = request.form.get('remark_out', '').strip()
            status_override = request.form.get('status_override', '').strip()
            
            # Validate inputs
            if len(remark_in) > 50:
                flash('Remark IN is too long (max 50 characters)', 'error')
                return render_template('edit_single.html', log=log, is_locked=is_locked)
            
            if len(remark_out) > 50:
                flash('Remark OUT is too long (max 50 characters)', 'error')
                return render_template('edit_single.html', log=log, is_locked=is_locked)
            
            # Update database with transaction
            cursor.execute('''
                UPDATE attendance_logs 
                SET remark_in = ?,
                    remark_out = ?,
                    status_override = ?,
                    is_manually_edited = 1,
                    last_edited_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (remark_in if remark_in else None, 
                  remark_out if remark_out else None, 
                  status_override if status_override else None,
                  log_id))
            
            conn.commit()
            conn.close()
            
            flash('Record updated successfully!', 'success')
            return redirect(url_for('edit_list'))
            
        except sqlite3.Error as e:
            conn.rollback()
            conn.close()
            flash(f'Database error: {str(e)}', 'error')
            return redirect(url_for('edit_single', log_id=log_id))
    
    conn.commit()  # Release lock if GET request
    conn.close()
    
    return render_template('edit_single.html', log=log, is_locked=is_locked)

@app.route('/edit/bulk', methods=['GET', 'POST'])
@license_required
def edit_bulk():
    """Bulk edit multiple records"""
    if 'access_code' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        # Get selected record IDs and field to update
        record_ids = request.form.getlist('record_ids')
        field = request.form.get('field')
        value = request.form.get('value', '').strip()
        
        if not record_ids:
            flash('No records selected', 'error')
            return redirect(url_for('edit_list'))
        
        if not field:
            flash('No field selected to update', 'error')
            return redirect(url_for('edit_list'))
        
        # Validate field
        allowed_fields = ['remark_in', 'remark_out', 'status_override']
        if field not in allowed_fields:
            flash('Invalid field selected', 'error')
            return redirect(url_for('edit_list'))
        
        conn = get_db()
        cursor = conn.cursor()
        
        try:
            # Check if any selected records are manually edited
            placeholders = ','.join(['?'] * len(record_ids))
            cursor.execute(f'''
                SELECT COUNT(*) as count FROM attendance_logs 
                WHERE id IN ({placeholders}) AND is_manually_edited = 1
            ''', record_ids)
            locked_count = cursor.fetchone()['count']
            
            if locked_count > 0:
                conn.close()
                flash(f'{locked_count} selected record(s) are locked (manually edited) and cannot be modified', 'error')
                return redirect(url_for('edit_list'))
            
            # Update all selected records
            cursor.execute(f'''
                UPDATE attendance_logs 
                SET {field} = ?,
                    is_manually_edited = 1,
                    last_edited_at = CURRENT_TIMESTAMP
                WHERE id IN ({placeholders})
            ''', [value if value else None] + record_ids)
            
            conn.commit()
            conn.close()
            
            flash(f'{len(record_ids)} record(s) updated successfully!', 'success')
            
        except sqlite3.Error as e:
            conn.rollback()
            conn.close()
            flash(f'Database error: {str(e)}', 'error')
        
        return redirect(url_for('edit_list'))
    
    # GET request - show bulk edit page
    conn = get_db()
    cursor = conn.cursor()
    
    # Get recent edits
    cursor.execute('''
        SELECT a.*, u.name as user_name
        FROM attendance_logs a
        LEFT JOIN users u ON a.user_id = u.user_id AND a.device_floor = u.device_floor
        WHERE a.is_manually_edited = 1
        ORDER BY a.last_edited_at DESC
        LIMIT 10
    ''')
    recent_edits = cursor.fetchall()
    conn.close()
    
    return render_template('edit_bulk.html', recent_edits=recent_edits)

# ============ CONTEXT PROCESSOR ============

@app.context_processor
def utility_processor():
    """Make variables available to all templates"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as total FROM attendance_logs')
    total_logs = cursor.fetchone()['total']
    cursor.execute('SELECT COUNT(*) as total FROM users')
    total_users = cursor.fetchone()['total']
    conn.close()
    return {'total_logs': total_logs, 'total_users': total_users}

# ============ MAIN ============

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5015)
[edit_app.py](https://github.com/user-attachments/files/29495739/edit_app.py)













{"expiry": "2026-08-28 16:16:00", "created": "2026-06-29 16:16:00"}

[activate.py# Open in Notepad.txt](https://github.com/user-attachments/files/29495753/activate.py.Open.in.Notepad.txt)


#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
License Activation System
Hidden - Only accessible via direct URL
Activation Key: azeem@taizco.com
"""

import os
import json
from datetime import datetime, timedelta
from flask import Flask, render_template_string, request, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = 'activation-secret-key-2026'

LICENSE_FILE = '.license'
ACTIVATION_KEY = 'azeem@taizco.com'
EXPIRY_DAYS = 60

# HTML Templates (embedded for standalone)
LOGIN_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>License Activation</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f0f0f0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }
        .container {
            background: white;
            padding: 40px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
            max-width: 400px;
            width: 100%;
        }
        .logo {
            font-size: 48px;
            margin-bottom: 10px;
        }
        .company {
            font-size: 20px;
            font-weight: bold;
            color: #333;
            margin-bottom: 20px;
        }
        .status {
            font-size: 18px;
            color: #666;
            margin-bottom: 20px;
        }
        .input-group {
            margin: 15px 0;
        }
        .input-group input {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
        }
        .btn {
            background: #1a237e;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 4px;
            font-size: 16px;
            cursor: pointer;
            width: 100%;
        }
        .btn:hover {
            background: #0d1442;
        }
        .error {
            color: #c62828;
            margin: 10px 0;
        }
        .success {
            color: #2e7d32;
            margin: 10px 0;
        }
        .footer {
            margin-top: 20px;
            font-size: 12px;
            color: #999;
        }
        .info {
            background: #e3f2fd;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
            font-size: 14px;
            color: #0d47a1;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">🅐∞🅐</div>
        <div class="company">USMAN ENTERPRISE</div>
        <div class="status">{{ status }}</div>
        
        {% if expired %}
            <div class="error">⚠️ License has expired</div>
            <div class="info">Please enter activation key to continue</div>
        {% endif %}
        
        <form method="POST">
            <div class="input-group">
                <input type="password" name="key" placeholder="Enter Activation Key" required>
            </div>
            <button type="submit" class="btn">Activate</button>
        </form>
        
        {% if error %}
            <div class="error">{{ error }}</div>
        {% endif %}
        {% if success %}
            <div class="success">{{ success }}</div>
        {% endif %}
        
        <div class="footer">All Rights Reserved ® 2026</div>
    </div>
</body>
</html>
'''

def get_license_data():
    """Read license file"""
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
                return data
        except:
            return None
    return None

def save_license_data(expiry_date):
    """Save license data"""
    data = {
        'expiry': expiry_date.strftime('%Y-%m-%d %H:%M:%S'),
        'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    with open(LICENSE_FILE, 'w') as f:
        json.dump(data, f)

def is_license_valid():
    """Check if license is valid"""
    data = get_license_data()
    if not data:
        return False, None
    
    try:
        expiry = datetime.strptime(data['expiry'], '%Y-%m-%d %H:%M:%S')
        now = datetime.now()
        if now <= expiry:
            return True, expiry
        else:
            return False, expiry
    except:
        return False, None

@app.route('/', methods=['GET', 'POST'])
def activate():
    """Main activation page"""
    valid, expiry = is_license_valid()
    
    if request.method == 'POST':
        key = request.form.get('key', '').strip()
        
        if key == ACTIVATION_KEY:
            # Activation successful - reset expiry
            new_expiry = datetime.now() + timedelta(days=EXPIRY_DAYS)
            save_license_data(new_expiry)
            return render_template_string(LOGIN_TEMPLATE,
                                        status='✅ License Active',
                                        expired=False,
                                        success=f'License activated successfully until {new_expiry.strftime("%Y-%m-%d")}',
                                        error=None)
        else:
            return render_template_string(LOGIN_TEMPLATE,
                                        status='❌ License Expired' if not valid else '✅ License Active',
                                        expired=not valid,
                                        success=None,
                                        error='Invalid activation key')
    
    if valid:
        days_left = (expiry - datetime.now()).days
        status = f'✅ License Active ({days_left} days remaining)'
        return render_template_string(LOGIN_TEMPLATE,
                                    status=status,
                                    expired=False,
                                    success=None,
                                    error=None)
    else:
        return render_template_string(LOGIN_TEMPLATE,
                                    status='❌ License Expired',
                                    expired=True,
                                    success=None,
                                    error=None)

@app.route('/check')
def check():
    """Hidden check endpoint - returns status"""
    valid, expiry = get_license_data()
    if valid:
        days_left = (expiry - datetime.now()).days
        return f'Active - {days_left} days remaining'
    else:
        return 'Expired'

@app.route('/reset')
def reset():
    """Hidden reset endpoint - delete license file"""
    if os.path.exists(LICENSE_FILE):
        os.remove(LICENSE_FILE)
        return 'License reset successfully'
    return 'No license file found'

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5020)










Flask==2.3.3
Flask-SQLAlchemy==3.1.1
openpyxl==3.1.2
reportlab==4.2.5
zk


[requirements.txt](https://github.com/user-attachments/files/29495761/requirements.txt)






[fetch_7days.py](https://github.com/user-attachments/files/29495770/fetch_7days.py)
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Attendance Data Fetcher - Last 7 Days
Fetches attendance data from 2nd and 3rd floor devices
Preserves all manually edited records
"""

import subprocess
import time
import os
import sys
import sqlite3
from datetime import datetime, timedelta
from collections import defaultdict
from zk import ZK

# Import configuration
from config import WIFI_CONFIG, FETCH_DAYS_BACK, DATABASE

# =================================================
# HELPER FUNCTIONS
# =================================================

def print_status(message, status="info"):
    """Print formatted status messages"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if status == "success":
        print(f"[{timestamp}] [OK] {message}")
    elif status == "error":
        print(f"[{timestamp}] [ERROR] {message}")
    elif status == "warning":
        print(f"[{timestamp}] [WARN] {message}")
    elif status == "info":
        print(f"[{timestamp}] [INFO] {message}")
    else:
        print(f"[{timestamp}] {message}")

def get_current_wifi():
    """Get current WiFi SSID on Windows"""
    try:
        result = subprocess.run(
            ['netsh', 'wlan', 'show', 'interfaces'],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        for line in result.stdout.split('\n'):
            if 'SSID' in line and 'BSSID' not in line:
                ssid = line.split(':')[1].strip()
                if ssid and ssid != '':
                    print_status(f"Current WiFi: {ssid}", "info")
                    return ssid
        return None
    except Exception as e:
        print_status(f"Error getting current WiFi: {e}", "error")
        return None

def switch_wifi(ssid, password=None):
    """Switch WiFi connection to specified SSID"""
    try:
        # Check if already connected
        current = get_current_wifi()
        if current == ssid:
            print_status(f"Already connected to {ssid}", "success")
            return True

        print_status(f"Switching from '{current}' to '{ssid}'...", "info")

        # Disconnect current WiFi
        subprocess.run(['netsh', 'wlan', 'disconnect'], capture_output=True)
        time.sleep(3)

        if password:
            # Create temp directory
            temp_dir = os.path.join(os.environ.get('TEMP', 'C:\\temp'), 'attendance_wifi')
            os.makedirs(temp_dir, exist_ok=True)

            # Create proper XML profile
            profile_xml = f'''<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{ssid}</name>
    <SSIDConfig>
        <SSID>
            <name>{ssid}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{password}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>'''

            profile_path = os.path.join(temp_dir, f"{ssid}.xml")
            with open(profile_path, 'w') as f:
                f.write(profile_xml)

            # Add profile
            subprocess.run(
                ['netsh', 'wlan', 'add', 'profile', f'filename={profile_path}'],
                capture_output=True,
                text=True
            )
            time.sleep(2)

            # Connect
            subprocess.run(
                ['netsh', 'wlan', 'connect', f'name={ssid}'],
                capture_output=True,
                text=True
            )
            time.sleep(8)

            # Clean up temp file
            try:
                os.remove(profile_path)
            except:
                pass

        else:
            # Open network (no password)
            subprocess.run(
                ['netsh', 'wlan', 'connect', f'name={ssid}'],
                capture_output=True,
                text=True
            )
            time.sleep(8)

        # Verify connection
        new_current = get_current_wifi()
        if new_current == ssid:
            print_status(f"Successfully connected to {ssid}", "success")
            return True
        else:
            print_status(f"Failed to connect to {ssid}", "error")
            return False

    except Exception as e:
        print_status(f"WiFi switch error: {e}", "error")
        return False

def get_fetch_cutoff_time():
    """Return datetime to fetch records from (last FETCH_DAYS_BACK days)"""
    cutoff = datetime.now() - timedelta(days=FETCH_DAYS_BACK)
    print_status(f"Fetching records from: {cutoff}", "info")
    return cutoff

def init_database():
    """Initialize database tables if they don't exist"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    cursor = conn.cursor()

    # Users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            name TEXT NOT NULL,
            device_floor TEXT NOT NULL,
            last_sync TIMESTAMP,
            is_active INTEGER DEFAULT 1,
            UNIQUE(user_id, device_floor)
        )
    ''')

    # Attendance logs table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attendance_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            timestamp TIMESTAMP NOT NULL,
            device_floor TEXT NOT NULL,
            punch_type INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            is_manually_edited INTEGER DEFAULT 0,
            last_edited_at TIMESTAMP,
            remark_in TEXT,
            remark_out TEXT,
            status_override TEXT,
            UNIQUE(user_id, timestamp, device_floor)
        )
    ''')

    # Create indexes for performance
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_attendance_timestamp ON attendance_logs(timestamp)')
    cursor.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_unique_attendance ON attendance_logs(user_id, timestamp, device_floor)')

    conn.commit()
    conn.close()
    print_status("Database ready", "success")

def save_user_to_db(user_id, name, floor):
    """Save or update user in database"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT is_active FROM users 
        WHERE user_id = ? AND device_floor = ?
    ''', (user_id, floor))
    existing = cursor.fetchone()

    if existing:
        cursor.execute('''
            UPDATE users 
            SET name = ?, last_sync = ?
            WHERE user_id = ? AND device_floor = ?
        ''', (name, datetime.now(), user_id, floor))
        print_status(f"Updated user: {name} (ID:{user_id}, Floor:{floor})", "info")
    else:
        cursor.execute('''
            INSERT INTO users (user_id, name, device_floor, last_sync, is_active)
            VALUES (?, ?, ?, ?, 1)
        ''', (user_id, name, floor, datetime.now()))
        print_status(f"Added new user: {name} (ID:{user_id}, Floor:{floor})", "success")

    conn.commit()
    conn.close()

def auto_calculate_remark_in(time_str, day_name):
    """Auto-calculate IN remark based on time"""
    if not time_str or time_str == '--':
        return 'MISSING IN'

    try:
        if ':' in str(time_str):
            parts = str(time_str).split(':')
            hour = int(parts[0]) if parts[0].isdigit() else 0
            minute = int(parts[1][:2]) if parts[1][:2].isdigit() else 0
        else:
            return 'ON TIME'

        check_in = datetime.strptime(f"{hour:02d}:{minute:02d}", '%H:%M').time()

        if day_name == 'Saturday':
            start = datetime.strptime('10:00', '%H:%M').time()
            grace_end = datetime.strptime('10:30', '%H:%M').time()
        else:
            start = datetime.strptime('09:15', '%H:%M').time()
            grace_end = datetime.strptime('09:30', '%H:%M').time()

        if check_in <= start:
            return 'ON TIME'
        elif check_in <= grace_end:
            return 'GRACE'
        else:
            return 'LATE'
    except:
        return 'ON TIME'

def auto_calculate_remark_out(time_str, day_name):
    """Auto-calculate OUT remark based on time"""
    if not time_str or time_str == '--':
        return 'MISSING OUT'

    try:
        if ':' in str(time_str):
            parts = str(time_str).split(':')
            hour = int(parts[0]) if parts[0].isdigit() else 0
            minute = int(parts[1][:2]) if parts[1][:2].isdigit() else 0
        else:
            return 'ON TIME'

        check_out = datetime.strptime(f"{hour:02d}:{minute:02d}", '%H:%M').time()

        if day_name == 'Saturday':
            early = datetime.strptime('14:00', '%H:%M').time()
            overtime = datetime.strptime('16:00', '%H:%M').time()
        else:
            early = datetime.strptime('17:00', '%H:%M').time()
            overtime = datetime.strptime('18:00', '%H:%M').time()

        if check_out > overtime:
            return 'OVERTIME'
        elif check_out < early:
            return 'EARLY DEPT'
        else:
            return 'ON TIME'
    except:
        return 'ON TIME'

def clean_and_save_attendance(records_by_day, floor):
    """Clean and save attendance records to database"""
    conn = sqlite3.connect(DATABASE, timeout=20)
    cursor = conn.cursor()

    total_kept = 0
    total_skipped_edited = 0

    for key, record_data in records_by_day.items():
        user_id = record_data['user_id']
        date_str = record_data['date']

        # Check if any manually edited record exists for this user/date/floor
        cursor.execute('''
            SELECT COUNT(*) FROM attendance_logs 
            WHERE user_id = ? AND device_floor = ? AND DATE(timestamp) = ? AND is_manually_edited = 1
        ''', (user_id, floor, date_str))
        edited_count = cursor.fetchone()[0]

        if edited_count > 0:
            print_status(f"Skipping {user_id} on {date_str} - has manually edited records", "warning")
            total_skipped_edited += 1
            continue

        # Delete ALL existing records for this user/date/floor (only auto records)
        cursor.execute('''
            DELETE FROM attendance_logs 
            WHERE user_id = ? AND device_floor = ? AND DATE(timestamp) = ? AND is_manually_edited = 0
        ''', (user_id, floor, date_str))

        punches = record_data.get('punches', [])

        # Find earliest IN and latest OUT
        check_in_time = None
        check_out_time = None
        check_in_raw = None
        check_out_raw = None

        for punch in punches:
            if punch['punch_type'] == 0:  # IN
                if check_in_time is None or punch['timestamp'] < check_in_time:
                    check_in_time = punch['timestamp']
                    check_in_raw = punch['raw_time']
            else:  # OUT
                if check_out_time is None or punch['timestamp'] > check_out_time:
                    check_out_time = punch['timestamp']
                    check_out_raw = punch['raw_time']

        day_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%A')

        # Insert check-in
        if check_in_time:
            remark_in = auto_calculate_remark_in(check_in_raw, day_name)
            cursor.execute('''
                INSERT OR REPLACE INTO attendance_logs 
                (user_id, timestamp, device_floor, punch_type, is_manually_edited, remark_in)
                VALUES (?, ?, ?, 0, 0, ?)
            ''', (user_id, check_in_time, floor, remark_in))
            total_kept += 1

        # Insert check-out
        if check_out_time:
            remark_out = auto_calculate_remark_out(check_out_raw, day_name)
            cursor.execute('''
                INSERT OR REPLACE INTO attendance_logs 
                (user_id, timestamp, device_floor, punch_type, is_manually_edited, remark_out)
                VALUES (?, ?, ?, 1, 0, ?)
            ''', (user_id, check_out_time, floor, remark_out))
            total_kept += 1

    conn.commit()
    conn.close()
    return total_kept, total_skipped_edited

def fetch_device_data(floor):
    """Fetch attendance data from a specific floor device"""
    config = WIFI_CONFIG[floor]

    print(f"\n{'='*60}")
    print(f"[FETCHING] {floor.upper()}")
    print(f"  WiFi: {config['ssid']}")
    print(f"  Device IP: {config['device_ip']}")
    print(f"  Fetching last {FETCH_DAYS_BACK} days of data")
    print(f"{'='*60}")

    # Switch WiFi
    if not switch_wifi(config['ssid'], config['password']):
        raise Exception(f"Cannot connect to WiFi '{config['ssid']}' for {floor}")

    print_status(f"Connecting to device at {config['device_ip']}...", "info")

    try:
        zk = ZK(
            config['device_ip'],
            port=config['device_port'],
            timeout=10,
            password=config['device_password'],
            force_udp=False,
            omit_ping=True
        )
        device_conn = zk.connect()

        if not device_conn:
            raise Exception(f"Cannot connect to device at {config['device_ip']} for {floor}")

        print_status("Connected to device successfully!", "success")

        # Disable device for data transfer
        try:
            device_conn.disable_device()
        except:
            pass

        # Fetch users
        print_status("Fetching users...", "info")
        users = device_conn.get_users()
        user_count = 0
        for user in users:
            if user.name:
                save_user_to_db(str(user.user_id), user.name, floor)
                user_count += 1
        print_status(f"Saved/Updated {user_count} users", "success")

        # Get cutoff time (last FETCH_DAYS_BACK days)
        cutoff_time = get_fetch_cutoff_time()

        # Fetch all logs
        print_status("Fetching attendance logs from device...", "info")
        all_logs = device_conn.get_attendance()

        # Group logs by user_id and date (only records after cutoff)
        records_to_clean = {}
        total_fetched = 0

        for log in all_logs:
            if log.timestamp >= cutoff_time:
                user_id = str(log.user_id)
                date_str = log.timestamp.strftime('%Y-%m-%d')
                key = f"{user_id}_{date_str}"

                if key not in records_to_clean:
                    records_to_clean[key] = {
                        'user_id': user_id,
                        'date': date_str,
                        'punches': []
                    }

                raw_time = log.timestamp.strftime('%H:%M')
                records_to_clean[key]['punches'].append({
                    'timestamp': log.timestamp,
                    'punch_type': log.punch,
                    'raw_time': raw_time
                })
                total_fetched += 1

        print_status(f"Fetched {total_fetched} records from device", "success")

        # Process and save
        cleaned_count, skipped_clean = clean_and_save_attendance(records_to_clean, floor)

        # Re-enable device
        try:
            device_conn.enable_device()
        except:
            pass

        device_conn.disconnect()

        print(f"\n[RESULTS] for {floor}:")
        print_status(f"Records fetched and cleaned: {cleaned_count} kept", "success")
        if skipped_clean > 0:
            print_status(f"Skipped (manually edited): {skipped_clean} users/days", "warning")

        return cleaned_count, user_count, total_fetched

    except Exception as e:
        print_status(f"Error fetching from {floor}: {e}", "error")
        raise

def main():
    """Main execution function"""
    print("\n" + "="*60)
    print("ATTENDANCE FETCHER - 7 DAYS INCREMENTAL MODE")
    print(f"Fetching last {FETCH_DAYS_BACK} days of data")
    print("Rule: Keep earliest IN, latest OUT per day")
    print("Manually edited records will NOT be overwritten")
    print("="*60)

    # Save original WiFi
    original_wifi = get_current_wifi()
    if original_wifi:
        print_status(f"Original WiFi saved: {original_wifi}", "info")
    else:
        print_status("No original WiFi detected", "warning")

    # Initialize database
    init_database()

    try:
        # Fetch 2nd floor
        print("\n" + "-" * 30)
        print_status("STEP 1: Fetching 2nd Floor Data", "info")
        floor2_total, floor2_users, floor2_fetched = fetch_device_data('2nd_floor')

        # Fetch 3rd floor
        print("\n" + "-" * 30)
        print_status("STEP 2: Fetching 3rd Floor Data", "info")
        floor3_total, floor3_users, floor3_fetched = fetch_device_data('3rd_floor')

    except Exception as e:
        print(f"\n{'='*60}")
        print_status(f"FETCH STOPPED: {e}", "error")
        print(f"{'='*60}")

        # Restore original WiFi
        if original_wifi:
            print_status(f"Attempting to restore original WiFi: {original_wifi}", "info")
            switch_wifi(original_wifi)

        sys.exit(1)

    # Restore original WiFi
    if original_wifi:
        print_status(f"Restoring original WiFi: {original_wifi}", "info")
        switch_wifi(original_wifi)
    else:
        print_status("No original WiFi to restore", "warning")

    # Final statistics
    conn = sqlite3.connect(DATABASE, timeout=20)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM attendance_logs")
    total_logs = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM attendance_logs WHERE is_manually_edited = 1")
    edited_count = cursor.fetchone()[0]
    conn.close()

    print("\n" + "="*60)
    print_status("FETCH COMPLETE!", "success")
    print("="*60)
    print(f"2nd Floor: {floor2_total} records added, {floor2_users} users, {floor2_fetched} fetched")
    print(f"3rd Floor: {floor3_total} records added, {floor3_users} users, {floor3_fetched} fetched")
    print(f"Total NEW records added: {floor2_total + floor3_total}")
    print(f"Total Users in DB: {total_users}")
    print(f"Total Attendance Logs in DB: {total_logs}")
    print(f"Manually Edited Records (protected): {edited_count}")
    print(f"Original WiFi restored: {original_wifi}")
    print(f"Fetching last {FETCH_DAYS_BACK} days of data")
    print("="*60)

    sys.exit(0)

if __name__ == '__main__':
    main()



















# ================= CONFIGURATION =================

# WiFi configurations for both floors
WIFI_CONFIG = {
    '2nd_floor': {
        'ssid': '35170741',
        'password': 'usman1234',
        'device_ip': '192.168.100.250',
        'device_port': 4370,
        'device_password': 0
    },
    '3rd_floor': {
        'ssid': '35170742',
        'password': '1122',
        'device_ip': '192.168.68.75',
        'device_port': 4370,
        'device_password': 1122
    }
}

# Fetch last 7 days of data
FETCH_DAYS_BACK = 7

# Database file
DATABASE = 'attendance.db'

# =================================================
    [config.py](https://github.com/user-attachments/files/29495777/config.py)















    
