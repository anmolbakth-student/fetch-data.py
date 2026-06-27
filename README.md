# fetch-data.py
attendance data fetch with wifi switch
import subprocess
import time
import os
from datetime import datetime, timedelta
from zk import ZK
import sqlite3
import sys
from collections import defaultdict

# ================= CONFIGURATION =================
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

# Fetch records from last 48 hours (ensures no missed records)
FETCH_HOURS_BACK = 48
# =================================================

def print_status(message, status="info"):
    if status == "success":
        print(f"[OK] {message}")
    elif status == "error":
        print(f"[ERROR] {message}")
    elif status == "warning":
        print(f"[WARN] {message}")
    elif status == "info":
        print(f"[INFO] {message}")
    else:
        print(message)

def get_current_wifi():
    try:
        result = subprocess.run(['netsh', 'wlan', 'show', 'interfaces'], 
                                capture_output=True, text=True, encoding='utf-8')
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
    try:
        current = get_current_wifi()
        if current == ssid:
            print_status(f"Already connected to {ssid}", "success")
            return True

        print_status(f"Switching from '{current}' to '{ssid}'...", "info")
        
        subprocess.run(['netsh', 'wlan', 'disconnect'], capture_output=True)
        time.sleep(2)
        
        if password:
            os.makedirs("C:\\temp", exist_ok=True)
            
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
            
            profile_path = f"C:\\temp\\{ssid}.xml"
            with open(profile_path, 'w') as f:
                f.write(profile_xml)
            
            subprocess.run(f'netsh wlan add profile filename="{profile_path}"', 
                          shell=True, capture_output=True, text=True)
            time.sleep(2)
            subprocess.run(f'netsh wlan connect name="{ssid}"', 
                          shell=True, capture_output=True, text=True)
            time.sleep(5)
        else:
            cmd = f'netsh wlan connect name="{ssid}" ssid="{ssid}" interface="Wi-Fi"'
            subprocess.run(cmd, shell=True, capture_output=True, text=True)
            time.sleep(5)

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

def init_database():
    conn = sqlite3.connect('attendance.db', timeout=20)
    cursor = conn.cursor()
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
    conn.commit()
    conn.close()
    print_status("Database ready", "success")

def get_fetch_cutoff_time():
    """Return datetime to fetch records from (last FETCH_HOURS_BACK hours)"""
    cutoff = datetime.now() - timedelta(hours=FETCH_HOURS_BACK)
    print_status(f"Fetching records from: {cutoff}", "info")
    return cutoff

def save_user_to_db(user_id, name, floor):
    conn = sqlite3.connect('attendance.db', timeout=20)
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
        print_status(f"Updated user: {name} (ID:{user_id}, Floor:{floor}) - is_active preserved", "info")
    else:
        cursor.execute('''
            INSERT INTO users (user_id, name, device_floor, last_sync, is_active)
            VALUES (?, ?, ?, ?, 1)
        ''', (user_id, name, floor, datetime.now()))
        print_status(f"Added new user: {name} (ID:{user_id}, Floor:{floor})", "success")
    
    conn.commit()
    conn.close()

def auto_calculate_remark_in(time_str, day_name):
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
    conn = sqlite3.connect('attendance.db', timeout=20)
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
        
        # Delete ALL existing records for this user/date/floor
        cursor.execute('''
            DELETE FROM attendance_logs 
            WHERE user_id = ? AND device_floor = ? AND DATE(timestamp) = ?
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
                INSERT INTO attendance_logs (user_id, timestamp, device_floor, punch_type, is_manually_edited, remark_in)
                VALUES (?, ?, ?, 0, 0, ?)
            ''', (user_id, check_in_time, floor, remark_in))
            total_kept += 1
            print_status(f"  Kept IN: {check_in_time} ({remark_in}) for {user_id}", "info")
        
        # Insert check-out
        if check_out_time:
            remark_out = auto_calculate_remark_out(check_out_raw, day_name)
            cursor.execute('''
                INSERT INTO attendance_logs (user_id, timestamp, device_floor, punch_type, is_manually_edited, remark_out)
                VALUES (?, ?, ?, 1, 0, ?)
            ''', (user_id, check_out_time, floor, remark_out))
            total_kept += 1
            print_status(f"  Kept OUT: {check_out_time} ({remark_out}) for {user_id}", "info")
    
    conn.commit()
    conn.close()
    return total_kept, total_skipped_edited

def fetch_device_data(floor):
    config = WIFI_CONFIG[floor]
    print(f"\n{'='*60}")
    print(f"[FETCHING] {floor.upper()}")
    print(f"   WiFi: {config['ssid']}")
    print(f"   Device IP: {config['device_ip']}")
    print(f"   Fetching last {FETCH_HOURS_BACK} hours of data")
    print(f"{'='*60}")

    if not switch_wifi(config['ssid'], config['password']):
        raise Exception(f"Cannot connect to WiFi '{config['ssid']}' for {floor}")

    print_status(f"Connecting to device at {config['device_ip']}...", "info")
    zk = ZK(config['device_ip'], port=config['device_port'], timeout=5, 
             password=config['device_password'], force_udp=False, ommit_ping=True)
    device_conn = zk.connect()

    if not device_conn:
        raise Exception(f"Cannot connect to device at {config['device_ip']} for {floor}")

    print_status("Connected to device successfully!", "success")

    try:
        device_conn.disable_device()
    except:
        pass

    print_status("Fetching users...", "info")
    users = device_conn.get_users()
    user_count = 0
    for user in users:
        if user.name:
            save_user_to_db(str(user.user_id), user.name, floor)
            user_count += 1
    print_status(f"Saved/Updated {user_count} users", "success")

    # Get cutoff time (last FETCH_HOURS_BACK hours)
    cutoff_time = get_fetch_cutoff_time()
    
    print_status(f"Fetching ALL logs from device...", "info")
    all_logs = device_conn.get_attendance()

    # Group logs by user_id and date (only for records after cutoff_time)
    records_to_clean = {}
    
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
            
            # Store raw time string for remark calculation
            raw_time = log.timestamp.strftime('%H:%M')
            
            records_to_clean[key]['punches'].append({
                'timestamp': log.timestamp,
                'punch_type': log.punch,
                'raw_time': raw_time
            })

    # Process with deduplication (earliest IN, latest OUT)
    cleaned_count, skipped_clean = clean_and_save_attendance(records_to_clean, floor)

    try:
        device_conn.enable_device()
    except:
        pass

    device_conn.disconnect()

    print(f"\n[RESULTS] for {floor}:")
    print_status(f"Records fetched and cleaned: {cleaned_count} kept", "success")
    if skipped_clean > 0:
        print_status(f"Skipped (manually edited): {skipped_clean} users/days", "warning")
    
    return cleaned_count, user_count

def main():
    print("\n" + "="*60)
    print("ATTENDANCE FETCHER - INCREMENTAL MODE")
    print(f"Fetching last {FETCH_HOURS_BACK} hours of data")
    print("Rule: Keep earliest IN, latest OUT per day")
    print("Manually edited records will NOT be overwritten")
    print("="*60)

    original_wifi = get_current_wifi()
    if original_wifi:
        print_status(f"Original WiFi saved: {original_wifi}", "info")
    else:
        print_status("No original WiFi detected", "warning")

    init_database()

    try:
        print("\n" + "-" * 30)
        print_status("STEP 1: Fetching 2nd Floor Data", "info")
        floor2_total, floor2_users = fetch_device_data('2nd_floor')
        
        print("\n" + "-" * 30)
        print_status("STEP 2: Fetching 3rd Floor Data", "info")
        floor3_total, floor3_users = fetch_device_data('3rd_floor')
        
    except Exception as e:
        print(f"\n{'='*60}")
        print_status(f"FETCH STOPPED: {e}", "error")
        print(f"{'='*60}")
        
        if original_wifi:
            print_status(f"Attempting to restore original WiFi: {original_wifi}", "info")
            switch_wifi(original_wifi)
        
        sys.exit(1)

    if original_wifi:
        print_status(f"Restoring original WiFi: {original_wifi}", "info")
        switch_wifi(original_wifi)
    else:
        print_status("No original WiFi to restore", "warning")

    conn = sqlite3.connect('attendance.db', timeout=20)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM attendance_logs")
    total_logs = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM attendance_logs WHERE is_manually_edited = 1")
    edited_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM users WHERE is_active = 0")
    inactive_count = cursor.fetchone()[0]
    conn.close()

    print("\n" + "="*60)
    print_status("FETCH COMPLETE!", "success")
    print("="*60)
    print(f"2nd Floor: {floor2_total} records added, {floor2_users} users")
    print(f"3rd Floor: {floor3_total} records added, {floor3_users} users")
    print(f"Total NEW records added: {floor2_total + floor3_total}")
    print(f"Total Users in DB: {total_users}")
    print(f"  - Active users: {total_users - inactive_count}")
    print(f"  - Inactive users: {inactive_count}")
    print(f"Total Attendance Logs in DB: {total_logs}")
    print(f"Manually Edited Records (protected): {edited_count}")
    print(f"Original WiFi restored: {original_wifi}")
    print(f"Fetching last {FETCH_HOURS_BACK} hours of data")
    print("="*60)
    
    sys.exit(0)

if __name__ == '__main__':
    os.makedirs("C:\\temp", exist_ok=True)
    main()
