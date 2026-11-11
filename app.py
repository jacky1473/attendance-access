#!/usr/bin/env python3
"""
easyWDMS Web GUI Dashboard
Runs on port 8083, connects to easyWDMS, provides In-Out punch analytics and Excel download.
"""

import sys
import os
import re
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify, send_file
import tempfile

# Add local directory and /root to sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
sys.path.insert(0, '/root')

from easywdms_report import EasyWDMSClient, calculate_in_out, export_to_excel, parse_date_range, safe_emp_sort_key

app = Flask(__name__, template_folder=os.path.join(SCRIPT_DIR, 'templates'))

# Global client and cached employees
client = None
cached_employees = {}
last_cache_time = None

def find_creds_path():
    candidates = [
        os.environ.get('EASYWDMS_CREDS'),
        os.path.join(SCRIPT_DIR, 'easyWDMS.txt'),
        '/root/easyWDMS.txt',
        os.path.expanduser('~/easyWDMS.txt')
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return os.path.join(SCRIPT_DIR, 'easyWDMS.txt')

def get_client():
    global client
    if client is None or not getattr(client, 'is_logged_in', False):
        host = os.environ.get('EASYWDMS_HOST', "http://122.176.49.74:8082")
        creds = find_creds_path()
        c = EasyWDMSClient(host=host, creds_path=creds)
        c.login()
        client = c
    return client

def get_cached_employees(force_refresh=False):
    global cached_employees, last_cache_time
    now = datetime.now()
    if force_refresh or not cached_employees or (last_cache_time and (now - last_cache_time).total_seconds() > 3600):
        cl = get_client()
        try:
            emps = cl.get_employees()
            if emps:
                cached_employees = emps
                last_cache_time = now
        except Exception as e:
            # Re-login and retry
            print(f"Failed to fetch employees: {e}, retrying with fresh login...")
            cl.login()
            emps = cl.get_employees()
            if emps:
                cached_employees = emps
                last_cache_time = now
    return cached_employees

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/status')
def api_status():
    try:
        cl = get_client()
        emps = get_cached_employees()
        return jsonify({
            'status': 'connected',
            'host': cl.host,
            'company': cl.company_name,
            'username': cl.username,
            'total_employees': len(emps)
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'error': str(e)
        }), 500

@app.route('/api/employees')
def api_employees():
    try:
        emps = get_cached_employees(force_refresh=request.args.get('refresh') == 'true')
        emp_list = sorted(list(emps.values()), key=lambda e: safe_emp_sort_key(e['emp_code']))
        return jsonify({
            'success': True,
            'count': len(emp_list),
            'employees': emp_list
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/areas')
def api_areas():
    try:
        cl = get_client()
        terminals = cl.get_terminals()
        area_map = {}
        for t in terminals:
            alias = (t.get('alias') or t.get('area_name') or 'Unknown').strip()
            if alias and alias not in area_map:
                area_map[alias] = {
                    'code': alias,
                    'name': alias,
                    'alias': alias,
                    'sn': t.get('sn', ''),
                    'ip': t.get('ip_address', ''),
                    'user_count': t.get('user_count', 0)
                }
        area_list = list(area_map.values())
        return jsonify({
            'success': True,
            'count': len(area_list),
            'areas': area_list
        })
    except Exception as e:
        fallback_areas = [
            {'code': 'MAIN1', 'name': 'MAIN1', 'alias': 'MAIN1', 'ip': '192.168.12.223'},
            {'code': 'MAIN 2', 'name': 'MAIN 2', 'alias': 'MAIN 2', 'ip': '192.168.12.222'},
            {'code': 'E & TELESALES', 'name': 'E & TELESALES', 'alias': 'E & TELESALES', 'ip': '192.168.12.224'},
            {'code': 'E & CEASED', 'name': 'E & CEASED', 'alias': 'E & CEASED', 'ip': '192.168.12.221'}
        ]
        return jsonify({
            'success': True,
            'count': len(fallback_areas),
            'areas': fallback_areas
        })

@app.route('/api/preview', methods=['POST'])
def api_preview():
    try:
        data = request.get_json() or {}
        emp_code = data.get('emp_code')
        if emp_code == 'ALL' or emp_code == '' or not emp_code:
            emp_code = None
        else:
            emp_code = str(emp_code).strip()

        # Parse allowed areas/machines
        raw_areas = data.get('areas')
        allowed_areas = None
        if raw_areas:
            if isinstance(raw_areas, list):
                allowed_areas = [str(x).strip() for x in raw_areas if str(x).strip()]
            elif isinstance(raw_areas, str):
                allowed_areas = [x.strip() for x in raw_areas.split(',') if x.strip()]

        start_date_str = data.get('start_date')
        end_date_str = data.get('end_date')

        if not start_date_str or not end_date_str:
            return jsonify({'success': False, 'error': 'start_date and end_date are required'}), 400

        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()

        if start_date > end_date:
            return jsonify({'success': False, 'error': 'From Date cannot be after To Date'}), 400

        cl = get_client()
        emps = get_cached_employees()

        single_emp = None
        if emp_code:
            single_emp = emps.get(emp_code)
            if not single_emp:
                single_emp = {'emp_code': emp_code, 'name': f"Employee {emp_code}", 'department': '-'}

        # Fetch punches with auto re-login fallback
        try:
            transactions = cl.fetch_transactions(start_date_str, end_date_str, emp_code=emp_code)
        except Exception:
            cl.login()
            transactions = cl.fetch_transactions(start_date_str, end_date_str, emp_code=emp_code)

        daily_records, summary_records = calculate_in_out(
            transactions, emps, start_date, end_date, single_emp=single_emp, allowed_areas=allowed_areas
        )

        # Aggregate stats (Clean present / single punches, no salary penalty metrics)
        total_present_days = 0
        total_single_punch_days = 0

        for s in summary_records.values():
            total_present_days += s.get('present_days', 0)
            total_single_punch_days += s.get('single_punch_days', 0)

        return jsonify({
            'success': True,
            'count': len(daily_records),
            'records': daily_records,
            'summary': list(summary_records.values()),
            'stats': {
                'total_records': len(daily_records),
                'present_days': total_present_days,
                'single_punch_days': total_single_punch_days,
                'employees_count': len(summary_records)
            }
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/export')
def api_export():
    try:
        emp_code = request.args.get('emp_code')
        is_all = request.args.get('all') == 'true' or not emp_code or emp_code == 'ALL'
        if is_all:
            emp_code = None
        else:
            emp_code = str(emp_code).strip()

        start_date_str = request.args.get('start_date')
        end_date_str = request.args.get('end_date')

        raw_areas = request.args.get('areas')
        allowed_areas = None
        if raw_areas:
            allowed_areas = [x.strip() for x in raw_areas.split(',') if x.strip()]

        if not start_date_str or not end_date_str:
            return "start_date and end_date are required", 400

        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()

        cl = get_client()
        emps = get_cached_employees()

        single_emp = None
        if emp_code:
            single_emp = emps.get(emp_code)
            if not single_emp:
                single_emp = {'emp_code': emp_code, 'name': f"Employee {emp_code}", 'department': '-'}

        try:
            transactions = cl.fetch_transactions(start_date_str, end_date_str, emp_code=emp_code)
        except Exception:
            cl.login()
            transactions = cl.fetch_transactions(start_date_str, end_date_str, emp_code=emp_code)

        daily_records, summary_records = calculate_in_out(
            transactions, emps, start_date, end_date, single_emp=single_emp, allowed_areas=allowed_areas
        )

        s_str = start_date.strftime('%Y%m%d')
        e_str = end_date.strftime('%Y%m%d')

        if single_emp:
            clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', single_emp['name'])
            filename = f"attendance_{single_emp['emp_code']}_{clean_name}_{s_str}_{e_str}.xlsx"
        else:
            filename = f"attendance_ALL_{s_str}_{e_str}.xlsx"

        export_dir = "/tmp/easywdms_exports"
        os.makedirs(export_dir, exist_ok=True)
        output_path = os.path.join(export_dir, filename)

        machines_display = ", ".join(allowed_areas) if allowed_areas and 'ALL' not in allowed_areas else ('All Machines' if allowed_areas and 'ALL' in allowed_areas else 'MAIN1 & MAIN 2')

        meta_info = {
            'company': cl.company_name,
            'period': f"{start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}",
            'employee_filter': f"[{single_emp['emp_code']}] {single_emp['name']}" if single_emp else "All Employees",
            'machines': machines_display
        }

        export_to_excel(daily_records, summary_records, output_path, meta_info)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return f"Export error: {e}", 500

if __name__ == '__main__':
    # Initialize client on startup
    try:
        cl = get_client()
        print(f"easyWDMS Client connected to {cl.host} ({cl.company_name})")
        get_cached_employees()
    except Exception as e:
        print(f"Initial login warning: {e}")

    # Listen on all interfaces on port 8083
    print("Starting easyWDMS Web GUI on http://0.0.0.0:8083 ...")
    app.run(host='0.0.0.0', port=8083, debug=False)
