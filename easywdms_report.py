#!/usr/bin/env python3
"""
easyWDMS Biometric Attendance In-Out Report Tool
Fetches raw biometric punches from easyWDMS and exports clean First-In & Last-Out reports to Excel.
Customized for:
1. Devices: MAIN & MAIN 2 only (MAIN1, MAIN 2)
2. Shift: Morning Shift (10:00 AM - 07:00 PM) with Full Day (>=8h) and Half Day (>=4h & <8h) logic.
"""

import sys
import os
import re
import base64
import urllib.parse
import argparse
from datetime import datetime, date, timedelta
from collections import defaultdict
import requests

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: openpyxl is required. Run 'python3 -m pip install openpyxl'")
    sys.exit(1)


# ==========================================
# RC4 Encryption used by easyWDMS login
# ==========================================
def zk_encrypt(key: str, data: str) -> str:
    s = list(range(256))
    j = 0
    key_len = len(key)
    for i in range(256):
        j = (j + s[i] + ord(key[i % key_len])) % 256
        s[i], s[j] = s[j], s[i]

    i = 0
    j = 0
    res = bytearray()
    for char in data:
        i = (i + 1) % 256
        j = (j + s[i]) % 256
        s[i], s[j] = s[j], s[i]
        res.append((ord(char) & 0xff) ^ s[(s[i] + s[j]) % 256])
    return base64.b64encode(res).decode('ascii')


def safe_emp_sort_key(c):
    c_str = str(c)
    return (0, int(c_str)) if c_str.isdigit() else (1, c_str)


def is_main_terminal(t):
    """Filter to ensure only MAIN1 and MAIN 2 machines are counted."""
    alias = (t.get('terminal_alias') or '').strip().upper()
    sn = (t.get('terminal_sn') or '').strip()
    return 'MAIN' in alias or sn in ['TDBD254200096', 'TDBD241101196']


def is_terminal_allowed(t, allowed_areas=None):
    """
    Filter to check if a punch transaction comes from an allowed terminal / area.
    If allowed_areas is None: defaults to MAIN1 and MAIN 2.
    If allowed_areas contains 'ALL' or '__ALL__': allows all machines.
    Otherwise, checks if terminal alias, area, or SN matches any in allowed_areas.
    """
    if allowed_areas is None:
        return is_main_terminal(t)

    if any(str(a).strip().upper() in ['ALL', '__ALL__'] for a in allowed_areas if a):
        return True

    alias = (t.get('terminal_alias') or '').strip().upper()
    sn = (t.get('terminal_sn') or '').strip().upper()
    area = (t.get('area_alias') or '').strip().upper()

    for a in allowed_areas:
        if not a:
            continue
        ac = str(a).strip().upper()
        if ac == alias or ac == area or ac == sn or ac in alias or ac in area:
            return True
        if ac == 'MAIN' and ('MAIN1' in alias or 'MAIN 2' in alias):
            return True
        if ac in ['TELESALES', 'E & TELESALES'] and 'TELESALES' in alias:
            return True
        if ac in ['CEASED', 'E & CEASED'] and 'CEASED' in alias:
            return True

    return False


# ==========================================
# easyWDMS API Client
# ==========================================
class EasyWDMSClient:
    def __init__(self, host="http://122.176.49.74:8082", creds_path="/root/easyWDMS.txt"):
        self.host = host.rstrip('/')
        self.creds_path = creds_path
        self.session = requests.Session()
        self.company_name = "easyWDMS"
        self.username = None
        self.password = None
        self.is_logged_in = False

    def load_credentials(self, username=None, password=None):
        if username and password:
            self.username = username
            self.password = password
            return

        path_to_try = self.creds_path
        if not path_to_try or not os.path.exists(path_to_try):
            candidates = [
                os.environ.get("EASYWDMS_CREDS"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "easyWDMS.txt"),
                "/root/easyWDMS.txt",
                os.path.expanduser("~/easyWDMS.txt")
            ]
            for p in candidates:
                if p and os.path.exists(p):
                    path_to_try = p
                    break

        if path_to_try and os.path.exists(path_to_try):
            with open(path_to_try, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if '=' in line:
                        k, v = line.split('=', 1)
                        k = k.strip().lower()
                        v = v.strip()
                        if 'user' in k:
                            self.username = v
                        elif 'pass' in k:
                            self.password = v

        if not self.username or not self.password:
            raise ValueError(f"Credentials not found in {self.creds_path} or arguments.")

    def login(self, username=None, password=None):
        self.load_credentials(username, password)
        self.is_logged_in = False
        login_url = f"{self.host}/login/?next=/"
        r = self.session.get(login_url, timeout=15)
        if r.status_code != 200:
            raise ConnectionError(f"Could not reach {login_url} (HTTP {r.status_code})")

        csrf = self.session.cookies.get('csrftoken')
        if not csrf:
            m = re.search(r'name=["\']csrfmiddlewaretoken["\']\s+value=["\']([^"\']+)["\']', r.text)
            if m:
                csrf = m.group(1)
            else:
                raise ValueError("Could not extract CSRF token from login page.")

        form_fields = [
            ('csrfmiddlewaretoken', csrf),
            ('username', self.username),
            ('password', self.password),
            ('template10', ''),
            ('login_type', 'pwd')
        ]
        serialized = urllib.parse.urlencode(form_fields)
        encrypted_data = zk_encrypt(csrf, serialized)

        payload = {
            'encrypt_data': encrypted_data,
            'csrfmiddlewaretoken': csrf
        }
        headers = {
            'X-Requested-With': 'XMLHttpRequest',
            'Referer': login_url
        }
        resp = self.session.post(f"{self.host}/login/", data=payload, headers=headers, timeout=15)
        try:
            res_json = resp.json()
            if res_json.get('ret') != 0:
                raise PermissionError(f"Login failed: {res_json.get('message', 'Invalid credentials')}")
        except Exception as e:
            if resp.status_code != 200:
                raise PermissionError(f"Login failed with status {resp.status_code}")

        try:
            home = self.session.get(f"{self.host}/", timeout=10)
            m = re.search(r"<span style=['\"][^'\"]*['\"][^>]*>[^<]*</span>/([^<]+)</li>", home.text)
            if m:
                self.company_name = m.group(1).strip()
        except Exception:
            pass

        self.is_logged_in = True
        return True

    def _request(self, method, url, auto_relogin=True, **kwargs):
        """Helper to send HTTP requests with automatic re-authentication if session expired."""
        if not self.is_logged_in:
            self.login()

        resp = self.session.request(method, url, **kwargs)

        is_auth_error = (
            resp.status_code in (401, 403) or
            (resp.status_code == 200 and '/login/' in str(resp.url) and 'login' in resp.text.lower())
        )

        if is_auth_error and auto_relogin:
            print("Session expired or unauthorized. Re-authenticating with easyWDMS...")
            self.login()
            resp = self.session.request(method, url, **kwargs)

        return resp

    def get_terminals(self):
        """Returns list of biometric terminals / machines configured in easyWDMS"""
        headers = {'X-Requested-With': 'XMLHttpRequest'}
        try:
            r = self._request('GET', f"{self.host}/iclock/api/terminals/", headers=headers, timeout=15)
            if r.status_code == 200:
                data = r.json().get('data', [])
                if data:
                    return data
        except Exception as e:
            print(f"Warning: Could not fetch terminals from API: {e}")

        # Fallback to known devices
        return [
            {'id': 2674620, 'alias': 'MAIN1', 'sn': 'TDBD254200096', 'ip_address': '192.168.12.223', 'area_name': 'MAIN1'},
            {'id': 2674642, 'alias': 'MAIN 2', 'sn': 'TDBD241101196', 'ip_address': '192.168.12.222', 'area_name': 'MAIN 2'},
            {'id': 2674641, 'alias': 'E & TELESALES', 'sn': 'TDBD242100172', 'ip_address': '192.168.12.224', 'area_name': 'E & TELESALES'},
            {'id': 2674640, 'alias': 'E & CEASED', 'sn': 'TDBD241100793', 'ip_address': '192.168.12.221', 'area_name': 'E & CEASED'},
        ]

    def get_employees(self):
        """Returns a dict of emp_code -> employee info dict"""
        headers = {'X-Requested-With': 'XMLHttpRequest'}
        employees = {}

        # 1. Active employees
        r = self._request('GET', f"{self.host}/personnel/api/employees/?page_size=1000", headers=headers, timeout=20)
        if r.status_code == 200:
            data = r.json().get('data', [])
            for item in data:
                code = str(item.get('emp_code', '')).strip()
                dept = item.get('department', {})
                dept_name = dept.get('dept_name', '-') if isinstance(dept, dict) else str(dept)
                pos = item.get('position', {})
                pos_name = pos.get('position_name', '-') if isinstance(pos, dict) else str(pos)

                first_name = (item.get('first_name') or '').strip()
                last_name = (item.get('last_name') or '').strip()
                if last_name and last_name != '-':
                    full_name = f"{first_name} {last_name}".strip()
                else:
                    full_name = first_name

                employees[code] = {
                    'emp_code': code,
                    'name': full_name,
                    'department': dept_name,
                    'position': pos_name,
                    'status': 'Active'
                }
        else:
            raise ConnectionError(f"Failed to fetch active employees: HTTP {r.status_code} - {r.text[:200]}")

        # 2. Resigned employees (for historical lookups)
        try:
            r_res = self._request('GET', f"{self.host}/personnel/resign/table/?page=1&limit=500", headers=headers, timeout=20)
            if r_res.status_code == 200:
                data = r_res.json().get('data', [])
                for item in data:
                    code = str(item.get('emp_code', '')).strip()
                    if code and code not in employees:
                        first_name = (item.get('first_name') or '').strip()
                        last_name = (item.get('last_name') or '').strip()
                        if last_name and last_name != '-':
                            full_name = f"{first_name} {last_name}".strip()
                        else:
                            full_name = first_name

                        employees[code] = {
                            'emp_code': code,
                            'name': full_name,
                            'department': item.get('department', '-'),
                            'position': item.get('position', '-'),
                            'status': 'Resigned'
                        }
        except Exception:
            pass

        return employees

    def fetch_transactions(self, start_date_str, end_date_str, emp_code=None):
        headers = {'X-Requested-With': 'XMLHttpRequest'}
        start_time = f"{start_date_str} 00:00:00"
        end_time = f"{end_date_str} 23:59:59"

        all_transactions = []
        page = 1
        page_size = 5000

        while True:
            params = {
                'start_time': start_time,
                'end_time': end_time,
                'page': page,
                'page_size': page_size
            }
            if emp_code:
                params['emp_code'] = emp_code

            r = self._request('GET', f"{self.host}/iclock/api/transactions/", params=params, headers=headers, timeout=30)
            if r.status_code != 200:
                raise ConnectionError(f"Error fetching transactions: HTTP {r.status_code} - {r.text[:200]}")

            data = r.json()
            rows = data.get('data', [])
            if not rows:
                break

            all_transactions.extend(rows)
            total_count = data.get('count', len(all_transactions))

            print(f"\rFetching punches: {len(all_transactions)} / {total_count}...", end="", flush=True)

            if not data.get('next') or len(all_transactions) >= total_count:
                break

            page += 1

        print(f"\rFetched total {len(all_transactions)} punches successfully.      ")
        return all_transactions


# ==========================================
# In-Out Calculation Logic (Clean First-In & Last-Out)
# ==========================================
def calculate_in_out(transactions, employees_map, start_date, end_date, single_emp=None, allowed_areas=None):
    """
    Groups transactions by (emp_code, date) and calculates First-In and Last-Out.
    FILTERS: punches from allowed_areas (defaults to MAIN1 & MAIN 2 if allowed_areas is None).
    Clean report: First-IN & Last-OUT punches (no salary deduction status or duration penalties).
    """
    grouped = defaultdict(lambda: defaultdict(list))
    for t in transactions:
        # FILTER: check against allowed areas/terminals
        if not is_terminal_allowed(t, allowed_areas):
            continue

        code = str(t.get('emp_code', '')).strip()
        punch_dt = t.get('punch_time', '')
        if not code or not punch_dt:
            continue
        date_part, time_part = punch_dt.split(' ', 1)
        grouped[code][date_part].append({
            'time': time_part,
            'datetime': datetime.strptime(punch_dt, '%Y-%m-%d %H:%M:%S'),
            'terminal': t.get('terminal_alias') or t.get('terminal_sn') or '-',
            'area': t.get('area_alias') or '-'
        })

    cur = start_date
    date_list = []
    while cur <= end_date:
        date_list.append(cur)
        cur += timedelta(days=1)

    daily_records = []
    summary_records = {}

    target_emp_codes = [single_emp['emp_code']] if single_emp else sorted(list(grouped.keys()), key=safe_emp_sort_key)

    # 1. Single Employee Report
    if single_emp:
        emp_code = single_emp['emp_code']
        emp_name = single_emp.get('name', 'Unknown')
        emp_dept = single_emp.get('department', '-')

        present_days = 0
        single_punch_days = 0

        for d in date_list:
            d_str = d.strftime('%Y-%m-%d')
            day_name = d.strftime('%A')

            punches = grouped[emp_code].get(d_str, [])
            if punches:
                punches.sort(key=lambda x: x['datetime'])
                first_punch = punches[0]
                last_punch = punches[-1]

                in_time = first_punch['time']
                in_term = first_punch['terminal']
                punch_count = len(punches)

                if punch_count > 1:
                    out_time = last_punch['time']
                    out_term = last_punch['terminal']
                    present_days += 1
                else:
                    out_time = "---"
                    out_term = "---"
                    single_punch_days += 1

                daily_records.append({
                    'emp_code': emp_code,
                    'name': emp_name,
                    'department': emp_dept,
                    'date': d_str,
                    'day': day_name,
                    'in_time': in_time,
                    'out_time': out_time,
                    'in_terminal': in_term,
                    'out_terminal': out_term,
                    'total_punches': punch_count
                })
            else:
                daily_records.append({
                    'emp_code': emp_code,
                    'name': emp_name,
                    'department': emp_dept,
                    'date': d_str,
                    'day': day_name,
                    'in_time': "---",
                    'out_time': "---",
                    'in_terminal': "---",
                    'out_terminal': "---",
                    'total_punches': 0
                })

        summary_records[emp_code] = {
            'emp_code': emp_code,
            'name': emp_name,
            'department': emp_dept,
            'total_days_in_period': len(date_list),
            'present_days': present_days,
            'single_punch_days': single_punch_days
        }

    # 2. Multi-Employee Report
    else:
        emp_stats = defaultdict(lambda: {'present': 0, 'single': 0})

        for emp_code in target_emp_codes:
            emp_info = employees_map.get(emp_code, {})
            emp_name = emp_info.get('name') or f"Employee {emp_code}"
            emp_dept = emp_info.get('department', '-')

            for d_str, punches in grouped[emp_code].items():
                punches.sort(key=lambda x: x['datetime'])
                first_punch = punches[0]
                last_punch = punches[-1]
                punch_count = len(punches)

                d_obj = datetime.strptime(d_str, '%Y-%m-%d').date()
                day_name = d_obj.strftime('%A')

                in_time = first_punch['time']
                in_term = first_punch['terminal']

                if punch_count > 1:
                    out_time = last_punch['time']
                    out_term = last_punch['terminal']
                    emp_stats[emp_code]['present'] += 1
                else:
                    out_time = "---"
                    out_term = "---"
                    emp_stats[emp_code]['single'] += 1

                daily_records.append({
                    'emp_code': emp_code,
                    'name': emp_name,
                    'department': emp_dept,
                    'date': d_str,
                    'day': day_name,
                    'in_time': in_time,
                    'out_time': out_time,
                    'in_terminal': in_term,
                    'out_terminal': out_term,
                    'total_punches': punch_count
                })

            summary_records[emp_code] = {
                'emp_code': emp_code,
                'name': emp_name,
                'department': emp_dept,
                'total_days_in_period': len(date_list),
                'present_days': emp_stats[emp_code]['present'],
                'single_punch_days': emp_stats[emp_code]['single']
            }

        daily_records.sort(key=lambda r: (r['date'], safe_emp_sort_key(r['emp_code'])))

    return daily_records, summary_records


# ==========================================
# Excel Generation (Styled openpyxl)
# ==========================================
def export_to_excel(daily_records, summary_records, output_path, meta_info):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    font_title = Font(name='Calibri', size=16, bold=True, color='1F4E78')
    font_subtitle = Font(name='Calibri', size=10, bold=True, color='595959')
    font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    font_bold = Font(name='Calibri', size=10, bold=True)
    font_regular = Font(name='Calibri', size=10)

    fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_zebra = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')

    thin_border = Side(style='thin', color='D9D9D9')
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # -----------------------------------------------
    # SHEET 1: Daily In-Out Attendance
    # -----------------------------------------------
    ws1 = wb.create_sheet(title="Daily In-Out Attendance")
    ws1.views.sheetView[0].showGridLines = True

    ws1.cell(row=1, column=1, value=f"{meta_info.get('company', 'Company')} - Attendance In & Out Report").font = font_title
    machines_str = meta_info.get('machines', 'MAIN1 & MAIN 2')
    subtitle = f"Period: {meta_info.get('period', '')} | Machines: {machines_str}"
    if meta_info.get('employee_filter'):
        subtitle += f" | Filter: {meta_info.get('employee_filter')}"
    ws1.cell(row=2, column=1, value=subtitle).font = font_subtitle

    headers = [
        "Emp Code", "Employee Name", "Department", "Date", "Day",
        "IN Time", "OUT Time", "IN Machine", "OUT Machine", "Punches"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, 1):
        cell = ws1.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell
    ws1.row_dimensions[header_row].height = 25

    row_idx = header_row + 1
    for r in daily_records:
        ws1.row_dimensions[row_idx].height = 20
        zebra = (row_idx % 2 == 0)

        vals = [
            (r['emp_code'], align_center),
            (r['name'], align_left),
            (r['department'], align_left),
            (r['date'], align_center),
            (r['day'], align_center),
            (r['in_time'], align_center),
            (r['out_time'], align_center),
            (r['in_terminal'], align_left),
            (r['out_terminal'], align_left),
            (r['total_punches'], align_center)
        ]

        for col_idx, (val, alignment) in enumerate(vals, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.alignment = alignment
            cell.border = border_cell
            if zebra:
                cell.fill = fill_zebra

        row_idx += 1

    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row < 4:
                continue
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # -----------------------------------------------
    # SHEET 2: Employee Summary
    # -----------------------------------------------
    ws2 = wb.create_sheet(title="Employee Summary")
    ws2.views.sheetView[0].showGridLines = True

    ws2.cell(row=1, column=1, value=f"{meta_info.get('company', 'Company')} - Attendance Summary").font = font_title
    ws2.cell(row=2, column=1, value=f"Period: {meta_info.get('period', '')} | Machines: {machines_str}").font = font_subtitle

    summary_headers = [
        "Emp Code", "Employee Name", "Department",
        "Total Days in Period", "Present Days (IN & OUT)", "Single Punch Days"
    ]

    header_row_2 = 4
    for col_idx, h in enumerate(summary_headers, 1):
        cell = ws2.cell(row=header_row_2, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell
    ws2.row_dimensions[header_row_2].height = 25

    row_idx_2 = header_row_2 + 1
    sorted_summaries = sorted(summary_records.values(), key=lambda s: safe_emp_sort_key(s['emp_code']))

    for s in sorted_summaries:
        ws2.row_dimensions[row_idx_2].height = 20
        zebra = (row_idx_2 % 2 == 0)

        s_vals = [
            (s['emp_code'], align_center),
            (s['name'], align_left),
            (s['department'], align_left),
            (s['total_days_in_period'], align_center),
            (s['present_days'], align_center),
            (s['single_punch_days'], align_center)
        ]

        for col_idx, (val, alignment) in enumerate(s_vals, 1):
            cell = ws2.cell(row=row_idx_2, column=col_idx, value=val)
            cell.font = font_regular
            cell.alignment = alignment
            cell.border = border_cell
            if zebra:
                cell.fill = fill_zebra

        row_idx_2 += 1

    for col in ws2.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row < 4:
                continue
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_path)
    return output_path


# ==========================================
# CLI Entry Point
# ==========================================
def parse_date_range(month_str=None, from_date_str=None, to_date_str=None):
    if month_str:
        try:
            parts = month_str.strip().split('-')
            year = int(parts[0])
            month = int(parts[1])
            start_date = date(year, month, 1)
            if month == 12:
                next_month = date(year + 1, 1, 1)
            else:
                next_month = date(year, month + 1, 1)
            end_date = next_month - timedelta(days=1)
            return start_date, end_date
        except Exception:
            raise ValueError(f"Invalid month format '{month_str}'. Please use YYYY-MM (e.g. 2026-10).")

    if from_date_str and to_date_str:
        try:
            start_date = datetime.strptime(from_date_str.strip(), '%Y-%m-%d').date()
            end_date = datetime.strptime(to_date_str.strip(), '%Y-%m-%d').date()
            if start_date > end_date:
                raise ValueError("from_date cannot be after to_date.")
            return start_date, end_date
        except Exception as e:
            raise ValueError(f"Invalid date format: {e}. Please use YYYY-MM-DD (e.g. 2026-11-01).")

    today = date.today()
    start_date = date(today.year, today.month, 1)
    end_date = today
    return start_date, end_date


def main():
    parser = argparse.ArgumentParser(description="easyWDMS Biometric Attendance In-Out Report Tool (MAIN & MAIN 2 Only)")
    parser.add_argument('--month', help="Specific month (YYYY-MM, e.g. 2026-10)")
    parser.add_argument('--from', dest='from_date', help="Start date (YYYY-MM-DD)")
    parser.add_argument('--to', dest='to_date', help="End date (YYYY-MM-DD)")
    parser.add_argument('--emp', help="Employee code (e.g. 10190)")
    parser.add_argument('--name', help="Employee name search query (e.g. Adnan)")
    parser.add_argument('--all', action='store_true', help="Export all employees")
    parser.add_argument('-o', '--output', help="Output Excel filename (.xlsx)")
    parser.add_argument('--creds', default="/root/easyWDMS.txt", help="Path to credentials file")
    parser.add_argument('--host', default="http://122.176.49.74:8082", help="Portal base URL")

    args = parser.parse_args()

    client = EasyWDMSClient(host=args.host, creds_path=args.creds)
    print(f"Connecting to portal {client.host}...")
    try:
        client.login()
        print(f"Logged in successfully as '{client.username}' | Company: '{client.company_name}'")
    except Exception as e:
        print(f"Login failed: {e}")
        sys.exit(1)

    print("Fetching employee directory...")
    employees_map = client.get_employees()
    print(f"Loaded {len(employees_map)} employee records.")

    start_date, end_date = parse_date_range(month_str=args.month, from_date_str=args.from_date, to_date_str=args.to_date)
    single_emp = None

    if args.emp:
        single_emp = employees_map.get(str(args.emp).strip())
        if not single_emp:
            single_emp = {'emp_code': str(args.emp).strip(), 'name': f"Employee {args.emp}", 'department': '-'}
    elif args.name:
        query = args.name.lower()
        matched = [emp for emp in employees_map.values() if query in emp['name'].lower()]
        if not matched:
            print(f"Error: No employee found matching name '{args.name}'.")
            sys.exit(1)
        single_emp = matched[0]
        print(f"Matched Employee: [{single_emp['emp_code']}] {single_emp['name']}")
    elif not args.all:
        args.all = True

    s_str = start_date.strftime('%Y%m%d')
    e_str = end_date.strftime('%Y%m%d')
    if args.output:
        output_file = args.output
    elif single_emp:
        clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', single_emp['name'])
        output_file = f"attendance_{single_emp['emp_code']}_{clean_name}_{s_str}_{e_str}.xlsx"
    else:
        output_file = f"attendance_MAIN_ALL_{s_str}_{e_str}.xlsx"

    target_code = single_emp['emp_code'] if single_emp else None
    print(f"\nQuerying punch logs for {start_date} to {end_date} (Machines: MAIN & MAIN 2)...")

    transactions = client.fetch_transactions(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'), emp_code=target_code)

    daily_records, summary_records = calculate_in_out(transactions, employees_map, start_date, end_date, single_emp=single_emp)

    meta_info = {
        'company': client.company_name,
        'period': f"{start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}",
        'employee_filter': f"[{single_emp['emp_code']}] {single_emp['name']}" if single_emp else "All Employees"
    }

    if not os.path.isabs(output_file):
        output_file = os.path.abspath(output_file)

    print(f"Generating formatted Excel file: {output_file}...")
    export_to_excel(daily_records, summary_records, output_file, meta_info)

    print("\n" + "=" * 60)
    print(" SUCCESS: Excel Attendance Report Generated!")
    print(" (Filter: MAIN1 & MAIN 2 Biometric First-IN & Last-OUT)")
    print("=" * 60)
    print(f"File Path    : {output_file}")
    print(f"Total Records: {len(daily_records)} daily records")
    if single_emp:
        summ = summary_records.get(single_emp['emp_code'], {})
        print(f"Employee     : [{single_emp['emp_code']}] {single_emp['name']}")
        print(f"Present Days : {summ.get('present_days', 0)} days (IN & OUT)")
        print(f"Single Punch : {summ.get('single_punch_days', 0)} days")
    else:
        print(f"Employees    : {len(summary_records)} employees processed")
    print("=" * 60)


if __name__ == '__main__':
    main()
