# attendance-access
# easyWDMS Biometric Attendance & In-Out Access Portal

A modern, responsive web dashboard and automated reporting tool for **easyWDMS biometric attendance systems**.
Connects to easyWDMS, processes raw punch logs, calculates clean First-IN and Last-OUT attendance, calculates working duration, supports multi-area machine filtering (MAIN1, MAIN 2, Telesales, Ceased), and exports styled multi-sheet Excel reports (.xlsx).

---

## ✨ Features

- **Multi-Area / Machine Filtering:**
  - `MAIN1` (`192.168.12.223`)
  - `MAIN 2` (`192.168.12.222`)
  - `E & TELESALES` (`192.168.12.224`)
  - `E & CEASED` (`192.168.12.221`)
  - 1-click presets: *MAIN Only*, *All Machines*, *Telesales Only*, *Ceased Only*, or custom checkboxes.
- **Clean In-Out Punch Logic:**
  - Accurate First-IN and Last-OUT punch extraction.
  - Automatically calculates Working Duration (e.g. `9h 44m`).
  - Attendance Status Badges: *Full Day (≥ 8h)*, *Half Day (≥ 4h)*, *Single Punch*, *No Punch / Absent / Sunday*.
- **Modern Responsive Web Dashboard:**
  - **Mobile First:** Touch-friendly card view for smartphones & tablets + Offcanvas slide-out filter drawer.
  - **Desktop View:** Sticky parameters panel + sortable data table.
  - **Dark / Light Mode:** 1-click theme switch persisted in `localStorage`.
  - **1-Click Today's Live:** Instant preview of today's live biometric punches.
  - **Interactive Analytics:** Clickable KPI cards (Total Records, Present Days, Single Punches, Total Employees) to filter records on the fly.
  - **Search & Pagination:** Instant column search + selectable page sizes (25, 50, 100, All).
- **Export & Integration:**
  - **Formatted Excel (.xlsx):** Generates two styled worksheets (`Daily In-Out Attendance` & `Employee Summary`) with auto column widths, borders, and headers.
  - **Copy TSV:** 1-click copy formatted for pasting directly into Excel, WhatsApp, or Google Sheets.
  - **Quick CSV Download:** Fast client-side CSV export.
  - **Print / PDF Sheet:** Clean print view with `@media print` styling.
- **Automated Systemd Service:** Self-recovering background service with auto-restart on boot.

---

## 📁 Repository Structure

```text
attendance-access/
├── app.py                      # Flask Web GUI Server (Port 8083)
├── easywdms_report.py          # easyWDMS API Client & Report Engine
├── templates/
│   └── index.html              # Responsive Web UI (Bootstrap 5.3 + FontAwesome)
├── easyWDMS.example.txt        # Template for easyWDMS credentials
├── requirements.txt            # Python dependencies (flask, requests, openpyxl)
├── systemd/
│   └── easywdms-gui.service    # Systemd service configuration
├── nginx/
│   └── easywdms.conf           # Nginx reverse proxy configuration
├── install.sh                  # 1-click automated installer script
├── .gitignore                  # Git ignore rules (protects credentials & exports)
└── README.md                   # Documentation
```

---

## 🚀 Installation on Office Linux Server

### Step 1: Clone Repository
```bash
cd /opt
git clone https://github.com/jacky1473/attendance-access.git
cd attendance-access
```

### Step 2: Configure easyWDMS Credentials
Create your `easyWDMS.txt` from the example template:
```bash
cp easyWDMS.example.txt easyWDMS.txt
nano easyWDMS.txt
```
Enter your easyWDMS login credentials:
```ini
username = your_easywdms_username
password = your_easywdms_password
```

### Step 3: Run Automated Installer
Run the installer script (installs python dependencies, configures and starts the systemd service):
```bash
chmod +x install.sh
sudo ./install.sh
```

---

## 🌐 Nginx Reverse Proxy Setup (Optional but Recommended)

To serve the portal under `http://your-server-ip/attendance/`:

### For CentOS / RHEL / Rocky Linux:
```bash
sudo cp nginx/easywdms.conf /etc/nginx/default.d/easywdms.conf
sudo nginx -t
sudo systemctl reload nginx
```

### For Ubuntu / Debian:
Add the contents of `nginx/easywdms.conf` inside your `/etc/nginx/sites-available/default` server block, then reload nginx:
```bash
sudo nginx -t
sudo systemctl reload nginx
```

Now you can access:
- **Web UI via Nginx:** `http://<your-server-ip>/attendance/`
- **Direct Web UI:** `http://<your-server-ip>:8083/`

---

## 🛠️ Standalone CLI Usage

You can also run reports directly from the terminal without the web GUI:

```bash
# Report for single employee for November 2026:
python3 easywdms_report.py --emp 10248 --month 2026-11

# Report for all employees for custom date range:
python3 easywdms_report.py --all --from-date 2026-11-01 --to-date 2026-11-10

# Search employee by name:
python3 easywdms_report.py --name "Christina" --month 2026-11

# Specify output Excel filename:
python3 easywdms_report.py --emp 10248 --month 2026-11 -o /tmp/november_report.xlsx
```

---

## 🔧 Service Management

```bash
# Check service status
sudo systemctl status easywdms-gui

# Restart service
sudo systemctl restart easywdms-gui

# View live service logs
sudo journalctl -u easywdms-gui -f
```

---

## 🔒 Security Note
Never commit `easyWDMS.txt` or personal access tokens to GitHub. The included `.gitignore` protects credentials and exported report files.
