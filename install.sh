#!/usr/bin/env bash
# ==============================================================================
# easyWDMS Attendance Portal - Automated Setup Script
# ==============================================================================
set -e

INSTALL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SERVICE_NAME="easywdms-gui"

echo "=========================================================="
echo " Installing easyWDMS Attendance Portal in: $INSTALL_DIR"
echo "=========================================================="

# 1. Check Python3
if ! command -v python3 &> /dev/null; then
    echo "[-] Python3 not found. Installing python3..."
    if command -v dnf &> /dev/null; then
        dnf install -y python3 python3-pip
    elif command -v apt-get &> /dev/null; then
        apt-get update && apt-get install -y python3 python3-pip
    elif command -v yum &> /dev/null; then
        yum install -y python3 python3-pip
    else
        echo "[!] Please install python3 and pip manually."
        exit 1
    fi
fi

# 2. Install Python Dependencies
echo "[+] Installing Python dependencies..."
python3 -m pip install --upgrade pip
python3 -m pip install -r "$INSTALL_DIR/requirements.txt"

# 3. Setup Credentials File
if [ ! -f "$INSTALL_DIR/easyWDMS.txt" ]; then
    echo "[+] Creating easyWDMS.txt from template..."
    cp "$INSTALL_DIR/easyWDMS.example.txt" "$INSTALL_DIR/easyWDMS.txt"
    echo "[!] IMPORTANT: Please edit $INSTALL_DIR/easyWDMS.txt with your easyWDMS username & password."
fi

# 4. Setup Systemd Service
echo "[+] Configuring systemd service: /etc/systemd/system/${SERVICE_NAME}.service"
cat <<EOF > "/etc/systemd/system/${SERVICE_NAME}.service"
[Unit]
Description=easyWDMS Web GUI Attendance Service
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=${INSTALL_DIR}
ExecStart=$(which python3) ${INSTALL_DIR}/app.py
Restart=always
RestartSec=3
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable "${SERVICE_NAME}.service"
systemctl restart "${SERVICE_NAME}.service"

echo "[+] Service status:"
systemctl status "${SERVICE_NAME}.service" --no-pager || true

# 5. Nginx Configuration guidance
echo ""
echo "=========================================================="
echo " Setup Completed Successfully!"
echo "=========================================================="
echo " - Dashboard Port : http://<server-ip>:8083"
echo " - Config File    : $INSTALL_DIR/easyWDMS.txt"
echo " - Service Command: systemctl restart $SERVICE_NAME"
echo ""
echo " To configure Nginx reverse proxy:"
echo " Copy '$INSTALL_DIR/nginx/easywdms.conf' to your Nginx configuration."
echo "=========================================================="
