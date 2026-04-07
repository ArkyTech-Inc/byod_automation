"""
NITDA BYOD — Security Gate Server
Handles QR code scan requests from the gate_scanner.html terminal.
Runs separately from the approval server (default port 5001).

Usage:
    python gate_server.py

Install dependencies:
    pip install flask openpyxl flask-cors
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

# Allow requests from the gate_scanner.html page (open for LAN access)
CORS(app)

# ── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE = 'NITDA_BYOD_Database.xlsx'

# Only these IP addresses are allowed to use the /gate/scan endpoint.
# Add the security post PC's local IP here.
# Leave empty [] to allow any device (not recommended for production).
ALLOWED_IPS = [
    '127.0.0.1',         # localhost (for testing)
    # '192.168.1.XX',    # ← replace with security post PC's LAN IP
]

# ── HELPERS ─────────────────────────────────────────────────────────────────

def is_allowed():
    """Check if the request comes from an authorised IP."""
    if not ALLOWED_IPS:
        return True
    client_ip = request.remote_addr
    return client_ip in ALLOWED_IPS


def get_device_record(reg_id: str) -> dict | None:
    """
    Look up a registration ID in the Device Registrations sheet.
    Returns a dict of relevant fields, or None if not found.
    """
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
        sheet = wb['Device Registrations']

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if str(row[0]).strip() == reg_id.strip():
                wb.close()
                return {
                    'reg_id':     str(row[0]),
                    'timestamp':  str(row[1]) if row[1] else '',
                    'name':       str(row[2]) if row[2] else '',
                    'department': str(row[3]) if row[3] else '',
                    'device_type':str(row[5]) if row[5] else '',
                    'device':     str(row[6]) if row[6] else '',   # Make/Model
                    'serial':     str(row[9]) if row[9] else '',
                    'email':      str(row[10]) if row[10] else '',
                    'status':     str(row[14]) if row[14] else 'Pending',
                    'compliant':  str(row[13]) if len(row) > 13 and row[13] else '',
                }
        wb.close()
        return None

    except FileNotFoundError:
        print(f"[ERROR] Excel file not found: {EXCEL_FILE}")
        return None
    except Exception as e:
        print(f"[ERROR] Reading Excel: {e}")
        return None


def log_gate_entry(reg_id: str, name: str, device: str, action: str,
                   officer: str, status: str, remarks: str = '') -> bool:
    """
    Append a new row to the Security Gate Log sheet.
    Columns: Log ID | Date | Time | Reg ID | Name | Device | Action | Officer | Status | Remarks
    """
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb['Security Gate Log']

        # Find next empty row
        next_row = sheet.max_row + 1
        # Edge case: if max_row has no data in col A, scan back
        while next_row > 2 and sheet.cell(next_row - 1, 1).value is None:
            next_row -= 1

        now = datetime.now()
        log_id = f"LOG-{now.strftime('%Y%m%d')}-{next_row - 1:04d}"

        row_data = [
            log_id,                           # A: Log ID
            now.strftime('%Y-%m-%d'),         # B: Date
            now.strftime('%H:%M:%S'),         # C: Time
            reg_id,                           # D: Registration ID
            name,                             # E: Name
            device,                           # F: Device Model
            action,                           # G: Action (Check In / Check Out)
            officer,                          # H: Security Officer
            status,                           # I: Status (Authorised / Denied)
            remarks                           # J: Remarks
        ]

        for col, value in enumerate(row_data, 1):
            sheet.cell(row=next_row, column=col).value = value

        wb.save(EXCEL_FILE)
        print(f"[LOG] {log_id} | {action} | {name} | {status}")
        return True

    except PermissionError:
        print("[ERROR] Excel file is open. Close it before scanning.")
        return False
    except Exception as e:
        print(f"[ERROR] Writing gate log: {e}")
        return False


def check_it_compliance(reg_id: str) -> bool:
    """
    Check if the device has passed IT inspection (Compliance Status = Compliant).
    Returns True if compliant, False otherwise.
    """
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
        sheet = wb['IT Inspection']

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and str(row[0]).strip() == reg_id.strip():
                compliance = str(row[13]).strip().lower() if row[13] else ''
                wb.close()
                return compliance == 'compliant'

        wb.close()
        return False  # No inspection record = not cleared

    except Exception as e:
        print(f"[ERROR] Checking IT compliance: {e}")
        return False


# ── ROUTES ───────────────────────────────────────────────────────────────────

@app.route('/gate/ping', methods=['GET'])
def ping():
    """Health check — lets the HTML page confirm the server is reachable."""
    return jsonify({'ok': True, 'server': 'NITDA Gate Server', 'time': datetime.now().isoformat()})


@app.route('/gate/scan', methods=['POST'])
def gate_scan():
    """
    Main scan endpoint.
    Expects JSON: { "reg_id": "BYOD-...", "action": "in"|"out", "officer": "..." }
    Returns JSON with ok, device details, or denial reason.
    """

    # ── IP check ──────────────────────────────────────────────────
    if not is_allowed():
        return jsonify({
            'ok': False,
            'reason': 'UNAUTHORISED TERMINAL',
            'message': f'This device ({request.remote_addr}) is not permitted to access the gate system.'
        }), 403

    # ── Parse body ────────────────────────────────────────────────
    body = request.get_json(silent=True)
    if not body or 'reg_id' not in body:
        return jsonify({'ok': False, 'reason': 'BAD REQUEST', 'message': 'Missing reg_id in request.'}), 400

    reg_id  = body.get('reg_id', '').strip()
    action  = body.get('action', 'in').lower()          # 'in' or 'out'
    officer = body.get('officer', 'Security Officer')

    action_label = 'Check In' if action == 'in' else 'Check Out'

    # ── Look up device ────────────────────────────────────────────
    device = get_device_record(reg_id)

    if device is None:
        log_gate_entry(reg_id, 'UNKNOWN', '—', action_label, officer,
                       'Denied', 'Registration ID not found')
        return jsonify({
            'ok': False,
            'reason': 'NOT FOUND',
            'message': f'No device registered with ID: {reg_id}. '
                       f'Contact the BYOD administrator.'
        })

    # ── Check approval status ─────────────────────────────────────
    if device['status'].lower() != 'approved':
        log_gate_entry(reg_id, device['name'], device['device'], action_label,
                       officer, 'Denied', f"Status is '{device['status']}' — not Approved")
        return jsonify({
            'ok': False,
            'reason': f"DEVICE {device['status'].upper()}",
            'message': f"This device has not been approved by administration. "
                       f"Current status: {device['status']}."
        })

    # ── Check IT compliance ───────────────────────────────────────
    if not check_it_compliance(reg_id):
        log_gate_entry(reg_id, device['name'], device['device'], action_label,
                       officer, 'Denied', 'No IT compliance record — inspection not completed')
        return jsonify({
            'ok': False,
            'reason': 'IT INSPECTION INCOMPLETE',
            'message': 'This device has not completed the required IT security inspection. '
                       'The owner must visit the IT department before bringing this device on-site.'
        })

    # ── All checks passed — log it ────────────────────────────────
    now = datetime.now()
    logged = log_gate_entry(
        reg_id, device['name'], device['device'],
        action_label, officer, 'Authorised'
    )

    remarks_log = '' if logged else 'Warning: failed to write to Excel'

    return jsonify({
        'ok': True,
        'reg_id':     device['reg_id'],
        'name':       device['name'],
        'department': device['department'],
        'device':     device['device'],
        'serial':     device['serial'],
        'status':     device['status'],
        'action':     action_label,
        'time':       now.strftime('%H:%M:%S'),
        'date':       now.strftime('%Y-%m-%d'),
        'excel_saved': logged,
        'remarks':    remarks_log
    })


@app.route('/gate/log', methods=['GET'])
def get_todays_log():
    """
    Returns today's gate log entries as JSON.
    Optionally filter by date: /gate/log?date=2025-06-01
    Only accessible from allowed IPs.
    """
    if not is_allowed():
        return jsonify({'ok': False, 'message': 'Unauthorised'}), 403

    target_date = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
        sheet = wb['Security Gate Log']
        entries = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and str(row[1]) == target_date:
                entries.append({
                    'log_id':  row[0],
                    'date':    str(row[1]),
                    'time':    str(row[2]),
                    'reg_id':  row[3],
                    'name':    row[4],
                    'device':  row[5],
                    'action':  row[6],
                    'officer': row[7],
                    'status':  row[8],
                    'remarks': row[9]
                })

        wb.close()
        return jsonify({'ok': True, 'date': target_date, 'entries': entries, 'count': len(entries)})

    except Exception as e:
        return jsonify({'ok': False, 'message': str(e)}), 500


# ── STARTUP ───────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print("\n" + "=" * 65)
    print("  NITDA BYOD — SECURITY GATE SERVER")
    print("=" * 65)

    if not os.path.exists(EXCEL_FILE):
        print(f"\n  [WARNING] '{EXCEL_FILE}' not found in current directory.")
        print("  Make sure this server runs from the same folder as the Excel file.\n")
    else:
        print(f"\n  Database  : {os.path.abspath(EXCEL_FILE)}")

    print(f"\n  Gate UI   : Open gate_scanner.html on the security post PC")
    print(f"  Server    : http://0.0.0.0:5001")
    print(f"\n  Allowed IPs: {ALLOWED_IPS if ALLOWED_IPS else 'ALL (set ALLOWED_IPS to restrict)'}")
    print("\n  Press Ctrl+C to stop")
    print("=" * 65 + "\n")

    app.run(host='0.0.0.0', port=5001, debug=False)
