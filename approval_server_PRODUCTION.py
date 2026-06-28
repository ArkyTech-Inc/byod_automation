"""
NITDA BYOD Management Portal Server - Production Ready
Handles supervisor approval/rejection with email notifications
"""
import os
import logging
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, request, render_template_string
from dotenv import load_dotenv
from supabase import create_client, Client

from config import Config

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Initialize Supabase client
try:
    supabase: Client = create_client(Config.SUPABASE_URL, Config.SUPABASE_KEY)
    logger.info("✅ Supabase connection initialized")
except Exception as e:
    logger.error(f"❌ Failed to connect to Supabase: {e}")
    raise RuntimeError("Failed to initialize Supabase connection")


def send_approval_email(to_email: str, subject: str, body: str) -> bool:
    """Send email notification"""
    try:
        msg = MIMEMultipart()
        msg['From'] = Config.SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Send email
        if Config.SMTP_PORT == 465:
            server = smtplib.SMTP_SSL(Config.SMTP_SERVER, Config.SMTP_PORT, timeout=10)
        else:
            server = smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT, timeout=10)
            server.starttls()
        
        server.login(Config.SENDER_EMAIL, Config.SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        
        logger.info(f"✅ Email sent to {to_email}: {subject}")
        return True
    except Exception as e:
        logger.error(f"❌ Failed to send email to {to_email}: {e}")
        return False


# HTML Templates
APPROVAL_PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NITDA BYOD Endorsement Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    </style>
</head>
<body class="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 flex items-center justify-center p-4">
    <div class="bg-white/10 backdrop-blur-xl p-8 rounded-2xl shadow-2xl max-w-md w-full border border-white/20">
        <div class="text-center mb-8">
            <h1 class="text-3xl font-bold text-white mb-2">NITDA BYOD Gateway</h1>
            <p class="text-slate-300 text-sm">Device Registration Approval Portal</p>
        </div>

        {% if error %}
            <div class="bg-red-500/20 border border-red-500/50 text-red-200 px-4 py-3 rounded-lg mb-6">
                <p class="font-semibold">⚠️ Error</p>
                <p class="text-sm">{{ error }}</p>
            </div>
        {% elif success %}
            <div class="bg-green-500/20 border border-green-500/50 text-green-200 px-4 py-3 rounded-lg mb-6">
                <p class="font-semibold">✅ Success</p>
                <p class="text-sm">{{ message }}</p>
            </div>

            {% if device_info %}
            <div class="bg-white/5 border border-white/10 rounded-lg p-4 mb-6 space-y-2 text-sm">
                <div class="flex justify-between">
                    <span class="text-slate-400">Registration ID:</span>
                    <span class="text-white font-mono">{{ device_info.registration_id }}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Name:</span>
                    <span class="text-white">{{ device_info.name }}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Department:</span>
                    <span class="text-white">{{ device_info.department }}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Device:</span>
                    <span class="text-white">{{ device_info.device_make_model or device_info.device_model or 'N/A' }}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Status:</span>
                    <span class="text-white font-semibold">{{ device_info.status }}</span>
                </div>
            </div>
            {% endif %}

            <a href="/" class="block w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-lg text-center transition">
                Return to Home
            </a>
        {% elif show_form %}
            <!-- Approval Form -->
            <div class="bg-white/5 border border-white/10 rounded-lg p-6 mb-6">
                <h2 class="text-white text-lg font-bold mb-4">
                    {% if action == 'approve' %}✅ Approve Device{% else %}❌ Reject Device{% endif %}
                </h2>

                {% if device_info %}
                <div class="bg-white/5 rounded-lg p-4 mb-6 space-y-2 text-sm text-slate-300">
                    <div><span class="text-slate-400">Name:</span> <span class="text-white font-semibold">{{ device_info.name }}</span></div>
                    <div><span class="text-slate-400">Department:</span> <span class="text-white">{{ device_info.department }}</span></div>
                    <div><span class="text-slate-400">Device:</span> <span class="text-white">{{ device_info.device_make_model or device_info.device_model }}</span></div>
                    <div><span class="text-slate-400">Serial:</span> <span class="font-mono text-white">{{ device_info.serial_number }}</span></div>
                </div>
                {% endif %}

                <form method="POST" class="space-y-4">
                    <div>
                        <label class="block text-white text-sm font-semibold mb-2">Comments (Optional)</label>
                        <textarea name="remarks" rows="3" placeholder="Add any additional comments..." 
                                  class="w-full px-4 py-3 bg-white/10 border border-white/20 rounded-lg text-white placeholder-slate-500 focus:outline-none focus:border-blue-500 resize-none"></textarea>
                    </div>

                    <button type="submit" name="action" value="{{ action }}"
                            class="w-full {% if action == 'approve' %}bg-green-600 hover:bg-green-700{% else %}bg-red-600 hover:bg-red-700{% endif %} text-white font-bold py-3 px-4 rounded-lg transition">
                        {% if action == 'approve' %}✅ Confirm Approval{% else %}❌ Confirm Rejection{% endif %}
                    </button>
                </form>
            </div>
        {% else %}
            <!-- Loading or initial state -->
            <div class="text-center text-slate-300">
                <p>Loading...</p>
            </div>
        {% endif %}
    </div>
</body>
</html>
"""

HOME_PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NITDA BYOD Gateway</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 flex items-center justify-center p-4">
    <div class="bg-white/10 backdrop-blur-xl p-8 rounded-2xl shadow-2xl max-w-md w-full border border-white/20 text-center">
        <h1 class="text-3xl font-bold text-white mb-2">NITDA BYOD Gateway</h1>
        <p class="text-slate-300 text-sm mb-6">Device Registration Management System</p>
        
        <div class="flex items-center justify-center gap-2 text-sm text-emerald-400 bg-emerald-500/10 px-4 py-3 rounded-lg border border-emerald-500/30">
            <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
            System Status: Online
        </div>
        
        <p class="text-slate-400 text-xs mt-6">
            This portal handles supervisor approvals for BYOD device registrations.
        </p>
    </div>
</body>
</html>
"""


@app.route('/')
def index():
    """Home page"""
    return render_template_string(HOME_PAGE_TEMPLATE)


@app.route('/endorse', methods=['GET', 'POST'])
def handle_endorsement():
    """Handle device approval/rejection"""
    reg_id = request.args.get('id', '').strip()
    action = request.args.get('action', '').strip().lower()

    # Validate parameters
    if not reg_id or action not in ['approve', 'reject']:
        return render_template_string(
            APPROVAL_PAGE_TEMPLATE,
            error="Invalid request parameters",
            success=False,
            show_form=False,
            device_info=None
        )

    try:
        # Fetch device record from Supabase
        response = supabase.table('device_registrations') \
            .select('*') \
            .eq('registration_id', reg_id) \
            .execute()

        if not response.data or len(response.data) == 0:
            return render_template_string(
                APPROVAL_PAGE_TEMPLATE,
                error=f"Device registration '{reg_id}' not found",
                success=False,
                show_form=False,
                device_info=None
            )

        device = response.data[0]

        # Check if already processed
        if device['status'] != 'Pending':
            return render_template_string(
                APPROVAL_PAGE_TEMPLATE,
                error=f"This request has already been processed. Status: {device['status']}",
                success=True,
                message="No further action needed",
                show_form=False,
                device_info=device
            )

        if request.method == 'GET':
            # Show approval form
            return render_template_string(
                APPROVAL_PAGE_TEMPLATE,
                success=False,
                show_form=True,
                device_info=device,
                action=action,
                error=None
            )

        # Handle POST (form submission)
        remarks = request.form.get('remarks', '').strip()
        new_status = 'Approved' if action == 'approve' else 'Rejected'

        logger.info(f"Processing {action} for device {reg_id}")

        # Update device status in Supabase
        update_data = {
            'status': new_status,
            'approved_by': 'Supervisor',
            'approval_date': datetime.now().isoformat(),
            'admin_remarks': remarks if remarks else f"Action: {new_status} via web portal"
        }

        supabase.table('device_registrations') \
            .update(update_data) \
            .eq('registration_id', reg_id) \
            .execute()

        logger.info(f"✅ Updated {reg_id} status to {new_status}")

        # Send confirmation email to applicant
        applicant_email = device.get('email')
        subject = f"BYOD Registration - {new_status}" if action == 'approve' else f"BYOD Registration - {new_status}"
        
        if action == 'approve':
            email_body = f"""Dear {device.get('name')},

Your BYOD device registration has been APPROVED by your supervisor.

Registration ID: {reg_id}
Device: {device.get('device_make_model') or device.get('device_model')}

Your device will now proceed to IT security compliance inspection.
You will receive further instructions about scheduling the inspection.

Best regards,
NITDA BYOD System"""
        else:
            email_body = f"""Dear {device.get('name')},

We regret to inform you that your BYOD device registration has been REJECTED.

Registration ID: {reg_id}
Reason: {remarks if remarks else 'Not specified'}

If you believe this is in error, please contact your IT department.

Best regards,
NITDA BYOD System"""

        email_sent = send_approval_email(applicant_email, subject, email_body)

        # Send notification to IT department (if approved)
        if action == 'approve':
            it_subject = f"New Device Awaiting Inspection - {reg_id}"
            it_body = f"""A new device has been approved and is ready for IT inspection:

Registration ID: {reg_id}
Name: {device.get('name')}
Device: {device.get('device_make_model') or device.get('device_model')}
Serial: {device.get('serial_number')}
Supervisor Approval: Yes
Date Approved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Please schedule an inspection with the applicant at {applicant_email}"""
            
            send_approval_email(Config.IT_DEPARTMENT_EMAIL, it_subject, it_body)

        # Show success page
        return render_template_string(
            APPROVAL_PAGE_TEMPLATE,
            success=True,
            show_form=False,
            device_info=device,
            message=f"Device {action}al processed successfully. Confirmation email has been sent.",
            error=None
        )

    except Exception as e:
        logger.error(f"❌ Error processing endorsement: {e}", exc_info=True)
        return render_template_string(
            APPROVAL_PAGE_TEMPLATE,
            error=f"System error: {str(e)}",
            success=False,
            show_form=False,
            device_info=None
        ), 500


@app.errorhandler(404)
def not_found(error):
    """Handle 404 errors"""
    return render_template_string(
        APPROVAL_PAGE_TEMPLATE,
        error="Page not found",
        success=False,
        show_form=False,
        device_info=None
    ), 404


@app.errorhandler(500)
def server_error(error):
    """Handle 500 errors"""
    logger.error(f"❌ Server error: {error}")
    return render_template_string(
        APPROVAL_PAGE_TEMPLATE,
        error="Internal server error. Please try again later.",
        success=False,
        show_form=False,
        device_info=None
    ), 500


if __name__ == '__main__':
    logger.info("\n" + "=" * 70)
    logger.info("🚀 NITDA BYOD APPROVAL SERVER STARTING")
    logger.info("=" * 70)
    logger.info(f"Server: {Config.APPROVAL_SERVER_HOST}:{Config.APPROVAL_SERVER_PORT}")
    logger.info(f"Supabase: {Config.SUPABASE_URL[:50]}...")
    logger.info("=" * 70 + "\n")
    
    app.run(
        host=Config.APPROVAL_SERVER_HOST,
        port=Config.APPROVAL_SERVER_PORT,
        debug=Config.APPROVAL_SERVER_DEBUG
    )
