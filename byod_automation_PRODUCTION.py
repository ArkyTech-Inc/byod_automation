"""
NITDA BYOD Automation System - Complete Supabase Implementation
Handles email notifications, IT inspection scheduling, and QR code generation
Production-ready with comprehensive error handling and logging
"""
import os
import logging
import qrcode
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
from io import BytesIO
from typing import Optional, Dict, List, Tuple
from pathlib import Path

from supabase import create_client, Client
from dotenv import load_dotenv

from config import Config

# Load environment variables
load_dotenv()

# Set up logging
logger = logging.getLogger(__name__)


class BYODAutomationException(Exception):
    """Custom exception for BYOD automation errors"""
    pass


class DatabaseConnectionError(BYODAutomationException):
    """Raised when database connection fails"""
    pass


class EmailSendError(BYODAutomationException):
    """Raised when email sending fails"""
    pass


class BYODAutomation:
    """
    Main automation engine for BYOD system
    Handles all automation workflows including approvals, inspections, and QR codes
    """

    def __init__(self):
        """Initialize the automation engine with Supabase connection"""
        try:
            logger.info("🚀 Initializing BYOD Automation Engine...")
            
            # Initialize Supabase client
            self.supabase: Client = create_client(Config.SUPABASE_URL, Config.SUPABASE_KEY)
            
            # Test connection
            self._test_database_connection()
            
            # Load settings from database
            self.smtp_server = Config.SMTP_SERVER
            self.smtp_port = Config.SMTP_PORT
            self.sender_email = Config.SENDER_EMAIL
            self.sender_password = Config.SENDER_PASSWORD
            self.it_email = self._get_setting("IT Email") or Config.IT_DEPARTMENT_EMAIL
            self.inspection_lead_time = int(self._get_setting("Inspection Lead Time (days)") or Config.INSPECTION_LEAD_TIME)
            
            logger.info(f"✅ Automation engine initialized successfully")
            logger.info(f"   SMTP Server: {self.smtp_server}:{self.smtp_port}")
            logger.info(f"   Sender Email: {self.sender_email}")
            logger.info(f"   IT Email: {self.it_email}")
            
        except Exception as e:
            logger.error(f"❌ Failed to initialize automation engine: {e}", exc_info=True)
            raise DatabaseConnectionError(f"Failed to connect to Supabase: {e}")

    def _test_database_connection(self) -> bool:
        """Test connection to Supabase database"""
        try:
            # Try to fetch one row to test connection
            self.supabase.table('device_registrations').select('id').limit(1).execute()
            logger.debug("✓ Database connection test successful")
            return True
        except Exception as e:
            logger.error(f"❌ Database connection test failed: {e}")
            raise DatabaseConnectionError(f"Cannot connect to Supabase: {e}")

    def _get_setting(self, key: str) -> Optional[str]:
        """Retrieve a setting from the automation_settings table"""
        try:
            response = self.supabase.table('automation_settings') \
                .select('value') \
                .eq('key', key) \
                .execute()
            
            if response.data and len(response.data) > 0:
                return response.data[0]['value']
            return None
        except Exception as e:
            logger.warning(f"⚠️ Could not retrieve setting '{key}': {e}")
            return None

    def _get_email_template(self, template_name: str) -> Optional[Dict[str, str]]:
        """Retrieve email template from database"""
        try:
            response = self.supabase.table('email_templates') \
                .select('subject, body') \
                .eq('template_name', template_name) \
                .execute()
            
            if response.data and len(response.data) > 0:
                return {
                    'subject': response.data[0]['subject'],
                    'body': response.data[0]['body']
                }
            logger.warning(f"⚠️ Email template '{template_name}' not found")
            return None
        except Exception as e:
            logger.error(f"❌ Error retrieving email template '{template_name}': {e}")
            return None

    def send_email(self, to_email: str, subject: str, body: str, 
                   qr_image_bytes: Optional[bytes] = None) -> Tuple[bool, Optional[str]]:
        """
        Send email with optional QR code attachment
        Returns (success, error_message)
        """
        try:
            if not Config.ENABLE_EMAIL_NOTIFICATIONS:
                logger.warning(f"⚠️ Email notifications disabled. Skipping email to {to_email}")
                return False, "Email notifications disabled"
            
            logger.debug(f"📧 Sending email to: {to_email}")
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # Add body
            msg.attach(MIMEText(body, 'plain'))
            
            # Add QR code attachment if provided
            if qr_image_bytes:
                try:
                    img = MIMEImage(qr_image_bytes)
                    img.add_header('Content-Disposition', 'attachment', filename="BYOD_Pass.png")
                    msg.attach(img)
                    logger.debug("✓ QR code attached to email")
                except Exception as e:
                    logger.error(f"❌ Failed to attach QR code: {e}")
                    # Don't fail entirely, just skip the attachment
            
            # Send email
            try:
                if self.smtp_port == 465:
                    # SSL connection
                    server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, timeout=10)
                else:
                    # TLS connection
                    server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=10)
                    server.starttls()
                
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
                server.quit()
                
                logger.info(f"✅ Email sent successfully to {to_email}: {subject}")
                return True, None
                
            except smtplib.SMTPAuthenticationError as e:
                error_msg = "SMTP Authentication failed. Check email credentials."
                logger.error(f"❌ {error_msg}: {e}")
                return False, error_msg
            except smtplib.SMTPException as e:
                error_msg = f"SMTP error: {str(e)}"
                logger.error(f"❌ {error_msg}")
                return False, error_msg
            except Exception as e:
                error_msg = f"Unexpected error sending email: {str(e)}"
                logger.error(f"❌ {error_msg}")
                return False, error_msg
                
        except Exception as e:
            logger.error(f"❌ Failed to prepare email: {e}", exc_info=True)
            return False, str(e)

    def process_new_registrations(self) -> int:
        """
        Process new pending registrations and send supervisor approval emails
        Returns count of registrations processed
        """
        try:
            logger.info("📋 Processing new device registrations...")
            
            # Query for pending registrations that haven't had emails sent
            response = self.supabase.table('device_registrations') \
                .select('*') \
                .eq('status', 'Pending') \
                .not_.ilike('admin_remarks', '%Emails Sent%') \
                .execute()
            
            registrations = response.data or []
            logger.debug(f"Found {len(registrations)} new registrations to process")
            
            processed_count = 0
            
            for reg in registrations:
                try:
                    reg_id = reg.get('registration_id')
                    supervisor_email = reg.get('supervisor_email')
                    supervisor_name = reg.get('supervisor_name')
                    applicant_name = reg.get('name')
                    department = reg.get('department')
                    device_model = reg.get('device_make_model') or reg.get('device_model') or 'Unknown'
                    
                    if not supervisor_email:
                        logger.warning(f"⚠️ Registration {reg_id} missing supervisor email. Skipping.")
                        continue
                    
                    # Build approval links
                    base_url = Config.APPROVAL_SERVER_URL
                    approve_link = f"{base_url}/endorse?id={reg_id}&action=approve"
                    reject_link = f"{base_url}/endorse?id={reg_id}&action=reject"
                    
                    # Get email template
                    template = self._get_email_template('Supervisor Approval Request')
                    if not template:
                        logger.warning(f"⚠️ Email template 'Supervisor Approval Request' not found")
                        continue
                    
                    # Format email
                    subject = template['subject'].format(
                        supervisor_name=supervisor_name,
                        name=applicant_name
                    )
                    
                    body = template['body'].format(
                        supervisor_name=supervisor_name,
                        name=applicant_name,
                        department=department,
                        device_model=device_model,
                        registration_id=reg_id,
                        approval_link=approve_link,
                        rejection_link=reject_link
                    )
                    
                    # Send email
                    success, error = self.send_email(supervisor_email, subject, body)
                    
                    if success:
                        # Mark as processed
                        remarks = reg.get('admin_remarks') or ''
                        new_remarks = f"{remarks} | Emails Sent".strip(" | ")
                        
                        self.supabase.table('device_registrations') \
                            .update({'admin_remarks': new_remarks}) \
                            .eq('registration_id', reg_id) \
                            .execute()
                        
                        logger.info(f"✅ Approval request sent for {reg_id}")
                        processed_count += 1
                    else:
                        logger.error(f"❌ Failed to send approval request for {reg_id}: {error}")
                        # Don't mark as processed if email failed
                
                except Exception as e:
                    logger.error(f"❌ Error processing registration: {e}", exc_info=True)
                    continue
            
            logger.info(f"Processed {processed_count}/{len(registrations)} new registrations")
            return processed_count
            
        except Exception as e:
            logger.error(f"❌ Error in process_new_registrations: {e}", exc_info=True)
            return 0

    def process_approved_devices(self) -> int:
        """
        Process approved devices and schedule IT inspections
        Returns count of devices processed
        """
        try:
            logger.info("✅ Processing approved devices...")
            
            # Query for approved devices without inspection records
            response = self.supabase.table('device_registrations') \
                .select('*') \
                .eq('status', 'Approved') \
                .execute()
            
            approved_devices = response.data or []
            logger.debug(f"Found {len(approved_devices)} approved devices")
            
            processed_count = 0
            
            for device in approved_devices:
                try:
                    reg_id = device.get('registration_id')
                    
                    # Check if inspection already scheduled
                    inspection_check = self.supabase.table('it_inspections') \
                        .select('id') \
                        .eq('registration_id', reg_id) \
                        .execute()
                    
                    if inspection_check.data and len(inspection_check.data) > 0:
                        logger.debug(f"   Inspection already scheduled for {reg_id}")
                        continue
                    
                    # Calculate inspection date
                    inspection_date = (datetime.now() + timedelta(days=self.inspection_lead_time)).strftime('%Y-%m-%d')
                    inspection_id = f"INS-{reg_id.replace('BYOD-', '')}"
                    
                    # Create inspection record
                    self.supabase.table('it_inspections').insert({
                        'registration_id': reg_id,
                        'name': device.get('name'),
                        'device_model': device.get('device_make_model') or device.get('device_model'),
                        'serial_number': device.get('serial_number'),
                        'inspection_id': inspection_id,
                        'inspection_date': inspection_date,
                        'compliance_status': 'Pending',
                        'qr_code_generated': 'No'
                    }).execute()
                    
                    logger.debug(f"✓ Created inspection record {inspection_id}")
                    
                    # Send inspection scheduled email to applicant
                    template = self._get_email_template('IT Inspection Schedule')
                    if template:
                        subject = template['subject'].format(registration_id=reg_id)
                        body = template['body'].format(
                            name=device.get('name'),
                            registration_id=reg_id,
                            device_model=device.get('device_make_model') or device.get('device_model'),
                            inspection_date=inspection_date
                        )
                        
                        applicant_email = device.get('email')
                        success, error = self.send_email(applicant_email, subject, body)
                        if not success:
                            logger.error(f"❌ Failed to send inspection email: {error}")
                    
                    # Send notification to IT department
                    it_subject = f"IT Inspection Scheduled - {reg_id}"
                    it_body = f"""New device scheduled for IT security inspection:

Registration ID: {reg_id}
Name: {device.get('name')}
Device: {device.get('device_make_model') or device.get('device_model')}
Serial Number: {device.get('serial_number')}
Scheduled Date: {inspection_date}

Please contact the applicant at {applicant_email} to confirm inspection time."""
                    
                    success, error = self.send_email(self.it_email, it_subject, it_body)
                    if not success:
                        logger.warning(f"⚠️ Failed to notify IT department: {error}")
                    
                    logger.info(f"✅ Scheduled inspection for {reg_id}")
                    processed_count += 1
                    
                except Exception as e:
                    logger.error(f"❌ Error processing approved device: {e}", exc_info=True)
                    continue
            
            logger.info(f"📊 Processed {processed_count} approved devices")
            return processed_count
            
        except Exception as e:
            logger.error(f"❌ Error in process_approved_devices: {e}", exc_info=True)
            return 0

    def process_compliant_devices(self) -> int:
        """
        Process compliant devices and generate QR code passes
        Returns count of devices processed
        """
        try:
            logger.info("🔐 Processing compliant devices and generating QR codes...")
            
            # Query for compliant devices without QR codes
            response = self.supabase.table('it_inspections') \
                .select('*') \
                .eq('compliance_status', 'Compliant') \
                .eq('qr_code_generated', 'No') \
                .execute()
            
            compliant_devices = response.data or []
            logger.debug(f"Found {len(compliant_devices)} compliant devices needing QR codes")
            
            # Create QR codes directory if it doesn't exist
            Path(Config.QR_STORAGE_PATH).mkdir(parents=True, exist_ok=True)
            
            processed_count = 0
            
            for inspection in compliant_devices:
                try:
                    reg_id = inspection.get('registration_id')
                    
                    # Get device registration details
                    reg_response = self.supabase.table('device_registrations') \
                        .select('*') \
                        .eq('registration_id', reg_id) \
                        .execute()
                    
                    if not reg_response.data:
                        logger.warning(f"⚠️ Device registration not found for {reg_id}")
                        continue
                    
                    device = reg_response.data[0]
                    
                    # Generate QR code data
                    qr_data = f"ID:{reg_id}|NAME:{inspection.get('name')}|MODEL:{inspection.get('device_model')}|MAC:{device.get('mac_address')}|SER:{inspection.get('serial_number')}|STATUS:COMPLIANT"
                    
                    # Create QR code image
                    qr = qrcode.QRCode(
                        version=Config.QR_CODE_VERSION,
                        box_size=Config.QR_CODE_BOX_SIZE,
                        border=Config.QR_CODE_BORDER
                    )
                    qr.add_data(qr_data)
                    qr.make(fit=True)
                    img = qr.make_image(fill_color="black", back_color="white")
                    
                    # Save QR code to file
                    qr_filename = f"{Config.QR_STORAGE_PATH}{reg_id}_pass.png"
                    img.save(qr_filename)
                    logger.debug(f"✓ Saved QR code to {qr_filename}")
                    
                    # Get QR code bytes for email
                    buffer = BytesIO()
                    img.save(buffer, format="PNG")
                    qr_bytes = buffer.getvalue()
                    
                    # Send QR code email to applicant
                    template = self._get_email_template('Compliance Pass Issued')
                    if template:
                        subject = template['subject'].format(registration_id=reg_id)
                        body = template['body'].format(
                            name=device.get('name'),
                            registration_id=reg_id,
                            device_model=inspection.get('device_model')
                        )
                        
                        success, error = self.send_email(
                            device.get('email'),
                            subject,
                            body,
                            qr_image_bytes=qr_bytes
                        )
                        if not success:
                            logger.error(f"❌ Failed to send QR code email: {error}")
                            continue
                    
                    # Update inspection record
                    self.supabase.table('it_inspections') \
                        .update({
                            'qr_code_generated': 'Yes',
                            'pass_issued_date': datetime.now().isoformat()
                        }) \
                        .eq('registration_id', reg_id) \
                        .execute()
                    
                    logger.info(f"✅ Generated QR code pass for {reg_id}")
                    processed_count += 1
                    
                except Exception as e:
                    logger.error(f"❌ Error processing compliant device: {e}", exc_info=True)
                    continue
            
            logger.info(f"📊 Generated {processed_count} QR code passes")
            return processed_count
            
        except Exception as e:
            logger.error(f"❌ Error in process_compliant_devices: {e}", exc_info=True)
            return 0

    def run_automation(self) -> None:
        """
        Execute all automation workflows
        This is the main entry point called by the background daemon
        """
        try:
            logger.info("=" * 70)
            logger.info(f"🚀 BYOD AUTOMATION RUN - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info("=" * 70)
            
            # Run all processing stages
            new_regs = self.process_new_registrations()
            approved_devs = self.process_approved_devices()
            qr_codes = self.process_compliant_devices()
            
            logger.info("=" * 70)
            logger.info(f"✅ Automation cycle complete")
            logger.info(f"   New registrations: {new_regs}")
            logger.info(f"   Approved devices: {approved_devs}")
            logger.info(f"   QR codes generated: {qr_codes}")
            logger.info("=" * 70)
            
        except Exception as e:
            logger.error(f"❌ CRITICAL ERROR in automation: {e}", exc_info=True)
            # Don't raise, just log so background daemon can continue
