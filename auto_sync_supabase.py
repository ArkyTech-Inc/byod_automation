"""
Google Sheets to Supabase Auto-Sync Script
Syncs Google Forms responses directly to Supabase database
Runs continuously in the background
"""
import os
import sys
import time
import logging
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from typing import Optional, List, Dict

import gspread
from google.oauth2.service_account import Credentials
from supabase import create_client, Client

from config import Config

# Load environment variables
load_dotenv()

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class GoogleSheetsSupabaseSync:
    """Syncs Google Forms responses to Supabase database"""

    def __init__(self, credentials_file: str, sheet_name: str):
        """
        Initialize the sync system
        
        Args:
            credentials_file: Path to Google service account JSON
            sheet_name: Name of the Google Sheet (where Form responses go)
        """
        self.credentials_file = credentials_file
        self.sheet_name = sheet_name
        self.gc = None
        self.supabase: Client = None
        self.last_row_count = 0
        
        # Initialize connections
        self._setup_google_api()
        self._setup_supabase()

    def _setup_google_api(self):
        """Initialize Google Sheets API connection"""
        try:
            if not os.path.exists(self.credentials_file):
                raise FileNotFoundError(f"Credentials file not found: {self.credentials_file}")
            
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = Credentials.from_service_account_file(self.credentials_file, scopes=scopes)
            self.gc = gspread.authorize(creds)
            logger.info("✅ Google Sheets API authenticated")
        except Exception as e:
            logger.error(f"❌ Failed to authenticate Google Sheets: {e}")
            raise

    def _setup_supabase(self):
        """Initialize Supabase connection"""
        try:
            self.supabase = create_client(Config.SUPABASE_URL, Config.SUPABASE_KEY)
            # Test connection
            self.supabase.table('device_registrations').select('count', count='exact').execute()
            logger.info("✅ Supabase connection established")
        except Exception as e:
            logger.error(f"❌ Failed to connect to Supabase: {e}")
            raise

    def get_form_responses(self) -> List[Dict]:
        """Fetch all form responses from Google Sheet"""
        try:
            sheet = self.gc.open(self.sheet_name).sheet1
            records = sheet.get_all_records()
            logger.debug(f"Fetched {len(records)} records from Google Sheet")
            return records
        except Exception as e:
            logger.error(f"❌ Error reading Google Sheet: {e}")
            return []

    def sync_to_supabase(self):
        """Sync new Google Form responses to Supabase"""
        try:
            # Get current form responses
            form_responses = self.get_form_responses()
            current_count = len(form_responses)
            
            if current_count == 0:
                logger.debug("No form responses found")
                return 0
            
            # Check if there are new responses
            if current_count <= self.last_row_count:
                logger.debug(f"No new responses ({current_count} total, {self.last_row_count} synced)")
                return 0
            
            # Process only new responses
            new_count = current_count - self.last_row_count
            logger.info(f"🟢 Found {new_count} new response(s). Syncing...")
            
            synced = 0
            for i, record in enumerate(form_responses[self.last_row_count:], 1):
                try:
                    # Generate registration ID
                    current_year = datetime.now().year
                    reg_id = f"BYOD-{current_year}{str(current_count - new_count + i).zfill(5)}"
                    
                    # Map Google Form columns to Supabase columns
                    # Adjust column names based on your actual Google Form
                    data = {
                        'registration_id': reg_id,
                        'timestamp': record.get('Timestamp', datetime.now().isoformat()),
                        'name': record.get('Full Name', ''),
                        'email': record.get('Email Address', ''),
                        'phone': record.get('Phone Number', ''),
                        'department': record.get('Department', ''),
                        'device_type': record.get('Device Type', ''),
                        'device_make_model': record.get('Device Make/Model', ''),
                        'operating_system': record.get('Operating System', ''),
                        'mac_address': record.get('MAC Address', ''),
                        'serial_number': record.get('Serial Number', ''),
                        'supervisor_name': record.get('Supervisor Name', ''),
                        'supervisor_email': record.get('Supervisor Email', ''),
                        'status': 'Pending'
                    }
                    
                    # Insert into Supabase
                    self.supabase.table('device_registrations').insert(data).execute()
                    logger.info(f"   ✓ Synced: {reg_id} - {data['name']}")
                    synced += 1
                    
                except Exception as e:
                    logger.error(f"   ❌ Error syncing record {i}: {e}")
                    continue
            
            # Update last synced count
            self.last_row_count = current_count
            logger.info(f"✅ Synced {synced}/{new_count} new registrations")
            return synced
            
        except Exception as e:
            logger.error(f"❌ Error in sync_to_supabase: {e}")
            return 0

    def run(self, check_interval: int = 60):
        """Run the sync continuously"""
        logger.info("\n" + "=" * 70)
        logger.info("🚀 GOOGLE FORMS → SUPABASE SYNC ENGINE STARTED")
        logger.info("=" * 70)
        logger.info(f"Sheet Name: {self.sheet_name}")
        logger.info(f"Check Interval: {check_interval} seconds")
        logger.info("Press Ctrl+C to stop")
        logger.info("=" * 70 + "\n")
        
        try:
            while True:
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                logger.debug(f"[{timestamp}] Checking for new form responses...")
                
                self.sync_to_supabase()
                
                logger.debug(f"Sleeping for {check_interval} seconds...\n")
                time.sleep(check_interval)
                
        except KeyboardInterrupt:
            logger.info("\n🛑 Sync stopped by user")
            sys.exit(0)
        except Exception as e:
            logger.error(f"❌ Fatal error: {e}")
            sys.exit(1)


def main():
    """Main entry point"""
    
    # Get configuration
    credentials_file = os.getenv('CREDENTIALS_JSON', 'credentials.json')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'NITDA BYOD Database')
    check_interval = int(os.getenv('CHECK_INTERVAL', '60'))
    
    logger.info("Google Forms → Supabase Sync Tool")
    logger.info("=" * 50)
    
    # Verify credentials file exists
    if not os.path.exists(credentials_file):
        logger.error(f"❌ Credentials file not found: {credentials_file}")
        logger.info(f"   Please place credentials.json in the current directory")
        sys.exit(1)
    
    # Create and run syncer
    try:
        syncer = GoogleSheetsSupabaseSync(credentials_file, sheet_name)
        syncer.run(check_interval=check_interval)
    except Exception as e:
        logger.error(f"❌ Failed to start sync: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
