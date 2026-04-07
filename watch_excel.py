"""
Excel File Watcher
Monitors Excel file for changes and runs automation
i used this to monitor the Excel file for changes and automatically run the automation script whenever new data is added.
It checks the file every 30 seconds and runs the automation if it detects any changes. This way, you don't have to manually run the automation script every time you add new registrations to the Excel file. Just run: python watch_excel.py
Although now implemented in byod_automation.py, you can still use this as a standalone script if you want to keep the automation and file watching separate.
"""

import time
import os
from datetime import datetime
import subprocess
import sys

class ExcelWatcher:
    def __init__(self, excel_file, check_interval=30):
        self.excel_file = excel_file
        self.check_interval = check_interval
        self.last_modified = 0
        
    def get_file_modified_time(self):
        """Get last modified time of Excel file"""
        try:
            return os.path.getmtime(self.excel_file)
        except:
            return 0
    
    def run(self):
        """Monitor Excel file and run automation on changes"""
        print("\n" + "="*70)
        print("EXCEL FILE WATCHER - STARTED")
        print("="*70)
        print(f"Monitoring: {self.excel_file}")
        print(f"Check interval: {self.check_interval} seconds")
        print("Press Ctrl+C to stop")
        print("="*70 + "\n")
        
        # Get initial modified time
        self.last_modified = self.get_file_modified_time()
        print(f"Initial file time: {datetime.fromtimestamp(self.last_modified)}")
        
        try:
            while True:
                time.sleep(self.check_interval)
                
                current_modified = self.get_file_modified_time()
                
                if current_modified > self.last_modified:
                    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] File changed detected!")
                    print("Running automation...")
                    
                    # Run automation
                    try:
                        result = subprocess.run(
                            [sys.executable, 'byod_automation.py'],
                            capture_output=True,
                            text=True
                        )
                        
                        print(result.stdout)
                        
                        if result.returncode == 0:
                            print("✓ Automation completed")
                        else:
                            print(f"✗ Automation failed: {result.stderr}")
                    
                    except Exception as e:
                        print(f"✗ Error running automation: {e}")
                    
                    self.last_modified = current_modified
                
                else:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] No changes detected")
        
        except KeyboardInterrupt:
            print("\n\nExcel watcher stopped")

if __name__ == "__main__":
    watcher = ExcelWatcher('NITDA_BYOD_Database.xlsx', check_interval=30)
    watcher.run()