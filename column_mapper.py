"""
Google Form Column Mapper
Helps identify the correct column mapping from your Google Form.
Used it when there was a data integration mismatch with google form and excel data propagation
"""

import gspread
from google.oauth2.service_account import Credentials

def analyze_google_form_columns():
    """Analyze Google Form columns and show what's in each one"""
    
    print("\n" + "="*70)
    print("GOOGLE FORM COLUMN ANALYZER")
    print("="*70)
    
    # Connect to Google Sheets
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    client = gspread.authorize(creds)
    
    # Open sheet
    spreadsheet_name = input("\nEnter your Google Sheet name: ").strip()
    sheet = client.open(spreadsheet_name).sheet1
    
    # Get all data
    all_values = sheet.get_all_values()
    
    if len(all_values) < 2:
        print("No data in sheet!")
        return
    
    headers = all_values[0]
    first_row = all_values[1] if len(all_values) > 1 else []
    
    print("\n" + "="*70)
    print("GOOGLE FORM COLUMNS (with sample data)")
    print("="*70)
    
    for i, header in enumerate(headers):
        sample = first_row[i] if i < len(first_row) else "No data"
        print(f"\nColumn {i}:")
        print(f"  Header: {header}")
        print(f"  Sample: {sample[:60]}..." if len(sample) > 60 else f"  Sample: {sample}")
    
    print("\n" + "="*70)
    print("EXPECTED EXCEL COLUMNS")
    print("="*70)
    
    expected = [
        ("Timestamp", "When the form was submitted"),
        ("Email", "User's email address"),
        ("Name", "Full name"),
        ("Department", "Department (DLCR, ITSM, etc.)"),
        ("Classification", "Staff, Intern, Guest, etc."),
        ("Phone", "Phone number"),
        ("Supervisor Name", "Name of supervisor"),
        ("Device Type", "Laptop, Smartphone, etc."),
        ("Device Make/Model", "HP EliteBook, etc."),
        ("Operating System", "Windows 11, macOS, etc."),
        ("MAC Address", "XX:XX:XX:XX:XX:XX"),
        ("Serial Number", "Device serial number"),
        ("Supervisor Email", "Supervisor's email"),
        ("Policy Checkboxes", "Agreement checkboxes - SKIP THIS")
    ]
    
    for i, (name, desc) in enumerate(expected):
        print(f"\n{i+1}. {name}")
        print(f"   {desc}")
    
    print("\n" + "="*70)
    print("MAPPING INSTRUCTIONS")
    print("="*70)
    print("""
To create the correct mapping:

For each Excel column, find which Google Form column (number) contains that data.

Example:
  If "Email" is in Google Form column 1, write: email_col = 1
  If "Name" is in Google Form column 2, write: name_col = 2
  
Write down your mapping, then update auto_sync.py
""")

if __name__ == "__main__":
    analyze_google_form_columns()
