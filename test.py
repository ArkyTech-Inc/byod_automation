import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os

load_dotenv()
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
gc = gspread.authorize(creds)

# Try to list all sheets the service account can access
try:
    sheets = gc.list_spreadsheet_files()
    print("Accessible sheets:")
    for sheet in sheets[:10]:
        print(f"  - {sheet['name']}")
except Exception as e:
    print(f"Error: {e}")
