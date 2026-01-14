"""
Configuration file for Gmail to Sheets automation
"""

import os

# Google API Scopes
GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# File paths
CREDENTIALS_FILE = os.path.join('credentials', 'credentials.json')
TOKEN_FILE = os.path.join('credentials', 'token.json')
STATE_FILE = 'state.json'

# Google Sheet configuration
# You'll need to create a Google Sheet and paste its ID here
SPREADSHEET_ID = '18aCaoZZ4hblyqy9lRiYJWiKOyXg35Z4yUci6cgqAUdg'

SHEET_NAME = 'EmailLog'  # Name of the sheet tab

# Email configuration
MARK_AS_READ = True  # Mark emails as read after processing
MAX_RESULTS = 100  # Maximum emails to fetch per run

# Column headers for the sheet
SHEET_HEADERS = ['From', 'Subject', 'Date', 'Content']