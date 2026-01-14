"""
Google Sheets Service Module - Handles Sheets API authentication and operations
"""

import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import config


def authenticate_sheets():
    """
    Authenticate with Google Sheets API using OAuth 2.0
    Returns: Sheets API service object
    """
    creds = None
    
    # Combine all scopes (Gmail + Sheets) for single authentication
    ALL_SCOPES = config.GMAIL_SCOPES + config.SHEETS_SCOPES
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists(config.TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(config.TOKEN_FILE, ALL_SCOPES)
        except Exception as e:
            print(f"Error loading credentials: {e}")
            creds = None
    
    # If no valid credentials, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print("Refreshing access token...")
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}")
                creds = None
        
        if not creds:
            if not os.path.exists(config.CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"Credentials file not found at {config.CREDENTIALS_FILE}\n"
                    "Please download credentials.json from Google Cloud Console"
                )
            
            print("Starting OAuth 2.0 flow for Gmail AND Sheets...")
            flow = InstalledAppFlow.from_client_secrets_file(
                config.CREDENTIALS_FILE,
                ALL_SCOPES  # Request both Gmail and Sheets scopes
            )
            creds = flow.run_local_server(port=0)
            print("Authentication successful!")
        
        # Save credentials for future runs
        with open(config.TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    try:
        service = build('sheets', 'v4', credentials=creds)
        print("Sheets service initialized successfully")
        return service
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def initialize_sheet(service, spreadsheet_id, sheet_name):
    """
    Initialize the sheet with headers if not already present
    """
    try:
        # Check if headers exist
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A1:D1'
        ).execute()
        
        values = result.get('values', [])
        
        if not values or values[0] != config.SHEET_HEADERS:
            # Headers don't exist or are incorrect, write them
            print("Writing headers to sheet...")
            body = {
                'values': [config.SHEET_HEADERS]
            }
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f'{sheet_name}!A1:D1',
                valueInputOption='RAW',
                body=body
            ).execute()
            print("Headers written successfully")
        else:
            print("Headers already exist")
            
    except HttpError as error:
        if error.resp.status == 404:
            print(f"Sheet '{sheet_name}' not found. Please create it first.")
        else:
            print(f"Error initializing sheet: {error}")


def append_to_sheet(service, spreadsheet_id, sheet_name, email_data):
    """
    Append email data to the sheet
    email_data: dict with keys From, Subject, Date, Content
    """
    try:
        # Prepare row data
        row = [
            email_data.get('From', ''),
            email_data.get('Subject', ''),
            email_data.get('Date', ''),
            email_data.get('Content', '')
        ]
        
        body = {
            'values': [row]
        }
        
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:D',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
        return True
        
    except HttpError as error:
        print(f"Error appending to sheet: {error}")
        return False


def batch_append_to_sheet(service, spreadsheet_id, sheet_name, emails_data):
    """
    Batch append multiple emails to the sheet
    emails_data: list of dicts with keys From, Subject, Date, Content
    """
    if not emails_data:
        print("No data to append")
        return 0
    
    try:
        # Prepare rows data
        rows = []
        for email_data in emails_data:
            row = [
                email_data.get('From', ''),
                email_data.get('Subject', ''),
                email_data.get('Date', ''),
                email_data.get('Content', '')
            ]
            rows.append(row)
        
        body = {
            'values': rows
        }
        
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:D',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
        updates = result.get('updates', {})
        rows_added = updates.get('updatedRows', 0)
        print(f"Successfully added {rows_added} row(s) to sheet")
        
        return rows_added
        
    except HttpError as error:
        print(f"Error batch appending to sheet: {error}")
        return 0


def get_existing_emails(service, spreadsheet_id, sheet_name):
    """
    Get all existing email subjects and dates from the sheet to check for duplicates
    Returns: set of tuples (from, subject, date)
    """
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A2:D'  # Skip header row
        ).execute()
        
        values = result.get('values', [])
        
        # Create set of unique identifiers (from + subject + date)
        existing = set()
        for row in values:
            if len(row) >= 3:
                # Create tuple of (From, Subject, Date)
                identifier = (
                    row[0] if len(row) > 0 else '',
                    row[1] if len(row) > 1 else '',
                    row[2] if len(row) > 2 else ''
                )
                existing.add(identifier)
        
        print(f"Found {len(existing)} existing email(s) in sheet")
        return existing
        
    except HttpError as error:
        if error.resp.status == 404:
            print(f"Sheet '{sheet_name}' not found")
        else:
            print(f"Error reading existing emails: {error}")
        return set()