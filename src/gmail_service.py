"""
Gmail Service Module - Handles Gmail API authentication and operations
"""

import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import config


def authenticate_gmail():
    """
    Authenticate with Gmail API using OAuth 2.0
    Returns: Gmail API service object
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
        service = build('gmail', 'v1', credentials=creds)
        print("Gmail service initialized successfully")
        return service
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


def get_unread_emails(service, max_results=100):
    """
    Fetch unread emails from inbox
    Returns: List of email messages
    """
    try:
        print(f"Fetching unread emails (max: {max_results})...")
        
        # Query for unread emails in inbox
        results = service.users().messages().list(
            userId='me',
            q='is:unread in:inbox',
            maxResults=max_results
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            print("No unread emails found.")
            return []
        
        print(f"Found {len(messages)} unread email(s)")
        
        # Fetch full message details for each email
        full_messages = []
        for i, message in enumerate(messages, 1):
            try:
                print(f"Fetching email {i}/{len(messages)}...", end='\r')
                msg = service.users().messages().get(
                    userId='me',
                    id=message['id'],
                    format='full'
                ).execute()
                full_messages.append(msg)
            except HttpError as error:
                print(f"\nError fetching message {message['id']}: {error}")
                continue
        
        print(f"\nSuccessfully fetched {len(full_messages)} email(s)")
        return full_messages
        
    except HttpError as error:
        print(f"An error occurred while fetching emails: {error}")
        return []


def mark_email_as_read(service, message_id):
    """
    Mark an email as read
    """
    try:
        service.users().messages().modify(
            userId='me',
            id=message_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
        return True
    except HttpError as error:
        print(f"Error marking email as read: {error}")
        return False


def mark_emails_as_read(service, message_ids):
    """
    Mark multiple emails as read
    """
    success_count = 0
    for msg_id in message_ids:
        if mark_email_as_read(service, msg_id):
            success_count += 1
    
    print(f"Marked {success_count}/{len(message_ids)} email(s) as read")
    return success_count