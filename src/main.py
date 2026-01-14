"""
Main script - Orchestrates Gmail to Sheets automation
"""

import json
import os
import sys
from datetime import datetime

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import config
from src.gmail_service import authenticate_gmail, get_unread_emails, mark_emails_as_read
from src.sheets_service import (
    authenticate_sheets, 
    initialize_sheet, 
    batch_append_to_sheet,
    get_existing_emails
)
from src.email_parser import parse_email


def load_state():
    """
    Load state from state.json
    Returns: dict with processed_message_ids
    """
    if os.path.exists(config.STATE_FILE):
        try:
            with open(config.STATE_FILE, 'r') as f:
                state = json.load(f)
                print(f"Loaded state: {len(state.get('processed_message_ids', []))} processed email(s)")
                return state
        except Exception as e:
            print(f"Error loading state: {e}")
            return {'processed_message_ids': []}
    else:
        print("No state file found, starting fresh")
        return {'processed_message_ids': []}


def save_state(state):
    """
    Save state to state.json
    """
    try:
        with open(config.STATE_FILE, 'w') as f:
            json.dump(state, f, indent=2)
        print(f"State saved: {len(state.get('processed_message_ids', []))} total processed email(s)")
    except Exception as e:
        print(f"Error saving state: {e}")


def main():
    """
    Main execution function
    """
    print("=" * 60)
    print("Gmail to Google Sheets Automation")
    print("=" * 60)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Check if spreadsheet ID is configured
    if config.SPREADSHEET_ID == 'YOUR_SPREADSHEET_ID_HERE':
        print("ERROR: Please set SPREADSHEET_ID in config.py")
        print("Create a Google Sheet and copy its ID from the URL")
        return
    
    # Load state
    state = load_state()
    processed_ids = set(state.get('processed_message_ids', []))
    
    # Authenticate services
    print("\n[Step 1] Authenticating with Google APIs...")
    gmail_service = authenticate_gmail()
    sheets_service = authenticate_sheets()
    
    if not gmail_service or not sheets_service:
        print("ERROR: Authentication failed")
        return
    
    # Initialize sheet with headers
    print("\n[Step 2] Initializing Google Sheet...")
    initialize_sheet(sheets_service, config.SPREADSHEET_ID, config.SHEET_NAME)
    
    # Get existing emails from sheet for duplicate check
    print("\n[Step 3] Checking for existing emails in sheet...")
    existing_emails = get_existing_emails(sheets_service, config.SPREADSHEET_ID, config.SHEET_NAME)
    
    # Fetch unread emails
    print("\n[Step 4] Fetching unread emails from Gmail...")
    messages = get_unread_emails(gmail_service, config.MAX_RESULTS)
    
    if not messages:
        print("\nNo new emails to process")
        print("=" * 60)
        return
    
    # Filter out already processed messages
    print("\n[Step 5] Filtering new emails...")
    new_messages = []
    for msg in messages:
        msg_id = msg['id']
        if msg_id not in processed_ids:
            new_messages.append(msg)
        else:
            print(f"Skipping already processed email: {msg_id}")
    
    if not new_messages:
        print("All emails have already been processed")
        print("=" * 60)
        return
    
    print(f"Found {len(new_messages)} new email(s) to process")
    
    # Parse emails
    print("\n[Step 6] Parsing email data...")
    parsed_emails = []
    message_ids_to_mark = []
    
    for i, msg in enumerate(new_messages, 1):
        print(f"Parsing email {i}/{len(new_messages)}...", end='\r')
        
        email_data = parse_email(msg)
        
        if email_data:
            # Create identifier for duplicate check
            identifier = (
                email_data.get('From', ''),
                email_data.get('Subject', ''),
                email_data.get('Date', '')
            )
            
            # Check if this email already exists in sheet
            if identifier not in existing_emails:
                parsed_emails.append(email_data)
                message_ids_to_mark.append(msg['id'])
            else:
                print(f"\nSkipping duplicate: {email_data.get('Subject', 'No Subject')}")
    
    print(f"\nSuccessfully parsed {len(parsed_emails)} unique email(s)")
    
    if not parsed_emails:
        print("No new unique emails to add to sheet")
        print("=" * 60)
        return
    
    # Append to Google Sheet
    print("\n[Step 7] Adding emails to Google Sheet...")
    rows_added = batch_append_to_sheet(
        sheets_service,
        config.SPREADSHEET_ID,
        config.SHEET_NAME,
        parsed_emails
    )
    
    if rows_added > 0:
        # Mark emails as read
        if config.MARK_AS_READ and message_ids_to_mark:
            print("\n[Step 8] Marking emails as read...")
            mark_emails_as_read(gmail_service, message_ids_to_mark)
        
        # Update state
        print("\n[Step 9] Updating state...")
        for msg_id in message_ids_to_mark:
            processed_ids.add(msg_id)
        
        state['processed_message_ids'] = list(processed_ids)
        state['last_run'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        save_state(state)
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total unread emails found: {len(messages)}")
    print(f"New emails processed: {len(parsed_emails)}")
    print(f"Rows added to sheet: {rows_added}")
    print(f"Total processed (all time): {len(processed_ids)}")
    print(f"Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScript interrupted by user")
    except Exception as e:
        print(f"\n\nERROR: {e}")
        import traceback
        traceback.print_exc()