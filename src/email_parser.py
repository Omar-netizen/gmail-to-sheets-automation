"""
Email parser module - extracts and cleans email data
"""

import base64
import email
from bs4 import BeautifulSoup
from datetime import datetime


def parse_email_date(date_string):
    """
    Parse email date string to readable format
    """
    try:
        # Parse the date string
        date_obj = email.utils.parsedate_to_datetime(date_string)
        # Format as: YYYY-MM-DD HH:MM:SS
        return date_obj.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        print(f"Error parsing date: {e}")
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def decode_base64(data):
    """
    Decode base64 encoded data
    """
    try:
        if data:
            return base64.urlsafe_b64decode(data).decode('utf-8')
        return ""
    except Exception as e:
        print(f"Error decoding base64: {e}")
        return ""


def html_to_text(html_content):
    """
    Convert HTML content to plain text
    """
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.decompose()
        # Get text and clean it
        text = soup.get_text()
        # Break into lines and remove leading/trailing space
        lines = (line.strip() for line in text.splitlines())
        # Break multi-headlines into a line each
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        # Remove blank lines
        text = '\n'.join(chunk for chunk in chunks if chunk)
        return text
    except Exception as e:
        print(f"Error converting HTML to text: {e}")
        return html_content


def get_email_body(payload):
    """
    Extract email body from payload (handles both plain text and HTML)
    """
    body = ""
    
    if 'parts' in payload:
        # Multi-part email
        for part in payload['parts']:
            mime_type = part.get('mimeType', '')
            
            if mime_type == 'text/plain':
                data = part.get('body', {}).get('data', '')
                body = decode_base64(data)
                break  # Prefer plain text
            elif mime_type == 'text/html':
                data = part.get('body', {}).get('data', '')
                html_body = decode_base64(data)
                body = html_to_text(html_body)
            elif mime_type.startswith('multipart/'):
                # Recursive for nested parts
                body = get_email_body(part)
                if body:
                    break
    else:
        # Single part email
        data = payload.get('body', {}).get('data', '')
        body = decode_base64(data)
        
        # If it's HTML, convert to text
        if payload.get('mimeType', '') == 'text/html':
            body = html_to_text(body)
    
    return body.strip()


def parse_email(message):
    """
    Parse email message and extract required fields
    Returns: dict with From, Subject, Date, Content
    """
    try:
        payload = message.get('payload', {})
        headers = payload.get('headers', [])
        
        # Extract headers
        email_data = {
            'From': '',
            'Subject': '',
            'Date': '',
            'Content': ''
        }
        
        for header in headers:
            name = header.get('name', '')
            value = header.get('value', '')
            
            if name == 'From':
                email_data['From'] = value
            elif name == 'Subject':
                email_data['Subject'] = value
            elif name == 'Date':
                email_data['Date'] = parse_email_date(value)
        
        # Extract body
        email_data['Content'] = get_email_body(payload)
        
        # Truncate content if too long (Google Sheets cell limit)
        if len(email_data['Content']) > 50000:
            email_data['Content'] = email_data['Content'][:50000] + "... [truncated]"
        
        return email_data
        
    except Exception as e:
        print(f"Error parsing email: {e}")
        return None