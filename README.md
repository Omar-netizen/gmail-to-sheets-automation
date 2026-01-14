# Gmail to Google Sheets Automation

A Python automation system that reads unread emails from Gmail and logs them into Google Sheets using official Google APIs.

---

## 📋 Table of Contents

- [Architecture](#architecture)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Duplicate Prevention](#duplicate-prevention)
- [State Management](#state-management)
- [Challenges & Solutions](#challenges--solutions)
- [Limitations](#limitations)
- [Project Structure](#project-structure)

---

## 🏗️ Architecture

```
┌─────────────┐
│   Gmail     │
│   Inbox     │
└──────┬──────┘
       │
       │ Gmail API (OAuth 2.0)
       │ Fetch unread emails
       ▼
┌─────────────────────┐
│   Python Script     │
│                     │
│  ┌───────────────┐  │
│  │ gmail_service │  │ ← Fetch & mark as read
│  └───────────────┘  │
│         │           │
│         ▼           │
│  ┌───────────────┐  │
│  │ email_parser  │  │ ← Parse email data
│  └───────────────┘  │
│         │           │
│         ▼           │
│  ┌───────────────┐  │
│  │sheets_service │  │ ← Write to sheet
│  └───────────────┘  │
│         │           │
│  ┌───────────────┐  │
│  │  state.json   │  │ ← Track processed IDs
│  └───────────────┘  │
└─────────┬───────────┘
          │
          │ Sheets API (OAuth 2.0)
          │ Append email data
          ▼
   ┌─────────────┐
   │   Google    │
   │   Sheet     │
   └─────────────┘
```

**Flow:**
1. Authenticate with Gmail & Sheets APIs using OAuth 2.0
2. Fetch unread emails from Gmail inbox
3. Check state.json for already processed email IDs
4. Parse email data (From, Subject, Date, Content)
5. Check Google Sheet for duplicate entries
6. Append new emails to Google Sheet
7. Mark processed emails as read in Gmail
8. Update state.json with processed email IDs

---

## ✨ Features

- ✅ OAuth 2.0 authentication (no service accounts)
- ✅ Reads unread emails from Gmail inbox
- ✅ Extracts: Sender, Subject, Date, Email body
- ✅ Appends data to Google Sheets automatically
- ✅ Marks processed emails as read
- ✅ **Triple-layer duplicate prevention**
- ✅ State persistence across runs
- ✅ HTML to plain text conversion (Bonus)
- ✅ Batch processing for efficiency
- ✅ Error handling with retry logic
- ✅ Progress indicators in terminal

---

## 📷 Screenshots
<img width="1917" height="861" alt="image" src="https://github.com/user-attachments/assets/821f4a99-0354-4d8e-871b-988d82cc03a6" />
<img width="1918" height="846" alt="image" src="https://github.com/user-attachments/assets/2d868cc6-802b-4486-8970-9f1a45b2e5d7" />

<img width="1917" height="901" alt="image" src="https://github.com/user-attachments/assets/43d88e91-9ed2-47ae-88e0-403a5165dd98" />


## 📦 Prerequisites

- Python 3.8 or higher
- Gmail account
- Google Cloud Project with:
  - Gmail API enabled
  - Google Sheets API enabled
  - OAuth 2.0 credentials

---

## 🚀 Installation

### 1. Clone the repository

```bash
git clone <your-repo-url>
cd gmail-to-sheets
```

### 2. Create virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Set up Google Cloud Console

#### a. Create a Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project: `Gmail-Sheets-Automation`
3. Select the project

#### b. Enable APIs
1. Search for "Gmail API" → Enable
2. Search for "Google Sheets API" → Enable

#### c. Configure OAuth Consent Screen
1. Go to "APIs & Services" → "OAuth consent screen"
2. Choose "External" → Create
3. Fill in:
   - App name: `Gmail to Sheets`
   - User support email: your email
   - Developer contact: your email
4. Click "Save and Continue"
5. Add scopes:
   - `https://www.googleapis.com/auth/gmail.modify`
   - `https://www.googleapis.com/auth/spreadsheets`
6. Add test users: your Gmail address
7. Save and continue

#### d. Create OAuth Credentials
1. Go to "Credentials" → "+ CREATE CREDENTIALS"
2. Choose "OAuth client ID"
3. Application type: "Desktop app"
4. Name: `Gmail Sheets Desktop Client`
5. Click "Create"
6. **Download JSON** file
7. Rename to `credentials.json`
8. Move to `credentials/` folder in project

### 5. Create Google Sheet

1. Go to [Google Sheets](https://sheets.google.com/)
2. Create a new blank spreadsheet
3. Name it: `Email Log`
4. Rename the tab to: `EmailLog` (case-sensitive!)
5. Copy the Spreadsheet ID from URL:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit
   ```

### 6. Configure the project

Edit `config.py`:

```python
SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'  # Paste your actual ID
SHEET_NAME = 'EmailLog'  # Must match your sheet tab name exactly
```

---

## ⚙️ Configuration

Edit `config.py` to customize:

```python
# Maximum emails to fetch per run
MAX_RESULTS = 100

# Mark emails as read after processing
MARK_AS_READ = True

# Google Sheet configuration
SPREADSHEET_ID = 'your-spreadsheet-id'
SHEET_NAME = 'EmailLog'
```

---

## 🎮 Usage

### First Run

```bash
python src/main.py
```

**What happens:**
1. Browser opens for OAuth authentication
2. Sign in with your Gmail account
3. Grant permissions (you'll see "Google hasn't verified this app" - click Advanced → Continue)
4. Browser shows "Authentication complete"
5. Script fetches and processes unread emails
6. Check your Google Sheet - data should be there!

### Subsequent Runs

```bash
python src/main.py
```

- No browser authentication needed (token saved)
- Processes only new unread emails
- No duplicates created

---

## 🔄 How It Works

### OAuth 2.0 Flow

1. **First time:**
   - Script requests Gmail & Sheets scopes
   - Opens browser for user consent
   - Saves `token.json` with access & refresh tokens

2. **Future runs:**
   - Loads existing token from `token.json`
   - Auto-refreshes if expired
   - No user interaction needed

### Email Processing Pipeline

```
Fetch Unread → Filter by State → Parse Data → Check Duplicates → Append to Sheet → Mark as Read → Update State
```

---

## 🛡️ Duplicate Prevention

**Three-layer approach:**

### Layer 1: State File Check
- `state.json` stores processed Gmail message IDs
- Before processing, checks if email ID already exists
- Prevents reprocessing same emails

### Layer 2: Sheet Content Check
- Reads existing rows from Google Sheet
- Creates unique identifier: `(From + Subject + Date)`
- Compares new emails against existing entries

### Layer 3: Timestamp-based Uniqueness
- Even identical emails have different timestamps
- Date field includes seconds precision
- Ensures truly duplicate emails are still logged

**Why this approach?**
- If `state.json` is deleted, Layer 2 prevents duplicates
- If sheet is manually edited, Layer 1 prevents reprocessing
- Robust against various failure scenarios

---

## 💾 State Management

### State File Structure

`state.json`:
```json
{
  "processed_message_ids": [
    "18d4f2a1b2c3d4e5",
    "18d4f2a1b2c3d4e6"
  ],
  "last_run": "2026-01-14 20:31:48"
}
```

**Why JSON file?**
- ✅ Simple, human-readable
- ✅ No database required
- ✅ Easy to debug/inspect
- ✅ Portable across systems
- ✅ Version control friendly (though gitignored)

**Alternative considered:** SQLite database
- ❌ Overkill for this use case
- ❌ Adds complexity
- ❌ Harder to inspect/debug

---

## 🧩 Challenges & Solutions

### Challenge 1: Insufficient Authentication Scopes

**Problem:** Initial token was created with only Gmail scope, causing `ACCESS_TOKEN_SCOPE_INSUFFICIENT` error when accessing Sheets.

**Solution:** 
- Modified both `gmail_service.py` and `sheets_service.py` to request combined scopes (`GMAIL_SCOPES + SHEETS_SCOPES`) during authentication
- Ensured single OAuth flow requests all required permissions upfront
- Delete old token and re-authenticate when scopes change

**Code:**
```python
ALL_SCOPES = config.GMAIL_SCOPES + config.SHEETS_SCOPES
flow = InstalledAppFlow.from_client_secrets_file(
    config.CREDENTIALS_FILE, 
    ALL_SCOPES
)
```

### Challenge 2: HTML Email Parsing

**Problem:** Many emails contain HTML content that's hard to read when logged directly.

**Solution:**
- Implemented HTML-to-text conversion using BeautifulSoup
- Extracts plain text while preserving readability
- Handles both plain text and HTML emails seamlessly

**Code snippet:**
```python
def html_to_text(html_content):
    soup = BeautifulSoup(html_content, 'lxml')
    for script in soup(["script", "style"]):
        script.decompose()
    text = soup.get_text()
    # Clean and format text
    return text.strip()
```

### Challenge 3: Multi-part Email Structure

**Problem:** Emails can have complex nested structures (plain text, HTML, attachments).

**Solution:**
- Recursive function to traverse email parts
- Prioritizes plain text over HTML
- Gracefully handles missing or malformed parts

### Challenge 4: Rate Limiting

**Problem:** Processing 1400+ emails could hit API rate limits.

**Solution:**
- Batch operations where possible
- Limit to 100 emails per run (`MAX_RESULTS`)
- User can run multiple times to process all emails
- Progress indicators keep user informed

---

## ⚠️ Limitations

1. **Rate Limits:**
   - Gmail API: 1 billion quota units/day (more than enough for typical use)
   - Sheets API: 100 requests/100 seconds per user
   - Current implementation: Processes max 100 emails per run

2. **Cell Size Limit:**
   - Google Sheets cell limit: 50,000 characters
   - Very long emails are truncated with `[truncated]` marker

3. **No Attachment Handling:**
   - Only processes email body text
   - Attachments are ignored
   - Could be added in future versions

4. **No Real-time Sync:**
   - Script must be run manually
   - Not a continuously running service
   - Could be scheduled with cron/Task Scheduler

5. **OAuth Token Expiry:**
   - Token auto-refreshes, but if offline for 6+ months, re-auth needed
   - Not an issue for regular use

6. **Single Account:**
   - Processes one Gmail account at a time
   - Would need separate credentials for multiple accounts

---

## 📁 Project Structure

```
gmail-to-sheets/
│
├── src/
│   ├── gmail_service.py      # Gmail API authentication & operations
│   ├── sheets_service.py     # Google Sheets API operations
│   ├── email_parser.py       # Email parsing & HTML conversion
│   └── main.py               # Main orchestration script
│
├── credentials/
│   ├── credentials.json      # OAuth client secrets (DO NOT COMMIT)
│   └── token.json            # Access tokens (auto-generated, DO NOT COMMIT)
│
├── proof/
│   └── (screenshots & video) # Proof of execution
│
├── .gitignore                # Security: blocks sensitive files
├── requirements.txt          # Python dependencies
├── config.py                 # Configuration settings
├── state.json                # State persistence (auto-generated)
└── README.md                 # This file
```

---

## 🔐 Security Notes

**NEVER commit these files:**
- `credentials/credentials.json` - Contains OAuth client secrets
- `credentials/token.json` - Contains access tokens
- `state.json` - May contain sensitive email IDs

These are protected by `.gitignore`.

---

## 📊 Example Output

```
============================================================
Gmail to Google Sheets Automation
============================================================
Started at: 2026-01-14 20:27:59

[Step 1] Authenticating with Google APIs...
Gmail service initialized successfully
Sheets service initialized successfully

[Step 2] Initializing Google Sheet...
Headers already exist

[Step 3] Checking for existing emails in sheet...
Found 100 existing email(s) in sheet

[Step 4] Fetching unread emails from Gmail...
Found 100 unread email(s)
Successfully fetched 100 email(s)

[Step 5] Filtering new emails...
Found 100 new email(s) to process

[Step 6] Parsing email data...
Successfully parsed 100 unique email(s)

[Step 7] Adding emails to Google Sheet...
Successfully added 100 row(s) to sheet

[Step 8] Marking emails as read...
Marked 100/100 email(s) as read

[Step 9] Updating state...
State saved: 100 total processed email(s)

============================================================
SUMMARY
============================================================
Total unread emails found: 100
New emails processed: 100
Rows added to sheet: 100
Total processed (all time): 100
Completed at: 2026-01-14 20:31:48
============================================================
```

---

## 👤 Author

Md. Omar Khan - mdomarkhan314@gmail.com

---

🙏 Acknowledgments

- Google APIs Documentation
- BeautifulSoup for HTML parsing
- Python community for excellent libraries
