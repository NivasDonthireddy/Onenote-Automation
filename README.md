# OneNote Automation

A simple Python automation tool to interact with Microsoft OneNote using the Microsoft Graph API.

## Features

- 🔐 Persistent authentication with token caching
- 📓 Create and manage OneNote notebooks, sections, and pages
- 🖼️ Insert images from clipboard or files
- ⌨️ Hotkey support for quick page creation (Ctrl+Shift+N)
- 🎯 Default notebook/section/page configuration for faster workflows
- 🎓 **NEW:** Import and create pages from Udemy course structure files
- 🔄 **NEW:** Automatic retry mechanism for failed pages (handles auth expiration)

## Prerequisites

- Python 3.7 or higher
- Microsoft account (personal or organizational)
- Azure AD App Registration (Client ID)

## Setup

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App Registrations
2. Create a new registration:
   - Name: OneNote Automation
   - Supported account types: Personal Microsoft accounts only (for personal accounts)
   - Redirect URI: Public client/native → `http://localhost`
3. Copy the **Application (client) ID**
4. Go to API Permissions:
   - Add `Notes.ReadWrite` and `Notes.Read` permissions
   - Grant admin consent if needed

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure Environment

1. Copy `.env.example` to `.env`
2. Add your `CLIENT_ID` from Azure
3. Optionally set default notebook, section, and page names

```bash
copy .env.example .env
```

Edit `.env` with your settings:
```
CLIENT_ID=your-client-id-here
DEFAULT_NOTEBOOK=My Notebook
DEFAULT_SECTION=My Section
DEFAULT_PAGE=My Page
```

## Usage

### Basic Usage

```python
from onenote_automation import OneNoteAutomation

# Initialize
onenote = OneNoteAutomation()

# Authenticate (browser will open once, then cached)
onenote.authenticate()

# Create a page with content
onenote.create_page_with_title_and_content(
    title="My New Page",
    content="This is my content",
    notebook_name="Python Learnings",
    section_name="Django"
)

# Insert image from clipboard
onenote.insert_image_into_page()
```

### Hotkey Mode

Run the script with hotkey support to quickly create pages:

```python
from onenote_automation import OneNoteAutomation

onenote = OneNoteAutomation()
onenote.authenticate()
onenote.start_hotkey_listener()  # Press Ctrl+Shift+N to create page
```

### Interactive Menu

Simply run the script to get an interactive menu:

```bash
python onenote_automation.py
```

### Import from Udemy Course Outputs

If you have Udemy course structure files from the `Udemy-Titles-Fetcher` project:

1. Run the script and select option 5
2. Choose a course output file (e.g., "AI Engineer Core Track - Clean.txt")
3. Select which sections to import (or press Enter for all)
4. Choose the destination notebook and section in OneNote
5. Pick a naming strategy for pages
6. Confirm and watch as all pages are created automatically!

This feature automatically creates organized OneNote pages from Udemy course structures, perfect for note-taking during courses.

## Configuration

### Environment Variables (.env)

- `CLIENT_ID` (required): Azure App Registration Client ID
- `TENANT_ID`: Default is 'common' for personal accounts
- `ACCOUNT_TYPE`: 'personal' or 'organizational'
- `USER_EMAIL`: Optional, your Microsoft account email
- `DEFAULT_NOTEBOOK`: Default notebook name
- `DEFAULT_SECTION`: Default section name
- `DEFAULT_PAGE`: Default page name
- `TOKEN_CACHE_FILE`: Token cache file path (default: .token_cache.json)

## Authentication

The first time you run the script, it will open a browser for authentication. Your token is then cached in `.token_cache.json`, so subsequent runs won't require re-authentication unless the token expires.

## Troubleshooting

### Authentication Issues
- Make sure your Azure app has the correct redirect URI: `http://localhost`
- Verify API permissions are granted
- Try deleting `.token_cache.json` and re-authenticating

### Authentication Expires During Import
- **Solution**: Use the retry mechanism!
- After import shows failures, choose "y" to retry
- Choose "y" to re-authenticate
- Only failed pages will be retried (progress is saved)
- See `RETRY_MECHANISM_GUIDE.md` for details

### Image Insertion Issues
- Pillow library is required for clipboard image support
- Make sure you've copied an image to clipboard before inserting

## License

MIT License - Feel free to use and modify as needed.

## Author

Created for automating OneNote workflows via Microsoft Graph API.

