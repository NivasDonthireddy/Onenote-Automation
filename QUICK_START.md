# OneNote Automation - Quick Start

## Quick Setup Guide

1. **Install Python packages:**
   ```
   pip install -r requirements.txt
   ```

2. **Setup your Azure App:**
   - Go to https://portal.azure.com
   - Navigate to Azure Active Directory → App Registrations
   - Create a new app with redirect URI: http://localhost
   - Copy your Client ID
   - Add API permissions: Notes.ReadWrite and Notes.Read

3. **Configure your environment:**
   - Copy `.env.example` to `.env`
   - Paste your Client ID in the `.env` file
   - Optionally set default notebook, section, and page

4. **Run the program:**
   ```
   python onenote_automation.py
   ```

5. **First run:**
   - A browser will open for authentication
   - Sign in with your Microsoft account
   - Your authentication will be cached for future use

That's it! You're ready to automate OneNote.

## Features You Can Use:

- Create pages with content
- Insert images from clipboard
- Set up hotkeys for quick page creation (Ctrl+Shift+N)
- Navigate notebooks, sections, and pages
- **🎓 Import entire Udemy courses** - Automatically create pages from course structure files

### Udemy Course Import

If you're using the companion `Udemy-Titles-Fetcher` project:

1. Run `python onenote_automation.py`
2. Select option **5** (Create pages from Udemy course output)
3. Choose your course file
4. Select sections to import
5. Pick your OneNote destination
6. Watch as all pages are created automatically!

For detailed documentation, see README.md

