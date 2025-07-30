"""
STEP-BY-STEP AZURE APP REGISTRATION FIX
======================================

URGENT: You need to configure redirect URIs in Azure Portal

1. Go to https://portal.azure.com
2. Navigate to "Azure Active Directory" → "App registrations"
3. Find your app: Client ID 93d09ad1-eee7-41cd-b10d-b665e0963861
4. Click on your app registration
5. In the left menu, click "Authentication"

6. CRITICAL STEP - Add Redirect URIs:
   - Click "Add a platform"
   - Select "Mobile and desktop applications"
   - Add these EXACT redirect URIs:
     ✓ http://localhost
     ✓ https://login.microsoftonline.com/common/oauth2/nativeclient
     ✓ msal93d09ad1-eee7-41cd-b10d-b665e0963861://auth

7. IMPORTANT SETTINGS:
   - Supported account types: "Personal Microsoft accounts only" OR
     "Accounts in any organizational directory and personal Microsoft accounts"
   - Allow public client flows: YES
   - Enable the following mobile and desktop flows: YES

8. API Permissions:
   - Microsoft Graph → Delegated permissions:
     ✓ Notes.Read
     ✓ Notes.ReadWrite
   - Click "Grant admin consent for [your directory]"

9. Click "Save" at the top

WAIT TIME: After saving, wait 5-10 minutes for changes to propagate.

TEST: Run python create_pages.py again after making these changes.
"""

print(__doc__)
