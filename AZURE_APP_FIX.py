"""
URGENT: Azure App Registration Fix for Personal Microsoft Accounts

The error "invalid_request: The provided value for the input parameter 'redirect_uri' is not valid"
means your app registration needs specific redirect URIs configured.

STEP-BY-STEP FIX:
================

1. Go to Azure Portal (https://portal.azure.com)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Find your app registration (Client ID: 93d09ad1-eee7-41cd-b10d-b665e0963861)
4. Click on "Authentication" in the left menu

5. CRITICAL - Add these Redirect URIs:
   Platform: "Mobile and desktop applications"

   Required URIs to add:
   ✓ http://localhost
   ✓ https://login.microsoftonline.com/common/oauth2/nativeclient
   ✓ msal93d09ad1-eee7-41cd-b10d-b665e0963861://auth (replace with your actual client ID)

6. CRITICAL - Supported account types:
   ✓ Must be set to "Personal Microsoft accounts only" OR
   ✓ "Accounts in any organizational directory and personal Microsoft accounts"

7. Advanced settings (at bottom of Authentication page):
   ✓ "Allow public client flows" = YES
   ✓ "Enable the following mobile and desktop flows" = YES

8. API Permissions:
   ✓ Microsoft Graph > Delegated permissions:
     - Notes.Read
     - Notes.ReadWrite
   ✓ Click "Grant admin consent" if available

9. Click "Save" at the top

AFTER MAKING THESE CHANGES:
==========================
- Wait 5-10 minutes for changes to propagate
- Run the script again: python create_pages.py

If you still get redirect URI errors, the most common issue is:
- Missing http://localhost redirect URI
- App not configured for personal accounts
- Public client flows not enabled

QUICK TEST:
==========
Run: python check_config.py
This will verify your configuration and provide additional guidance.
"""

print(__doc__)
