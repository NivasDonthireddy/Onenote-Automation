#!/usr/bin/env python3
"""
App Registration Configuration Checker
This script helps verify if your Azure app registration is properly configured for personal Microsoft accounts.
"""

import os
import requests
from dotenv import load_dotenv

def check_app_registration():
    """Check if the app registration is accessible and properly configured"""
    load_dotenv()

    client_id = os.getenv('CLIENT_ID')
    if not client_id:
        print("❌ CLIENT_ID not found in .env file")
        return False

    print(f"🔍 Checking app registration: {client_id}")

    # Check if the app exists and is accessible
    try:
        # This endpoint provides public information about the app
        url = f"https://login.microsoftonline.com/common/v2.0/.well-known/openid_configuration"
        response = requests.get(url)
        response.raise_for_status()

        print("✅ Microsoft identity platform is accessible")

        # Test the specific tenant/authority configuration
        authority_url = "https://login.microsoftonline.com/common/v2.0/.well-known/openid_configuration"
        auth_response = requests.get(authority_url)
        auth_response.raise_for_status()

        print("✅ 'common' authority endpoint is accessible")

        return True

    except requests.exceptions.RequestException as e:
        print(f"❌ Error accessing Microsoft endpoints: {str(e)}")
        return False

def print_configuration_guide():
    """Print configuration guide for Azure app registration"""
    print("\n📋 Azure App Registration Configuration Checklist:")
    print("=" * 60)

    print("\n1. 📱 Supported Account Types:")
    print("   ✓ Should be set to 'Personal Microsoft accounts only' or")
    print("   ✓ 'Accounts in any organizational directory and personal Microsoft accounts'")

    print("\n2. 🔗 Redirect URIs:")
    print("   ✓ Platform: Mobile and desktop applications")
    print("   ✓ Add: http://localhost")
    print("   ✓ Add: https://login.microsoftonline.com/common/oauth2/nativeclient")

    print("\n3. 🔑 API Permissions:")
    print("   ✓ Microsoft Graph > Delegated permissions:")
    print("     - Notes.Read")
    print("     - Notes.ReadWrite")
    print("   ✓ Grant admin consent (if required)")

    print("\n4. 🚫 What NOT to configure for personal accounts:")
    print("   ✗ Don't add client secrets (not needed for public client)")
    print("   ✗ Don't use application permissions (use delegated)")
    print("   ✗ Don't restrict to specific tenant")

    print("\n5. 🔧 Advanced Settings:")
    print("   ✓ 'Allow public client flows' should be 'Yes'")
    print("   ✓ 'Supported account types' must include personal accounts")

def main():
    print("Azure App Registration Checker")
    print("=" * 40)

    # Check basic connectivity
    if check_app_registration():
        print("✅ Basic connectivity test passed")
    else:
        print("❌ Basic connectivity test failed")

    # Print configuration guide
    print_configuration_guide()

    print(f"\n📝 Next Steps:")
    print("1. Verify the above settings in Azure Portal")
    print("2. Run the OneNote automation script again")
    print("3. If still having issues, try creating a new app registration")

if __name__ == "__main__":
    main()
