import os
import requests
import json
import webbrowser
import urllib.parse
from msal import PublicClientApplication
from dotenv import load_dotenv

class OneNoteAutomation:
    def __init__(self):
        load_dotenv()
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANT_ID', 'common')
        self.account_type = os.getenv('ACCOUNT_TYPE', 'personal')
        self.user_email = os.getenv('USER_EMAIL')

        if not self.client_id:
            raise ValueError("Missing CLIENT_ID in environment variables. Please check your .env file.")

        # For personal accounts, use common endpoint
        if self.account_type == 'personal':
            self.authority = "https://login.microsoftonline.com/common"
            self.scope = [
                "https://graph.microsoft.com/Notes.ReadWrite",
                "https://graph.microsoft.com/Notes.Read"
            ]
        else:
            self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            self.scope = ["https://graph.microsoft.com/.default"]

        self.graph_url = "https://graph.microsoft.com/v1.0"

        # Use PublicClientApplication for personal accounts
        self.app = PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority
        )

        self.access_token = None
        self.account = None

    def authenticate(self):
        """Authenticate using interactive browser flow for personal accounts"""
        try:
            # Try to get token silently first (if user has authenticated before)
            accounts = self.app.get_accounts()
            if accounts:
                print("Found existing account, attempting silent authentication...")
                result = self.app.acquire_token_silent(self.scope, account=accounts[0])
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    self.account = accounts[0]
                    print("✅ Silent authentication successful!")
                    return True

            # If silent auth fails, use interactive browser flow
            print("Starting interactive authentication...")
            print("This will open a browser window for authentication.")

            try:
                # Try interactive browser flow first (more reliable)
                result = self.app.acquire_token_interactive(
                    scopes=self.scope,
                    prompt="select_account"
                    # Remove explicit redirect_uri - MSAL handles this automatically
                )

                if "access_token" in result:
                    self.access_token = result["access_token"]
                    print("✅ Authentication successful!")
                    return True
                else:
                    print(f"❌ Interactive authentication failed: {result.get('error_description', 'Unknown error')}")

            except Exception as interactive_error:
                print(f"Interactive flow failed: {str(interactive_error)}")
                print("Falling back to device code flow...")

                # Fallback to device code flow
                flow = self.app.initiate_device_flow(scopes=self.scope)

                if "user_code" not in flow:
                    print(f"Device flow error: {flow.get('error', 'Unknown error')}")
                    print(f"Error description: {flow.get('error_description', 'No description')}")
                    raise ValueError("Failed to create device flow")

                print(f"\n🔐 Please visit: {flow['verification_uri']}")
                print(f"📱 Enter code: {flow['user_code']}")
                print("\nOpening browser automatically...")

                # Open browser automatically
                webbrowser.open(flow['verification_uri'])

                input("\nPress Enter after completing authentication in the browser...")

                # Complete the flow
                result = self.app.acquire_token_by_device_flow(flow)

                if "access_token" in result:
                    self.access_token = result["access_token"]
                    print("✅ Authentication successful!")
                    return True
                else:
                    print(f"❌ Device flow authentication failed: {result.get('error_description', 'Unknown error')}")
                    return False

        except Exception as e:
            print(f"❌ Authentication error: {str(e)}")
            print("\n🔧 Troubleshooting tips:")
            print("1. Make sure your app registration supports 'Personal Microsoft accounts'")
            print("2. Check that your CLIENT_ID is correct")
            print("3. Verify the app has proper redirect URIs configured")
            print("4. Ensure the app has Notes.ReadWrite permissions")
            print("5. Add 'http://localhost' as a redirect URI in Azure Portal")
            return False

    def get_headers(self):
        """Get headers for API requests"""
        if not self.access_token:
            raise ValueError("Not authenticated. Call authenticate() first.")

        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

    def get_notebooks(self):
        """Get all notebooks for the authenticated user"""
        try:
            # For personal accounts, use /me endpoint
            url = f"{self.graph_url}/me/onenote/notebooks"

            response = requests.get(url, headers=self.get_headers())
            response.raise_for_status()

            notebooks = response.json().get('value', [])
            print(f"📚 Found {len(notebooks)} notebooks:")
            for nb in notebooks:
                print(f"  📖 {nb['displayName']} (ID: {nb['id']})")

            return notebooks
        except requests.exceptions.RequestException as e:
            print(f"❌ Error getting notebooks: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response text: {e.response.text}")
            return []

    def get_sections(self, notebook_id):
        """Get all sections in a notebook"""
        try:
            url = f"{self.graph_url}/me/onenote/notebooks/{notebook_id}/sections"

            response = requests.get(url, headers=self.get_headers())
            response.raise_for_status()

            sections = response.json().get('value', [])
            print(f"📂 Found {len(sections)} sections:")
            for section in sections:
                print(f"  📄 {section['displayName']} (ID: {section['id']})")

            return sections
        except requests.exceptions.RequestException as e:
            print(f"❌ Error getting sections: {str(e)}")
            return []

    def create_page(self, section_id, page_title, page_content=""):
        """Create a new page in a specific section"""
        try:
            url = f"{self.graph_url}/me/onenote/sections/{section_id}/pages"

            # Escape HTML characters in title to preserve exact formatting
            import html
            escaped_title = html.escape(page_title)

            # Create HTML content for the page with preserved formatting
            html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{escaped_title}</title>
    <meta name="created" content="{self._get_current_datetime()}" />
</head>
<body>
    <div>
        <p style="font-size: 18pt; font-weight: bold; margin: 0;">{escaped_title}</p>
    </div>
    {page_content if page_content else ''}
</body>
</html>"""

            headers = self.get_headers()
            headers['Content-Type'] = 'text/html'

            response = requests.post(url, headers=headers, data=html_content.encode('utf-8'))
            response.raise_for_status()

            page_data = response.json()
            print(f"✅ Page '{page_title}' created successfully!")
            print(f"   📄 Page ID: {page_data.get('id')}")

            # Try to get the web URL
            web_url = page_data.get('links', {}).get('oneNoteWebUrl', {}).get('href')
            if web_url:
                print(f"   🌐 Page URL: {web_url}")

            return page_data
        except requests.exceptions.RequestException as e:
            print(f"❌ Error creating page '{page_title}': {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response text: {e.response.text}")
            return None

    def create_multiple_pages(self, section_id, page_titles):
        """Create multiple pages with given titles"""
        created_pages = []

        print(f"\n🚀 Creating {len(page_titles)} pages...")
        for i, title in enumerate(page_titles, 1):
            print(f"\n📝 Creating page {i}/{len(page_titles)}: {title}")
            page = self.create_page(section_id, title)
            if page:
                created_pages.append(page)

        print(f"\n✅ Summary: {len(created_pages)} out of {len(page_titles)} pages created successfully.")
        return created_pages

    def find_notebook_by_name(self, notebook_name):
        """Find a notebook by name"""
        notebooks = self.get_notebooks()
        for notebook in notebooks:
            if notebook['displayName'].lower() == notebook_name.lower():
                return notebook
        return None

    def find_section_by_name(self, notebook_id, section_name):
        """Find a section by name within a notebook"""
        sections = self.get_sections(notebook_id)
        for section in sections:
            if section['displayName'].lower() == section_name.lower():
                return section
        return None

    def _get_current_datetime(self):
        """Get current datetime in ISO format"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def main():
    """Example usage"""
    onenote = OneNoteAutomation()

    # Authenticate
    print("🔐 Starting authentication for personal Microsoft account...")
    if not onenote.authenticate():
        print("❌ Failed to authenticate. Please check your app registration and permissions.")
        return

    # Example: Create pages in a specific notebook and section
    notebook_name = "My Notebook"  # Replace with your notebook name
    section_name = "General"       # Replace with your section name

    page_titles = [
        "Meeting Notes - July 30, 2025",
        "Project Planning Ideas",
        "Daily Task List",
        "Brainstorming Session"
    ]

    # Find notebook
    print(f"\n🔍 Looking for notebook: {notebook_name}")
    notebook = onenote.find_notebook_by_name(notebook_name)

    if not notebook:
        print(f"❌ Notebook '{notebook_name}' not found.")
        print("\nℹ️  Available notebooks:")
        onenote.get_notebooks()
        return

    # Find section
    print(f"\n🔍 Looking for section: {section_name}")
    section = onenote.find_section_by_name(notebook['id'], section_name)

    if not section:
        print(f"❌ Section '{section_name}' not found.")
        print(f"\nℹ️  Available sections in '{notebook_name}':")
        onenote.get_sections(notebook['id'])
        return

    # Create pages
    print(f"\n📋 Creating pages in notebook '{notebook_name}' > section '{section_name}'")
    onenote.create_multiple_pages(section['id'], page_titles)

if __name__ == "__main__":
    main()
