
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

    def get_supported_image_formats(self):
        """Get list of supported image formats"""
        return ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.svg']

    def validate_image_file(self, image_path):
        """Validate if the image file is supported and exists"""
        if not os.path.exists(image_path):
            print(f"❌ Image file not found: {image_path}")
            return False

        _, ext = os.path.splitext(image_path.lower())
        if ext not in self.get_supported_image_formats():
            print(f"❌ Unsupported image format: {ext}")
            print(f"Supported formats: {', '.join(self.get_supported_image_formats())}")
            return False

        # Check file size (OneNote has a 100MB limit per attachment)
        file_size = os.path.getsize(image_path)
        max_size = 100 * 1024 * 1024  # 100MB in bytes
        if file_size > max_size:
            print(f"❌ Image file too large: {file_size / (1024*1024):.2f}MB (max: 100MB)")
            return False

        print(f"✅ Image file validated: {image_path} ({file_size / 1024:.2f}KB)")
        return True

    def _get_current_datetime(self):
        """Get current datetime in ISO format"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
import os
import requests
import json
import webbrowser
import urllib.parse
import base64
import mimetypes
import io
import tempfile
from msal import PublicClientApplication
from dotenv import load_dotenv

try:
    from PIL import ImageGrab, Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

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
    <h1>{escaped_title}</h1>
    {f'<p>{page_content}</p>' if page_content else ''}
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

    def create_page_with_image(self, section_id, page_title, image_path=None, image_url=None, page_content=""):
        """Create a new page with an embedded image"""
        try:
            url = f"{self.graph_url}/me/onenote/sections/{section_id}/pages"

            # Escape HTML characters in title to preserve exact formatting
            import html
            escaped_title = html.escape(page_title)

            # Create HTML content with image
            if image_path and os.path.exists(image_path):
                # Local image file
                html_content = self._create_html_with_local_image(escaped_title, image_path, page_content)
                return self._create_page_multipart(url, html_content, image_path)

            elif image_url:
                # Remote image URL
                html_content = self._create_html_with_remote_image(escaped_title, image_url, page_content)

                headers = self.get_headers()
                headers['Content-Type'] = 'text/html'

                response = requests.post(url, headers=headers, data=html_content.encode('utf-8'))
                response.raise_for_status()

                page_data = response.json()
                print(f"✅ Page '{page_title}' with image created successfully!")
                return page_data
            else:
                print("❌ No valid image path or URL provided")
                return None

        except requests.exceptions.RequestException as e:
            print(f"❌ Error creating page with image '{page_title}': {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response text: {e.response.text}")
            return None

    def create_page_with_clipboard_image(self, section_id, page_title, page_content=""):
        """Create a new page with an image from clipboard"""
        if not PIL_AVAILABLE:
            print("❌ PIL (Pillow) library not installed. Please install it with: pip install Pillow")
            return None

        try:
            # Get image from clipboard
            clipboard_image = ImageGrab.grabclipboard()

            if clipboard_image is None:
                print("❌ No image found in clipboard. Please copy an image first.")
                return None

            if not isinstance(clipboard_image, Image.Image):
                print("❌ Clipboard content is not an image.")
                return None

            print("✅ Image found in clipboard!")
            print(f"   📏 Size: {clipboard_image.size[0]}x{clipboard_image.size[1]} pixels")
            print(f"   🎨 Mode: {clipboard_image.mode}")

            # Save clipboard image to temporary file
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            temp_path = temp_file.name
            temp_file.close()

            # Convert to RGB if necessary (for PNG compatibility)
            if clipboard_image.mode in ('RGBA', 'LA'):
                # Create white background for transparent images
                background = Image.new('RGB', clipboard_image.size, (255, 255, 255))
                if clipboard_image.mode == 'RGBA':
                    background.paste(clipboard_image, mask=clipboard_image.split()[-1])
                else:
                    background.paste(clipboard_image)
                clipboard_image = background
            elif clipboard_image.mode != 'RGB':
                clipboard_image = clipboard_image.convert('RGB')

            # Save as PNG
            clipboard_image.save(temp_path, 'PNG', optimize=True)

            print(f"💾 Saved clipboard image to temporary file: {temp_path}")

            # Create page with the temporary image
            result = self.create_page_with_image(
                section_id=section_id,
                page_title=page_title,
                image_path=temp_path,
                page_content=page_content
            )

            # Clean up temporary file
            try:
                os.unlink(temp_path)
                print("🗑️ Cleaned up temporary file")
            except:
                pass  # Ignore cleanup errors

            return result

        except Exception as e:
            print(f"❌ Error creating page with clipboard image: {str(e)}")
            return None

    def check_clipboard_for_image(self):
        """Check if clipboard contains an image"""
        if not PIL_AVAILABLE:
            return False, "PIL (Pillow) library not installed"

        try:
            clipboard_image = ImageGrab.grabclipboard()
            if clipboard_image is None:
                return False, "No content in clipboard"

            if not isinstance(clipboard_image, Image.Image):
                return False, "Clipboard content is not an image"

            return True, f"Image found: {clipboard_image.size[0]}x{clipboard_image.size[1]} pixels, {clipboard_image.mode} mode"

        except Exception as e:
            return False, f"Error checking clipboard: {str(e)}"

    def _create_html_with_local_image(self, title, image_path, content=""):
        """Create HTML content with local image reference"""
        filename = os.path.basename(image_path)

        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{self._get_current_datetime()}" />
</head>
<body>
    <h1>{title}</h1>
    {f'<p>{content}</p>' if content else ''}
    <img src="name:{filename}" alt="{title}" style="max-width: 100%; height: auto;" />
</body>
</html>"""
        return html_content

    def _create_html_with_remote_image(self, title, image_url, content=""):
        """Create HTML content with remote image URL"""
        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{self._get_current_datetime()}" />
</head>
<body>
    <h1>{title}</h1>
    {f'<p>{content}</p>' if content else ''}
    <img src="{image_url}" alt="{title}" style="max-width: 100%; height: auto;" />
</body>
</html>"""
        return html_content

    def _create_page_multipart(self, url, html_content, image_path):
        """Create a page with multipart request for local image"""
        import uuid

        # Generate boundary for multipart request
        boundary = f"Part_{uuid.uuid4().hex}"

        # Get image content type
        content_type, _ = mimetypes.guess_type(image_path)
        if not content_type:
            content_type = 'application/octet-stream'

        # Read image file
        with open(image_path, 'rb') as image_file:
            image_data = image_file.read()

        filename = os.path.basename(image_path)

        # Create multipart body
        multipart_body = f"""--{boundary}\r
Content-Disposition: form-data; name="Presentation"\r
Content-Type: text/html\r
\r
{html_content}\r
--{boundary}\r
Content-Disposition: form-data; name="{filename}"\r
Content-Type: {content_type}\r
\r
""".encode('utf-8')

        multipart_body += image_data
        multipart_body += f"\r\n--{boundary}--\r\n".encode('utf-8')

        # Set headers for multipart request
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': f'multipart/form-data; boundary={boundary}'
        }

        response = requests.post(url, headers=headers, data=multipart_body)
        response.raise_for_status()

        page_data = response.json()
        print(f"✅ Page with local image created successfully!")
        print(f"   📄 Page ID: {page_data.get('id')}")

        # Try to get the web URL
        web_url = page_data.get('links', {}).get('oneNoteWebUrl', {}).get('href')
        if web_url:
            print(f"   🌐 Page URL: {web_url}")

        return page_data
