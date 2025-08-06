import os
import requests
import json
import webbrowser
import urllib.parse
import base64
import mimetypes
import io
import tempfile
from msal import PublicClientApplication, SerializableTokenCache
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

        # Default settings from .env
        self.default_notebook = os.getenv('DEFAULT_NOTEBOOK', '').strip()
        self.default_section = os.getenv('DEFAULT_SECTION', '').strip()
        self.default_page = os.getenv('DEFAULT_PAGE', '').strip()
        self.token_cache_file = os.getenv('TOKEN_CACHE_FILE', '.token_cache.json')

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

        # Initialize token cache
        self.cache = SerializableTokenCache()
        if os.path.exists(self.token_cache_file):
            with open(self.token_cache_file, 'r') as f:
                self.cache.deserialize(f.read())

        # Use PublicClientApplication with token cache
        self.app = PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache
        )

        self.access_token = None
        self.account = None

        # Cache for default objects
        self._cached_notebook = None
        self._cached_section = None
        self._cached_page = None

    def _save_token_cache(self):
        """Save token cache to file for persistence"""
        if self.cache.has_state_changed:
            with open(self.token_cache_file, 'w') as f:
                f.write(self.cache.serialize())

    def authenticate(self, force_reauth=False):
        """Authenticate with persistent token caching"""
        try:
            # Try to get token silently first (if user has authenticated before)
            accounts = self.app.get_accounts()
            if accounts and not force_reauth:
                print("🔄 Using cached authentication...")
                result = self.app.acquire_token_silent(self.scope, account=accounts[0])
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    self.account = accounts[0]
                    self._save_token_cache()
                    print("✅ Authentication successful!")
                    return True

            # If silent auth fails, use interactive browser flow
            print("🔐 Starting authentication...")
            print("This will open a browser window for authentication.")

            try:
                result = self.app.acquire_token_interactive(
                    scopes=self.scope,
                    prompt="select_account"
                )

                if "access_token" in result:
                    self.access_token = result["access_token"]
                    self.account = result.get("account")
                    self._save_token_cache()
                    print("✅ Authentication successful!")
                    return True
                else:
                    print(f"❌ Interactive authentication failed: {result.get('error_description', 'Unknown error')}")

            except Exception as interactive_error:
                print(f"Interactive flow failed: {str(interactive_error)}")
                return False

        except Exception as e:
            print(f"❌ Authentication error: {str(e)}")
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

    def get_or_create_default_page(self):
        """Get or create the default page based on .env settings"""
        try:
            # Get default notebook
            if not self._cached_notebook:
                if self.default_notebook:
                    self._cached_notebook = self.find_notebook_by_name(self.default_notebook)
                    if not self._cached_notebook:
                        print(f"❌ Default notebook '{self.default_notebook}' not found.")
                        return None
                else:
                    print("📚 No default notebook specified. Please set DEFAULT_NOTEBOOK in .env file.")
                    return None

            # Get default section
            if not self._cached_section:
                if self.default_section:
                    self._cached_section = self.find_section_by_name(self._cached_notebook['id'], self.default_section)
                    if not self._cached_section:
                        print(f"❌ Default section '{self.default_section}' not found.")
                        return None
                else:
                    print("📂 No default section specified. Please set DEFAULT_SECTION in .env file.")
                    return None

            # Get or create default page
            if not self._cached_page:
                if self.default_page:
                    self._cached_page = self.find_page_by_name(self._cached_section['id'], self.default_page)
                    if not self._cached_page:
                        print(f"📝 Creating new page '{self.default_page}'...")
                        self._cached_page = self.create_page(self._cached_section['id'], self.default_page, "")
                    else:
                        print(f"📄 Using existing page '{self.default_page}'")
                else:
                    print("📝 No default page specified. Please set DEFAULT_PAGE in .env file.")
                    return None

            return self._cached_page

        except Exception as e:
            print(f"❌ Error getting/creating default page: {str(e)}")
            return None

    def find_page_by_name(self, section_id, page_name):
        """Find a page by name within a section"""
        try:
            url = f"{self.graph_url}/me/onenote/sections/{section_id}/pages"
            response = requests.get(url, headers=self.get_headers())
            response.raise_for_status()

            pages = response.json().get('value', [])
            for page in pages:
                if page['title'].lower() == page_name.lower():
                    return page
            return None
        except requests.exceptions.RequestException as e:
            print(f"❌ Error finding page: {str(e)}")
            return None

    def add_image_to_page(self, page_id, image_path=None, image_url=None):
        """Add an image to an existing OneNote page"""
        try:
            url = f"{self.graph_url}/me/onenote/pages/{page_id}/content"

            if image_path and os.path.exists(image_path):
                return self._add_local_image_to_page(url, image_path)
            elif image_url:
                return self._add_remote_image_to_page(url, image_url)
            else:
                print("❌ No valid image path or URL provided")
                return False

        except requests.exceptions.RequestException as e:
            print(f"❌ Error adding image to page: {str(e)}")
            return False

    def _add_local_image_to_page(self, url, image_path):
        """Add a local image to an existing page using PATCH request"""
        import uuid

        boundary = f"Part_{uuid.uuid4().hex}"
        content_type, _ = mimetypes.guess_type(image_path)
        if not content_type:
            content_type = 'application/octet-stream'

        # Optimize image size for faster upload
        optimized_path = self._optimize_image_for_upload(image_path)

        with open(optimized_path, 'rb') as image_file:
            image_data = image_file.read()

        filename = os.path.basename(image_path)

        # Create patch content with simple image and single line break
        patch_content = f"""[{{
    "target": "body",
    "action": "append",
    "content": "<img src=\\"name:{filename}\\" alt=\\"{filename}\\" style=\\"max-width: 100%; height: auto;\\" /><br/>"
}}]"""

        # Create multipart body
        multipart_body = f"""--{boundary}\r
Content-Disposition: form-data; name="Commands"\r
Content-Type: application/json\r
\r
{patch_content}\r
--{boundary}\r
Content-Disposition: form-data; name="{filename}"\r
Content-Type: {content_type}\r
\r
""".encode('utf-8')

        multipart_body += image_data
        multipart_body += f"\r\n--{boundary}--\r\n".encode('utf-8')

        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': f'multipart/form-data; boundary={boundary}'
        }

        response = requests.patch(url, headers=headers, data=multipart_body)
        response.raise_for_status()

        # Clean up optimized file if it's different from original
        if optimized_path != image_path:
            try:
                os.unlink(optimized_path)
            except:
                pass

        print(f"✅ Image '{filename}' added to page successfully!")
        return True

    def _add_remote_image_to_page(self, url, image_url):
        """Add a remote image to an existing page using PATCH request"""
        patch_content = [{
            "target": "body",
            "action": "append",
            "content": f'<img src="{image_url}" alt="Remote Image" style="max-width: 100%; height: auto;" /><br/>'
        }]

        headers = self.get_headers()
        response = requests.patch(url, headers=headers, json=patch_content)
        response.raise_for_status()

        print(f"✅ Remote image added to page successfully!")
        return True

    def quick_add_clipboard_image(self):
        """Quick add clipboard image to default page - main hotkey function"""
        if not PIL_AVAILABLE:
            print("❌ PIL (Pillow) library not installed. Please install it with: pip install Pillow")
            return False

        # Ensure authentication
        if not self.access_token:
            if not self.authenticate():
                print("❌ Authentication failed")
                return False

        # Get clipboard image
        try:
            clipboard_image = ImageGrab.grabclipboard()

            if clipboard_image is None:
                print("❌ No image found in clipboard")
                return False

            if not isinstance(clipboard_image, Image.Image):
                print("❌ Clipboard content is not an image")
                return False

            print(f"📋 Image found: {clipboard_image.size[0]}x{clipboard_image.size[1]} pixels")

            # Get or create default page
            page = self.get_or_create_default_page()
            if not page:
                print("❌ Could not get/create default page")
                return False

            # Save clipboard image to temporary file
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            temp_path = temp_file.name
            temp_file.close()

            # Convert image if necessary
            if clipboard_image.mode in ('RGBA', 'LA'):
                background = Image.new('RGB', clipboard_image.size, (255, 255, 255))
                if clipboard_image.mode == 'RGBA':
                    background.paste(clipboard_image, mask=clipboard_image.split()[-1])
                else:
                    background.paste(clipboard_image)
                clipboard_image = background
            elif clipboard_image.mode != 'RGB':
                clipboard_image = clipboard_image.convert('RGB')

            clipboard_image.save(temp_path, 'PNG', optimize=True)

            # Add image to page
            result = self.add_image_to_page(page['id'], image_path=temp_path)

            # Clean up
            try:
                os.unlink(temp_path)
            except:
                pass

            if result:
                print(f"✅ Image added to page '{page['title']}' successfully!")
                return True
            else:
                print("❌ Failed to add image to page")
                return False

        except Exception as e:
            print(f"❌ Error: {str(e)}")
            return False

    def _optimize_image_for_upload(self, image_path):
        """Optimize image for OneNote web version compatibility with fixed 800px width"""
        if not PIL_AVAILABLE:
            return image_path

        try:
            with Image.open(image_path) as img:
                # Get original size
                original_size = os.path.getsize(image_path)

                # Fixed width for OneNote web version consistency
                target_width = 800  # Match OneNote web version default width

                # Always resize to target width maintaining aspect ratio
                if img.width != target_width:
                    ratio = target_width / img.width
                    new_height = int(img.height * ratio)
                    new_size = (target_width, new_height)
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                    print(f"📏 Resized to OneNote web standard: {img.width}x{img.height} → {new_size[0]}x{new_size[1]}")

                # Convert to RGB if necessary
                if img.mode in ('RGBA', 'LA'):
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'RGBA':
                        background.paste(img, mask=img.split()[-1])
                    else:
                        background.paste(img)
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')

                # Create optimized temporary file
                temp_file = tempfile.NamedTemporaryFile(suffix='.jpg', delete=False)
                optimized_path = temp_file.name
                temp_file.close()

                # Use high quality for OneNote web compatibility
                quality = 90  # Higher quality for 800px standard
                img.save(optimized_path, 'JPEG', quality=quality, optimize=True)

                # Always use the resized version for consistency
                optimized_size = os.path.getsize(optimized_path)
                print(f"🗜️ Optimized for OneNote web: {original_size//1024}KB → {optimized_size//1024}KB")
                return optimized_path

        except Exception as e:
            print(f"⚠️ Image optimization failed, using original: {str(e)}")
            return image_path

