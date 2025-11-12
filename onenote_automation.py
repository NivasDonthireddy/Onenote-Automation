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

            # Create HTML content for the page - only add heading if there's content
            if page_content.strip():
                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{escaped_title}</title>
    <meta name="created" content="{self._get_current_datetime()}" />
</head>
<body>
    <h1>{escaped_title}</h1>
    <p>{page_content}</p>
</body>
</html>"""
            else:
                # For empty content, create minimal page without duplicate title in body
                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>{escaped_title}</title>
    <meta name="created" content="{self._get_current_datetime()}" />
</head>
<body>
    <p></p>
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

        try:
            response = requests.post(url, headers=headers, data=multipart_body)
            response.raise_for_status()

            page_data = response.json()
            print(f"✅ Page with image created successfully!")
            print(f"   📄 Page ID: {page_data.get('id')}")

            # Try to get the web URL
            web_url = page_data.get('links', {}).get('oneNoteWebUrl', {}).get('href')
            if web_url:
                print(f"   🌐 Page URL: {web_url}")

            return page_data

        except requests.exceptions.RequestException as e:
            print(f"❌ Error creating page with image: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response text: {e.response.text}")
            return None

    def _get_current_datetime(self):
        """Get current datetime in ISO format"""
        from datetime import datetime
        return datetime.now().isoformat()

    def find_notebook_by_name(self, notebook_name):
        """Find a notebook by name (case-insensitive)"""
        notebooks = self.get_notebooks()
        for notebook in notebooks:
            if notebook['displayName'].lower() == notebook_name.lower():
                return notebook
        return None

    def find_section_by_name(self, notebook_id, section_name):
        """Find a section by name in a specific notebook (case-insensitive)"""
        sections = self.get_sections(notebook_id)
        for section in sections:
            if section['displayName'].lower() == section_name.lower():
                return section
        return None

    def get_pages(self, section_id):
        """Get all pages in a section"""
        try:
            url = f"{self.graph_url}/me/onenote/sections/{section_id}/pages"
            response = requests.get(url, headers=self.get_headers())
            response.raise_for_status()

            pages = response.json().get('value', [])
            print(f"📄 Found {len(pages)} pages:")
            for page in pages:
                print(f"  📝 {page['title']} (ID: {page['id']})")

            return pages
        except requests.exceptions.RequestException as e:
            print(f"❌ Error getting pages: {str(e)}")
            return []

    def find_page_by_title(self, section_id, page_title):
        """Find a page by title in a specific section (case-insensitive)"""
        pages = self.get_pages(section_id)
        for page in pages:
            if page['title'].lower() == page_title.lower():
                return page
        return None

    def create_multiple_pages(self, section_id, page_titles, page_content=""):
        """Create multiple pages with given titles"""
        created_pages = []
        failed_pages = []

        print(f"📝 Creating {len(page_titles)} pages...")

        for i, title in enumerate(page_titles, 1):
            print(f"Creating page {i}/{len(page_titles)}: '{title}'")

            page_data = self.create_page(section_id, title, page_content)
            if page_data:
                created_pages.append(page_data)
                print(f"✅ Created: '{title}'")
            else:
                failed_pages.append(title)
                print(f"❌ Failed: '{title}'")

        print(f"\n📊 Summary:")
        print(f"✅ Successfully created: {len(created_pages)} pages")
        if failed_pages:
            print(f"❌ Failed to create: {len(failed_pages)} pages")
            for title in failed_pages:
                print(f"   - {title}")

        return created_pages, failed_pages

    def get_default_notebook(self):
        """Get or cache the default notebook"""
        if self._cached_notebook is None:
            if self.default_notebook:
                print(f"🔍 Looking for default notebook: '{self.default_notebook}'")
                self._cached_notebook = self.find_notebook_by_name(self.default_notebook)
                if self._cached_notebook:
                    print(f"✅ Found default notebook: '{self._cached_notebook['displayName']}'")
                else:
                    print(f"❌ Default notebook '{self.default_notebook}' not found")
            else:
                print("❌ No default notebook specified in .env file")

        return self._cached_notebook

    def get_default_section(self):
        """Get or cache the default section"""
        if self._cached_section is None:
            notebook = self.get_default_notebook()
            if notebook and self.default_section:
                print(f"🔍 Looking for default section: '{self.default_section}'")
                self._cached_section = self.find_section_by_name(notebook['id'], self.default_section)
                if self._cached_section:
                    print(f"✅ Found default section: '{self._cached_section['displayName']}'")
                else:
                    print(f"❌ Default section '{self.default_section}' not found")
            else:
                if not notebook:
                    print("❌ Cannot find default section without default notebook")
                else:
                    print("❌ No default section specified in .env file")

        return self._cached_section

    def quick_create_page(self, page_title, page_content=""):
        """Quickly create a page using default notebook and section"""
        section = self.get_default_section()
        if section:
            return self.create_page(section['id'], page_title, page_content)
        else:
            print("❌ Cannot create page: no default section available")
            return None

    def reset_cache(self):
        """Reset cached notebook, section, and page objects"""
        self._cached_notebook = None
        self._cached_section = None
        self._cached_page = None
        print("🔄 Cache reset")

    def list_all_structure(self):
        """List all notebooks, sections, and pages in a hierarchical view"""
        print("🏗️ OneNote Structure:")
        print("=" * 50)

        notebooks = self.get_notebooks()

        for notebook in notebooks:
            print(f"📚 {notebook['displayName']}")
            sections = self.get_sections(notebook['id'])

            for section in sections:
                print(f"  📂 {section['displayName']}")
                pages = self.get_pages(section['id'])

                for page in pages:
                    print(f"    📄 {page['title']}")

                if not pages:
                    print("    (no pages)")

            if not sections:
                print("  (no sections)")

        if not notebooks:
            print("(no notebooks)")

    def select_notebook_interactive(self):
        """Interactively select a notebook"""
        notebooks = self.get_notebooks()
        if not notebooks:
            print("❌ No notebooks found")
            return None

        print("\n📚 Available Notebooks:")
        for i, notebook in enumerate(notebooks, 1):
            print(f"  {i}. {notebook['displayName']}")

        while True:
            try:
                choice = input(f"\nSelect notebook (1-{len(notebooks)}) or press Enter for default: ").strip()

                if not choice:
                    # Try to use default
                    default_notebook = self.get_default_notebook()
                    if default_notebook:
                        print(f"✅ Using default notebook: '{default_notebook['displayName']}'")
                        return default_notebook
                    else:
                        print("❌ No default notebook available, please select one")
                        continue

                choice_num = int(choice)
                if 1 <= choice_num <= len(notebooks):
                    selected = notebooks[choice_num - 1]
                    print(f"✅ Selected notebook: '{selected['displayName']}'")
                    return selected
                else:
                    print(f"❌ Please enter a number between 1 and {len(notebooks)}")
            except ValueError:
                print("❌ Please enter a valid number")
            except KeyboardInterrupt:
                print("\n❌ Selection cancelled")
                return None

    def select_section_interactive(self, notebook_id):
        """Interactively select a section from a notebook"""
        sections = self.get_sections(notebook_id)
        if not sections:
            print("❌ No sections found in this notebook")
            return None

        print("\n📂 Available Sections:")
        for i, section in enumerate(sections, 1):
            print(f"  {i}. {section['displayName']}")

        while True:
            try:
                choice = input(f"\nSelect section (1-{len(sections)}) or press Enter for default: ").strip()

                if not choice:
                    # Try to use default
                    default_section = self.get_default_section()
                    if default_section:
                        # Check if the default section is in this notebook
                        for section in sections:
                            if section['id'] == default_section['id']:
                                print(f"✅ Using default section: '{default_section['displayName']}'")
                                return default_section
                        print("❌ Default section not found in selected notebook, please select one")
                        continue
                    else:
                        print("❌ No default section available, please select one")
                        continue

                choice_num = int(choice)
                if 1 <= choice_num <= len(sections):
                    selected = sections[choice_num - 1]
                    print(f"✅ Selected section: '{selected['displayName']}'")
                    return selected
                else:
                    print(f"❌ Please enter a number between 1 and {len(sections)}")
            except ValueError:
                print("❌ Please enter a valid number")
            except KeyboardInterrupt:
                print("\n❌ Selection cancelled")
                return None

    def interactive_create_page(self):
        """Interactive page creation with notebook/section selection"""
        print("🎯 Interactive Page Creation Mode")
        print("=" * 40)

        # Select notebook
        notebook = self.select_notebook_interactive()
        if not notebook:
            return None

        # Select section
        section = self.select_section_interactive(notebook['id'])
        if not section:
            return None

        # Get page details
        print("\n📝 Page Details:")
        page_title = input("Enter page title: ").strip()
        if not page_title:
            print("❌ Page title cannot be empty")
            return None

        page_content = input("Enter page content (optional): ").strip()

        # Check for clipboard image
        if PIL_AVAILABLE:
            has_image, image_info = self.check_clipboard_for_image()
            if has_image:
                use_clipboard = input(f"\n🖼️ {image_info}\nInclude clipboard image? (y/n): ").strip().lower()
                if use_clipboard in ['y', 'yes']:
                    return self.create_page_with_clipboard_image(section['id'], page_title, page_content)

        # Create regular page
        return self.create_page(section['id'], page_title, page_content)

    def interactive_create_multiple_pages(self):
        """Interactive creation of multiple pages with notebook/section selection"""
        print("🎯 Interactive Multiple Page Creation Mode")
        print("=" * 45)

        # Select notebook
        notebook = self.select_notebook_interactive()
        if not notebook:
            return None

        # Select section
        section = self.select_section_interactive(notebook['id'])
        if not section:
            return None

        # Get page titles
        print("\n📝 Multiple Page Creation:")
        print("Enter page titles (one per line). Press Enter twice when done, or type 'done' on a new line:")

        page_titles = []
        while True:
            title = input(f"Page {len(page_titles) + 1} title: ").strip()
            if not title or title.lower() == 'done':
                break
            page_titles.append(title)

        if not page_titles:
            print("❌ No page titles provided")
            return None

        # Get optional common content
        page_content = input(f"\nEnter common content for all {len(page_titles)} pages (optional): ").strip()

        # Confirm creation
        print(f"\n📋 About to create {len(page_titles)} pages:")
        for i, title in enumerate(page_titles, 1):
            print(f"  {i}. {title}")

        confirm = input(f"\nProceed with creating {len(page_titles)} pages? (y/n): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("❌ Operation cancelled")
            return None

        # Create pages
        return self.create_multiple_pages(section['id'], page_titles, page_content)

    def get_udemy_output_files(self):
        """Get list of clean output files from Udemy-Titles-Fetcher"""
        # Get parent directory of current file's directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        udemy_outputs_path = os.path.join(parent_dir, 'Udemy-Titles-Fetcher', 'outputs')

        if not os.path.exists(udemy_outputs_path):
            print(f"❌ Udemy outputs directory not found: {udemy_outputs_path}")
            return []

        # Get all *Clean.txt files
        clean_files = []
        try:
            for filename in os.listdir(udemy_outputs_path):
                if filename.endswith('- Clean.txt'):
                    full_path = os.path.join(udemy_outputs_path, filename)
                    clean_files.append({
                        'filename': filename,
                        'path': full_path,
                        'course_name': filename.replace(' - Clean.txt', '')
                    })
        except Exception as e:
            print(f"❌ Error reading Udemy outputs directory: {str(e)}")
            return []

        return clean_files

    def parse_udemy_output_file(self, file_path):
        """Parse a Udemy output file and extract course structure"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Extract course name from the header
            course_name = ""
            lines = content.split('\n')
            for line in lines:
                if line.strip().startswith('COURSE:'):
                    course_name = line.replace('COURSE:', '').strip()
                    break

            # Parse the structure - looking for numbered items
            sections = []
            current_section = None

            for line in lines:
                # Keep track of leading whitespace
                stripped = line.strip()
                if not stripped or stripped.startswith('#') or stripped.startswith('='):
                    continue

                # Check if line starts with a number pattern (after stripping)
                if stripped and stripped[0].isdigit():
                    # Check if line has leading whitespace (indented items are pages)
                    has_leading_space = line[0].isspace() if line else False

                    # Split by first period
                    parts = stripped.split('.', 1)
                    if len(parts) >= 2:
                        number_part = parts[0].strip()
                        rest = parts[1].strip()

                        # Main sections have no leading space and single digit (e.g., "1. Week 1")
                        if not has_leading_space and number_part.isdigit():
                            # This is a main section
                            current_section = {
                                'number': number_part,
                                'title': rest,
                                'pages': []
                            }
                            sections.append(current_section)

                        # Sub-items have leading space and dot in number (e.g., "   1.1. Page Title")
                        elif has_leading_space and current_section and '.' in rest:
                            # Extract page number and title (e.g., "1" from "1.1. Title")
                            # The rest is like "1. Title" so split again
                            sub_parts = rest.split('.', 1)
                            if len(sub_parts) >= 2:
                                page_number = number_part + '.' + sub_parts[0].strip()
                                page_title = sub_parts[1].strip()
                                if page_title:
                                    current_section['pages'].append({
                                        'number': page_number,
                                        'title': page_title
                                    })

            return {
                'course_name': course_name,
                'sections': sections
            }

        except Exception as e:
            print(f"❌ Error parsing file: {str(e)}")
            return None

    def select_udemy_file_interactive(self):
        """Interactively select a Udemy output file"""
        files = self.get_udemy_output_files()
        if not files:
            print("❌ No Udemy output files found")
            print("💡 Make sure Udemy-Titles-Fetcher/outputs contains *Clean.txt files")
            return None

        print("\n📚 Available Udemy Course Output Files:")
        for i, file_info in enumerate(files, 1):
            print(f"  {i}. {file_info['course_name']}")

        while True:
            try:
                choice = input(f"\nSelect file (1-{len(files)}) or 'q' to quit: ").strip()

                if choice.lower() == 'q':
                    return None

                choice_num = int(choice)
                if 1 <= choice_num <= len(files):
                    selected = files[choice_num - 1]
                    print(f"✅ Selected: '{selected['course_name']}'")
                    return selected
                else:
                    print(f"❌ Please enter a number between 1 and {len(files)}")
            except ValueError:
                print("❌ Please enter a valid number")
            except KeyboardInterrupt:
                print("\n❌ Selection cancelled")
                return None

    def create_pages_from_udemy_file(self):
        """Interactive creation of pages from Udemy output file"""
        print("🎓 Create Pages from Udemy Course Output")
        print("=" * 45)

        # Select Udemy file
        file_info = self.select_udemy_file_interactive()
        if not file_info:
            return None

        # Parse the file
        print(f"\n📖 Parsing file: {file_info['filename']}")
        parsed_data = self.parse_udemy_output_file(file_info['path'])

        if not parsed_data:
            print("❌ Failed to parse file")
            return None

        course_name = parsed_data['course_name']
        sections = parsed_data['sections']

        print(f"\n📚 Course: {course_name}")
        print(f"📂 Found {len(sections)} sections with pages")

        # Show sections
        print("\n📋 Available Sections:")
        for i, section in enumerate(sections, 1):
            print(f"  {i}. {section['title']} ({len(section['pages'])} pages)")

        # Ask which sections to create
        print("\nOptions:")
        print("  • Press Enter to create ALL sections")
        print("  • Enter section numbers (comma-separated, e.g., 1,3,4)")
        print("  • Type 'q' to quit")

        selection = input("\nYour choice: ").strip()

        if selection.lower() == 'q':
            print("❌ Operation cancelled")
            return None

        # Determine which sections to create
        if not selection:
            selected_sections = sections
            print(f"✅ Creating all {len(sections)} sections")
        else:
            try:
                section_indices = [int(x.strip()) for x in selection.split(',')]
                selected_sections = [sections[i-1] for i in section_indices if 1 <= i <= len(sections)]
                print(f"✅ Creating {len(selected_sections)} selected sections")
            except (ValueError, IndexError):
                print("❌ Invalid section selection")
                return None

        # Select notebook and section in OneNote
        print("\n📓 Select destination in OneNote:")
        notebook = self.select_notebook_interactive()
        if not notebook:
            return None

        section = self.select_section_interactive(notebook['id'])
        if not section:
            return None

        # Ask for page creation strategy
        print("\n📝 Page Creation Strategy:")
        print("  1. With numbering (default) - '1. Week 1', '1.1. Day 1 - Lesson Title'")
        print("  2. Without numbering - 'Week 1', 'Day 1 - Lesson Title'")
        print("  3. Numbering + section prefix - '1. Week 1', 'Week 1 - 1.1. Lesson Title'")

        strategy = input("\nSelect strategy (1-3, default=1): ").strip()
        if not strategy:
            strategy = '1'

        # Count total pages INCLUDING section pages
        total_lesson_pages = sum(len(s['pages']) for s in selected_sections)
        total_section_pages = len(selected_sections)
        total_pages = total_lesson_pages + total_section_pages

        # Confirm
        print(f"\n📊 Summary:")
        print(f"  Course: {course_name}")
        print(f"  Sections to create: {total_section_pages}")
        print(f"  Lesson pages: {total_lesson_pages}")
        print(f"  Total Pages: {total_pages}")
        print(f"  Destination: {notebook['displayName']} → {section['displayName']}")
        print(f"  Strategy: {['With numbering', 'Without numbering', 'Numbering + prefix'][int(strategy)-1] if strategy in ['1','2','3'] else 'With numbering'}")

        confirm = input(f"\n🚀 Create {total_pages} pages? (y/n): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("❌ Operation cancelled")
            return None

        # Create pages with retry mechanism
        print(f"\n🔨 Creating {total_pages} pages...")
        created_count = 0
        failed_count = 0
        failed_pages = []  # Track failed pages for retry

        for sect in selected_sections:
            # Create section page first
            if strategy == '1':
                section_title = f"{sect['number']}. {sect['title']}"
            elif strategy == '2':
                section_title = sect['title']
            elif strategy == '3':
                section_title = f"{sect['number']}. {sect['title']}"
            else:
                section_title = f"{sect['number']}. {sect['title']}"

            print(f"\n📂 Creating section: {section_title}")
            result = self.create_page(section['id'], section_title, "")

            if result:
                created_count += 1
                print(f"  ✅ [{created_count}/{total_pages}] {section_title}")
            else:
                failed_count += 1
                failed_pages.append({
                    'type': 'section',
                    'title': section_title,
                    'section_info': sect
                })
                print(f"  ❌ Failed: {section_title}")

            # Create lesson pages
            for page_info in sect['pages']:
                # Build page title based on strategy
                if strategy == '1':
                    # Default: With numbering (e.g., "1.1. Day 1 - Lesson Title")
                    page_title = f"{page_info['number']}. {page_info['title']}"
                elif strategy == '2':
                    # Without numbering (e.g., "Day 1 - Lesson Title")
                    page_title = page_info['title']
                elif strategy == '3':
                    # With section prefix (e.g., "Week 1 - 1.1. Lesson Title")
                    page_title = f"{sect['title']} - {page_info['number']}. {page_info['title']}"
                else:
                    # Default fallback
                    page_title = f"{page_info['number']}. {page_info['title']}"

                # Create the page
                result = self.create_page(section['id'], page_title, "")

                if result:
                    created_count += 1
                    print(f"  ✅ [{created_count}/{total_pages}] {page_title}")
                else:
                    failed_count += 1
                    failed_pages.append({
                        'type': 'lesson',
                        'title': page_title,
                        'page_info': page_info,
                        'section_info': sect
                    })
                    print(f"  ❌ Failed: {page_title}")

        # Final summary
        print(f"\n{'='*50}")
        print(f"✅ Successfully created: {created_count} pages")
        if failed_count > 0:
            print(f"❌ Failed to create: {failed_count} pages")
        print(f"{'='*50}")

        # Offer retry for failed pages
        if failed_pages:
            print(f"\n🔄 Retry Options:")
            print(f"  {len(failed_pages)} pages failed to create")
            retry = input(f"\nRetry failed pages? (y/n): ").strip().lower()

            if retry in ['y', 'yes']:
                return self._retry_failed_pages(section['id'], failed_pages, strategy, created_count, total_pages)

        return {
            'created': created_count,
            'failed': failed_count,
            'total': total_pages,
            'failed_pages': failed_pages
        }

    def _retry_failed_pages(self, section_id, failed_pages, strategy, initial_created_count, total_pages):
        """Retry creating failed pages with authentication refresh"""
        print(f"\n🔄 Retrying {len(failed_pages)} failed pages...")
        print("💡 Tip: If authentication expired, you may need to re-authenticate")

        # Ask if user wants to re-authenticate
        reauth = input("\nRe-authenticate before retry? (y/n, recommended if auth expired): ").strip().lower()
        if reauth in ['y', 'yes']:
            print("\n🔐 Re-authenticating...")
            if not self.authenticate(force_reauth=True):
                print("❌ Re-authentication failed")
                return {
                    'created': initial_created_count,
                    'failed': len(failed_pages),
                    'total': total_pages,
                    'failed_pages': failed_pages
                }
            print("✅ Re-authentication successful!\n")

        retry_created = 0
        retry_failed = 0
        still_failed = []

        for i, page_data in enumerate(failed_pages, 1):
            page_title = page_data['title']
            print(f"  🔄 [{i}/{len(failed_pages)}] Retrying: {page_title}")

            result = self.create_page(section_id, page_title, "")

            if result:
                retry_created += 1
                print(f"  ✅ Success!")
            else:
                retry_failed += 1
                still_failed.append(page_data)
                print(f"  ❌ Still failed")

        # Retry summary
        print(f"\n{'='*50}")
        print(f"📊 Retry Results:")
        print(f"  ✅ Successfully created: {retry_created} pages")
        print(f"  ❌ Still failed: {retry_failed} pages")
        print(f"\n📈 Overall Progress:")
        print(f"  ✅ Total created: {initial_created_count + retry_created}/{total_pages}")
        print(f"  ❌ Total failed: {retry_failed}/{total_pages}")
        print(f"{'='*50}")

        # Offer another retry if still have failures
        if still_failed:
            print(f"\n⚠️ {len(still_failed)} pages still failed")
            another_retry = input(f"Retry again? (y/n): ").strip().lower()

            if another_retry in ['y', 'yes']:
                return self._retry_failed_pages(section_id, still_failed, strategy,
                                               initial_created_count + retry_created, total_pages)

        return {
            'created': initial_created_count + retry_created,
            'failed': retry_failed,
            'total': total_pages,
            'failed_pages': still_failed
        }

    def _save_failed_pages(self, failed_pages, course_name, section_id):
        """Save failed pages to a JSON file for later retry"""
        try:
            import json
            from datetime import datetime

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"failed_pages_{timestamp}.json"

            save_data = {
                'course_name': course_name,
                'section_id': section_id,
                'timestamp': timestamp,
                'failed_count': len(failed_pages),
                'failed_pages': failed_pages
            }

            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, indent=2, ensure_ascii=False)

            print(f"\n💾 Failed pages saved to: {filename}")
            print(f"   You can retry these pages later using option 6")
            return filename
        except Exception as e:
            print(f"❌ Error saving failed pages: {str(e)}")
            return None

    def _load_failed_pages(self, filename):
        """Load failed pages from a JSON file"""
        try:
            import json

            with open(filename, 'r', encoding='utf-8') as f:
                save_data = json.load(f)

            print(f"\n📂 Loaded failed pages:")
            print(f"   Course: {save_data['course_name']}")
            print(f"   Failed pages: {save_data['failed_count']}")
            print(f"   Saved at: {save_data['timestamp']}")

            return save_data
        except FileNotFoundError:
            print(f"❌ File not found: {filename}")
            return None
        except Exception as e:
            print(f"❌ Error loading failed pages: {str(e)}")
            return None

    def retry_failed_pages_from_file(self):
        """Retry failed pages from a saved file"""
        print("🔄 Retry Failed Pages from File")
        print("=" * 40)

        # List available failed page files
        import os
        failed_files = [f for f in os.listdir('.') if f.startswith('failed_pages_') and f.endswith('.json')]

        if not failed_files:
            print("❌ No saved failed page files found")
            return None

        print("\n📋 Available retry files:")
        for i, filename in enumerate(failed_files, 1):
            print(f"  {i}. {filename}")

        try:
            choice = input(f"\nSelect file (1-{len(failed_files)}): ").strip()
            choice_num = int(choice)

            if 1 <= choice_num <= len(failed_files):
                filename = failed_files[choice_num - 1]
                save_data = self._load_failed_pages(filename)

                if not save_data:
                    return None

                # Re-authenticate before retry
                print("\n🔐 Re-authenticating for retry...")
                if not self.authenticate(force_reauth=True):
                    print("❌ Authentication failed")
                    return None

                section_id = save_data['section_id']
                failed_pages = save_data['failed_pages']

                # Retry the failed pages
                result = self._retry_failed_pages(section_id, failed_pages, '1', 0, len(failed_pages))

                # If all succeeded, delete the file
                if result['failed'] == 0:
                    try:
                        os.remove(filename)
                        print(f"\n🗑️ Removed retry file (all pages created successfully)")
                    except:
                        pass

                return result
            else:
                print(f"❌ Invalid selection")
                return None

        except (ValueError, KeyboardInterrupt):
            print("\n❌ Operation cancelled")
            return None

    def quick_create_multiple_pages(self, page_titles, page_content=""):
        """Create multiple pages with given titles"""
        created_pages = []
        failed_pages = []

        print(f"📝 Creating {len(page_titles)} pages...")

        for i, title in enumerate(page_titles, 1):
            print(f"Creating page {i}/{len(page_titles)}: '{title}'")

            page_data = self.create_page(self.default_section_id, title, page_content)
            if page_data:
                created_pages.append(page_data)
                print(f"✅ Created: '{title}'")
            else:
                failed_pages.append(title)
                print(f"❌ Failed: '{title}'")

        print(f"\n📊 Summary:")
        print(f"✅ Successfully created: {len(created_pages)} pages")
        if failed_pages:
            print(f"❌ Failed to create: {len(failed_pages)} pages")
            for title in failed_pages:
                print(f"   - {title}")

        return created_pages, failed_pages

def main():
    """Enhanced main function with interactive options"""
    try:
        onenote = OneNoteAutomation()

        # Authenticate
        if not onenote.authenticate():
            print("❌ Authentication failed")
            return

        print("\n🎯 OneNote Automation Tool")
        print("=" * 30)

        while True:
            print("\nChoose an option:")
            print("1. Quick create page (using defaults)")
            print("2. Interactive create page (choose notebook/section)")
            print("3. Quick create multiple pages (using defaults)")
            print("4. Interactive create multiple pages (choose notebook/section)")
            print("5. 🎓 Create pages from Udemy course output")
            print("6. List all structure")
            print("7. Reset cache")
            print("8. Exit")

            choice = input("\nEnter your choice (1-8): ").strip()

            if choice == '1':
                # Quick create using defaults
                page_title = input("Enter page title: ").strip()
                if page_title:
                    page_content = input("Enter page content (optional): ").strip()

                    # Check for clipboard image
                    if PIL_AVAILABLE:
                        has_image, image_info = onenote.check_clipboard_for_image()
                        if has_image:
                            use_clipboard = input(f"\n🖼️ {image_info}\nInclude clipboard image? (y/n): ").strip().lower()
                            if use_clipboard in ['y', 'yes']:
                                section = onenote.get_default_section()
                                if section:
                                    onenote.create_page_with_clipboard_image(section['id'], page_title, page_content)
                                else:
                                    print("❌ No default section available. Please use interactive mode.")
                                continue

                    result = onenote.quick_create_page(page_title, page_content)
                    if not result:
                        print("💡 Tip: Use interactive mode (option 2 or 4) to select notebook/section")
                else:
                    print("❌ Page title cannot be empty")

            elif choice == '2':
                # Interactive create single page
                onenote.interactive_create_page()

            elif choice == '3':
                # Quick create multiple pages using defaults
                print("\n📝 Quick Multiple Page Creation (using defaults)")
                print("Enter page titles. You can use multiple formats:")
                print("  • One per line")
                print("  • Comma-separated: Title1, Title2, Title3")
                print("  • Mixed: Title1, Title2")
                print("           Title3")
                print("\nPress Enter twice when done:")

                titles_input = ""
                empty_lines = 0

                while empty_lines < 2:
                    line = input()
                    if line.strip():
                        titles_input += line + "\n"
                        empty_lines = 0
                    else:
                        empty_lines += 1

                page_titles = onenote.parse_titles_input(titles_input)

                if page_titles:
                    page_content = input(f"\nEnter common content for all {len(page_titles)} pages (optional): ").strip()
                    print(f"\n📋 About to create {len(page_titles)} pages:")
                    for i, title in enumerate(page_titles, 1):
                        print(f"  {i}. {title}")

                    confirm = input(f"\nProceed with creating {len(page_titles)} pages? (y/n): ").strip().lower()
                    if confirm in ['y', 'yes']:
                        result = onenote.quick_create_multiple_pages(page_titles, page_content)
                        if not result:
                            print("💡 Tip: Use interactive mode (option 4) to select notebook/section")
                    else:
                        print("❌ Operation cancelled")
                else:
                    print("❌ No page titles provided")

            elif choice == '4':
                # Interactive create multiple pages
                onenote.interactive_create_multiple_pages()

            elif choice == '5':
                # Create pages from Udemy course output
                onenote.create_pages_from_udemy_file()

            elif choice == '6':
                # List structure
                onenote.list_all_structure()

            elif choice == '7':
                # Reset cache
                onenote.reset_cache()
                print("✅ Cache reset complete")

            elif choice == '8':
                # Exit
                print("👋 Goodbye!")
                break

            else:
                print("❌ Invalid choice. Please enter 1-8.")

    except KeyboardInterrupt:
        print("\n👋 Goodbye!")
    except Exception as e:
        print(f"❌ An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
