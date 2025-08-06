#!/usr/bin/env python3
"""
OneNote Single Page Creator with Image Support
Usage: python create_pages_with_images.py
"""

import os
from onenote_automation import OneNoteAutomation

def create_single_page_with_image():
    """Interactive script to create a single OneNote page with image"""
    onenote = OneNoteAutomation()

    # Authenticate
    print("🔐 Authenticating with Microsoft Graph...")
    if not onenote.authenticate():
        print("❌ Authentication failed. Please check your credentials in .env file.")
        return

    print("✅ Authentication successful!\n")

    # Get notebooks
    print("📚 Available notebooks:")
    notebooks = onenote.get_notebooks()

    if not notebooks:
        print("❌ No notebooks found.")
        return

    # Select notebook
    while True:
        notebook_name = input("\n📖 Enter notebook name (or type 'list' to see all): ").strip()
        if notebook_name.lower() == 'list':
            for nb in notebooks:
                print(f"  - {nb['displayName']}")
            continue

        notebook = onenote.find_notebook_by_name(notebook_name)
        if notebook:
            break
        print(f"❌ Notebook '{notebook_name}' not found. Please try again.")

    # Get sections
    print(f"\n📂 Available sections in '{notebook['displayName']}':")
    sections = onenote.get_sections(notebook['id'])

    if not sections:
        print("❌ No sections found in this notebook.")
        return

    # Select section
    while True:
        section_name = input("\n📄 Enter section name (or type 'list' to see all): ").strip()
        if section_name.lower() == 'list':
            for section in sections:
                print(f"  - {section['displayName']}")
            continue

        section = onenote.find_section_by_name(notebook['id'], section_name)
        if section:
            break
        print(f"❌ Section '{section_name}' not found. Please try again.")

    # Get page details
    page_title = input("\n📝 Enter page title: ").strip()
    if not page_title:
        print("❌ Page title cannot be empty.")
        return

    page_content = input("📄 Enter page content (optional): ").strip()

    # Image options
    print("\n🖼️ Image Options:")
    print("1. Add local image file")
    print("2. Add remote image URL")
    print("3. Paste image from clipboard")
    print("4. Create page without image")

    choice = input("\nSelect an option (1-4): ").strip()

    if choice == '1':
        # Local image
        image_path = input("🖼️ Enter path to image file: ").strip()

        # Validate image file
        if not onenote.validate_image_file(image_path):
            return

        print(f"\n🚀 Creating page '{page_title}' with local image...")
        result = onenote.create_page_with_image(
            section_id=section['id'],
            page_title=page_title,
            image_path=image_path,
            page_content=page_content
        )

    elif choice == '2':
        # Remote image
        image_url = input("🌐 Enter image URL: ").strip()
        if not image_url:
            print("❌ Image URL cannot be empty.")
            return

        print(f"\n🚀 Creating page '{page_title}' with remote image...")
        result = onenote.create_page_with_image(
            section_id=section['id'],
            page_title=page_title,
            image_url=image_url,
            page_content=page_content
        )

    elif choice == '3':
        # Clipboard image
        print("\n📋 Checking clipboard for image...")
        has_image, message = onenote.check_clipboard_for_image()

        if not has_image:
            print(f"❌ {message}")
            print("\n💡 Tips for using clipboard images:")
            print("   1. Take a screenshot (Windows Key + Shift + S)")
            print("   2. Copy an image from a website (right-click → Copy image)")
            print("   3. Copy an image from an image editor")
            return

        print(f"✅ {message}")
        print(f"\n🚀 Creating page '{page_title}' with clipboard image...")
        result = onenote.create_page_with_clipboard_image(
            section_id=section['id'],
            page_title=page_title,
            page_content=page_content
        )

    elif choice == '4':
        # No image
        print(f"\n🚀 Creating page '{page_title}' without image...")
        result = onenote.create_page(
            section_id=section['id'],
            page_title=page_title,
            page_content=page_content
        )

    else:
        print("❌ Invalid choice.")
        return

    if result:
        print("\n✅ Page created successfully!")
        # Display page details
        if 'id' in result:
            print(f"📄 Page ID: {result['id']}")
        if 'links' in result and 'oneNoteWebUrl' in result['links']:
            web_url = result['links']['oneNoteWebUrl'].get('href')
            if web_url:
                print(f"🌐 Page URL: {web_url}")
    else:
        print("\n❌ Failed to create page.")

if __name__ == "__main__":
    print("🖼️ OneNote Single Page Creator with Image")
    print("=" * 45)

    # Show supported image formats
    onenote_temp = OneNoteAutomation()
    print(f"📋 Supported image formats: {', '.join(onenote_temp.get_supported_image_formats())}")
    print()

    create_single_page_with_image()
