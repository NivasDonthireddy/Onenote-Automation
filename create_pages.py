#!/usr/bin/env python3
"""
OneNote Page Creator Utility
Usage: python create_pages.py
"""

from onenote_automation import OneNoteAutomation

def create_pages_interactive():
    """Interactive script to create OneNote pages"""
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
            for sec in sections:
                print(f"  - {sec['displayName']}")
            continue

        section = onenote.find_section_by_name(notebook['id'], section_name)
        if section:
            break
        print(f"❌ Section '{section_name}' not found. Please try again.")

    # Get page titles
    print(f"\n📝 Enter page titles to create (one per line, empty line to finish):")
    page_titles = []
    while True:
        title = input("Page title: ").strip()
        if not title:
            break
        page_titles.append(title)

    if not page_titles:
        print("❌ No page titles provided.")
        return

    # Confirm creation
    print(f"\n📋 Summary:")
    print(f"  Notebook: {notebook['displayName']}")
    print(f"  Section: {section['displayName']}")
    print(f"  Pages to create: {len(page_titles)}")
    for i, title in enumerate(page_titles, 1):
        print(f"    {i}. {title}")

    confirm = input("\n❓ Create these pages? (y/N): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("❌ Operation cancelled.")
        return

    # Create pages
    print(f"\n🚀 Creating pages...")
    created_pages = onenote.create_multiple_pages(section['id'], page_titles)

    print(f"\n✅ Done! {len(created_pages)} pages created successfully.")

def create_pages_from_list():
    """Create pages from a predefined list"""
    # Example configuration - modify as needed
    NOTEBOOK_NAME = "My Notebook"
    SECTION_NAME = "General"
    PAGE_TITLES = [
        "Weekly Planning - July 30, 2025",
        "Project Milestones",
        "Meeting Notes Template",
        "Ideas & Brainstorming",
        "Action Items"
    ]

    onenote = OneNoteAutomation()

    # Authenticate
    print("🔐 Authenticating...")
    if not onenote.authenticate():
        print("❌ Authentication failed.")
        return

    # Find notebook
    print(f"📚 Finding notebook: {NOTEBOOK_NAME}")
    notebook = onenote.find_notebook_by_name(NOTEBOOK_NAME)
    if not notebook:
        print(f"❌ Notebook '{NOTEBOOK_NAME}' not found.")
        print("Available notebooks:")
        onenote.get_notebooks()
        return

    # Find section
    print(f"📂 Finding section: {SECTION_NAME}")
    section = onenote.find_section_by_name(notebook['id'], SECTION_NAME)
    if not section:
        print(f"❌ Section '{SECTION_NAME}' not found.")
        print("Available sections:")
        onenote.get_sections(notebook['id'])
        return

    # Create pages
    print(f"🚀 Creating {len(PAGE_TITLES)} pages...")
    created_pages = onenote.create_multiple_pages(section['id'], PAGE_TITLES)

    print(f"✅ Done! {len(created_pages)} pages created.")

if __name__ == "__main__":
    print("OneNote Page Creator")
    print("===================")
    print("1. Interactive mode")
    print("2. Create from predefined list")

    choice = input("\nSelect mode (1 or 2): ").strip()

    if choice == "1":
        create_pages_interactive()
    elif choice == "2":
        create_pages_from_list()
    else:
        print("❌ Invalid choice.")
