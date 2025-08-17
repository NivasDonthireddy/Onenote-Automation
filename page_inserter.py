#!/usr/bin/env python3
"""
OneNote Page Inserter - Specify notebook/section and add multiple pages
Usage: python page_inserter.py
"""

from onenote_automation import OneNoteAutomation

def insert_pages_to_section():
    """Interactive script to insert multiple pages into a specified notebook/section"""
    onenote = OneNoteAutomation()

    # Authenticate
    print("🔐 Authenticating with Microsoft Graph...")
    if not onenote.authenticate():
        print("❌ Authentication failed. Please check your credentials in .env file.")
        return False

    print("✅ Authentication successful!\n")

    # Step 1: Select Notebook
    print("📚 STEP 1: Select Notebook")
    print("-" * 30)
    
    # Check if user wants to use default notebook
    if onenote.default_notebook:
        use_default = input(f"Use default notebook '{onenote.default_notebook}'? (Y/n): ").strip().lower()
        if use_default in ['', 'y', 'yes']:
            notebook = onenote.get_default_notebook()
            if notebook:
                print(f"✅ Using default notebook: '{notebook['displayName']}'")
            else:
                print(f"❌ Default notebook '{onenote.default_notebook}' not found.")
                return False
        else:
            notebook = None
    else:
        notebook = None

    # If no default or user declined, show all notebooks
    if not notebook:
        notebooks = onenote.get_notebooks()
        if not notebooks:
            print("❌ No notebooks found.")
            return False

        # Interactive notebook selection
        while True:
            notebook_name = input("\n📖 Enter notebook name (or 'list' to see all): ").strip()
            if notebook_name.lower() == 'list':
                print("Available notebooks:")
                for nb in notebooks:
                    print(f"  - {nb['displayName']}")
                continue

            notebook = onenote.find_notebook_by_name(notebook_name)
            if notebook:
                print(f"✅ Selected notebook: '{notebook['displayName']}'")
                break
            print(f"❌ Notebook '{notebook_name}' not found. Please try again.")

    # Step 2: Select Section
    print(f"\n📂 STEP 2: Select Section in '{notebook['displayName']}'")
    print("-" * 50)
    
    # Check if user wants to use default section
    if onenote.default_section:
        use_default = input(f"Use default section '{onenote.default_section}'? (Y/n): ").strip().lower()
        if use_default in ['', 'y', 'yes']:
            section = onenote.find_section_by_name(notebook['id'], onenote.default_section)
            if section:
                print(f"✅ Using default section: '{section['displayName']}'")
            else:
                print(f"❌ Default section '{onenote.default_section}' not found.")
                section = None
        else:
            section = None
    else:
        section = None

    # If no default or user declined, show all sections
    if not section:
        sections = onenote.get_sections(notebook['id'])
        if not sections:
            print("❌ No sections found in this notebook.")
            return False

        # Interactive section selection
        while True:
            section_name = input("\n📄 Enter section name (or 'list' to see all): ").strip()
            if section_name.lower() == 'list':
                print("Available sections:")
                for sec in sections:
                    print(f"  - {sec['displayName']}")
                continue

            section = onenote.find_section_by_name(notebook['id'], section_name)
            if section:
                print(f"✅ Selected section: '{section['displayName']}'")
                break
            print(f"❌ Section '{section_name}' not found. Please try again.")

    # Step 3: Add Pages
    print(f"\n📝 STEP 3: Add Pages to '{section['displayName']}'")
    print("-" * 50)
    print(f"Target: {notebook['displayName']} > {section['displayName']}")
    
    # Choose page addition method
    print("\nPage addition options:")
    print("1. Enter page titles manually (one by one)")
    print("2. Enter multiple titles at once (comma-separated)")
    print("3. Use predefined templates")
    print("4. Paste from clipboard (one title per line)")
    
    while True:
        method = input("\nChoose method (1-4): ").strip()
        if method in ['1', '2', '3', '4']:
            break
        print("❌ Invalid choice. Please enter 1, 2, 3, or 4.")

    page_titles = []
    
    if method == '1':
        # Manual entry one by one
        print("\n📝 Enter page titles (press Enter on empty line to finish):")
        while True:
            title = input("Page title: ").strip()
            if not title:
                break
            page_titles.append(title)
            print(f"  ✓ Added: '{title}'")
    
    elif method == '2':
        # Comma-separated entry
        titles_input = input("\n📝 Enter page titles separated by commas: ").strip()
        if titles_input:
            page_titles = [title.strip() for title in titles_input.split(',') if title.strip()]
    
    elif method == '3':
        # Predefined templates
        templates = {
            '1': {
                'name': 'Daily Planning',
                'titles': ['Today\'s Goals', 'Priority Tasks', 'Meetings', 'Notes', 'Tomorrow\'s Prep']
            },
            '2': {
                'name': 'Weekly Review',
                'titles': ['Monday Plan', 'Tuesday Plan', 'Wednesday Plan', 'Thursday Plan', 'Friday Plan', 'Weekend Goals', 'Weekly Summary']
            },
            '3': {
                'name': 'Project Management',
                'titles': ['Project Overview', 'Requirements', 'Timeline', 'Resources', 'Milestones', 'Risk Assessment', 'Status Updates']
            },
            '4': {
                'name': 'Meeting Notes',
                'titles': ['Agenda', 'Attendees', 'Discussion Points', 'Decisions Made', 'Action Items', 'Follow-up Tasks']
            },
            '5': {
                'name': 'Study Notes',
                'titles': ['Key Concepts', 'Important Formulas', 'Examples', 'Practice Problems', 'Summary', 'Questions for Review']
            }
        }
        
        print("\nAvailable templates:")
        for key, template in templates.items():
            print(f"{key}. {template['name']} ({len(template['titles'])} pages)")
        
        template_choice = input("\nSelect template (1-5): ").strip()
        if template_choice in templates:
            page_titles = templates[template_choice]['titles'].copy()
            print(f"✅ Selected template: {templates[template_choice]['name']}")
        else:
            print("❌ Invalid template choice.")
            return False
    
    elif method == '4':
        # Clipboard paste
        print("\n📝 Paste page titles (one per line, press Enter then Ctrl+Z+Enter on Windows or Ctrl+D on Unix to finish):")
        try:
            import sys
            lines = []
            while True:
                try:
                    line = input().strip()
                    if line:
                        lines.append(line)
                except EOFError:
                    break
            page_titles = lines
        except KeyboardInterrupt:
            print("\n❌ Operation cancelled.")
            return False

    if not page_titles:
        print("❌ No page titles provided.")
        return False

    # Show summary and confirm
    print(f"\n📋 Summary:")
    print(f"Notebook: {notebook['displayName']}")
    print(f"Section: {section['displayName']}")
    print(f"Pages to create: {len(page_titles)}")
    print("\nPage titles:")
    for i, title in enumerate(page_titles, 1):
        print(f"  {i}. {title}")
    
    # Pages will be created with empty content
    page_content = ""

    # Final confirmation
    confirm = input(f"\n🚀 Create {len(page_titles)} pages with empty content? (Y/n): ").strip().lower()
    if confirm in ['', 'y', 'yes']:
        print(f"\n🔨 Creating {len(page_titles)} pages...")
        created_pages, failed_pages = onenote.create_multiple_pages(
            section['id'], 
            page_titles, 
            page_content
        )
        
        # Final summary
        print(f"\n🎉 OPERATION COMPLETE!")
        print(f"📊 Results:")
        print(f"✅ Successfully created: {len(created_pages)} pages")
        if failed_pages:
            print(f"❌ Failed to create: {len(failed_pages)} pages")
            print("Failed pages:")
            for title in failed_pages:
                print(f"   - {title}")
        
        return len(created_pages) > 0
    else:
        print("❌ Operation cancelled.")
        return False

def quick_insert_pages():
    """Quick insert using default notebook and section from .env"""
    onenote = OneNoteAutomation()
    
    # Authenticate
    if not onenote.authenticate():
        print("❌ Authentication failed")
        return False
    
    # Use defaults
    section = onenote.get_default_section()
    if not section:
        print("❌ No default section configured. Please set DEFAULT_NOTEBOOK and DEFAULT_SECTION in .env file.")
        return False
    
    print(f"🎯 Quick insert to: {onenote._cached_notebook['displayName']} > {section['displayName']}")
    
    # Get pages to insert
    print("\n📝 Enter page titles (press Enter on empty line to finish):")
    page_titles = []
    while True:
        title = input("Page title: ").strip()
        if not title:
            break
        page_titles.append(title)
    
    if not page_titles:
        print("❌ No page titles provided.")
        return False
    
    # Create pages
    print(f"\n🔨 Creating {len(page_titles)} pages...")
    created_pages, failed_pages = onenote.create_multiple_pages(section['id'], page_titles)
    
    return len(created_pages) > 0

if __name__ == "__main__":
    print("📝 OneNote Page Inserter")
    print("=" * 40)
    print("Choose mode:")
    print("1. Interactive mode (select notebook/section)")
    print("2. Quick mode (use defaults from .env)")
    print("3. Clear cache and refresh data")

    mode = input("\nSelect mode (1/2/3): ").strip()

    if mode == '1':
        success = insert_pages_to_section()
    elif mode == '2':
        success = quick_insert_pages()
    elif mode == '3':
        print("🔄 Clearing cache...")
        onenote = OneNoteAutomation()
        onenote.reset_cache()
        print("✅ Cache cleared! Fresh data will be loaded next time.")
        success = True
    else:
        print("❌ Invalid choice.")
        success = False
    
    if success and mode in ['1', '2']:
        print("\n✅ Page insertion completed successfully!")
    elif not success and mode in ['1', '2']:
        print("\n❌ Page insertion failed or was cancelled.")
