#!/usr/bin/env python3
"""
OneNote Hotkey Image Inserter
Keep this running in the background for instant clipboard image insertion
"""

import time
import keyboard
from onenote_automation import OneNoteAutomation

class OneNoteHotkey:
    def __init__(self):
        self.onenote = None
        self.is_running = False

    def initialize(self):
        """Initialize OneNote automation and authenticate"""
        print("🚀 Initializing OneNote Hotkey Tool...")

        try:
            self.onenote = OneNoteAutomation()

            # Authenticate once at startup
            if not self.onenote.authenticate():
                print("❌ Failed to authenticate. Please check your credentials.")
                return False

            # Validate default settings
            if not self.onenote.default_notebook:
                print("⚠️  Warning: No DEFAULT_NOTEBOOK set in .env file")
                print("   Please update your .env file with:")
                print("   DEFAULT_NOTEBOOK=YourNotebookName")
                print("   DEFAULT_SECTION=YourSectionName")
                print("   DEFAULT_PAGE=YourPageName")
                return False

            if not self.onenote.default_section:
                print("⚠️  Warning: No DEFAULT_SECTION set in .env file")
                return False

            if not self.onenote.default_page:
                print("⚠️  Warning: No DEFAULT_PAGE set in .env file")
                return False

            print("✅ OneNote Hotkey Tool initialized successfully!")
            print(f"📚 Target: {self.onenote.default_notebook} > {self.onenote.default_section} > {self.onenote.default_page}")
            print("\n🔥 HOT STATE ACTIVE!")
            print("📋 Press Ctrl+Shift+V to insert clipboard image")
            print("🛑 Press Ctrl+C to exit")

            return True

        except Exception as e:
            print(f"❌ Initialization error: {str(e)}")
            return False

    def on_hotkey_pressed(self):
        """Handle hotkey press event"""
        print("\n⚡ Hotkey pressed! Adding clipboard image...")

        try:
            success = self.onenote.quick_add_clipboard_image()
            if success:
                print("🎉 Image added successfully!")
            else:
                print("💔 Failed to add image")
        except Exception as e:
            print(f"❌ Error adding image: {str(e)}")

        print("🔥 Ready for next image...")

    def run(self):
        """Start the hotkey listener"""
        if not self.initialize():
            return

        # Register hotkey (Ctrl+Shift+V)
        keyboard.add_hotkey('ctrl+shift+v', self.on_hotkey_pressed)

        self.is_running = True

        try:
            print("\n" + "="*50)
            print("🎯 READY! Use Ctrl+Shift+V to insert images")
            print("="*50)

            # Keep the script running
            while self.is_running:
                time.sleep(0.1)

        except KeyboardInterrupt:
            print("\n\n👋 Exiting OneNote Hotkey Tool...")
            self.is_running = False

def main():
    """Main entry point"""
    hotkey_tool = OneNoteHotkey()
    hotkey_tool.run()

if __name__ == "__main__":
    main()
