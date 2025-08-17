#!/usr/bin/env python3
"""
Clear OneNote Cache - Reset cached notebook/section data
Usage: python clear_cache.py
"""

from onenote_automation import OneNoteAutomation

def clear_cache():
    """Clear the cached notebook and section data"""
    onenote = OneNoteAutomation()

    print("🔄 Clearing OneNote cache...")
    onenote.reset_cache()
    print("✅ Cache cleared successfully!")
    print("\nNext time you run the automation, it will:")
    print("- Re-fetch your notebooks")
    print("- Re-fetch your sections")
    print("- Use fresh data from OneNote")

if __name__ == "__main__":
    clear_cache()
