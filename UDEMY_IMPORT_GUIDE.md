# 🎓 Udemy Course Import Guide

## Overview

The OneNote Automation tool now includes a powerful feature to automatically create OneNote pages from Udemy course structure files. This is perfect for students who want to organize their course notes in OneNote.

## How It Works

1. **Detects Course Files**: Automatically finds `*- Clean.txt` files in the `Udemy-Titles-Fetcher/outputs` folder
2. **Parses Structure**: Reads the course structure (Weeks/Sections and individual lessons)
3. **Interactive Selection**: Lets you choose which sections to import
4. **Smart Naming**: Offers multiple page naming strategies
5. **Bulk Creation**: Creates all pages automatically in your chosen OneNote location

## Step-by-Step Usage

### 1. Run the Application

```bash
python onenote_automation.py
```

Or double-click `RUN.bat`

### 2. Select Option 5

```
Choose an option:
1. Quick create page (using defaults)
2. Interactive create page (choose notebook/section)
3. Quick create multiple pages (using defaults)
4. Interactive create multiple pages (choose notebook/section)
5. 🎓 Create pages from Udemy course output
6. List all structure
7. Reset cache
8. Exit

Enter your choice (1-8): 5
```

### 3. Select Your Course

The tool will show all available Udemy course files:

```
📚 Available Udemy Course Output Files:
  1. AI Engineer Agentic Track The Complete Agent & MCP Course
  2. AI Engineer Core Track LLM Engineering, RAG, QLoRA, Agents

Select file (1-2) or 'q' to quit:
```

### 4. Choose Sections to Import

After parsing, you'll see:

```
📚 Course: AI Engineer Agentic Track: The Complete Agent & MCP Course
📂 Found 6 sections with pages

📋 Available Sections:
  1. Week 1 (26 pages)
  2. Week 2 (21 pages)
  3. Week 3 (19 pages)
  4. Week 4 (23 pages)
  5. Week 5 (17 pages)
  6. Week 6 - MCP (23 pages)

Options:
  • Press Enter to create ALL sections
  • Enter section numbers (comma-separated, e.g., 1,3,4)
  • Type 'q' to quit
```

### 5. Select OneNote Destination

Choose your notebook and section where pages will be created:

```
📓 Select destination in OneNote:

📚 Available Notebooks:
  1. Python_Learnings
  2. Work Notes
  3. Personal

Select notebook (1-3):
```

### 6. Choose Naming Strategy

Three options for page titles:

```
📝 Page Creation Strategy:
  1. With numbering (default) - '1. Week 1', '1.1. Day 1 - Lesson Title'
  2. Without numbering - 'Week 1', 'Day 1 - Lesson Title'
  3. Numbering + section prefix - '1. Week 1', 'Week 1 - 1.1. Lesson Title'

Select strategy (1-3, default=1):
```

**Examples:**
- **Strategy 1** (Default): 
  - Section: `1. Week 1`
  - Lesson: `1.1. Day 1 - Autonomous AI Agent Demo: Using N8n to Control Smart Home Devices`
- **Strategy 2**: 
  - Section: `Week 1`
  - Lesson: `Day 1 - Autonomous AI Agent Demo: Using N8n to Control Smart Home Devices`
- **Strategy 3**: 
  - Section: `1. Week 1`
  - Lesson: `Week 1 - 1.1. Day 1 - Autonomous AI Agent Demo: Using N8n to Control Smart Home Devices`

**Note:** Strategy 1 is the default and matches the exact numbering from your Udemy course structure.

### 7. Confirm and Create

Review the summary and confirm:

```
📊 Summary:
  Course: AI Engineer Agentic Track: The Complete Agent & MCP Course
  Sections to create: 6
  Lesson pages: 129
  Total Pages: 135
  Destination: Python_Learnings → AI Engineering
  Strategy: With numbering

🚀 Create 135 pages? (y/n):
```

### 8. Watch It Create!

The tool will create all pages and show progress:

```
🔨 Creating 135 pages...

📂 Creating section: 1. Week 1
  ✅ [1/135] 1. Week 1
  ✅ [2/135] 1.1. Day 1 - Autonomous AI Agent Demo...
  ✅ [3/135] 1.2. Day 1 - AI Agent Frameworks Explained...
  ✅ [4/135] 1.3. Day 1 - Agent Engineering Setup...
  ...

📂 Creating section: 2. Week 2
  ✅ [27/135] 2. Week 2
  ✅ [28/135] 2.1. Day 1 - Understanding Async Python...
  ...

==================================================
✅ Successfully created: 135 pages
==================================================
```

## Tips & Best Practices

### 1. Organize Your Notebooks First

Create dedicated notebooks and sections in OneNote before importing:
- Example: Notebook "Udemy Courses" → Section "AI Engineering"

### 2. Import in Batches

For large courses, import one or two weeks at a time instead of all at once.

### 3. Use Consistent Naming

Choose a naming strategy that matches your note-taking style:
- **With numbering (default)**: Preserves course structure exactly, easy to navigate
- **Without numbering**: Cleaner look if you prefer minimal titles
- **Numbering + prefix**: Best for multiple courses in the same section, provides context

### 4. Verify Course Files

Make sure your Udemy course files are in:
```
Automation Projects/
  └── Udemy-Titles-Fetcher/
      └── outputs/
          ├── Course Name - Clean.txt
          └── Another Course - Clean.txt
```

### 5. Review Before Confirming

Always check the summary before confirming the creation of many pages.

## Troubleshooting

### No Udemy Files Found

**Problem**: "❌ No Udemy output files found"

**Solution**: 
- Check that files exist in `Udemy-Titles-Fetcher/outputs/`
- Ensure files end with `- Clean.txt`
- Verify the Udemy-Titles-Fetcher folder is in the same parent directory

### Parsing Failed

**Problem**: "❌ Failed to parse file"

**Solution**:
- Check that the file follows the expected format (see example below)
- Ensure the file has proper encoding (UTF-8)

### Pages Not Creating

**Problem**: Some pages fail to create

**Solution**:
- Check your OneNote authentication
- Verify you have write permissions to the notebook
- Try creating pages in smaller batches
- Check if page names are too long (OneNote has limits)

## Expected File Format

Your Udemy output files should look like this:

```
================================================================================
COURSE: Your Course Name
================================================================================

1. Section Name

   1.1. Lesson Title One
   1.2. Lesson Title Two
   1.3. Lesson Title Three

2. Another Section

   2.1. Another Lesson
   2.2. Yet Another Lesson
```

**Key Points:**
- Main sections start with a single number (1., 2., 3.)
- Sub-items (lessons) are indented and have two numbers (1.1., 1.2., 2.1.)
- The COURSE: line identifies the course name

## Integration with Udemy-Titles-Fetcher

This feature is designed to work seamlessly with the **Udemy-Titles-Fetcher** project:

1. Use Udemy-Titles-Fetcher to export course structure
2. It generates `*- Clean.txt` files in the outputs folder
3. OneNote Automation automatically detects these files
4. Import them into OneNote with one click!

## Benefits

✅ **Save Time**: No manual page creation  
✅ **Stay Organized**: Perfect structure matching course layout  
✅ **Focus on Learning**: Pages ready before you start the course  
✅ **Batch Processing**: Import entire courses at once  
✅ **Flexible**: Choose sections and naming as needed  

## Example Workflow

1. **Morning**: Export new Udemy course with Udemy-Titles-Fetcher
2. **Import**: Run OneNote Automation and import all weeks
3. **Study**: Take notes in pre-created pages as you watch videos
4. **Review**: All notes organized by course structure

## Video Courses Tested

✅ AI Engineer Agentic Track (135 items)  
✅ AI Engineer Core Track (Similar format)  
✅ Any course exported with Udemy-Titles-Fetcher  

---

**Happy Learning! 📚✨**

For more information, see the main README.md

