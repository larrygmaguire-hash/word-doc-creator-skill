# Client Document Automation with Claude Code

This guide walks you through setting up an automated workflow for generating professional Word documents from markdown using Claude Code. By the end, you'll have a system that creates branded documents with your letterhead, correct formatting, and consistent styling.

## What You'll Build

A Claude Code skill that:
- Takes markdown content and converts it to a professionally formatted Word document
- Preserves your letterhead (header, footer, logo)
- Applies consistent typography (fonts, sizes, spacing)
- Asks you questions to gather recipient details before generating
- Opens the finished document in Word for review

---

## Quick Start: What's Possible

### Can I just give Claude Code these files and ask it to set everything up?

**Yes, partially.** Here's what Claude Code can and cannot do:

| Task | Can Claude Code Do It? | Notes |
|------|------------------------|-------|
| Create folder structure | ✅ Yes | Just ask it to create the folders |
| Create virtual environment | ✅ Yes | Ask it to run the venv commands |
| Install python-docx | ✅ Yes | After venv is created |
| Copy/move files to correct locations | ✅ Yes | Specify source and destination |
| Update paths in the script | ✅ Yes | Tell it your project location |
| Update paths in the skill file | ✅ Yes | Tell it your project location |
| Create your letterhead | ❌ No | You must do this in Word manually |
| Design your header/footer | ❌ No | Requires Word's visual editor |
| Choose your fonts and branding | ❌ No | These are your design decisions |

### The Fastest Setup Path

1. **Give Claude Code all four files** (this document, the script, the skill, and the sample template)
2. **Tell Claude Code:**
   > "Set up this client document automation system in my project at `/path/to/my/project`. Create the folder structure, virtual environment, and update all paths in the script and skill file to match my system."
3. **Manually create your letterhead** in Word (see Step 3 below)
4. **Update your personal details** in the script (name, qualifications, title)
5. **Test with:** `/client-doc`

---

## Critical: Your Instructions Must Be Explicit

**Your template doesn't need to be perfect. Your instructions do.**

Claude Code can only implement what you specify. If you don't mention font sizes, you'll get defaults. If you don't specify heading styles, Claude will guess. The more explicit your formatting specification, the better your output.

### Formatting Specification Checklist

Before asking Claude Code to build or modify your document automation, work through this checklist. You don't need all of these—but you need to decide which ones matter for your documents.

#### Typography
| Element | Questions to Answer |
|---------|---------------------|
| **Body text font** | Which font? What size? (e.g., Calibri 11pt) |
| **Heading fonts** | Same as body or different? What sizes for H1, H2, H3? |
| **Bold/italic usage** | When should text be bold? Italic? Both? |
| **Qualifications/credentials** | Smaller font size? Different style? |

#### Paragraph Formatting
| Element | Questions to Answer |
|---------|---------------------|
| **Line spacing** | Single, 1.15, 1.5, double? |
| **Paragraph spacing** | Space before/after paragraphs? How much? |
| **First line indent** | Indented paragraphs or block style? |
| **Alignment** | Left-aligned, justified, or mixed? |

#### Lists
| Element | Questions to Answer |
|---------|---------------------|
| **Bullet style** | Round bullets, squares, dashes, custom? |
| **Numbered list style** | 1. 2. 3. or (1) (2) (3) or i. ii. iii.? |
| **Nested lists** | How should sub-items appear? Different bullet? |
| **List indentation** | How far indented from margin? |
| **Spacing within lists** | Space between items or tight? |

#### Document Structure
| Element | Questions to Answer |
|---------|---------------------|
| **Page breaks** | Where should pages break? After cover letter? Before appendices? |
| **Section breaks** | Different headers/footers for different sections? |
| **Table of contents** | Needed? Linked/clickable? What levels to include? |
| **Page numbers** | Position? Format? Start from 1 or skip cover page? |

#### Headers and Footers
| Element | Questions to Answer |
|---------|---------------------|
| **Header content** | Logo? Company name? Document title? |
| **Footer content** | Page numbers? Date? Company registration? |
| **First page different** | Different header/footer on page 1? |
| **Margins** | How much space for header/footer? |

#### Visual Elements
| Element | Questions to Answer |
|---------|---------------------|
| **Images/photos** | Where positioned? Size constraints? Captions? |
| **Charts/graphs** | Generated or embedded from Excel? Styling? |
| **Tables** | Border style? Header row formatting? Alternating row colours? |
| **Logos** | Size? Position? On every page or first only? |

#### Sign-off and Signature
| Element | Questions to Answer |
|---------|---------------------|
| **Closing phrase** | "Yours sincerely", "Kind regards", "Best wishes"? |
| **Name format** | Full name? With qualifications? |
| **Title/role** | Include job title? Company name? |
| **Signature image** | Digital signature? Space for handwritten? |

### Example: A Complete Formatting Specification

Here's what a thorough specification looks like. Give something like this to Claude Code:

```
Document Formatting Requirements:

TYPOGRAPHY
- Body text: Calibri 11pt, left-aligned
- H2 headings: Calibri 14pt bold
- H3 headings: Calibri 12pt bold
- Qualifications line: Calibri 10pt

SPACING
- Line spacing: 1.2
- Paragraph spacing: 6pt after each paragraph
- No first-line indent (block style)

LISTS
- Bullets: Round bullet character (•)
- Sub-bullets: Circle character (○)
- Numbered lists: 1. 2. 3. format
- List items: 3pt spacing between items

PAGE STRUCTURE
- Page break after cover letter (before main content)
- No table of contents needed
- Page numbers in footer, centred

HEADERS/FOOTERS
- Header: Company logo left, company name right
- Footer: Page number centre, registration details right
- First page: Same header/footer as other pages

SIGN-OFF
- Closing: "Yours sincerely" followed by blank line
- Name: Full name in bold
- Qualifications: On separate line, smaller font
- Title: Job title on final line
```

### What Happens Without Explicit Instructions

| If You Don't Specify... | Claude Code Will... |
|-------------------------|---------------------|
| Font name | Use a common default (often Calibri or Arial) |
| Font sizes | Use standard sizes (11pt body, 14pt headings) |
| Line spacing | Use single or 1.15 spacing |
| Bullet style | Use standard round bullets |
| Page breaks | Only break where you use `\newpage` |
| Table formatting | Use basic borders, no styling |
| Image placement | Place inline with text |

This may be fine for your needs—but if you want specific formatting, you must specify it.

### Updating the Script With Your Specifications

Once you've defined your formatting requirements, you have two options:

1. **Give them to Claude Code** and ask it to update `create_client_doc.py` to match
2. **Edit the script yourself** using the customisation section at the top

The script's customisation section handles common settings. For advanced formatting (tables, images, charts), you may need Claude Code to add new functions to the script.

---

## Prerequisites

Before starting, you need:

- **Claude Code** installed and working
- **Python 3** installed on your system (check with `python3 --version`)
- **Microsoft Word** (for creating your letterhead and viewing output)
- **A logo or branding elements** for your letterhead (optional but recommended)
- **Your formatting specification** (use the checklist above)

---

## Step 1: Set Up Your Project Folder

Create a folder structure for your document automation:

```
your-project/
├── .claude/
│   └── skills/
│       └── creating-client-documents.md
├── Scripts/
│   └── create_client_doc.py
├── Templates/
│   └── your-letterhead.docx
└── .venv/
```

**Ask Claude Code:**
> "Create this folder structure at `/path/to/my/project`"

Or manually:

```bash
mkdir -p your-project/.claude/skills
mkdir -p your-project/Scripts
mkdir -p your-project/Templates
```

---

## Step 2: Create a Python Virtual Environment

macOS prevents installing Python packages system-wide. A virtual environment isolates the dependencies.

**Ask Claude Code:**
> "Create a Python virtual environment at `/path/to/my/project/.venv` and install python-docx"

Or manually:

```bash
cd your-project
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx
```

**What is `.venv`?**
It's an isolated Python environment. The `python-docx` library lives inside it, separate from your system Python. Claude Code activates this environment automatically when running the script.

---

## Step 3: Prepare Your Letterhead Template (Manual Step)

**This step cannot be automated.** You must create your letterhead in Word.

The template is a Word document with your header and footer already configured. The script uses this as the base for all generated documents, preserving your branding.

### Option A: Use an Existing Letterhead

If you already have a Word document with your letterhead:
1. Open it in Word
2. Delete all body content (keep headers/footers intact)
3. Save it to `Templates/your-letterhead.docx`

### Option B: Create a New Letterhead

1. Open the included `sample-letterhead-template.docx`
2. **Edit the header:** Double-click the header area (or View → Header and Footer)
   - Add your logo
   - Add your company name
   - Add your address and contact details
3. **Edit the footer:** Double-click the footer area
   - Add registration details, VAT number, etc.
   - Add website/social links if desired
4. Delete any placeholder body text
5. Save as `Templates/your-letterhead.docx`

**Critical:** Headers and footers must be set using Word's Header & Footer tools, not just typed at the top/bottom of the page. The script preserves the document's section properties (`sectPr`) which contain header/footer references.

---

## Step 4: Install the Python Script

Copy `create_client_doc.py` to your `Scripts/` folder.

### Update the Paths

The script needs to know where your template lives. Find this section near the top:

```python
# Default template path - UPDATE THIS
DEFAULT_TEMPLATE = "/path/to/your-project/Templates/your-letterhead.docx"
```

**Ask Claude Code:**
> "Update the paths in `create_client_doc.py` to use my project at `/path/to/my/project`"

### Customise Your Branding

Find the customisation section:

```python
# =============================================================================
# CUSTOMISE THESE SETTINGS FOR YOUR BRAND
# =============================================================================

# Font settings
FONT_NAME = 'Calibri'           # Change to your preferred font
BODY_FONT_SIZE = 11             # Body text size in points
QUALIFICATIONS_FONT_SIZE = 10   # Smaller size for qualifications line
HEADING_2_SIZE = 14             # H2 heading size
HEADING_3_SIZE = 12             # H3 heading size

# Line spacing
LINE_SPACING = 1.2

# Your sign-off details - CHANGE THESE
YOUR_NAME = "Your Name"
YOUR_QUALIFICATIONS = "Your Qualifications Here"
YOUR_TITLE = "Your Professional Title"
```

Update these values to match your brand and personal details.

---

## Step 5: Install the Claude Code Skill

Copy `creating-client-documents.md` to `.claude/skills/`.

### Update the Paths in the Skill

The skill file contains the bash command that runs the script. You must update the paths:

```bash
source "/path/to/your-project/.venv/bin/activate" && python3 "/path/to/your-project/Scripts/create_client_doc.py" \
  --template "/path/to/your-project/Templates/your-letterhead.docx" \
  ...
```

**Ask Claude Code:**
> "Update the paths in `creating-client-documents.md` to use my project at `/path/to/my/project`"

**Important:** Use absolute paths (starting with `/`), not relative paths. On macOS, your path likely looks like:
```
/Users/yourname/Projects/your-project/
```

---

## Step 6: Create Your First Document

### Prepare Your Markdown Source

Create a markdown file with your document content:

```markdown
# Document Title

This is the first paragraph of your cover letter. It appears
after the salutation (Dear Name).

This is the second paragraph. Leave blank lines between
paragraphs.

\newpage

## Section Heading

Content that appears on page 2 and beyond goes here.

### Subsection

- Bullet points work
- Like this

1. Numbered lists too
2. Like this

**Bold text** is supported in paragraphs.
```

The `\newpage` marker tells the script where to insert a page break.

### Generate the Document

In Claude Code, type:

```
/client-doc
```

Claude will ask you:
1. Path to your markdown source file
2. Where to save the output
3. Recipient's name, title, organisation, and address
4. Document title
5. Date (or leave blank for today)

After confirming the details, Claude generates the document and opens it in Word.

---

## Step 7: Test and Refine

Generate a test document and check:

- [ ] Header appears correctly on all pages
- [ ] Footer appears correctly on all pages
- [ ] Fonts are correct throughout
- [ ] Spacing looks right (not too cramped or loose)
- [ ] Page break appears in the right place
- [ ] Your sign-off looks correct

If anything needs adjusting, edit the Python script settings and regenerate.

---

## Troubleshooting

### "python-docx not installed" error

The virtual environment isn't activated. The skill file should handle this automatically, but check the path to `.venv` is correct.

### Headers/footers not appearing

1. Check your template has actual headers/footers (not just text at top/bottom of page)
2. Ensure headers/footers were created via Word's Header & Footer tools (View → Header and Footer)
3. The script preserves `sectPr` (section properties) which reference headers/footers—if this element is missing from your template, headers won't appear

### Wrong font appearing

Word substitutes fonts if your specified font isn't installed. Either:
- Install the font on your system
- Choose a universally available font (Arial, Times New Roman, Calibri)

### Script path errors

Use absolute paths in the skill file, not relative paths. Find your full path with:
```bash
pwd
```

### Permission errors on macOS

If you see "externally-managed-environment" errors, the virtual environment wasn't activated. Check the skill file runs `source .venv/bin/activate` before calling Python.

---

## Advanced: Customising the Skill Questions

The skill file controls what questions Claude asks. You can modify it to:

- Add more fields (e.g., project reference number, invoice number)
- Change the order of questions
- Add validation (e.g., check file exists before proceeding)
- Skip questions you always answer the same way
- Add conditional logic (different questions for different document types)

Edit `.claude/skills/creating-client-documents.md` to customise the workflow.

---

## Files Included

| File | Purpose | Needs Customisation |
|------|---------|---------------------|
| `create_client_doc.py` | Python script that generates Word documents | Yes—paths and personal details |
| `creating-client-documents.md` | Claude Code skill file | Yes—paths only |
| `sample-letterhead-template.docx` | Starter template to customise | Yes—your branding |
| `SETUP-INSTRUCTIONS.md` | This guide | No |

---

## Summary

### What You Need To Do Manually
1. Create your letterhead in Word (header, footer, logo)
2. Update your personal details in the script (name, qualifications, title)

### What Claude Code Can Do For You
1. Create the folder structure
2. Set up the virtual environment
3. Install python-docx
4. Update all paths in the script and skill file
5. Generate documents on demand

### Quick Command Reference

| Command | What It Does |
|---------|--------------|
| `/client-doc` | Generate a new client document |
| Ask Claude to "update paths" | Reconfigure for your system |
| Ask Claude to "create a test document" | Verify everything works |

Every time you need to create a client document, just type `/client-doc` and answer the questions. The system handles formatting, letterhead, and styling automatically.
