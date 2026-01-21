# Creating Client Documents

Generate professional Word documents using your letterhead template with correct formatting.

## Invocation
- `/client-doc` - Create a new client document from markdown source

## When to Use
Use this skill when creating:
- Proposals
- Workshop outlines
- Reports
- Invoices
- Any official client-facing document

## Workflow

### Step 1: Gather Information
Before generating the document, ask for:

1. **Source file**: "What is the path to your markdown source file?"
2. **Output file**: "Where should I save the Word document?"
3. **Recipient details**:
   - "What is the recipient's full name (including title, e.g., Ms Jane Smith)?"
   - "What is their job title?"
   - "What organisation do they work for?"
   - "What is their address? (I'll need street address, city/postcode, and country)"
4. **Document title**: "What should the document title be?"
5. **Date**: "What date should appear on the document? (Leave blank for today)"

### Step 2: Confirm Details
Before generating, confirm all details with the user:

```
I'll create a document with these details:
- Recipient: [name], [title] at [organisation]
- Address: [full address]
- Document title: [title]
- Date: [date]
- Source: [source path]
- Output: [output path]

Shall I proceed?
```

### Step 3: Generate Document
Run the Python script:

```bash
source "[YOUR_VENV_PATH]/bin/activate" && python3 "[YOUR_SCRIPT_PATH]/create_client_doc.py" \
  --template "[YOUR_TEMPLATE_PATH]/letterhead.docx" \
  --source "[source_path]" \
  --output "[output_path]" \
  --recipient-name "[name]" \
  --recipient-title "[title]" \
  --recipient-org "[organisation]" \
  --recipient-address "[address]" \
  --recipient-city "[city_postcode]" \
  --recipient-country "[country]" \
  --doc-title "[document_title]" \
  --date "[date]"
```

### Step 4: Open for Review
```bash
open -a "Microsoft Word" "[output_path]"
```

## Customisation Required

Before using this skill, update the paths in Step 3:
- `[YOUR_VENV_PATH]` - Path to your Python virtual environment
- `[YOUR_SCRIPT_PATH]` - Path to the create_client_doc.py script
- `[YOUR_TEMPLATE_PATH]` - Path to your letterhead template

Also customise the Python script settings:
- `YOUR_NAME` - Your name for the sign-off
- `YOUR_QUALIFICATIONS` - Your qualifications
- `YOUR_TITLE` - Your professional title
- Font and spacing settings as needed

## Markdown Source Format

Your markdown source file should follow this structure:

```markdown
# Document Title

Content of the cover letter goes here. This will appear
as body paragraphs after the salutation.

Multiple paragraphs are supported. Just leave a blank
line between them.

\newpage

## Section Heading

Content that appears after the page break goes here.

- Bullet points are supported
- Like this

1. Numbered lists too
2. Like this

### Subsection

More content...
```

## Formatting Applied

| Element | Font Size | Spacing |
|---------|-----------|---------|
| Body text | 11pt | Space after, no space before |
| Qualifications | 10pt | — |
| Bulleted lists | 11pt | No space before/after |
| Headings | 12-14pt | Space before |
| Line spacing | — | 1.2 throughout |
