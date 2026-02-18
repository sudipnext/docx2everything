# docx2everything

Convert DOCX files to plain text or markdown format with preserved structure.

## Installation

```bash
pip install docx2everything
```

Or install from source:

```bash
# Modern way (recommended)
pip install .

# Or using setup.py (deprecated but still works)
python setup.py install
```

## Testing Without Installation

The CLI script works directly without installation - no PYTHONPATH needed!

**Using CLI (no installation required):**
```bash
# Extract text
python3 bin/docx2everything demo.docx

# Convert to markdown
python3 bin/docx2everything --markdown demo.docx > output.md

# With images
python3 bin/docx2everything --markdown -i images/ demo.docx > output.md
```

**Using Python:**
```bash
# Set PYTHONPATH to current directory
PYTHONPATH=. python3 -c "import docx2everything; print(docx2everything.process('demo.docx')[:100])"
```

**In Python script:**
```python
import sys
sys.path.insert(0, '/path/to/python-docx2txt')

import docx2everything
text = docx2everything.process('document.docx')
```

## Usage

### Command Line

**Extract plain text:**
```bash
docx2everything document.docx
```

**Convert to markdown:**
```bash
docx2everything --markdown document.docx > output.md
```

**Extract images:**
```bash
docx2everything -i images/ document.docx
```

**Markdown with images:**
```bash
docx2everything --markdown -i images/ document.docx > output.md
```

### Python API

```python
import docx2everything

# Extract plain text
text = docx2everything.process("document.docx")

# Convert to markdown
markdown = docx2everything.process_to_markdown("document.docx")

# Extract images
text = docx2everything.process("document.docx", img_dir="images/")

# Markdown with images
markdown = docx2everything.process_to_markdown("document.docx", img_dir="images/")
```

## Features

- ✅ Plain text extraction
- ✅ Markdown conversion with preserved structure:
  - Tables → Markdown tables (with merged cells support, alignment hints)
  - Lists → Bulleted/numbered lists (with proper sequence tracking)
  - Headings → Markdown headings (#, ##, ###, etc.) with custom style detection
  - Formatting → Bold, italic, strikethrough
  - Links → Markdown links
  - Images → Markdown image references
  - Footnotes → Markdown footnote references `[^1]`
  - Endnotes → Markdown endnote references `[^1]`
  - Comments → Inline HTML comments with author info
  - Charts → Chart placeholders with type and metadata `*[Chart: Title (Chart Type)]*`
  - Page breaks → HTML comments `<!-- Page Break -->`
  - Section breaks → HTML comments `<!-- Section Break -->`
- ✅ Image extraction
- ✅ Header and footer support
- ✅ Custom style detection (parses styles.xml for better heading detection)
- ✅ Table formatting (column alignment detection and hints)
- ✅ Robust error handling for malformed DOCX files

## Requirements

Python 3.6+

## License

MIT License - see LICENSE.txt
