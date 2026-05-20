# docx2md

CLI tool to convert between DOCX and Markdown.

## Install

```bash
uv tool install -e .
```

## Usage

```bash
# DOCX → Markdown
docx2md input.docx
docx2md input.docx output.md

# Markdown → DOCX
docx2md -r input.md
docx2md -r input.md output.docx

# Extract images to files instead of inline base64
docx2md --images input.docx
```

## Supported formatting

- Headings (h1–h6)
- Bold, italic
- Ordered and unordered lists
- Tables
- Code blocks and inline code
- Blockquotes
- Links
- Images (embedded base64 or extracted to files)
- Horizontal rules
- Superscript / subscript
