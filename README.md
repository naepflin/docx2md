# docx2md

CLI tool to convert between DOCX and Markdown.

## Install

```bash
uv tool install -e .
```

## Usage

The direction is chosen from the input file, so there is nothing to pass:

```bash
# DOCX → Markdown
docx2md input.docx
docx2md input.docx output.md

# Markdown → DOCX
docx2md input.md
docx2md input.md output.docx

# Extract images to files instead of inline base64
docx2md --images input.docx
```

Detection reads the file, not its name: anything that is a DOCX container
converts to Markdown, and anything else is read as Markdown text. A file with
a `.docx` name that is not a valid DOCX is reported as an error rather than
being treated as Markdown.

`-r` is still accepted for compatibility, but it is no longer needed.

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
