# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server for creating, reading, and manipulating Microsoft Word documents. It uses FastMCP to expose Word document operations as tools that AI assistants can invoke.

## Build and Run Commands

```bash
# Install dependencies
pip install -r requirements.txt
# or
pip install -e .

# Run the server (stdio transport - default)
python word_mcp_server.py
# or using the installed entry point
word_mcp_server

# Run with different transports (via environment variables)
MCP_TRANSPORT=streamable-http MCP_PORT=8000 python word_mcp_server.py
MCP_TRANSPORT=sse MCP_PORT=8000 python word_mcp_server.py

# Run tests
pytest tests/

# Run a single test
pytest tests/test_convert_to_pdf.py -v

# Enable debug logging
MCP_DEBUG=1 python word_mcp_server.py
```

## Architecture

The codebase follows a modular three-layer architecture:

```
word_document_server/
├── main.py           # Entry point: FastMCP server setup, transport config, tool registration
├── tools/            # MCP tool implementations (exposed to clients)
│   ├── document_tools.py      # create, copy, info, list, merge documents
│   ├── content_tools.py       # headings, paragraphs, tables, pictures, lists
│   ├── format_tools.py        # text formatting, styles, table formatting
│   ├── protection_tools.py    # password protection, restricted editing
│   ├── footnote_tools.py      # footnotes/endnotes with robust validation
│   ├── extended_document_tools.py  # PDF conversion, text search
│   └── comment_tools.py       # comment extraction
├── core/             # Core business logic (low-level document manipulation)
│   ├── styles.py     # heading/table style management
│   ├── tables.py     # table borders, styling, copying
│   ├── footnotes.py  # footnote/endnote XML manipulation
│   ├── protection.py # document protection internals
│   └── comments.py   # comment parsing
└── utils/            # Shared utilities
    ├── file_utils.py       # file operations, path handling
    ├── document_utils.py   # document property extraction, text search
    └── extended_document_utils.py  # additional utilities
```

**Data flow**: `main.py` registers tools → tools call core functions → core uses utils for common operations.

## Key Dependencies

- **python-docx**: Core Word document manipulation
- **FastMCP**: MCP protocol implementation (requires 2.8.1+)
- **msoffcrypto-tool**: Document password protection
- **docx2pdf**: PDF conversion (requires LibreOffice or MS Word)

## Transport Configuration

The server supports three transport modes configured via environment variables:
- `stdio` (default): Standard I/O for CLI tools like Claude Desktop
- `streamable-http`: HTTP transport at `http://{host}:{port}/mcp`
- `sse`: Server-Sent Events at `http://{host}:{port}/sse`

Environment variables: `MCP_TRANSPORT`, `MCP_HOST`, `MCP_PORT`, `MCP_PATH`, `MCP_SSE_PATH`

## Tool Registration Pattern

Tools are registered in `main.py` using FastMCP decorators. Each tool is a thin wrapper that delegates to the corresponding module in `tools/`:

```python
@mcp.tool()
def add_paragraph(filename: str, text: str, style: str = None, ...):
    return content_tools.add_paragraph(filename, text, style, ...)
```

When adding new tools: implement in the appropriate `tools/*.py` module, then register in `main.py`.

## Common Patterns

- All document operations use 0-based indexing for paragraphs, rows, and columns
- Colors are hex strings without '#' prefix (e.g., "FF0000" for red)
- Widths/sizes support "points" or "percentage" units
- Most functions accept `filename` as the first parameter (path to .docx file)
