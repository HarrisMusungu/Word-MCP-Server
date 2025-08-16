# Simple Word Document MCP Server

A simplified MCP (Model Context Protocol) server for basic Word document operations using FastMCP.

## Features

This simplified version provides only essential Read and Write operations:

### Read Operations
- **read_document** - Extract all text content from a Word document
- **get_document_info** - Get basic document metadata (title, author, creation date, etc.)
- **list_documents** - List all .docx files in a directory

### Write Operations
- **create_document** - Create a new Word document with optional metadata
- **write_text** - Add text content to a document (append or replace)
- **add_heading** - Add formatted headings (levels 1-6)
- **replace_text** - Find and replace text throughout the document

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the server:
```bash
python word_mcp_server.py
```

## Configuration

Use the provided `mcp-config.json` configuration file with your MCP client:

```json
{
  "mcpServers": {
    "word-server": {
      "command": "python",
      "args": ["word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

## File Path Handling

**Important**: All document operations work with **local file paths** relative to the directory where the MCP server is running:

- **Relative paths**: `"my_document.docx"` or `"./docs/report.docx"`
- **Absolute paths**: `"/Users/username/Documents/file.docx"`
- **Directory operations**: Use `list_documents("./my_folder")` to list documents in a specific folder

The server automatically adds `.docx` extension if not provided.

## Usage Examples

### Reading Documents
```python
# Read all text from a document
read_document("my_document.docx")

# Get document information
get_document_info("my_document.docx")

# List all Word documents in current directory
list_documents(".")
```

### Writing Documents
```python
# Create a new document
create_document("new_doc.docx", title="My Document", author="John Doe")

# Add text content
write_text("new_doc.docx", "This is my first paragraph.", append=True)

# Add a heading
add_heading("new_doc.docx", "Chapter 1", level=1)

# Replace text
replace_text("new_doc.docx", "old text", "new text")
```

## Differences from Original

This simplified version removes:
- Complex formatting tools (tables, images, styles)
- Advanced features (footnotes, comments, protection)
- Multiple transport options
- Extensive utility modules
- Complex document manipulation tools

## Dependencies

- `fastmcp>=2.8.1` - MCP server framework
- `python-docx>=1.1.2` - Word document manipulation

## License

MIT License (same as original project)