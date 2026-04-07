# Claude Writer - LibreOffice Writer MCP Integration

## Quick Start

```bash
./install.sh
```
Then open LibreOffice Writer and click the "Claude" button.

## Project Structure

```
claude-writer/
├── install.sh / uninstall.sh    # Install/remove scripts
├── launcher.sh                  # tmux session launcher
├── claude_pane.sh               # Claude Code pane startup
├── menu.sh                      # Persistent top menu pane
├── agents.json                  # Agent definitions (grammar, references, reviewer, comments, formatting)
├── system_prompt.txt            # System prompt with guardrails
├── server.py                    # MCP server (JSON-RPC over stdio)
├── uno_connection.py            # UNO bridge singleton
├── tools/
│   ├── __init__.py              # Tool registry
│   ├── document_read.py         # Read tools
│   ├── document_edit.py         # Edit tools
│   ├── document_nav.py          # Navigation: sections, save, undo/redo
│   ├── document_style.py        # Style tools
│   ├── tables.py                # Table reading
│   ├── comments.py              # Comment/annotation tools
│   ├── track_changes.py         # Track changes tools
│   └── report.py                # Report saving
└── extension/                   # LibreOffice extension (toolbar button)
    ├── META-INF/manifest.xml
    ├── Addons.xcu
    ├── description.xml / .txt
    └── python/claude_writer.py  # Macro (button handler)
```

## MCP Protocol

The server uses **newline-delimited JSON-RPC 2.0** over stdio.

## Tools (29)

### Navigate
- `list_sections` - List all headings with paragraph ranges
- `read_section` - Read a section by heading name (partial match)

### Read
- `read_document` - Read paragraphs with text and style names (paginated)
- `get_document_info` - Document metadata
- `search_text` - Regex/case-sensitive search
- `get_paragraph_details` - Character-level formatting

### Edit
- `insert_text`, `replace_text`, `delete_text`, `insert_paragraph`

### Style
- `set_paragraph_style`, `set_character_format`, `list_styles`

### Tables
- `list_tables`, `read_table`

### Comments
- `get_comments`, `add_comment`, `reply_to_comment`, `delete_comment`, `process_claude_comments`

### Track Changes
- `enable_track_changes`, `disable_track_changes`, `get_track_changes_status`
- `accept_all_changes`, `reject_all_changes`

### Safety
- `undo`, `redo`, `save_document`

### Reports
- `save_report` - Save a report file next to the document

## Agents

Defined in `agents.json`, invoked with `@agent_name`:
- `@grammar` - Grammar/spelling/style fixes with track changes
- `@references` - Citation validation
- `@reviewer` - Scientific peer review
- `@comments` - Systematic comment handler with report
- `@formatting` - Formatting consistency checker
