# Claude Writer - AI for LibreOffice

Give Claude full control over your LibreOffice Writer documents. Click a button, get a themed terminal with specialist AI agents for grammar, references, peer review, comment handling, and formatting.

All edits use **Track Changes** so you can accept or reject every modification.

## Install

```bash
git clone https://github.com/ot483/claude-libreoffice-writer.git
cd claude-libreoffice-writer
./install.sh
```

### Prerequisites

- **LibreOffice** (6.4+)
- **Claude Code** CLI - [install here](https://claude.ai/download)
- **tmux** (`sudo apt install tmux`)
- **Python 3** with UNO bridge (`sudo apt install python3-uno libreoffice-script-provider-python`)

The install script checks for all prerequisites and tells you what's missing.

## Usage

1. Open a document in LibreOffice Writer
2. Click the **Claude** button in the toolbar
3. A themed terminal opens with a menu bar and Claude Code below

### Agents

Press an F-key anytime to run a specialist agent:

| Key | Agent | What it does |
|-----|-------|-------------|
| **F1** | Grammar | Fixes spelling, grammar, and style issues with track changes |
| **F2** | References | Validates citations vs. reference list, flags mismatches |
| **F3** | Reviewer | Scientific peer review - adds constructive comments |
| **F4** | Comments | Reads all comments, categorizes them, proposes a plan, addresses each one, saves a report |
| **F5** | Formatting | Checks heading hierarchy, font consistency, figure/table numbering |
| **F6** | Selection | Reads your currently selected text and waits for instructions |

Or just type naturally in the Claude panel below the menu.

### Working with selected text

Select text in your document, then either:
- Press **F6** in the Claude terminal - Claude reads the selection and asks what to do
- **Tools > Ask Claude about selection** - a dialog asks for your instruction and sends it to Claude

### How edits work

1. Claude enables **Track Changes** before making any edits
2. Changes appear as colored markup in your document (insertions highlighted, deletions struck through)
3. You **accept or reject** each change in LibreOffice (right-click → Accept/Reject)

### @claude Comments

Write comments in your document starting with `@claude` to leave instructions:

> **@claude:** Rewrite this paragraph to be more concise

Then ask Claude to process them or press F4.

## Tools (31)

| Category | Tools |
|----------|-------|
| **Navigate** | `list_sections`, `read_section` |
| **Read** | `read_document`, `get_document_info`, `search_text`, `get_paragraph_details` |
| **Edit** | `insert_text`, `replace_text`, `delete_text`, `insert_paragraph` |
| **Style** | `set_paragraph_style`, `set_character_format`, `list_styles` |
| **Tables** | `list_tables`, `read_table` |
| **Comments** | `get_comments`, `add_comment`, `reply_to_comment`, `delete_comment`, `process_claude_comments` |
| **Track Changes** | `enable_track_changes`, `disable_track_changes`, `get_track_changes_status`, `accept_all_changes`, `reject_all_changes` |
| **Selection** | `get_selection` |
| **Safety** | `undo`, `redo`, `save_document` |
| **Reports** | `save_report` |

## Architecture

```
LibreOffice Writer
    ↓ click "Claude" button
    ├── Starts UNO acceptor (port 2002)
    └── Opens tmux session
            ┌──────────────────────────────────────────────────┐
            │                                                  │
            │   Claude Code (full window)                      │
            │   MCP connected to your document                 │
            │                                                  │
            ├──────────────────────────────────────────────────┤
            │ (◉_◉) Claude Writer │ F1 F2 F3 F4 F5 │ status  │
            └──────────────────────────────────────────────────┘
                    ↓ MCP (stdio, JSON-RPC)
                server.py → tools/ → UNO bridge → LibreOffice
```

## Uninstall

```bash
./uninstall.sh
```

## Troubleshooting

**Claude button doesn't appear:** Restart LibreOffice. If still missing: Tools > Customize > Toolbars > Add Command > Macros > claude_writer > StartClaude.

**MCP server not connecting:** Run `claude mcp list` - libreoffice should show "Connected". If not, run `./install.sh` again.

**tmux not found:** `sudo apt install tmux`

**"Python script provider not found":** `sudo apt install libreoffice-script-provider-python`

## Disclaimer

This software is provided as-is. **Always back up your documents before using Claude Writer.** AI-generated edits may be incorrect, incomplete, or destructive. The authors are not responsible for any data loss, document corruption, or unintended modifications. Review all Track Changes carefully before accepting them.

## License

MIT
