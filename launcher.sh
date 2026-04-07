#!/bin/bash
# Claude Writer launcher - tmux session with status bar menu
# Called by the LibreOffice macro to start Claude Code

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Find claude_pane.sh
CLAUDE_PANE_SH=""
for dir in "$SCRIPT_DIR" "$HOME/.config/libreoffice/4/user/Scripts/python"; do
    [ -f "$dir/claude_pane.sh" ] && CLAUDE_PANE_SH="$dir/claude_pane.sh" && break
done

# Check dependencies
if ! command -v tmux >/dev/null 2>&1; then
    echo "Error: tmux not found. Install it: sudo apt install tmux"
    exit 1
fi

if [ -z "$CLAUDE_PANE_SH" ]; then
    echo "Error: claude_pane.sh not found. Run install.sh again."
    exit 1
fi

# Kill existing session
tmux kill-session -t claude-writer 2>/dev/null

# Create session with just Claude
tmux new-session -d -s claude-writer "$CLAUDE_PANE_SH"

# Enable mouse scrolling and increase scrollback
tmux set-option -t claude-writer mouse on
tmux set-option -t claude-writer history-limit 10000

# Configure status bar as the menu (bottom of screen - away from Claude's logo)
tmux set-option -t claude-writer status on
tmux set-option -t claude-writer status-position bottom
tmux set-option -t claude-writer status-style "bg=#2a2a4e,fg=#e0e0f0"
tmux set-option -t claude-writer status-left-length 150
tmux set-option -t claude-writer status-right-length 0
tmux set-option -t claude-writer status-left \
    " #[fg=#c084fc,bold](◉_◉) Claude Writer#[fg=#666]  │  #[fg=#87ceeb]F1#[fg=#999] Grammar  #[fg=#98fb98]F2#[fg=#999] References  #[fg=#ffd700]F3#[fg=#999] Reviewer  #[fg=#c084fc]F4#[fg=#999] Comments  #[fg=#e0e0f0]F5#[fg=#999] Formatting#[fg=#666]  │  #[fg=#999]Just type below for free chat "
tmux set-option -t claude-writer status-right ""
tmux set-option -t claude-writer window-status-format ""
tmux set-option -t claude-writer window-status-current-format ""

# Bind F1-F5 keys to send agent commands
tmux bind-key -n F1 send-keys "@grammar Check the entire document for grammar, spelling, and style issues. Enable track changes first, then fix each issue." Enter
tmux bind-key -n F2 send-keys "@references Validate all references and citations in the document. Check for missing refs, broken numbering, and consistency." Enter
tmux bind-key -n F3 send-keys "@reviewer Review this document as a scientific peer reviewer. Comment on methodology, logic, structure, and clarity." Enter
tmux bind-key -n F4 send-keys "@comments Process all comments in the document. Analyze, categorize, and present a plan before making changes." Enter
tmux bind-key -n F5 send-keys "@formatting Check this document for formatting consistency and submission readiness. Report all issues." Enter

# Attach
exec tmux attach-session -t claude-writer
