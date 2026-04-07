#!/bin/bash
# Runs inside the bottom tmux pane - launches Claude Code

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE="$HOME/.claude-writer"
mkdir -p "$WORKSPACE"
cd "$WORKSPACE"

# Find files
PROMPT_FILE=""
AGENTS_JSON=""
for dir in "$SCRIPT_DIR" "$HOME/.config/libreoffice/4/user/Scripts/python"; do
    [ -f "$dir/system_prompt.txt" ] && [ -z "$PROMPT_FILE" ] && PROMPT_FILE="$dir/system_prompt.txt"
    [ -f "$dir/agents.json" ] && [ -z "$AGENTS_JSON" ] && AGENTS_JSON="$dir/agents.json"
done

# Find Claude
CLAUDE_PATH=""
if command -v claude >/dev/null 2>&1; then
    CLAUDE_PATH="claude"
else
    for p in "$HOME/.claude/local/claude" "$HOME/.local/bin/claude" "/usr/local/bin/claude"; do
        [ -x "$p" ] && CLAUDE_PATH="$p" && break
    done
fi

if [ -z "$CLAUDE_PATH" ]; then
    echo "Error: Claude Code CLI not found."
    echo "Install from: https://claude.ai/download"
    read -p "Press Enter to close..."
    exit 1
fi

# Build args
ARGS=(--name "Claude Writer")
[ -n "$PROMPT_FILE" ] && ARGS+=(--append-system-prompt-file "$PROMPT_FILE")
[ -n "$AGENTS_JSON" ] && ARGS+=(--agents "$(cat "$AGENTS_JSON")")

exec "$CLAUDE_PATH" "${ARGS[@]}"
