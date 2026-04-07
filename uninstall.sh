#!/bin/bash
# Uninstall Claude Writer from LibreOffice
#
# Usage:
#   ./uninstall.sh

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
NC='\033[0m'

ok() { echo -e "  ${GREEN}✓${NC} $1"; }

echo ""
echo "=== Uninstalling Claude Writer ==="
echo ""

# Find LibreOffice profile
LO_PROFILE=""
for candidate in \
    "$HOME/.config/libreoffice/4/user" \
    "$HOME/.config/libreoffice/3/user" \
    "$HOME/.config/libreoffice/user" \
    ; do
    if [ -d "$candidate" ]; then
        LO_PROFILE="$candidate"
        break
    fi
done

# Remove macro, launcher, and MCP server
if [ -n "$LO_PROFILE" ]; then
    rm -f "$LO_PROFILE/Scripts/python/claude_writer.py"
    rm -f "$LO_PROFILE/Scripts/python/launcher.sh"
    rm -f "$LO_PROFILE/Scripts/python/system_prompt.txt"
    rm -rf "$LO_PROFILE/Scripts/python/mcp_server"
    ok "Removed macro and MCP server files"
fi

# Remove gnome-terminal profile
if command -v dconf >/dev/null 2>&1; then
    for id in $(dconf list /org/gnome/terminal/legacy/profiles:/ 2>/dev/null); do
        name=$(dconf read "/org/gnome/terminal/legacy/profiles:/${id}visible-name" 2>/dev/null)
        if [ "$name" = "'Claude Writer'" ]; then
            # Remove from profile list
            PROFILE_ID=$(echo "$id" | tr -d ':/')
            PROFILES=$(dconf read /org/gnome/terminal/legacy/profiles:/list 2>/dev/null)
            NEW_PROFILES=$(echo "$PROFILES" | sed "s/, '$PROFILE_ID'//" | sed "s/'$PROFILE_ID', //" | sed "s/'$PROFILE_ID'//")
            dconf write /org/gnome/terminal/legacy/profiles:/list "$NEW_PROFILES"
            dconf reset -f "/org/gnome/terminal/legacy/profiles:/${id}"
            ok "Removed terminal profile"
            break
        fi
    done
fi

# Remove extension
unopkg remove com.claude.writer 2>/dev/null && \
    ok "Removed toolbar extension" || \
    ok "No toolbar extension to remove"

# Remove MCP server from Claude Code
CLAUDE_PATH=""
if command -v claude >/dev/null 2>&1; then
    CLAUDE_PATH="claude"
else
    for p in "$HOME/.claude/local/claude" "$HOME/.local/bin/claude" "/usr/local/bin/claude"; do
        if [ -x "$p" ]; then
            CLAUDE_PATH="$p"
            break
        fi
    done
fi

if [ -n "$CLAUDE_PATH" ]; then
    "$CLAUDE_PATH" mcp remove libreoffice 2>/dev/null && \
        ok "Removed MCP server from Claude Code" || \
        ok "No MCP server registration to remove"
fi

echo ""
echo -e "${GREEN}=== Uninstall complete ===${NC}"
echo ""
echo "Restart LibreOffice to remove the toolbar button."
echo ""
