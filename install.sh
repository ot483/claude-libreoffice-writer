#!/bin/bash
# Install Claude Writer for LibreOffice
#
# What this does:
#   1. Checks prerequisites (LibreOffice, Python script provider, Claude Code)
#   2. Installs the macro + MCP server into LibreOffice user scripts
#   3. Installs the toolbar button extension
#   4. Registers the MCP server with Claude Code
#
# Usage:
#   ./install.sh

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

ok()   { echo -e "  ${GREEN}✓${NC} $1"; }
warn() { echo -e "  ${YELLOW}!${NC} $1"; }
fail() { echo -e "  ${RED}✗${NC} $1"; }

echo ""
echo "=== Claude Writer for LibreOffice ==="
echo ""

# ---------- Check prerequisites ----------

echo "Checking prerequisites..."

# LibreOffice
if command -v soffice >/dev/null 2>&1; then
    ok "LibreOffice found"
else
    fail "LibreOffice not found"
    echo "  Install it from: https://www.libreoffice.org/download"
    exit 1
fi

# Python3
if command -v python3 >/dev/null 2>&1; then
    ok "Python3 found: $(python3 --version 2>&1)"
else
    fail "Python3 not found"
    exit 1
fi

# Python UNO bridge
if python3 -c "import uno" 2>/dev/null; then
    ok "Python UNO bridge available"
else
    # Try with dist-packages in path
    if PYTHONPATH="/usr/lib/python3/dist-packages" python3 -c "import uno" 2>/dev/null; then
        ok "Python UNO bridge available (via dist-packages)"
    else
        fail "Python UNO bridge not found"
        echo ""
        echo "  Install it:"
        echo "    Ubuntu/Debian: sudo apt install python3-uno"
        echo "    Fedora:        sudo dnf install libreoffice-pyuno"
        echo "    Arch:          sudo pacman -S libreoffice-fresh"
        exit 1
    fi
fi

# Python script provider for LibreOffice
SCRIPT_PROVIDER_OK=false
if dpkg -l libreoffice-script-provider-python >/dev/null 2>&1; then
    SCRIPT_PROVIDER_OK=true
elif rpm -q libreoffice-pyuno >/dev/null 2>&1; then
    SCRIPT_PROVIDER_OK=true
elif pacman -Q libreoffice-fresh >/dev/null 2>&1; then
    SCRIPT_PROVIDER_OK=true
elif [ -f /usr/lib/libreoffice/program/pythonloader.py ]; then
    SCRIPT_PROVIDER_OK=true
fi

if [ "$SCRIPT_PROVIDER_OK" = true ]; then
    ok "LibreOffice Python script provider"
else
    fail "LibreOffice Python script provider not found"
    echo ""
    echo "  Install it:"
    echo "    Ubuntu/Debian: sudo apt install libreoffice-script-provider-python"
    echo "    Fedora:        (included with libreoffice-pyuno)"
    echo "    Arch:          (included with libreoffice-fresh)"
    exit 1
fi

# Claude Code CLI
CLAUDE_PATH=""
if command -v claude >/dev/null 2>&1; then
    CLAUDE_PATH="$(which claude)"
    ok "Claude Code CLI found: $CLAUDE_PATH"
else
    for p in "$HOME/.claude/local/claude" "$HOME/.local/bin/claude" "/usr/local/bin/claude"; do
        if [ -x "$p" ]; then
            CLAUDE_PATH="$p"
            break
        fi
    done
    if [ -n "$CLAUDE_PATH" ]; then
        ok "Claude Code CLI found: $CLAUDE_PATH"
    else
        fail "Claude Code CLI not found"
        echo "  Install it from: https://claude.ai/download"
        exit 1
    fi
fi

# tmux
if command -v tmux >/dev/null 2>&1; then
    ok "tmux found"
else
    fail "tmux not found"
    echo "  Install it:"
    echo "    Ubuntu/Debian: sudo apt install tmux"
    echo "    Fedora:        sudo dnf install tmux"
    echo "    Arch:          sudo pacman -S tmux"
    exit 1
fi

echo ""

# ---------- Find LibreOffice user profile ----------

LO_PROFILE=""
for candidate in \
    "$HOME/.config/libreoffice/4/user" \
    "$HOME/.config/libreoffice/3/user" \
    "$HOME/.config/libreoffice/user" \
    "$HOME/Library/Application Support/LibreOffice/4/user" \
    ; do
    if [ -d "$candidate" ]; then
        LO_PROFILE="$candidate"
        break
    fi
done

if [ -z "$LO_PROFILE" ]; then
    fail "LibreOffice user profile not found"
    echo "  Try opening LibreOffice once to create it, then run this script again."
    exit 1
fi
ok "LibreOffice profile: $LO_PROFILE"

SCRIPTS_DIR="$LO_PROFILE/Scripts/python"
MCP_DIR="$SCRIPTS_DIR/mcp_server"

# ---------- Install macro ----------

echo ""
echo "Installing..."

mkdir -p "$SCRIPTS_DIR"
cp "$SCRIPT_DIR/extension/python/claude_writer.py" "$SCRIPTS_DIR/"
cp "$SCRIPT_DIR/launcher.sh" "$SCRIPTS_DIR/"
cp "$SCRIPT_DIR/claude_pane.sh" "$SCRIPTS_DIR/"
cp "$SCRIPT_DIR/menu.sh" "$SCRIPTS_DIR/"
cp "$SCRIPT_DIR/system_prompt.txt" "$SCRIPTS_DIR/"
cp "$SCRIPT_DIR/agents.json" "$SCRIPTS_DIR/"
chmod +x "$SCRIPTS_DIR/launcher.sh" "$SCRIPTS_DIR/claude_pane.sh" "$SCRIPTS_DIR/menu.sh"
ok "Installed macro, launcher, menu, agents, and system prompt"

# ---------- Install MCP server ----------

mkdir -p "$MCP_DIR/tools"
cp "$SCRIPT_DIR/server.py" "$MCP_DIR/"
cp "$SCRIPT_DIR/uno_connection.py" "$MCP_DIR/"
cp "$SCRIPT_DIR/tools/__init__.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/document_read.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/document_edit.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/document_style.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/comments.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/track_changes.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/report.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/document_nav.py" "$MCP_DIR/tools/"
cp "$SCRIPT_DIR/tools/tables.py" "$MCP_DIR/tools/"
ok "Installed MCP server"

# ---------- Create terminal profile (gnome-terminal) ----------

if command -v dconf >/dev/null 2>&1 && command -v gnome-terminal >/dev/null 2>&1; then
    echo ""
    echo "Setting up terminal theme..."

    # Check if profile already exists
    EXISTING_PROFILES=$(dconf read /org/gnome/terminal/legacy/profiles:/list 2>/dev/null)
    HAS_PROFILE=false
    if [ -n "$EXISTING_PROFILES" ]; then
        # Check all profiles for one named "Claude Writer"
        for id in $(dconf list /org/gnome/terminal/legacy/profiles:/ 2>/dev/null); do
            name=$(dconf read "/org/gnome/terminal/legacy/profiles:/${id}visible-name" 2>/dev/null)
            if [ "$name" = "'Claude Writer'" ]; then
                HAS_PROFILE=true
                break
            fi
        done
    fi

    if [ "$HAS_PROFILE" = true ]; then
        ok "Claude Writer terminal profile already exists"
    else
        PROFILE_ID=$(python3 -c "import uuid; print(uuid.uuid4())" 2>/dev/null || cat /proc/sys/kernel/random/uuid)
        PROFILE_PATH="/org/gnome/terminal/legacy/profiles:/:${PROFILE_ID}/"

        # Register profile
        if [ -z "$EXISTING_PROFILES" ] || [ "$EXISTING_PROFILES" = "@as []" ]; then
            dconf write /org/gnome/terminal/legacy/profiles:/list "['${PROFILE_ID}']"
        else
            NEW_LIST=$(echo "$EXISTING_PROFILES" | sed "s/]$/, '${PROFILE_ID}']/" )
            dconf write /org/gnome/terminal/legacy/profiles:/list "$NEW_LIST"
        fi

        # Theme: deep navy background, soft white text, purple accents
        dconf write "${PROFILE_PATH}visible-name" "'Claude Writer'"
        dconf write "${PROFILE_PATH}use-theme-colors" "false"
        dconf write "${PROFILE_PATH}background-color" "'rgb(26,26,46)'"
        dconf write "${PROFILE_PATH}foreground-color" "'rgb(224,224,240)'"
        dconf write "${PROFILE_PATH}cursor-colors-set" "true"
        dconf write "${PROFILE_PATH}cursor-background-color" "'rgb(192,132,252)'"
        dconf write "${PROFILE_PATH}cursor-foreground-color" "'rgb(26,26,46)'"
        dconf write "${PROFILE_PATH}palette" "['rgb(26,26,46)', 'rgb(255,107,107)', 'rgb(81,207,102)', 'rgb(255,212,59)', 'rgb(116,143,252)', 'rgb(192,132,252)', 'rgb(102,217,239)', 'rgb(224,224,240)', 'rgb(61,61,92)', 'rgb(255,135,135)', 'rgb(105,219,124)', 'rgb(255,224,102)', 'rgb(145,167,255)', 'rgb(208,168,255)', 'rgb(150,242,255)', 'rgb(255,255,255)']"
        dconf write "${PROFILE_PATH}bold-is-bright" "true"
        dconf write "${PROFILE_PATH}default-size-columns" "100"
        dconf write "${PROFILE_PATH}default-size-rows" "30"
        dconf write "${PROFILE_PATH}scrollback-lines" "10000"

        ok "Created Claude Writer terminal profile"
    fi
fi

# ---------- Build and install toolbar extension ----------

echo ""
echo "Installing toolbar button..."

# Close LibreOffice if running (extension install needs it)
if pgrep -x soffice.bin > /dev/null 2>&1; then
    warn "LibreOffice is running - toolbar button will appear after restart"
fi

EXT_DIR="$SCRIPT_DIR/extension"
OXT_FILE="/tmp/claude_writer_toolbar.oxt"
rm -f "$OXT_FILE"

cd "$EXT_DIR"
zip -q "$OXT_FILE" \
    META-INF/manifest.xml \
    description.xml \
    description.txt \
    Addons.xcu

# Try to install extension (may fail with python error - that's OK)
OUTPUT=$(unopkg add --force "$OXT_FILE" 2>&1) || true
if echo "$OUTPUT" | grep -qi "error" && ! echo "$OUTPUT" | grep -qi "enabling: python"; then
    warn "Extension install had issues - you can add the toolbar manually"
    warn "  Tools > Customize > Toolbars > Add Command > Macros > claude_writer > StartClaude"
else
    ok "Installed toolbar extension"
fi
rm -f "$OXT_FILE"

# ---------- Register MCP server with Claude Code ----------

echo ""
echo "Registering MCP server with Claude Code..."

# Find PYTHONPATH for uno module
UNO_PYTHONPATH=""
for p in "/usr/lib/python3/dist-packages" "/usr/lib/python3.*/dist-packages" "/usr/lib64/python3.*/site-packages"; do
    # Expand glob
    for expanded in $p; do
        if [ -f "$expanded/uno.py" ]; then
            UNO_PYTHONPATH="$expanded"
            break 2
        fi
    done
done

FULL_PYTHONPATH="$MCP_DIR"
if [ -n "$UNO_PYTHONPATH" ]; then
    FULL_PYTHONPATH="$MCP_DIR:$UNO_PYTHONPATH"
fi

# Remove existing entry if present
"$CLAUDE_PATH" mcp remove libreoffice 2>/dev/null || true

# Add MCP server
"$CLAUDE_PATH" mcp add libreoffice \
    -s user \
    -t stdio \
    -e "PYTHONPATH=$FULL_PYTHONPATH" \
    -- python3 "$MCP_DIR/server.py" 2>&1 && \
    ok "Registered MCP server with Claude Code" || \
    warn "Could not register MCP server automatically"

# ---------- Set up workspace ----------

echo ""
echo "Setting up workspace..."

WORKSPACE="$HOME/.claude-writer"
mkdir -p "$WORKSPACE"
cp "$SCRIPT_DIR/system_prompt.txt" "$WORKSPACE/CLAUDE.md"

# Trust the workspace directory in Claude Code
python3 -c "
import json, os
config_path = os.path.expanduser('~/.claude.json')
try:
    with open(config_path) as f:
        data = json.load(f)
except:
    data = {}

workspace = '$WORKSPACE'
if 'projects' not in data:
    data['projects'] = {}
if workspace not in data['projects']:
    data['projects'][workspace] = {}
data['projects'][workspace]['hasTrustDialogAccepted'] = True

with open(config_path, 'w') as f:
    json.dump(data, f, indent=2)
" 2>/dev/null && ok "Workspace trusted: $WORKSPACE" || warn "Could not auto-trust workspace"

# ---------- Done ----------

echo ""
echo -e "${GREEN}=== Installation complete! ===${NC}"
echo ""
echo "How to use:"
echo "  1. Open (or restart) LibreOffice Writer"
echo "  2. Click the 'Claude' button in the toolbar"
echo "     (or: Tools > Macros > claude_writer > StartClaude > Run)"
echo "  3. Claude Code opens in a terminal, connected to your document"
echo ""
echo "To uninstall: ./uninstall.sh"
echo ""
