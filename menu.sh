#!/bin/bash
# Persistent top menu pane for Claude Writer
# This runs in the top tmux pane and stays visible

PURPLE='\033[38;5;141m'
WHITE='\033[1;37m'
DIM='\033[2m'
CYAN='\033[38;5;117m'
GREEN='\033[38;5;114m'
YELLOW='\033[38;5;222m'
NC='\033[0m'

TARGET="claude-writer:0.1"

draw_menu() {
    clear
    echo -e "  ${PURPLE}(◉_◉)${NC}  ${WHITE}Claude Writer${NC}  ${DIM}- AI for your documents${NC}"
    echo -e "  ${CYAN}F1${NC} Grammar  ${GREEN}F2${NC} References  ${YELLOW}F3${NC} Reviewer  ${PURPLE}F4${NC} Comments  ${WHITE}F5${NC} Formatting"
    echo -e "  ${DIM}Press F1-F5 anytime to run an agent · type freely below · q to quit${NC}"
    echo -e "  ${DIM}─────────────────────────────────────────────────────────────────────────────────────────${NC}"
}

send_and_focus() {
    tmux send-keys -t "$TARGET" "$1" Enter
    sleep 0.2
    tmux select-pane -t "$TARGET"
}

draw_menu

while true; do
    read -rsn1 key
    case "$key" in
        1) send_and_focus "@grammar Check the entire document for grammar, spelling, and style issues. Enable track changes first, then fix each issue." ;;
        2) send_and_focus "@references Validate all references and citations in the document. Check for missing refs, broken numbering, and consistency." ;;
        3) send_and_focus "@reviewer Review this document as a scientific peer reviewer. Comment on methodology, logic, structure, and clarity." ;;
        4) send_and_focus "@comments Process all comments in the document. Analyze, categorize, and present a plan before making changes." ;;
        5) send_and_focus "@formatting Check this document for formatting consistency and submission readiness. Report all issues." ;;
        q|Q) tmux kill-session -t claude-writer 2>/dev/null; exit 0 ;;
    esac
done
