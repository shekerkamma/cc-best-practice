#!/usr/bin/env bash
# Demo PreToolUse hook: logs every Bash command Claude runs.
# Reads the hook payload JSON from stdin, appends a one-line entry to bash-log.txt.

LOG_FILE="$(dirname "$0")/bash-log.txt"
PAYLOAD="$(cat)"

TIMESTAMP="$(date '+%Y-%m-%d %H:%M:%S')"
TOOL_NAME="$(echo "$PAYLOAD" | python3 -c 'import sys,json; d=json.load(sys.stdin); print(d.get("tool_name",""))')"
COMMAND="$(echo "$PAYLOAD" | python3 -c 'import sys,json; d=json.load(sys.stdin); print(d.get("tool_input",{}).get("command",""))')"

echo "[$TIMESTAMP] $TOOL_NAME :: $COMMAND" >> "$LOG_FILE"

# Exit 0 = allow the tool call to proceed.
# Non-zero exit with stdout JSON {"decision":"block","reason":"..."} would block it.
exit 0
