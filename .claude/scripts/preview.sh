#!/bin/bash
# Usage: ./preview.sh <markdown-file>
# Renders a markdown file to HTML with GitHub styling and opens in Windows browser
set -e
if [ -z "$1" ]; then
  echo "Usage: $0 <markdown-file>"
  exit 1
fi
SRC="$(realpath "$1")"
DIR="$(dirname "$SRC")"
BASE="$(basename "$SRC" .md)"
OUT="$DIR/$BASE-preview.html"

cat > "$OUT" << 'HTMLEOF'
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Preview</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/github-markdown-css@5.5.1/github-markdown-light.css">
<style>
  body { box-sizing: border-box; max-width: 980px; margin: 0 auto; padding: 45px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  img { max-width: 100%; }
  pre { background: #f6f8fa; padding: 16px; border-radius: 6px; overflow-x: auto; }
  code { background: #f6f8fa; padding: 2px 6px; border-radius: 3px; font-size: 85%; }
  pre code { background: transparent; padding: 0; }
</style>
</head>
<body class="markdown-body">
HTMLEOF
npx -y marked "$SRC" >> "$OUT" 2>/dev/null
echo "</body></html>" >> "$OUT"
explorer.exe "$(wslpath -w "$OUT")" 2>/dev/null || true
echo "Opened: $OUT"
