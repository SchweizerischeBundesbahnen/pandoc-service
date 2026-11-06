#!/bin/bash
# Exports each DOCX of $1 to a PDF in $2 with Microsoft Word for Mac, as the references of the corpus.
# Usage: tools/corpus/word-macos/export.sh corpus/docx corpus/reference/macos
# Word asks to update the fields of a document of pandoc-service: the script answers Yes, and updates the whole table of contents.
src=$(cd "$1" && pwd); dst=$(cd "$2" && pwd)
osascript "$(dirname "$0")/quit-word.applescript" >/dev/null 2>&1
for f in "$src"/*.docx; do
  n=$(basename "$f" .docx); out="$dst/$n.pdf"
  [ -f "$out" ] && continue
  osascript -e "ignoring application responses" -e "tell application \"Microsoft Word\" to open (POSIX file \"$f\")" -e "end ignoring" >/dev/null 2>&1
  ok=0; quiet=0
  for i in $(seq 1 60); do
    sleep 1
    r=$(osascript "$(dirname "$0")/answer-dialogs.applescript" 2>/dev/null)
    [ "$r" = "0" ] || { quiet=0; continue; }
    quiet=$((quiet+1))
    c=$(osascript -e 'with timeout of 5 seconds' -e 'tell application "Microsoft Word" to count documents' -e 'end timeout' 2>/dev/null)
    [ "$c" = "1" ] && [ $quiet -ge 3 ] && { ok=1; break; }
  done
  [ $ok = 1 ] || { echo "$n: did not open"; continue; }
  sleep 3
  osascript -e 'with timeout of 120 seconds' -e "tell application \"Microsoft Word\" to save as active document file name \"$out\" file format format PDF" -e 'end timeout' >/dev/null 2>&1 || echo "$n: save failed"
  osascript "$(dirname "$0")/quit-word.applescript" >/dev/null 2>&1
  [ -f "$out" ] && echo "$n: ok" || echo "$n: no pdf"
done
