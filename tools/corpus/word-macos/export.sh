#!/bin/bash
# Exports each DOCX of $1 to a PDF in $2 with Microsoft Word for Mac, as the references of the corpus.
# Usage: tools/corpus/word-macos/export.sh corpus/docx corpus/reference/macos
# Word asks to update the fields of a document of pandoc-service: the script answers Yes, and updates the whole table of contents.
# Every call to Word has a time limit: a Word which hangs, on a dialog or otherwise, is killed and the next document goes on.
src=$(cd "$1" && pwd); dst=$(cd "$2" && pwd); here=$(dirname "$0")
limit() { perl -e 'alarm shift @ARGV; exec @ARGV' "$@"; }
quit_word() {
  limit 40 osascript "$here/quit-word.applescript" >/dev/null 2>&1 || pkill -9 -x "Microsoft Word"
  pgrep -x "Microsoft Word" >/dev/null && { pkill -9 -x "Microsoft Word"; sleep 2; }
  return 0
}
quit_word
for f in "$src"/[!~]*.docx; do
  n=$(basename "$f" .docx); out="$dst/$n.pdf"
  [ -f "$out" ] && continue
  limit 30 osascript -e "ignoring application responses" -e "tell application \"Microsoft Word\" to open (POSIX file \"$f\")" -e "end ignoring" >/dev/null 2>&1
  ok=0; quiet=0
  for _ in $(seq 1 60); do
    sleep 1
    r=$(limit 15 osascript "$here/answer-dialogs.applescript" 2>/dev/null)
    [ "$r" = "0" ] || { quiet=0; continue; }
    quiet=$((quiet+1))
    c=$(limit 10 osascript -e 'with timeout of 5 seconds' -e 'tell application "Microsoft Word" to count documents' -e 'end timeout' 2>/dev/null)
    [ "$c" = "1" ] && [ $quiet -ge 3 ] && { ok=1; break; }
  done
  [ $ok = 1 ] || { echo "$n: did not open"; quit_word; continue; }
  sleep 3
  limit 150 osascript -e 'with timeout of 120 seconds' -e "tell application \"Microsoft Word\" to save as active document file name \"$out\" file format format PDF" -e 'end timeout' >/dev/null 2>&1 || echo "$n: save failed"
  quit_word
  [ -f "$out" ] && echo "$n: ok" || echo "$n: no pdf"
done
