-- Quits Word without saving: its close command fails on this version, so each document gets a fresh Word.
tell application "Microsoft Word" to quit saving no
repeat 30 times
	if application "Microsoft Word" is not running then exit repeat
	delay 1
end repeat
