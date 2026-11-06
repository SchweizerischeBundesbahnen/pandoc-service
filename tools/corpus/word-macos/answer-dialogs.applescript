-- Answers the dialogs Word shows while it opens a DOCX of pandoc-service. Returns how many it answered.
tell application "System Events" to tell process "Microsoft Word"
	set answered to 0
	repeat with w in (every window)
		if exists (button "Yes" of w) then
			click button "Yes" of w
			set answered to answered + 1
		else if name of w is "Update Table of Contents" then
			click radio button "Update entire table" of radio group 1 of w
			click button "OK" of w
			set answered to answered + 1
		end if
	end repeat
	return answered
end tell
