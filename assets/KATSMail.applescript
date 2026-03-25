on createDebugMailInOutlook(paramString)
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "|||"
	set parts to every text item of paramString
	set AppleScript's text item delimiters to oldTIDs
	
	if (count of parts) is not 4 then error "Bad parameter payload."
	
	set recipientEmail to item 1 of parts
	set subjectText to item 2 of parts
	set bodyText to item 3 of parts
	set attachmentPath to item 4 of parts
	
	tell application "Microsoft Outlook"
		activate
		
		set newMessage to make new outgoing message with properties {subject:subjectText, content:bodyText}
		
		tell newMessage
			make new recipient at end of to recipients with properties {email address:{address:recipientEmail}}
			make new attachment with properties {file:POSIX file attachmentPath}
			open
		end tell
	end tell
	
	return "OK"
end createDebugMailInOutlook

