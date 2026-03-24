on installKatsTools(paramString)
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "|||"
	set parts to every text item of paramString
	set AppleScript's text item delimiters to oldTIDs
	
	if (count of parts) is not 3 then error "Bad parameter payload."
	
	set downloadUrl to item 1 of parts
	set destDir to item 2 of parts
	set targetName to item 3 of parts
	
	set tmpFile to (POSIX path of (path to temporary items)) & targetName
	
	set shellCmd to "(/usr/bin/curl -L --fail -o " & quoted form of tmpFile & " " & quoted form of downloadUrl & " && " & ¬
		"/bin/mkdir -p " & quoted form of destDir & " && " & ¬
		"while /usr/bin/pgrep -x " & quoted form of "Microsoft Word" & " >/dev/null 2>&1; do /bin/sleep 2; done && " & ¬
		"/bin/cp " & quoted form of tmpFile & " " & quoted form of (destDir & "/" & targetName) & " && " & ¬
		"/usr/bin/osascript -e " & quoted form of "display dialog \"KATS-Tools uppdaterad. Starta Word igen.\" buttons {\"OK\"} default button \"OK\" with title \"KATS-Tools\"" & ¬
		") >/dev/null 2>&1 &"
	
	do shell script "/bin/sh -c " & quoted form of shellCmd
	return "OK"
end installKatsTools
