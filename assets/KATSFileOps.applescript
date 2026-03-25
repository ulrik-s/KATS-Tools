on copyFileOutsideSandbox(paramString)
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "|||"
	set parts to every text item of paramString
	set AppleScript's text item delimiters to oldTIDs
	
	if (count of parts) is not 2 then error "Bad parameter payload."
	
	set sourcePath to item 1 of parts
	set targetPath to item 2 of parts
	
	set targetDir to do shell script "/usr/bin/dirname " & quoted form of targetPath
	
	do shell script "/bin/mkdir -p " & quoted form of targetDir
	do shell script "/bin/cp -f " & quoted form of sourcePath & " " & quoted form of targetPath
	
	return "OK"
end copyFileOutsideSandbox

on ping(dummyValue)
	return "OK"
end ping

