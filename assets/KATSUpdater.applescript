on checkAndInstallUpdate(paramString)
	try
		set oldTIDs to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "|||"
		set parts to every text item of paramString
		set AppleScript's text item delimiters to oldTIDs

		if (count of parts) is not 4 then error "Bad parameter payload."

		set repoOwner to item 1 of parts
		set repoName to item 2 of parts
		set currentVersion to item 3 of parts
		set installDir to item 4 of parts

		set appScriptsDir to POSIX path of (path to home folder) & "Library/Application Scripts/com.microsoft.Word/"
		set updaterPath to appScriptsDir & "KATSUpdater.sh"

		do shell script "/usr/bin/test -x " & quoted form of updaterPath

		set cmd to "/bin/bash " & quoted form of updaterPath & space & quoted form of repoOwner & space & quoted form of repoName & space & quoted form of currentVersion & space & quoted form of installDir
		set resultText to do shell script cmd

		if resultText is equal to "" then
			return "FAILED|Updater returned no result."
		end if

		return resultText
	on error errText number errNum
		return "FAILED|" & errText & " (" & errNum & ")"
	end try
end checkAndInstallUpdate

on ping(paramString)
	return "OK:" & paramString
end ping

