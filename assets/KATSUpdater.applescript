on checkAndInstallUpdate(paramString)
	try
		set appScriptsDir to POSIX path of (path to home folder) & "Library/Application Scripts/com.microsoft.Word/"
		set updaterPath to appScriptsDir & "KATSUpdater.sh"

		do shell script "/bin/test -x " & quoted form of updaterPath

		set cmd to "/bin/bash " & quoted form of updaterPath
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
