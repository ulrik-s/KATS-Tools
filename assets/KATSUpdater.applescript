on checkAndInstallUpdate(paramString)
	try
		set oldTIDs to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "|||"
		set parts to every text item of paramString
		set AppleScript's text item delimiters to oldTIDs

		if (count of parts) is not 4 then error "Bad parameter payload."

		set repoOwner to item 1 of parts
		set repoName to item 2 of parts
		set currentVersion to my normalizeVersion(item 3 of parts)
		set installDir to item 4 of parts

		set appScriptsDir to POSIX path of (path to home folder) & "Library/Application Scripts/com.microsoft.Word"
		set tempDir to do shell script "/usr/bin/mktemp -d /tmp/kats-update.XXXXXX"
		set jsonPath to tempDir & "/latest.json"
		set zipPath to tempDir & "/KATS-Tools-mac-update.zip"
		set unpackDir to tempDir & "/payload"

		set apiUrl to "https://api.github.com/repos/" & repoOwner & "/" & repoName & "/releases/latest"

		do shell script "/usr/bin/curl -fsSL " & quoted form of apiUrl & space & ¬
			"-H " & quoted form of "Accept: application/vnd.github+json" & space & ¬
			"-H " & quoted form of "X-GitHub-Api-Version: 2022-11-28" & space & ¬
			"-o " & quoted form of jsonPath

		set latestVersion to my normalizeVersion(do shell script "/usr/bin/python3 -c " & quoted form of "import json,sys; d=json.load(open(sys.argv[1], encoding='utf-8')); print(d.get('tag_name',''))" & space & quoted form of jsonPath)

		if latestVersion is equal to "" then error "Could not determine latest release version."

		if my compareVersions(currentVersion, latestVersion) ≥ 0 then
			return "UPTODATE"
		end if

		set assetName to "KATS-Tools-mac-update.zip"
		set downloadUrl to do shell script "/usr/bin/python3 -c " & quoted form of "import json,sys; d=json.load(open(sys.argv[1], encoding='utf-8')); name=sys.argv[2]; print(next((a.get('browser_download_url','') for a in d.get('assets',[]) if a.get('name')==name), ''))" & space & quoted form of jsonPath & space & quoted form of assetName

		if downloadUrl is equal to "" then error "Release asset " & assetName & " not found."

		do shell script "/bin/mkdir -p " & quoted form of unpackDir
		do shell script "/usr/bin/curl -fL " & quoted form of downloadUrl & " -o " & quoted form of zipPath
		do shell script "/usr/bin/unzip -oq " & quoted form of zipPath & " -d " & quoted form of unpackDir

		do shell script "/bin/mkdir -p " & quoted form of installDir
		do shell script "/bin/mkdir -p " & quoted form of appScriptsDir

		if my fileExists(unpackDir & "/KATS-Tools.dotm") then
			do shell script "/bin/cp -f " & quoted form of (unpackDir & "/KATS-Tools.dotm") & space & quoted form of (installDir & "/KATS-Tools.dotm")
		end if

		if my fileExists(unpackDir & "/KATS-Version.txt") then
			do shell script "/bin/cp -f " & quoted form of (unpackDir & "/KATS-Version.txt") & space & quoted form of (installDir & "/KATS-Version.txt")
		end if

		if my fileExists(unpackDir & "/KATSUpdater.applescript") then
			do shell script "/bin/cp -f " & quoted form of (unpackDir & "/KATSUpdater.applescript") & space & quoted form of (appScriptsDir & "/KATSUpdater.applescript")
		end if

		if my fileExists(unpackDir & "/KATSMail.applescript") then
			do shell script "/bin/cp -f " & quoted form of (unpackDir & "/KATSMail.applescript") & space & quoted form of (appScriptsDir & "/KATSMail.applescript")
		end if

		if my fileExists(unpackDir & "/KATSFileOps.applescript") then
			do shell script "/bin/cp -f " & quoted form of (unpackDir & "/KATSFileOps.applescript") & space & quoted form of (appScriptsDir & "/KATSFileOps.applescript")
		end if

		do shell script "/bin/chmod 644 " & quoted form of (installDir & "/KATS-Tools.dotm") || true
		do shell script "/bin/chmod 644 " & quoted form of (installDir & "/KATS-Version.txt") || true
		do shell script "/bin/chmod 644 " & quoted form of (appScriptsDir & "/KATSUpdater.applescript") || true
		do shell script "/bin/chmod 644 " & quoted form of (appScriptsDir & "/KATSMail.applescript") || true
		do shell script "/bin/chmod 644 " & quoted form of (appScriptsDir & "/KATSFileOps.applescript") || true

		return "INSTALLED|" & latestVersion
	on error errText number errNum
		return "FAILED|" & errText & " (" & errNum & ")"
	end try
end checkAndInstallUpdate

on normalizeVersion(v)
	set t to v as text
	if t starts with "v" or t starts with "V" then
		return text 2 thru -1 of t
	end if
	return t
end normalizeVersion

on compareVersions(a, b)
	set aNorm to my normalizeVersion(a)
	set bNorm to my normalizeVersion(b)

	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "."
	set aParts to every text item of aNorm
	set bParts to every text item of bNorm
	set AppleScript's text item delimiters to oldTIDs

	set maxLen to (count of aParts)
	if (count of bParts) > maxLen then set maxLen to (count of bParts)

	repeat with i from 1 to maxLen
		set av to 0
		set bv to 0

		if i ≤ (count of aParts) then
			try
				set av to (item i of aParts) as integer
			end try
		end if

		if i ≤ (count of bParts) then
			try
				set bv to (item i of bParts) as integer
			end try
		end if

		if av < bv then return -1
		if av > bv then return 1
	end repeat

	return 0
end compareVersions

on fileExists(posixPath)
	try
		do shell script "/usr/bin/test -f " & quoted form of posixPath
		return true
	on error
		return false
	end try
end fileExists
