on selfTest(paramString)
	return "OK"
end selfTest

on httpGet(urlText)
	try
		set userAgentText to "Mozilla/5.0 Excel-CookPropertyDueDiligence-Mac"
		set curlCommand to "/usr/bin/curl --fail --silent --show-error --location --connect-timeout 15 --max-time 30 --user-agent " & quoted form of userAgentText & " " & quoted form of urlText
		set responseText to do shell script curlCommand without altering line endings
		return responseText
	on error errMsg number errNum
		return "__ERROR__|" & errNum & "|" & errMsg
	end try
end httpGet

on documentsFolder(paramString)
	try
		return POSIX path of (path to documents folder)
	on error errMsg number errNum
		return "__ERROR__|" & errNum & "|" & errMsg
	end try
end documentsFolder

on normalizeFolder(pathText)
	try
		if pathText starts with "/" then return pathText
		set folderAlias to pathText as alias
		return POSIX path of folderAlias
	on error
		return pathText
	end try
end normalizeFolder

on openFile(posixPath)
	try
		do shell script "/usr/bin/open " & quoted form of posixPath
		return "OK"
	on error errMsg number errNum
		return "__ERROR__|" & errNum & "|" & errMsg
	end try
end openFile
