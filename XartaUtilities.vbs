Function FileNameFormattedDateNow()
	FileNameFormattedDateNow = Year(Date()) & Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2)
End Function

Function ZipToSambaShare(o, XartaScriptDir, _ 
							zipSource, _
							zipDestinationFileName, _
							zipPassword)

	Dim script, args, unc, zip

	unc = 	" -uncServer " & o("paths")("uncServer") & _
			" -uncFullPath " & o("paths")("uncServer") & o("paths")("uncPath") & _
			" -uncUser " & o("network")("User") & _
			" -uncPass " & o("network")("Password")

	zip = 	" -zipDestination " & zipDestinationFileName & _
			" -zipSource " & zipSource & _
			" -zipPassword " & zipPassword

	script = XartaScriptDir & "Xarta7zip.ps1"
	args = unc & zip & " " & XartaScriptDir & "Xarta7zip.bat"

	ZipToSambaShare = PowerShell(script, args)

End Function

Function PowerShell(script, args)
	Set objShell = CreateObject("WScript.Shell")
	' make sure args has leading space
	args = " " & LTrim(args)
	PowerShell = objShell.Run("powershell -ExecutionPolicy Bypass -noexit -file " & script & args)
End Function



Function hMailServer(startstop)
	
	Dim RetVal

	RetVal = -1

	If startstop = "start" Then
		RetVal = StartService("hMailServer")
	End If

	If startstop = "stop" Then
		RetVal = StopService("hMailServer")
	End If

	hMailServer = RetVal

End Function

Function StartService(servicename)
	Set ServiceSet = GetObject("winmgmts:").ExecQuery("select * from Win32_Service where Name='" & servicename & "'")
	
	For Each Service in ServiceSet
		RetVal = Service.StartService()
	Next

	StartService = RetVal
End Function

Function StopService(servicename)
	Set ServiceSet = GetObject("winmgmts:").ExecQuery("select * from Win32_Service where Name='" & servicename & "'")
	
	For Each Service in ServiceSet
		RetVal = Service.StopService()
	Next

	StopService = RetVal
End Function


' escape characters that create a problem in .bat files etc. EVEN when in quotes
Function Esc(p)
	Dim escaped
	escaped = ""
	For i=1 To Len(p)-1
		If Mid(p,i,1) = "%" Then
			escaped = escaped & "%%"
		Else
			escaped = escaped & Mid(p,i,1)
		End If
	Next

	Esc = escaped
End Function