' ----------------------------------------------------------------
' Back-up my hMailServer Windows 10 installation ...
' 	Aiming for:
'		hMailServer settings
'		Email data folder
'		MySQL dump
' -----------------------------------------------------------------

'------------------------------------------------------------------
' Haven't used VBScript in a decade or so I think ... seemed handy 
' for hMailServer ... just Google searched for "include", and "json"
' etc. just to use slightly more modern approaches / handy things
' -----------------------------------------------------------------


' *****************************************************************
' INCLUDE FILES
' *************
' safe for when elevated (elevated changes working directory I think)
Dim iFiles, XartaScriptDir
XartaScriptDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

Set iFiles = CreateObject("Scripting.Dictionary")
'iFiles.Add 	"XartaElevate.vbs", 		"Elevate (UAC) script"

iFiles.Add 	"VbsJson.vbs", 				"Someone's great class for " & _
										"JSON encoding/decoding"

iFiles.Add 	"XartaJson.vbs", 			"Depends on VbsJson.vbs: " & _
										"decode Xarta.json (settings)"

iFiles.Add 	"XartaADS_constants.vbs", 	" ... just constants"

iFiles.Add 	"XartaComputer.vbs",		"initialise & encapsulate " & _
										"in class host related " & _
										"objects/values for convenience " & _
										"and future extension"

iFiles.Add 	"XartaErrorCodes.vbs",		"Functions to return error " & _
										"descriptions"

iFiles.Add 	"XartaUtilities.vbs",		"Boring stuff grouped together"

iFiles.Add "XartaLog.vbs",				"Logging class I found on the net"

For Each iFile in iFiles
	With CreateObject("Scripting.FileSystemObject")
		executeGlobal .openTextFile(XartaScriptDir & iFile).readAll()
	End With	
Next

Set Logging = New Cls_Logging
logging.logevent = true
call logging.write("Script started",1)

' *****************************************************************



Dim XARTADEBUG : XARTADEBUG= False

SetTasksAndDoBkUps XartaScriptDir

Sub SetTasksAndDoBkUps(XartaScriptDir)
	Dim o, jsonTasks
	Dim success : success = False
	Set o = GetXartaJsonObject(XartaScriptDir)
	Set jsonTasks = o("tasks")

	For Each jsonObj in jsonTasks
		If (WScript.Arguments.Count = 0) Then
			SetScheduler 	o, _
							jsonTasks(jsonObj)("TN"), _
							jsonObj, _
							jsonTasks(jsonObj)("SC"), _
							jsonTasks(jsonObj)("D"), _
							jsonTasks(jsonObj)("ST")

			call logging.write("Scheduled: " & jsonObj, 1)
			success = True
			WScript.Sleep 100 ' allow SchTasks time to add task
		ElseIf (WScript.Arguments(0) = jsonObj) Then
			success = True
			With CreateObject("Scripting.FileSystemObject")
				call logging.write("Calling: " & jsonObj,1)
				executeGlobal jsonObj + " GetXartaJsonObject(XartaScriptDir), XartaScriptDir"
			End With
		End If
	Next
	If (WScript.Arguments.Count > 1) Then
		If (WScript.Arguments(1)="DEBUG") Then
			XARTADEBUG = True
		End If
	End If

	If success = False Then
		call logging.write("Unknown parameter: " & WScript.Arguments(0),1)
	End If
End Sub

'CopyHMSsettings GetXartaJsonObject(XartaScriptDir), XartaScriptDir
'DeleteHMSsettings GetXartaJsonObject(XartaScriptDir), XartaScriptDir
'DeleteSqlDump GetXartaJsonObject(XartaScriptDir), XartaScriptDir

Sub CopyHMSsettings(o, XartaScriptDir)
On error resume Next

	sFolder = o("paths")("hmsettingsbkup") & "\"
	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")

	For Each oFile In oFSO.GetFolder(sFolder).Files
		If UCase(oFSO.GetExtensionName(oFile.Name)) = "7Z" Then
			newName = Mid(oFile.Name,10,15) & "-HMsettings"
			newName = Replace(newName,"-","")
			newName = Replace(newName," ","-")
			oFSO.MoveFile sFolder & oFile.Name, sFolder & newName
			RetVal = ZipToSambaShare(o, XartaScriptDir, _
						sFolder & newName, _
						newName, _
						o("7zip")("Password"))
			Exit For
		End if
	Next

	Set oFSO = Nothing

End Sub

' make sure not to try to delete something while it's still
' being used!  e.g. if it's still background-copying to samba share
' TODO add something in PowerShell script I can reference (for job completion)
Sub DeleteHMSsettings(o, XartaScriptDir)
	DeleteTodaysFile o("paths")("hmsettingsbkup") & "\"
End Sub

Sub DeleteSqlDump(o, XartaScriptDir)
	DeleteTodaysFile o("paths")("mysqldumpoutput") & "\"
End Sub

Sub DeleteTodaysFile(sFolder)
	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")

	For Each oFile In oFSO.GetFolder(sFolder).Files
		If  Mid(oFile.Name,1,8) = FileNameFormattedDateNow() Then
			oFSO.DeleteFile sFolder & oFile.Name
		End if
	Next

	Set oFSO = Nothing
End Sub

' Test IsyyyymmddDate
' msgbox IsyyyymmddDate("20170229")
' TODO: Make IsyyyymmddDate more efficient
Function IsyyyymmddDate(yyyymmdd)
On error resume Next

	Dim IsAdate : IsAdate = True

	If Not (IsNumeric(yyyymmdd)) Then
		IsAdate = False
	End If

	If Not (Len(yyyymmdd) = 8) Then
		IsAdate = False
	End If

	Dim testDay, testMonth, testYear
	testDay = CInt(Mid(yyyymmdd,7,2)) 
	testMonth = CInt(Mid(yyyymmdd,5,2))
	testYear = CInt(Mid(yyyymmdd,1,4))

	If (testDay < 1) Or (testDay > 31) Then
		IsAdate = False
	End If

	If (testMonth < 1) Or (testMonth > 12) Then
		IsAdate = False
	End If

	If (testYear < 2000) Or (testYear > 2050) Then
		IsAdate = False
	End If

	Dim testDate
	testDate = DateSerial(testYear, testMonth, testDay)

	If (Year(testDate) <> testYear) Then
		IsAdate = False
	End If

	If (Month(testDate) <> testMonth) Then
		IsAdate = False
	End If

	If (Day(testDate) <> testDay) Then
		IsAdate = False
	End If


	If Err.Number <> 0 Then
		IsyyyymmddDate = False
		Err.clear
	Else
		IsyyyymmddDate = IsAdate
	End If

End Function

' Call XartaBackup.vbs ApprovedDeleteOldBackUps from powershell
' that has Execution policy set to bypass, and that connects
' to UNC path where backups are stored
Sub ApprovedDeleteOldBackUps(o, XartaScriptDir)
On error resume Next

	' TODO: check UNC path accessible!
	DeleteOldFiles o, o("paths")("uncServer") & o("paths")("uncPath") & "\"

	If err.Number <> 0 Then
		call logging.write("ApprovedDeleteOldBackUps error",3)
	End If
End Sub

'ScheduledDeleteOldBackUps GetXartaJsonObject(XartaScriptDir), XartaScriptDir
Sub ScheduledDeleteOldBackUps(o, XartaScriptDir)
On error resume Next

	Dim script, args, unc, RetVal

	unc = 	" -uncServer " & o("paths")("uncServer") & _
			" -uncFullPath " & o("paths")("uncServer") & o("paths")("uncPath") & _
			" -uncUser " & o("network")("User") & _
			" -uncPass " & o("network")("Password")


	script = XartaScriptDir & "XartaDeleteOldBkUps.ps1"
	args = unc & " " & XartaScriptDir & "XartaBackup.vbs"

	RetVal = PowerShell(script, args, True)

	If err.Number <> 0 Then
		call logging.write("ScheduledDeleteOldBackUps error",3)
	End If
End Sub

' TODO: ONLY ALLOW A FEW DELETIONS PER DAY OR SOMETHING?
' IN CASE OF PROBLEM WITH CLOCK?
Sub DeleteOldFiles(o, sFolder)
	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Dim fileDate

	For Each oFile In oFSO.GetFolder(sFolder).Files
		If Len(oFile.Name) > 7 Then
			 fileDate = Mid(oFile.Name,1,8)
			 If IsyyyymmddDate(fileDate) Then
				If AllowedDeleteFile(o, fileDate) Then
					oFSO.DeleteFile sFolder & oFile.Name
				End If
			End If
		End if
	Next

	Set oFSO = Nothing
End Sub


' Test AllowedDeleteFile function:
' msgbox (AllowedDeleteFile(GetXartaJsonObject(XartaScriptDir), "20090101"))
Function AllowedDeleteFile(o, yyyymmddFile)
	AllowedDeleteFile = _
		IsOld(	yyyymmddFile, _
				FileNameFormattedDateNow(), _
				o("bkupKeep")("keepYears"), _
				o("bkupKeep")("keepMonths"), _
				o("bkupKeep")("keepWeeks"), _
				o("bkupKeep")("keepDays") )
End Function


' Test IsOld function
' msgbox ( IsOld("20160305", "20170407", 4,1,6,30) )
Function IsOld(yyyymmddFile, yyyymmddNow, keepYears, keepMonths, keepWeeks, keepdays)

	' True, False:
	Dim oldYearly, oldMonthly, oldWeekly, oldDaily
	Dim isNotFirstJan : isNotFirstJan = True
	Dim isNotFirstOfMonth : isNotFirstOfMonth = True
	Dim isNotSunday : isNotSunday = True

	If Weekday(yyyymmddToDateSerial(yyyymmddFile)) = 1 Then
		isNotSunday = False
	End If

	If Day(yyyymmddToDateSerial(yyyymmddFile)) = 1 Then
		isNotFirstOfMonth = False
	End If

	If Day(yyyymmddToDateSerial(yyyymmddFile)) = 1 And _
		Month(yyyymmddToDateSerial(yyyymmddFile)) = 1 Then
		isNotFirstJan = False
	End If
	
	oldDaily = IsOldInterval(yyyymmddFile, yyyymmddNow, keepdays, "d")
	oldDaily = oldDaily And 	isNotSunday And _
								isNotFirstOfMonth And _
								isNotFirstJan

	oldWeekly = IsOldInterval(yyyymmddFile, yyyymmddNow, keepWeeks, "w")
	oldWeekly = oldWeekly And 	isNotFirstOfMonth And _
								isNotFirstJan

	oldMonthly = IsOldInterval(yyyymmddFile, yyyymmddNow, keepMonths, "m")
	oldMonthly = oldMonthly And	isNotFirstJan

	oldYearly = IsOldInterval(yyyymmddFile, yyyymmddNow, keepYears, "yyyy")
	
	IsOld = oldDaily Or oldWeekly Or oldMonthly Or oldYearly

End Function


Function IsOldInterval(yyyymmddFile, yyyymmddNow, keepInterval, intervalType)
	
	Dim fd, n, numInterval

	fd = yyyymmddToDateSerial(yyyymmddFile)
	n = yyyymmddToDateSerial(yyyymmddNow)

	numInterval = DateDiff(intervalType, fd, n)
	IsOldInterval = (numInterval > keepInterval)
	
End Function



Sub BkUpHMSsettings(o, XartaScriptDir)
On error resume Next

	' ------------------------------------------------------------------------------------
	' Creates a hMailServer backup using the settings 
	' specified in hMailAdmin.
	' NOTES:
	' I had problems with this, and permissions and scheduling
	' No issue using a normal user account for scheduling if set
	' to "only when user is logged on", but despite all the ideas
	' and research I put into it, including using indirect bat files
	' in the %USERPROFILE% and checking for absolute paths, and 
	' other "start in" related corrections (I think it worked if I put
	' in the optional "start in" manually), I could only get this hMailServer
	' backup script to work (when scheduled), and when set to
	' "whether logged on or not", when the selected user is an
	' Administrator (note hMailServer service set to an ordinary
	' user account - an account I originally hoped to use for the scripted
	' backups - setting it with batch-rights in local policy etc.)
	Dim oApp
	Set oApp = CreateObject("hMailServer.Application")

	' Authenticate the client.
	Call oApp.Authenticate (o("hMailServer")("User"), o("hMailServer")("Password"))
	Call oApp.BackupManager.StartBackup()
	' ------------------------------------------------------------------------------------

	If err.Number <> 0 Then
		call logging.write("BkUpHMSsettings error",3)
	End If
End Sub


Sub WriteSqlBackUpScripts(o, XartaScriptDir)
On error resume Next

	' not using here but keeping for reference:
	' mysqldump return codes:
	'	1	EX_USAGE
	'	2	EX_MYSQLERR
	'	3	EX_CONSCHECK
	'	4	EX_EOM
	'	5	EX_EOF
	'	6	EX_ILLEGAL_TABLE

	
	Const ForWriting = 2
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(XartaScriptDir & "XartaMySqlDump.bat", ForWriting)
	
	objFile.Writeline(	"REM dynamically generated by XartaBackupvbs")
	objFile.Writeline(	chr(34) & o("paths")("mysqldumpexe") & chr(34) & _
						" --defaults-extra-file=" & o("paths")("mysqldumpdefaultsextrafile") & _
						" --verbose -u" & o("mySQL")("backup")("User") & _
						" -hlocalhost hmailserver 2> " & o("paths")("mysqldumpoutput") & "\dump.log > " & _
						"%1")
	
	Set objFile = objFSO.OpenTextFile(o("paths")("mysqldumpdefaultsextrafile"), ForWriting)

	objFile.Writeline("[mysqldump]")
	objFile.Writeline("user=" & o("mySQL")("backup")("User"))
	objFile.Writeline("password=" & chr(34) & o("mySQL")("backup")("Password") & chr(34))

	objFile.Close

	If err.Number <> 0 Then
		call logging.write("WriteSqlBackUpScripts error",3)
	End If
End Sub

'BkUpMySql GetXartaJsonObject(XartaScriptDir), XartaScriptDir

Sub BkUpMySql(o, XartaScriptDir)
On error resume Next

	WriteSqlBackUpScripts o, XartaScriptDir
	
	Dim DumpDate, DumpFilename, RetVal
	DumpDate = FileNameFormattedDateNow()
	DumpFilename = o("paths")("mysqldumpoutput") & "\" & DumpDate & "hmdump.sql"

	Set objShell = CreateObject("WScript.Shell")
	objShell.Run XartaScriptDir & "XartaMySqlDump.bat " & DumpFilename, 4, True

	'objShell.Run("powershell -ExecutionPolicy Bypass -noexit -file G:\XARTA-SCRIPTS\Xarta7zipTEST.ps1")

	RetVal = ZipToSambaShare(o, XartaScriptDir, _
							DumpFileName, _
							DumpDate & "hmdump.sql.7z", _
							o("7zip")("Password"))
	
	If err.Number <> 0 Then
		call logging.write("BkUpMySql error",3)
	End If
End Sub

'BkUpHMSdata GetXartaJsonObject(XartaScriptDir), XartaScriptDir

Sub BkUpHMSdata(o, XartaScriptDir)
On error resume Next

	Dim RetVal, resultnum, resultdescription, filenameDest

	resultnum = hMailServer("stop")
	
	filenameDest = FileNameFormattedDateNow() & "hmdata.7z"

	RetVal = ZipToSambaShare(o, XartaScriptDir, _
								o("paths")("hmdata") & "\*", _
								filenameDest, _
								o("7zip")("Password"))

	resultnum = hMailServer("start")
	resultdescription = StartStopServiceError(resultnum)
	
	If err.Number <> 0 Then
		call logging.write("BkUpHMSdata error",3)
	End If
End Sub





Sub SetScheduler(o, taskName, taskArg, SC, D, ST)
On error resume Next
	' http://www.robvanderwoude.com/schtasks.php
	Dim taskActionPath
	taskActionPath = WScript.ScriptFullName & " " & Chr(34) & taskArg & Chr(34)

	Set wShell = CreateObject("Wscript.Shell")

	' only delete and re-add the scheduled task if this script is run without Arguments
	' (When scheduled there will be an argument as set here)

	' Looks like the delete instruction only works if the user trying to delete is the 
	' same user who authored the task before (doesn't matter who is set for run-as)
	wShell.Run "SchTasks /Delete /TN """ & taskName & """ /F", 0

	If SC = "ONCE" Then
		SC = " /SC " & SC
		D =  " /SD " & D 	' a little hack to reduce json and argument-parsing-in-this-script
							' complexity (SD = Start Date)
	ElseIf Not (SC="") Then
		SC = " /SC " & SC
		If Not (D="") Then
			D = " /D " & D
		End If
	End If


	wShell.Run "SchTasks /Create /RU """ & o("windowsAccounts")("scheduler")("User") & """ /RP """  & _
		o("windowsAccounts")("scheduler")("Password") & """" & SC & D & " /TN """ & _
		taskName & """ /TR """ & taskActionPath & """ /ST " & ST & """", 0
	
	If err.Number <> 0 Then
		call logging.write("SetScheduler error",3)
	End If
End Sub