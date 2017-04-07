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

For Each iFile in iFiles
	With CreateObject("Scripting.FileSystemObject")
		executeGlobal .openTextFile(XartaScriptDir & iFile).readAll()
	End With	
Next
' *****************************************************************



Const XARTADEBUG = false ' MAKE SURE false FOR TASK SCHEDULED EVENTS (ECHO USED)

SetTasksAndDoBkUps XartaScriptDir

Sub SetTasksAndDoBkUps(XartaScriptDir)
	Dim o, jsonTasks
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

			WScript.Sleep 100 ' allow SchTasks time to add task
		ElseIf (WScript.Arguments(0) = jsonObj) Then
			With CreateObject("Scripting.FileSystemObject")
				executeGlobal jsonObj + " GetXartaJsonObject(XartaScriptDir), XartaScriptDir"
			End With
		End If
	Next
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

' Test IsOld function
'msgbox ( IsOld("20160305", "20170407", 4,1,6,30) )

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
End Sub


Sub WriteSqlBackUpScripts(o, XartaScriptDir)
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

End Sub

'BkUpMySql GetXartaJsonObject(XartaScriptDir), XartaScriptDir

Sub BkUpMySql(o, XartaScriptDir)

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
End Sub

'BkUpHMSdata GetXartaJsonObject(XartaScriptDir), XartaScriptDir

Sub BkUpHMSdata(o, XartaScriptDir)

	Dim RetVal, resultnum, resultdescription, filenameDest

	resultnum = hMailServer("stop")
	
	filenameDest = FileNameFormattedDateNow() & "hmdata.7z"

	RetVal = ZipToSambaShare(o, XartaScriptDir, _
								o("paths")("hmdata") & "\*", _
								filenameDest, _
								o("7zip")("Password"))

	resultnum = hMailServer("start")
	resultdescription = StartStopServiceError(resultnum)

End Sub





Sub SetScheduler(o, taskName, taskArg, SC, D, ST)
	' http://www.robvanderwoude.com/schtasks.php
	Dim taskActionPath
	taskActionPath = WScript.ScriptFullName & " " & Chr(34) & taskArg & Chr(34)

	Set wShell = CreateObject("Wscript.Shell")

	' only delete and re-add the scheduled task if this script is run without Arguments
	' (When scheduled there will be an argument as set here)

	' Looks like the delete instruction only works if the user trying to delete is the 
	' same user who authored the task before (doesn't matter who is set for run-as)
	wShell.Run "SchTasks /Delete /TN """ & taskName & """ /F", 0
	wShell.Run "SchTasks /Create /RU """ & o("windowsAccounts")("scheduler")("User") & """ /RP """  & _
		o("windowsAccounts")("scheduler")("Password") & """ /SC " & SC & " /D " & D & " /TN """ & _
		taskName & """ /TR """ & taskActionPath & """ /ST " & ST & """", 0
End Sub