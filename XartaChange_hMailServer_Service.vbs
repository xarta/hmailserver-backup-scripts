' ----------------------------------------------------------------
' Just changes the log-on account for the hMailServer service
' TODO: make sure NTFS permissions set on my G: drive, and F: drive
' TODO: not sure if the account used for the hMailServer service
'       could do with batch rights?

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
iFiles.Add 	"XartaElevate.vbs", 		"Elevate (UAC) script"

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

For Each iFile in iFiles
	With CreateObject("Scripting.FileSystemObject")
		executeGlobal .openTextFile(XartaScriptDir & iFile).readAll()
	End With	
Next
' *****************************************************************


Dim PC, o
Set PC = New XartaComputer
Set o = GetXartaJsonObject(XartaScriptDir)

strSrv = "hMailServer"                                                 ' service name to look for
strSrvUser = ".\" & o("windowsAccounts")("mailservice")("User")        ' (local) user we want service to run as
strSrvUserPwd = o("windowsAccounts")("mailservice")("Password")


Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & PC.HostName & "\root\cimv2")
Set oServiceList = oWMIService.ExecQuery ("Select * from Win32_Service where Name ="   & """" & strSrv & """")

' https://msdn.microsoft.com/en-us/library/aa384901%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
' ... Change method of the Win32_Service class

For Each aService in oServiceList
    If aService.Name = strSrv Then
        ' msgbox aService.Name & " " & aService.Startname
        msgChangedService = "Changing service user of " & strSrv & " to " & strSrvUser & _
            ". Error description = "
        errService aService.Change( , , , , , , strSrvUser, strSrvUserPwd), msgChangedService
    End If
Next

Sub errService(errNum, msgService)
    
    Dim errDescription
    
    errDescription = StartStopServiceError(errnum)
    msgbox msgService & errDescription

End Sub