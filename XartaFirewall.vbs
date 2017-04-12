' ----------------------------------------------------------------
' Just a wrapper for XartaFirewall.ps1 to make sure Windows Firewall
' rules are set appropriately  TODO look for least priviledge 
' opportunities
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
iFiles.Add 	"XartaElevate.vbs", _
                                        "Elevate (UAC) script"

iFiles.Add	"VbsJson.vbs", _
                                        "Someone's great class for " & _
                                        "JSON encoding/decoding"

iFiles.Add 	"XartaJson.vbs", _
                                        "Depends on VbsJson.vbs: " & _
                                        "decode Xarta.json (settings)"

'iFiles.Add 	"XartaADS_constants.vbs", _
                                        '" ... just constants"

'iFiles.Add 	"XartaComputer.vbs", _
                                        '"initialise & encapsulate " & _
                                        '"in class host related " & _
                                        '"objects/values for convenience " & _
                                        '"and future extension"

'iFiles.Add 	"XartaErrorCodes.vbs", _
                                        '"Functions to return error " & _
                                        '"descriptions"

iFiles.Add 	"XartaUtilities.vbs", _
                                        "Boring stuff grouped together"

iFiles.Add "XartaLog.vbs", _
                                        "Logging class I found on the net"

For Each iFile in iFiles
    With CreateObject("Scripting.FileSystemObject")
        executeGlobal .openTextFile(XartaScriptDir & iFile).readAll()
    End With	
Next
' *****************************************************************

'TODO Add rules to Xarta.json and pass to XartaFirewall.ps1 
' (just a long string per rule)
UnsetSetFirewallRules GetXartaJsonObject(XartaScriptDir), XartaScriptDir
Sub UnsetSetFirewallRules(o, XartaScriptDir)
On error resume Next

    Dim script, args, RetVal


    script = XartaScriptDir & "XartaFirewall.ps1"
    args = ""
	
    RetVal = PowerShell(script, args, True)
	
    If err.Number <> 0 Then
        call logging.write("UnsetSetFirewallRules error",3)
    End If
End Sub