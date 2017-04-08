' ----------------------------------------------------------------
' Add XartaMail User  account for hMailServer service
' Add XartaTask Admin account for scheduler (user logged on or not)
' any others ... reading-in from Xarta.json

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

iFiles.Add 	"XartaADS_constants.vbs", _
                                        " ... just constants"

iFiles.Add 	"XartaComputer.vbs", _
                                        "initialise & encapsulate " & _
                                        "in class host related " & _
                                        "objects/values for convenience " & _
                                        "and future extension"

iFiles.Add 	"XartaErrorCodes.vbs", _
                                        "Functions to return error " & _
                                        "descriptions"

'iFiles.Add 	"XartaUtilities.vbs", _
                                        '"Boring stuff grouped together"

'iFiles.Add "XartaLog.vbs", _
                                        '"Logging class I found on the net"

For Each iFile in iFiles
    With CreateObject("Scripting.FileSystemObject")
        executeGlobal .openTextFile(XartaScriptDir & iFile).readAll()
    End With	
Next
' *****************************************************************


' read-in xarta.json and decode it, look for windows accounts objects
' ... iterate over them, and for each one, if it doesn't already exist,
' ... create it, setting properties as per json object
createUsers(GetXartaJsonObject(XartaScriptDir))


Sub createUsers(o)

    Dim PC, jsonAccount
    Set PC = New XartaComputer

    Set jsonAccount = o("windowsAccounts")
        For Each winAccount in jsonAccount
            createUser _
                jsonAccount(winAccount)("User"), _
                jsonAccount(winAccount)("Password"), _
                jsonAccount(winAccount)("Group"), _ 
                jsonAccount(winAccount)("Fullname"), _
                jsonAccount(winAccount)("Description"), _
                PC
        Next
End Sub

Sub createUser(userName, userPass, userType, userFullname, userDescription, PC)
	
    Dim accountExists, colAccounts, objUser

    ' check if local account already exists
    accountExists = 0
    Set colAccounts = GetObject("WinNT://" & PC.HostName & "")
    colAccounts.Filter = Array("user")
    For Each objUser In colAccounts
        If objUser.Name = userName Then
            accountExists = 1
        End If
    Next

    If accountExists = 0 Then

        Set colAccounts = GetObject("WinNT://" & PC.HostName & ",computer")
        Set objUser = colAccounts.Create("user", userName)

        objUser.SetPassword userPass
        objUser.Fullname = userFullname
        objUser.Description = userDescription
        objUser.SetInfo 

        intUserFlags = objUser.Get("UserFlags")
        objUser.Put "UserFlags",    intUserFlags Or _
                                    ADS_UF_DONT_EXPIRE_PASSWD Or _
                                    ADS_UF_PASSWD_CANT_CHANGE
        objUser.SetInfo 

        Set Group = GetObject("WinNT://" & PC.HostName & "/" & userType & ",group")
        Group.Add(objUser.ADspath)

    End If

End Sub