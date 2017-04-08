' decode Xarta.json
Function GetXartaJsonObject(xDir)
    Dim fso, json, str
    Set json = New VbsJson
    Set fso = WScript.CreateObject("Scripting.Filesystemobject")
    str = fso.OpenTextFile(xDir & "Xarta.json").ReadAll
    Set GetXartaJsonObject = json.Decode(str)
End Function

' Dim o
' Set o = GetXartaJsonObject(XartaScriptDir)

' o is a nested dictionary ... syntax gets a little tricky
' till finally we reach a string "item" rather than dictionary item
' e.g. o("windowsAccounts")("scheduler")("User") ...
'   all keys; takes the entire depth to return the item for key "User"

'msgbox TypeName(o)
'msgbox o.Keys()(0)
'msgbox TypeName(o.Items()(0))
'msgbox o.Items()(0).Keys()(0)