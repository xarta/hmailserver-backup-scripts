' https://gallery.technet.microsoft.com/scriptcenter/Logging-class-object-34a6c110#content

'================================================================================= 
'Logging Class =================================================================== 
'================================================================================= 
 
'Not required for WSF, but is when in standard VBS 
Const ForAppending = 8 
 
'Log and Event writer object 
Class cls_Logging 
'Class for logging to file and event viewer 
 
Private oWShell,oNet,oFSo,Filehandle,rFilehandle 
Private fPath,strRFPath,fMaxSize,fLogname,strRemoteErr,BoolEvent,BoolRemote,oDict 
 
Private Sub Class_Initialize() 
    'Object Init subroutine 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    Set oWShell     = CreateObject("Wscript.Shell") 
    Set oNet         = CreateObject("Wscript.Network") 
    Set oFSo         = CreateObject("Scripting.FileSystemObject") 
    Set oDict        = CreateObject("Scripting.Dictionary") 
        LogEvent     = False 
        RemoteLog    = False 
        Path         = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)_ 
                        -Len(WScript.ScriptName))) 
        File         = LCase(oNet.ComputerName) 
        MaxSize = 2 
End Sub 
 
'--------------------------------------------------------- 
 
Private Sub Class_Terminate() 
    'Object Termination subroutine 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    If bOpen Then 
        Filehandle.close 
    End If 
        Set oWShell = Nothing 
        Set oNet = Nothing 
        Set oFSo = Nothing 
        Set Filehandle = Nothing 
        Set fPath = Nothing 
        Set fMaxSize = Nothing 
        Set fLogname = Nothing 
        Set BoolEvent = Nothing 
End Sub 
 
'--------------------------------------------------------- 
'File name properties, for changing and retrieving the log file name 
Public Property Let File(strFile) 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    If (InStr(StrReverse(strFile),"gol.")) <> 0 Then 
        fLogname = strFile 
    Else 
        fLogname = strFile & ".log" 
    End If 
End Property 
        Public Property Get File() 
            If Debugmode Then On Error Goto 0 Else On Error Resume Next 
                File = fLogname 
        End Property 
 
'--------------------------------------------------------- 
'Path name properties, for changing and retrieving the path to logs 
Public Property Let Path(strPath) 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
        If (InStr(StrReverse(strpath),"\")) <> 1 Then 
            fPath = strPath & "\" 
        Else 
            fPath = strPath 
        End If 
End Property 
        Public Property Get Path() 
            If Debugmode Then On Error Goto 0 Else On Error Resume Next 
            Path = fPath 
        End Property 
         
'--------------------------------------------------------- 
'Fully concatenated file name property for retrival. 
        Public Property Get FullFileName() 
            FullFileName = Path & File 
        End Property 
 
'--------------------------------------------------------- 
'Property for setting maximum file size of log file 
Public Property Let MaxSize(strVal) 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    fMaxSize = Cint(strVal) * 1048576 
End Property 
        Public Property Get MaxSize() 
            If Debugmode Then On Error Goto 0 Else On Error Resume Next 
            MaxSize = fMaxSize 
        End Property 
 
'--------------------------------------------------------- 
'Boolean property to determine if the filehandle is in use 
Private Property Get bOpen() 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next     
        If IsObject(Filehandle) Then 
            bOpen = True 
        Else 
            bOpen = False 
        End If 
End Property 
 
'--------------------------------------------------------- 
Public Property Let LogEvent( blValue) 
'Bool property that dictates event viewer rights 
    BoolEvent = blValue 
End Property 
    Private Property Get LogEvent() 
        LogEvent = BoolEvent 
    End Property 
     
Public Property Let RemoteLog( blValue) 
'Bool property that dictates if logging occurs to remote location 
    BoolRemote = blValue 
End Property 
    Private Property Get RemoteLog() 
        RemoteLog = BoolRemote 
    End Property 
Public Property Let RemotePath( strPath) 
        If (InStr(StrReverse(strpath),"\")) <> 1 Then 
            strRFPath = strPath & "\" 
        Else 
            strRFPath = strPath 
        End If 
End Property 
    Public Property Get RemotePath() 
        RemotePath = strRFPath 
    End Property 
 
'--------------------------------------------------------- 
Private Sub RemoteErrBuffer( strKey,  strItem) 
'Method to concactenate new items under one key at the end of the string 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
        Dim concat 
 
    If Not oDict.Exists(strKey) Then 
        Call oDict.Add(strkey, stritem) 
    Else 
        concat = oDict.Item(strKey) 
        concat = concat & "|:|" & strItem 
            oDict.Remove(strKey) 
        Call oDict.Add(strKey,concat) 
    End If 
 
End Sub 
 
'--------------------------------------------------------- 
Public Function ErrBuffer() 
'Method to return contents of the error buffer 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
     
    Dim ItemToSplit, ItemArray, item 
     
    ItemToSplit = oDict.item("remotelog") 
     
    ItemArray = Split(ItemToSplit, "|:|") 
     
    ErrBuffer = ItemArray 
     
End Function 
'--------------------------------------------------------- 
Public Sub WriteRemote(strVal) 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
         
    If Not CreateRemote Then 
        Exit Sub 
    End If 
 
        rFilehandle.WriteLine strVal 
 
End Sub 
 
'--------------------------------------------------------- 
'Subroutine for creating the remote log file and instantiating the handle 
Private Function CreateRemote() 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
 
    Dim FileProperty,Logsize 
 
    CreateRemote = False 
 
    If Not oFso.FolderExists(RemotePath) Then  
        Call Write(RemotePath & " Does not exist, or is unreachable.",3) 
            Exit Function 
    End If  
 
    If Not oFSo.FileExists(RemotePath & File) Then 
        oFso.CreateTextFile(RemotePath & File) 
    Else 
        oFSo.DeleteFile(RemotePath & File) 
            oFso.CreateTextFile(RemotePath & File) 
    End If     
        If Not IsObject(rFilehandle) Then 
               Set rFileHandle = oFSo.OpenTextFile(RemotePath & File, _ 
                ForAppending, True) 
        End If 
    If oFSo.FolderExists(RemotePath) Then 
        CreateRemote = True 
    End If 
End Function 
 
'--------------------------------------------------------- 
'Subroutine for writing log entries 
Public Function Write( msg,  mtype) 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    Dim msgline, etype 
        Call Create() 
 
        If Not bOpen Then 
            Call Create() 
        End If 
    msgline = "<![LOG["&msg&"]LOG]!><time="&""""&DatePart("h",Time) _ 
    &":"&DatePart("n",Time)&":"&DatePart("s",Time)&".000+0"""&" date=""" _ 
    &Replace(Date,"/","-")&""""&" component="""&Left(WScript.ScriptName, _ 
    Len(WScript.ScriptName)-Len(".vbs"))&""" context="""" type="""&mtype _ 
    &""" thread="""" file="""&Left(WScript.ScriptName,Len(WScript.ScriptName)_ 
    -Len(".vbs"))& """>" 
 
    Filehandle.WriteLine msgline 
 
 
        Select Case Mtype 
            Case 1 
                etype = 0 
            Case 2 
                etype = 2 
                        If LogEvent Then 
                    oWShell.LogEvent etype, msg 
                        End If  
                     
                    Call RemoteErrBuffer("remotelog", msg  & "," & "2") 
                     
            Case Else 
                etype = 1 
                        If LogEvent Then 
                    oWShell.LogEvent etype, msg 
                        End If 
 
                    Call RemoteErrBuffer("remotelog", msg  & "," & "1") 
 
        End Select 
     
End Function 
 
'--------------------------------------------------------- 
'Subroutine for rolling over log file at file size limit 
Private Sub Rollover() 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
    If bOpen Then 
        Filehandle.Close 
    End If 
    oFso.CopyFile FullFileName, Left(FullFileName,(Len(FullFileName)-1)), True 
        oFSo.DeleteFile FullFileName 
    Set FileHandle = oFSo.OpenTextFile(FullFileName, ForAppending, True)             
End Sub 
 
'--------------------------------------------------------- 
'Subroutine for creating the log file and instantiating the handle 
Private Sub Create() 
    If Debugmode Then On Error Goto 0 Else On Error Resume Next 
 
    Dim FileProperty,Logsize 
 
    If Not oFSo.FileExists(FullFileName) Then 
        oFSo.CreateTextFile(FullFileName) 
    End If     
        If Not bOpen Then 
               Set FileHandle = oFSo.OpenTextFile(FullFileName, ForAppending, True) 
        End if 
    Set FileProperty = oFSo.GetFile(FullFileName) 
                Logsize = FileProperty.size 
    If Logsize > MaxSize Then 
            Filehandle.WriteLine "\\\\\\\\\\File Size Reached//////////" 
        Call Rollover() 
    End If 
End Sub 
End Class