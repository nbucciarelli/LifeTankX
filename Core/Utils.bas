Attribute VB_Name = "Utils"
Option Explicit

'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
Public Const CSIDL_MYDOCUMENTS = &HC             'My Documents
Public Const CSIDL_APPDATA = &H1A           'Users App Data?
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_PERSONAL = &H5

Public Function GetDataFolder()
   On Error GoTo GenericFolder
   Dim PathName  As String
   Dim strPath   As String
   Dim lngReturn As Long
   Dim ReturnVal As Long
   
   'PathName = String(260, 0)
   'ReturnVal = SHGetFolderPath(g_PluginSite.hwnd, CSIDL_APPDATA, 0, 0, PathName)
   'PathName = Left(PathName, InStr(PathName, vbNullChar) - 1)
   'GetDataFolder = PathName & "\LifeTankX"
   
    strPath = String(260, 0)
    lngReturn = SHGetFolderPath(0, CSIDL_PERSONAL, 0, &H0, strPath)
    PathName = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)
    GetDataFolder = PathName & "\LifeTankX"
    
   Exit Function
GenericFolder:
   If Err.Number = 453 Then GetDataFolder = App.Path
End Function

'Builds/Overwrites the file filename and inserts the message Title in the first line
Public Function xBuildLogFile(ByVal sPath As String, ByVal sFileName As String, sTitle As String) As TextStream
On Error GoTo ErrorHandler
    
    Dim oFS As New FileSystemObject
    Dim ts As TextStream
    
    Dim lngFileNr As Integer
    Dim sLine As String
    Dim oldFile As String
    
    'MyDebug "xBuildLogFile: sPath: " & sPath & "  sFileName: " & sFileName
    
    sFileName = sPath & "\" & sFileName
    oldFile = sFileName & "-old.txt"
    
    ' Make a backup of each one
    If FileExists(sFileName) Then
        If FileExists(oldFile) Then
            Call oFS.DeleteFile(oldFile, True)
        End If
        Call oFS.MoveFile(sFileName, oldFile)
    End If
    
    'MyDebug "Building " & sFileName
    
    Set ts = oFS.CreateTextFile(sFileName, False)
    
    If Not Valid(ts) Then
        ' Hmm, file is either aleady open or can't be written to, so bail!
        Set xBuildLogFile = Nothing
        Exit Function
    End If
    
    ts.WriteLine ("====================================================================")
    ts.WriteLine (sTitle)
    ts.WriteLine ("====================================================================")
    
    ' Return the new filehandle
    Set xBuildLogFile = ts
    
Fin:
    Exit Function
ErrorHandler:
    MsgBox "(Core) ERROR @ xBuildLogFile - " & Err.Description
    Resume Fin
End Function

Public Sub xBuildPluginLogFiles(ByVal sPath As String)
On Error GoTo ErrorHandler

    Dim VersionString As String
    
    VersionString = App.Major & "." & App.Minor & "." & App.Revision
    
    'MyDebug "xBuildPluginLogFiles: " & sPath & " " & FILE_DEBUG_LOG & " " & App.Title & " v" & VersionString & " - Debug Log - Session : " & Now
    
    Set g_debugLog = xBuildLogFile(sPath, FILE_DEBUG_LOG, App.Title & " v" & VersionString & " - Debug Log - Session : " & Now)
    Set g_errorLog = xBuildLogFile(sPath, FILE_ERROR_LOG, App.Title & " v" & VersionString & " - Error Log - Session : " & Now)
    Set g_chatLog = xBuildLogFile(sPath, FILE_CHAT_LOG, App.Title & " v" & VersionString & " - Chat Log - Session : " & Now)
    Set g_eventLog = xBuildLogFile(sPath, FILE_EVENTS_LOG, App.Title & " v" & VersionString & " - Events Log - Session : " & Now)

Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ BuildPluginLogFiles - " & Err.Description
    Resume Fin
End Sub

Public Sub xCloseLogFiles()
On Error GoTo ErrorHandler

    If Valid(g_debugLog) Then g_debugLog.Close
    If Valid(g_errorLog) Then g_errorLog.Close
    If Valid(g_chatLog) Then g_chatLog.Close
    If Valid(g_eventLog) Then g_eventLog.Close

Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ BuildPluginLogFiles - " & Err.Description
    Resume Fin
End Sub

Public Sub xWriteMessageToFile(ByVal aTS As TextStream, Msg As String)
On Error GoTo ErrorHandler

    If Valid(g_PluginSite) And Valid(g_Core) Then
        If Valid(g_Core.Engine) Then
            If Not g_Core.Engine.DisableLogs Then
                aTS.WriteLine ("[" & CurrentTime & "]  " & Msg)
            End If
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ xWriteMessageToFile - " & Err.Description
    Resume Fin
End Sub



'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------


'Builds/Overwrites the file filename and inserts the message Title in the first line
Public Sub BuildLogFile(ByVal sFileName As String, sTitle As String)
On Error GoTo ErrorHandler
    
    Dim lngFileNr As Long, sLine As String
    sFileName = GetDataFolder & "\" & sFileName
    MyDebug "Building " & sFileName
    lngFileNr = FreeFile(0)
    Open sFileName For Output As #lngFileNr
        Print #lngFileNr, "===================================================================="
        Print #lngFileNr, sTitle
        Print #lngFileNr, "===================================================================="
    Close #lngFileNr
    
Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ BuildLogFile (" & sFileName & ") - " & Err.Description
    Resume Fin
End Sub

Public Sub BuildPluginLogFiles()
On Error GoTo ErrorHandler

    Dim VersionString As String
    
    'Macro.LogOutputPath = OutputFolderPath
    VersionString = App.Major & "." & App.Minor & "." & App.Revision
    
    BuildLogFile GetDebugLogPath, App.Title & " v" & VersionString & " - Debug Log - Session : " & Now
    BuildLogFile GetErrorLogPath, App.Title & " v" & VersionString & " - Error Log - Session : " & Now
    BuildLogFile GetChatLogPath, App.Title & " v" & VersionString & " - Chat Log - Session : " & Now
    BuildLogFile GetEventsLogPath, App.Title & " v" & VersionString & " - Events Log - Session : " & Now

Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ BuildPluginLogFiles - " & Err.Description
    Resume Fin
End Sub

Public Function GetEventsLogPath(Optional ByVal bPath As String = PATH_LOGS) As String
    GetEventsLogPath = bPath & "\" & FILE_EVENTS_LOG
End Function

Public Function GetDebugLogPath(Optional ByVal bPath As String = PATH_LOGS) As String
    GetDebugLogPath = bPath & "\" & FILE_DEBUG_LOG
End Function

Public Function GetErrorLogPath(Optional ByVal bPath As String = PATH_LOGS) As String
    GetErrorLogPath = bPath & "\" & FILE_ERROR_LOG
End Function

Public Function GetChatLogPath(Optional ByVal bPath As String = PATH_LOGS) As String
    GetChatLogPath = bPath & "\" & FILE_CHAT_LOG
End Function

Public Sub MyDebug(strMsg As String, Optional bSilent As Boolean = False)
On Error GoTo ErrorHandler

    If Valid(g_debugLog) Then
        xWriteMessageToFile g_debugLog, strMsg
    Else
        WriteMessageToFile GetDebugLogPath, strMsg
    End If
    
    'If (Not bSilent) And Valid(g_PluginSite) And Valid(g_Core) Then
    '    If Valid(g_Core.Engine) Then
    '        If g_Core.Engine.DebugMode Then
    '            'g_Hooks.AddChatTextRaw "[LTx Dbg] ", COLOR_PURPLE, 0
    '            g_Hooks.AddChatText "[LTx] " & strMsg, COLOR_PURPLE, 0
    '        End If
    '    End If
    'End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "(Core) MyDebug - " & Err.Description & " - line: " & Erl & " [msg: " & strMsg & "]"
    Resume Fin
End Sub

Public Sub LogEvent(strMsg As String, Optional bSilent As Boolean = False)
On Error GoTo ErrorHandler

    If Valid(g_eventLog) Then
        xWriteMessageToFile g_eventLog, strMsg
    Else
        WriteMessageToFile GetEventsLogPath, strMsg
    End If
    
    If Not bSilent Then
        PrintMessage "<EVENT> " & strMsg, COLOR_BRIGHT_YELLOW
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "LogEvent - " & Err.Description
    Resume Fin
End Sub

Public Sub PrintMessage(strMsg As String, Optional Color As Long = COLOR_CYAN)
On Error GoTo ErrorHandler

    If Valid(g_debugLog) Then
        xWriteMessageToFile g_debugLog, strMsg
    Else
        WriteMessageToFile GetDebugLogPath, strMsg
    End If
    
    If Valid(g_PluginSite) Then
        'g_Hooks.AddChatTextRaw "[ " & PLUG_NAME & " ] ", COLOR_BRIGHT_YELLOW, 0
        g_Hooks.AddChatText "[" & PLUG_NAME & "] " & strMsg, Color, 0
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "(Core) PrintMessage - " & Err.Description & " - line: " & Erl & " [msg: " & strMsg & "]"
    Resume Fin
End Sub

Public Sub PrintWarning(strMsg As String)
On Error GoTo ErrorHandler

    If Valid(g_debugLog) Then
        xWriteMessageToFile g_debugLog, strMsg
    Else
        WriteMessageToFile GetDebugLogPath, strMsg
    End If
    
    If Valid(g_errorLog) Then
        xWriteMessageToFile g_errorLog, strMsg
    Else
        WriteMessageToFile GetErrorLogPath, strMsg
    End If

    If Valid(g_PluginSite) Then
        'g_Hooks.AddChatTextRaw "[ " & PLUG_NAME & " ] ", COLOR_BRIGHT_YELLOW, 0
        g_Hooks.AddChatText "[" & PLUG_NAME & "] WARNING : " & strMsg, COLOR_PURPLE, 0
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "(Core) PrintWarning - " & Err.Description & " - line: " & Erl & " [msg: " & strMsg & "]"
    Resume Fin
End Sub

Public Sub PrintErrorMessage(strMsg As String, Optional ByVal bShowErrorTag As Boolean = True)
On Error GoTo ErrorHandler

    If Valid(g_errorLog) Then
        xWriteMessageToFile g_errorLog, strMsg
    Else
        WriteMessageToFile GetErrorLogPath, strMsg
    End If
    
    If Valid(g_debugLog) Then
        xWriteMessageToFile g_debugLog, "<< ERROR >> " & strMsg
    Else
        WriteMessageToFile GetDebugLogPath, "<< ERROR >> " & strMsg
    End If
    
    If Valid(g_PluginSite) Then
        'g_Hooks.AddChatTextRaw "[ " & PLUG_NAME & " ] ", COLOR_BRIGHT_YELLOW, 0
        g_Hooks.AddChatText "[" & PLUG_NAME & "] " & strMsg, COLOR_RED, 0
    End If

Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ PrintErrorMessage - " & Err.Description & " - line: " & Erl & " [msg: " & strMsg & "]"
    Resume Fin
End Sub

Public Sub WriteMessageToFile(FileName As String, Msg As String)
On Error GoTo ErrorHandler

    Dim lngFileNr As Long, sLine As String
    lngFileNr = FreeFile(0)
    Open GetDataFolder & "\" & FileName For Append As #lngFileNr
        Print #lngFileNr, "[" & CurrentTime & "] ", Msg
    Close #lngFileNr
    
Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ WriteMessageToFile - " & Err.Description
    Resume Fin
End Sub


'Convert the XML file content (the user-interface) to a string
'So it can be stored in ViewShema
Public Function FileToString(sFile As String) As String
On Error GoTo ErrorHandler

  Dim lngFileNr As Long, sLine As String
  
  FileToString = ""
  lngFileNr = FreeFile(0)
  ' Here I'm opening from same directory that scribe.dll is installed,
  ' but you can open from anywhere you'd like.
  Open GetDataFolder & "\" & sFile For Input As #lngFileNr
  Do Until EOF(lngFileNr)
    Line Input #lngFileNr, sLine
    FileToString = FileToString & sLine
  Loop
  Close #lngFileNr
  
Fin:
    Exit Function
ErrorHandler:
    FileToString = ""
    PrintErrorMessage "FileToString(" & sFile & ")"
    Resume Fin
End Function

Public Function BoolToInteger(BoolValue As Boolean) As String
    If BoolValue = True Then
        BoolToInteger = 1
    Else
        BoolToInteger = 0
    End If
End Function

Public Function GetPercent(ByVal Source As Long, ByVal Percentage As Integer) As Long
    GetPercent = ((CLng(Percentage) * Source) / 100)
End Function

Public Function GetSquareRange(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
    GetSquareRange = ((x2 - x1) * (x2 - x1)) + ((y2 - y1) * (y2 - y1)) + ((z2 - z1) * (z2 - z1))
End Function

Public Function GetRange(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
    GetRange = Sqr(GetSquareRange(x1, y1, z1, x2, y2, z2))
End Function

Public Function GetIntegerMinutesFromSeconds(TimeInSeconds As Long) As Integer
Dim SecondsLeft As Long
Dim NumMinutes As Integer
    
    NumMinutes = 0
    SecondsLeft = TimeInSeconds - 60
    While (SecondsLeft > 0)
        NumMinutes = NumMinutes + 1
        SecondsLeft = SecondsLeft - 60
    Wend
    
    GetIntegerMinutesFromSeconds = NumMinutes
    
End Function

'0 based index
Public Sub BuildArgsList(ByVal Data As String, ByRef Args() As String, ByRef NumArgs As Integer, Optional TokenChar As String = " ")
Dim Pos As Integer
Dim DataCopy As String
Dim CurArg As String 'current argument

    NumArgs = 0
    
    Data = Trim(Data)
    DataCopy = Data
    
    If Len(DataCopy) > 0 Then
        Do
            DataCopy = Trim(DataCopy)
            Pos = InStr(1, DataCopy, TokenChar)
            If Pos > 0 Then
                CurArg = Mid(DataCopy, 1, Pos - 1)
                ReDim Preserve Args(NumArgs + 1)
                Args(NumArgs) = CurArg
                NumArgs = NumArgs + 1
                'remove this argument from the data list, so we can extract the remaining ones
                DataCopy = Mid(DataCopy, Pos + Len(TokenChar))
            Else 'Last arg
                CurArg = Mid(DataCopy, 1)
                ReDim Preserve Args(NumArgs + 1)
                Args(NumArgs) = CurArg
                NumArgs = NumArgs + 1
            End If
        Loop While (Pos > 0 And DataCopy <> "" And CurArg <> "")
    End If
    
End Sub

Public Function CreateTimer() As clsTimer
On Error GoTo ErrorHandler

    If Valid(g_Timers) Then
        Set CreateTimer = g_Timers.CreateTimer
    Else
        Set CreateTimer = Nothing
    End If

Fin:
    Exit Function
ErrorHandler:
    Set CreateTimer = Nothing
    PrintErrorMessage "CreateTimer - " & Err.Description
    Resume Fin
End Function

'------------------------------------------------------------------
Public Function myDateFormat(ByVal aDate As Date) As String
    myDateFormat = Format(aDate, "dddd dd mmm yyyy")
End Function
