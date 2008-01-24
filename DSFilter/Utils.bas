Attribute VB_Name = "Utils"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Public Declare Function timeGetTime Lib "kernel32" () As Long 'elapsed time in msec since system boot
Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long


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
    
    Set g_debugLog = xBuildLogFile(sPath, "DSFilter-Debug.txt", App.Title & " v" & VersionString & " - Debug Log - Session : " & Now)
    Set g_errorLog = xBuildLogFile(sPath, "DSFilter-Error.txt", App.Title & " v" & VersionString & " - Error Log - Session : " & Now)

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

Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ BuildPluginLogFiles - " & Err.Description
    Resume Fin
End Sub

Public Sub xWriteMessageToFile(ByVal aTS As TextStream, Msg As String)
On Error GoTo ErrorHandler

    aTS.WriteLine ("[" & DS_FormatedTimeStamp & "]  " & Msg)
    
Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Core) ERROR @ xWriteMessageToFile - " & Err.Description
    Resume Fin
End Sub

Public Function FileExists(szFile As String) As Boolean
    FileExists = False
    On Error GoTo LExit
    FileExists = (Dir(szFile) <> "")
LExit:
End Function

'---------------------------------------------------------------------
'---------------------------------------------------------------------

Public Function DS_CreateOutputFile(ByVal sFileName As String) As Long
On Error GoTo ErrorHandler
    Dim lFileNum As Long
    
    lFileNum = FreeFile(0)
    Open sFileName For Output As #lFileNum
    
Fin:
    DS_CreateOutputFile = lFileNum
    Exit Function
ErrorHandler:
    lFileNum = -1
    DS_ErrorMsgBox "DS_CreateOutputFile(" & sFileName & ") : " & Err.Description
    GoTo Fin
End Function

Public Function DS_CloseOutputFile(ByVal lFileNum As Long) As Boolean
On Error GoTo ErrorHandler

    Close #lFileNum
    DS_CloseOutputFile = True
    
Fin:
    Exit Function
ErrorHandler:
    DS_CloseOutputFile = False
    GoTo Fin
End Function

Public Function DS_WriteToFile(ByVal lFileNum As Long, ByVal sText As String) As Boolean
On Error GoTo ErrorHandler

    Print #lFileNum, DS_FormatedTimeStamp & " " & sText
    DS_WriteToFile = True

Fin:
    Exit Function
ErrorHandler:
    DS_WriteToFile = False
    DS_ErrorMsgBox "DS_WriteToFile : " & Err.Description
    GoTo Fin
End Function

Public Function DS_FormatedTimeStamp() As String
    DS_FormatedTimeStamp = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Function

Public Function DS_FilterVersion() As String
    DS_FilterVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function DS_MsgBox(ByVal sMsg As String)
    MsgBox sMsg, , FILTER_FULL_NAME
End Function

Public Function DS_ErrorMsgBox(ByVal sMsg As String)
    DS_MsgBox "Error - " & sMsg
End Function

Public Sub myDebug(ByVal sMsg As String)
On Error Resume Next
    Call g_Filter.LogDebugMsg(sMsg)
End Sub

Public Sub myError(ByVal sMsg As String)
On Error Resume Next
    Call g_Filter.LogErrorMsg(sMsg)
End Sub
 
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

Public Function dmg2vuln(ByVal DamageType As eDamageType) As Integer
    Select Case DamageType
        Case DMG_SLASHING
            dmg2vuln = FL_SLASHING
        Case DMG_PIERCING
            dmg2vuln = FL_PIERCING
        Case DMG_BLUDGEONING
            dmg2vuln = FL_BLUDGEONING
        Case DMG_FIRE
            dmg2vuln = FL_FIRE
        Case DMG_COLD
            dmg2vuln = FL_COLD
        Case DMG_ACID
            dmg2vuln = FL_ACID
        Case DMG_LIGHTNING
            dmg2vuln = FL_LIGHTNING
        Case Else
            myError "dmg2vuln: unknown damage type " & DamageType
            dmg2vuln = -1
    End Select
End Function

Public Function ToRaceString(ByVal race As Long) As String
    
    Select Case race
    Case 1
       ToRaceString = "Aluvian"
    Case 2
       ToRaceString = "Gharu'ndim"
    Case 3
       ToRaceString = "Sho"
    Case 4
       ToRaceString = "Viamontian"
    End Select
    
End Function

Public Function FormatXp(Val As Variant, Optional ShowXpTag As Boolean = True, Optional szSeparator As String = ",") As String
    Dim Tag As String
    Dim tmp As String
    Dim sLen As Integer
    Dim i As Integer
    
    If ShowXpTag Then
        Tag = " xp"
    Else
        Tag = ""
    End If
    
    
    If Val = 0 Then
        FormatXp = "0" & Tag
    Else
        tmp = Format(Val, "##")
        sLen = Len(tmp)
        FormatXp = ""
        For i = sLen To 1 Step -3
            If i > 3 Then
                FormatXp = szSeparator & Mid(tmp, i - 2, 3) & FormatXp
            Else
                FormatXp = Mid(tmp, 1, i) & FormatXp
            End If
        Next i
        
        FormatXp = FormatXp & Tag
    End If
    
End Function

'FROM SkunkWork
' We've been given an unsigned DWORD potentially exceeding 2147483647.
' To VB it looks like a signed long.  Convert to unsigned Double.
Public Function UDblFromDw(ByVal dw As Long) As Double
    
    UDblFromDw = dw
    If (dw And &H80000000) <> 0 Then
        UDblFromDw = UDblFromDw + 4294967296#
    End If
    
End Function

'Double from QWORD
Public Function DblFromQw(qw() As Byte) As Double

    Dim d As Double, i As Long
    
    d = 0
    For i = 7 To 0 Step -1
        d = d * &H100 + qw(i)
    Next i

    DblFromQw = d
End Function


Public Function TimeRemaining(ByVal fTimer As Long) As Long
    TimeRemaining = fTimer - g_Time
End Function

Public Function TimerExpired(ByVal fTimer As Long) As Boolean
    TimerExpired = ((fTimer - g_Time) <= 0)
End Function
