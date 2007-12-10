Attribute VB_Name = "shUtils"
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                 SHARED MODULE                       [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                                                     [[
'[[             Utility Functions                       [[
'[[                                                     [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Option Explicit

Public Function Extract(ByVal Msg As String, PosStart As Integer, PosEnd As Integer) As String
Dim sRet As String
Dim iLen As Integer

    'default
    sRet = ""
    iLen = Len(Msg)
    
    If (PosStart < 1) Or (PosStart > iLen) Or (PosEnd < PosStart) Or (PosEnd > iLen) Then
        Exit Function
    Else
        sRet = Mid(Msg, PosStart, PosEnd - PosStart + 1)
    End If
    Extract = sRet
End Function

Public Function FirstWord(ByVal strSource As String) As String
Dim Pos As Integer
Dim sRet As String

    'default ret val
    sRet = ""
    
    Pos = InStr(1, strSource, " ") 'look for space
    If Pos < 1 Then 'single-word string
        sRet = strSource
    Else
        sRet = Mid(strSource, 1, Pos - 1)
    End If

    FirstWord = sRet
    
End Function

Public Function SameText(ByVal str1 As String, ByVal str2 As String) As Boolean
    If LCase(Trim(str1)) = LCase(Trim(str2)) Then
        SameText = True
    Else
        SameText = False
    End If
End Function

Public Function Valid(ByVal obj As Variant) As Boolean
    Valid = Not (obj Is Nothing)
End Function

'Replace any strRemove character/sub-string from the strSource by strReplace
Public Sub CleanString(ByRef strSource As String, ByVal strRemove As String, Optional ByVal strReplace As String = "")
    strSource = Replace(strSource, strRemove, strReplace)
End Sub

Public Sub AddFlag(ByRef lVar, ByVal lFlag As Long)
    lVar = lVar Or lFlag
End Sub

Public Sub RemoveFlag(ByRef lVar, ByVal lFlag As Long)
    lVar = lVar And (Not lFlag)
End Sub

Public Function HasFlag(ByRef lVar As Long, ByVal lFlag As Long) As Boolean
    HasFlag = (lVar And lFlag)
End Function

Public Sub CondAddFlag(ByVal lFrom As Long, ByVal lFlag As Long, ByRef lDest As Long)
    If HasFlag(lFrom, lFlag) Then Call AddFlag(lDest, lFlag)
End Sub

Public Function GetDataFolder() As String
   On Error GoTo GenericFolder
   Dim PathName  As String
   Dim strPath   As String
   Dim lngReturn As Long
   Dim ReturnVal As Long
   
   strPath = String(260, 0)
   lngReturn = SHGetFolderPath(0, CSIDL_PERSONAL, 0, &H0, strPath)
   PathName = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)
   GetDataFolder = PathName & "\LifeTankX"
    
   Exit Function
GenericFolder:
   If Err.Number = 453 Then GetDataFolder = App.Path
End Function
