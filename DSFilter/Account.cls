VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colChars As Collection    'CharName -> acAccountChar

Private Sub Class_Initialize()
    Set m_colChars = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colChars = Nothing
End Sub

Public Function Count() As Long
   Count = m_colChars.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colChars.[_NewEnum]
End Function

Public Function Exists(ByVal sCharName As String) As Boolean
On Error GoTo ErrorHandler
    Dim objChar As acAccountChar
    
    Set objChar = m_colChars(sCharName)
    Exists = True

Fin:
    Set objChar = Nothing
    Exit Function
ErrorHandler:
    Exists = False
    Resume Fin
End Function

Friend Function Add(ByVal sCharName As String, Optional ByVal lCharGUID As Long) As acAccountChar
On Error GoTo ErrorHandler
    Dim objChar As New acAccountChar

    If Exists(sCharName) Then
        myDebug "Account.Add(" & sCharName & ") - character already in collection - Ignoring add"
        Set Add = Nothing
    Else
        objChar.Name = sCharName
        If Not IsEmpty(lCharGUID) Then
            objChar.GUID = lCharGUID
        End If
        Call m_colChars.Add(objChar, sCharName)
        myDebug "Account: added character " & objChar.Name & " [" & objChar.GUID & "]"
        Set Add = objChar
    End If
        
Fin:
    Set objChar = Nothing
    Exit Function
ErrorHandler:
    myError "acAccount.Add - " & Err.Description
    Set Add = Nothing
    Resume Fin
End Function


