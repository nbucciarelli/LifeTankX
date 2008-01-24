VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colDatas As Collection

Private Sub Class_Initialize()
    Set m_colDatas = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colDatas = Nothing
End Sub

Public Property Get Count() As Long
   Count = m_colDatas.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colDatas.[_NewEnum]
End Function

Public Sub Clear()
    Set m_colDatas = New Collection
End Sub

Public Function Item(ByVal lKey As Long) As DataItem
On Error GoTo ErrorHandler
    
    Dim oData As DataItem
    If Not Exists(lKey, oData) Then
        Set oData = Nothing
    End If
    
Fin:
    Set Item = oData
    Set oData = Nothing
    Exit Function
ErrorHandler:
    Set oData = Nothing
    myError "colDatas.Item - " & Err.Description
    Resume Fin
End Function

Public Function AddDataItem(ByVal oData As DataItem) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oData) Then
        If Not Exists(oData.Key) Then
            Call m_colDatas.Add(oData, CStr(oData.Key))
            bRet = True
        End If
    End If
    
Fin:
    AddDataItem = bRet
    Exit Function
ErrorHandler:
    myError "colDatas.AddDataItem - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function Add(ByVal lKey As Long, ByVal sVal As String, Optional ByVal lExtraFlag As Long = 0) As DataItem
On Error GoTo ErrorHandler
    
    If Not Exists(lKey) Then
        Dim oData As New DataItem
        oData.Key = lKey
        oData.Val = sVal
        oData.Flag = lExtraFlag
        
        If AddDataItem(oData) Then
            Set Add = oData
        Else
            Set Add = Nothing
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    myError "colDatas.Add(" & lKey & ", " & sVal & ") - " & Err.Description & " - line: " & Erl
    Set Add = Nothing
    Resume Fin
End Function

Public Function Exists(ByVal lKey As Long, Optional ByRef oDataOut As DataItem) As Boolean
On Error GoTo ErrorHandler
    Set oDataOut = m_colDatas(CStr(lKey))
    Exists = True
Fin:
    Exit Function
ErrorHandler:
    Set oDataOut = Nothing
    Exists = False
    Resume Fin
End Function

'Returns empty string if matching value not found
Public Function GetValue(ByVal lKey As Long) As String
Attribute GetValue.VB_UserMemId = 0
On Error GoTo ErrorHandler
    Dim oData As DataItem
    If Exists(lKey, oData) Then
        GetValue = oData.Val
    Else
        GetValue = ""
    End If
Fin:
    Set oData = Nothing
    Exit Function
ErrorHandler:
    GetValue = ""
    Resume Fin
End Function

'Returns 0 if matching Key not found
Public Function GetKey(ByVal sVal As String) As Long
On Error GoTo ErrorHandler
    Dim lRet As Long
    
    Dim oData As DataItem
    For Each oData In m_colDatas
        If SameText(sVal, oData.Val) Then
            lRet = oData.Key
            GoTo Fin
        End If
    Next oData
    
Fin:
    GetKey = lRet
    Set oData = Nothing
    Exit Function
ErrorHandler:
    lRet = 0
    Resume Fin
End Function

Public Function GetFlag(ByVal lKey As Long) As Long
On Error GoTo ErrorHandler

    Dim oData As DataItem
    If Exists(lKey, oData) Then
        GetFlag = oData.Flag
    Else
        GetFlag = 0
    End If

Fin:
    Exit Function
ErrorHandler:
    GetFlag = 0
    Resume Fin
End Function

