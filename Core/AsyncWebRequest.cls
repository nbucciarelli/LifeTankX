VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncWebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Event OnReadyStateChange(ByVal iNewState As Integer)
Public Event OnRequestComplete(ByVal sData As String)
Public Event OnRequestFailed()

Private m_HttpReq As MSXML2.XMLHTTP40

Private Sub Class_Initialize()
    Call Randomize(Time)
    Set m_HttpReq = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_HttpReq = Nothing
End Sub

Public Function FunctionReadyStateChange()
Attribute FunctionReadyStateChange.VB_UserMemId = 0
Attribute FunctionReadyStateChange.VB_MemberFlags = "40"
    'Debug.Print "AsyncRequest ReadyStateChange : " & m_HttpReq.readyState
    If m_HttpReq.readyState = 4 Then
        RaiseEvent OnRequestComplete(m_HttpReq.responseText)
    End If
End Function

Public Function Send(ByVal sUrl As String) As Boolean
    Set m_HttpReq = New MSXML2.XMLHTTP40
    m_HttpReq.onreadystatechange = Me
    Call m_HttpReq.Open("GET", sUrl & "&random=" & Str(Rnd), True)
    Call m_HttpReq.setRequestHeader("Cache-control", "no-cache")    'Disable caching
    Call m_HttpReq.Send
End Function
