VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Container for the macro timers
Private m_colTimers As New Collection
Private m_lIndexCounter As Long

Private Sub Class_Initialize()
    m_lIndexCounter = 0
End Sub

Private Sub Class_Terminate()
    m_lIndexCounter = 0
    Set m_colTimers = Nothing
End Sub

Public Function CreateTimer() As clsTimer
On Error GoTo ErrorHandler
    
    Set CreateTimer = New clsTimer
    'CreateTimer.Index = m_lIndexCounter
    Call m_colTimers.Add(CreateTimer)
    'm_lIndexCounter = m_lIndexCounter + 1
    
Fin:
    Exit Function
ErrorHandler:
    Set CreateTimer = Nothing
    PrintErrorMessage "clsTimers.CreateTimer"
    Resume Fin
End Function

Public Sub Update()
On Error GoTo ErrorHandler
    
    Dim objTimer As clsTimer
    
    For Each objTimer In m_colTimers
        If Valid(objTimer) And objTimer.Enabled Then
            If objTimer.Expired Then Call objTimer.TriggerTimeout
        End If
    Next objTimer
    
Fin:
    Set objTimer = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsTimers.Update"
    Resume Fin
End Sub

Public Sub ResetAll()
On Error GoTo ErrorHandler

    Dim objTimer As clsTimer
    
    For Each objTimer In m_colTimers
        Call objTimer.Reset
    Next objTimer
    
Fin:
    Set objTimer = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsTimers.ResetAll"
    Resume Fin
End Sub



