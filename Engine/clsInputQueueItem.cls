VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputQueueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public InputType As Integer
Public Pos As New clsPoint  'Point to click, or start point
Public Pos2 As New clsPoint 'End point
Public Delay As Double      'Delay in msec
Public KeyCode As Long      'Key asci code
Public Text As String

Private Sub Class_Initialize()
    InputType = -1
    Delay = 0
    KeyCode = 0
End Sub

Private Sub Class_Terminate()
    Set Pos = Nothing
    Set Pos2 = Nothing
End Sub

