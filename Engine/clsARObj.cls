VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsARObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class holds the AutoResponse info

Option Explicit

Public ar_guid As Long
Public ar_delay As Double
Public ar_regex As String
Public ar_response As String
Public ar_used As Boolean
Public ar_usedTime As Double
Public ar_talkTime As Double
Public ar_matchString As String
Public ar_repeat As Integer
Public ar_numDone As Integer
Public ar_extraResponse As String
Public SetClear As Boolean


Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    ar_guid = -1
    ar_delay = 0
    ar_regex = ""
    ar_response = ""
    ar_used = False
    ar_usedTime = 0
    ar_talkTime = 0
    ar_repeat = 1
    ar_numDone = 0
    ar_matchString = ""
    ar_extraResponse = ""

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsARObj_Initialize - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub Class_Terminate()
   ' Hmm
End Sub

