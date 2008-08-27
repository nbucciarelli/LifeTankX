VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const FILE_DEBUGLOG = "DSFilter_Debug.log"
Private Const FILE_ERROR = "DSFilter_Error.log"

Private Const DEBUG_ME = False

Private m_lFileDebugLog As Long
Private m_lFileErrorLog As Long

Public bDebugMode As Boolean

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    m_lFileDebugLog = -1
    m_lFileErrorLog = -1
    bDebugMode = True
    Call CreateLogFiles
    
Fin:
    Exit Sub
ErrorHandler:
    DS_ErrorMsgBox "clsLog.Class_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    DS_CloseOutputFile m_lFileDebugLog
    DS_CloseOutputFile m_lFileErrorLog
    ' Close all the log files
    Call xCloseLogFiles
End Sub

Private Function CreateLogFiles() As Long
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    
    'Tighter log files, but breaks dual log
    'Call Utils.xBuildPluginLogFiles(GetDataFolder)
    
    m_lFileDebugLog = DS_CreateOutputFile(GetDataFolder & "\Logs\" & FILE_DEBUGLOG)
    bOk = (m_lFileDebugLog <> -1)
    
    If bOk Then
        m_lFileErrorLog = DS_CreateOutputFile(GetDataFolder & "\Logs\" & FILE_ERROR)
        bOk = (m_lFileErrorLog <> -1)
    End If
    
    If bOk Then
        DS_WriteToFile m_lFileDebugLog, "=================================================="
        DS_WriteToFile m_lFileDebugLog, FILTER_FULL_NAME & " v" & DS_FilterVersion & " - Debug Log File - Session: [" & DS_FormatedTimeStamp & "]"
        DS_WriteToFile m_lFileDebugLog, "=================================================="
        
        DS_WriteToFile m_lFileErrorLog, "=================================================="
        DS_WriteToFile m_lFileErrorLog, FILTER_FULL_NAME & " v" & DS_FilterVersion & " - Error Log File - Session: [" & DS_FormatedTimeStamp & "]"
        DS_WriteToFile m_lFileErrorLog, "=================================================="
    End If
    
Fin:
    CreateLogFiles = bOk
    Exit Function
ErrorHandler:
    bOk = False
    GoTo Fin
End Function

Public Sub DebugMsg(ByVal sMsg As String)
On Error GoTo ErrorHandler

    If DEBUG_ME Or bDebugMode Then
        If Valid(g_debugLog) Then
            Call xWriteMessageToFile(g_debugLog, sMsg)
        ElseIf m_lFileDebugLog <> -1 Then
            Call DS_WriteToFile(m_lFileDebugLog, sMsg)
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    Resume Fin
End Sub


Public Sub ErrorMsg(ByVal sMsg As String)
On Error GoTo ErrorHandler

    If Valid(g_errorLog) Then
        Call xWriteMessageToFile(g_errorLog, sMsg)
    ElseIf m_lFileErrorLog <> -1 Then
        DebugMsg "<<ERROR>> " & sMsg
        Call DS_WriteToFile(m_lFileErrorLog, sMsg)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    Resume Fin
End Sub

