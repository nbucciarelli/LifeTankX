VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIMsgLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========================================
'User Interface for tab : Messages Logger
'===========================================
Option Explicit

Private Const INTERFACE_NAME = "UIMsgLog"

Public WithEvents lstMessageLog As DecalControls.list
Attribute lstMessageLog.VB_VarHelpID = -1
Public chkMonitorTells As DecalControls.CheckBox
Attribute chkMonitorTells.VB_VarHelpID = -1
Public chkMonitorPublic As DecalControls.CheckBox
Attribute chkMonitorPublic.VB_VarHelpID = -1
Public chkMonitorFellowship As DecalControls.CheckBox
Attribute chkMonitorFellowship.VB_VarHelpID = -1

Public WithEvents btnClearMessageLog As DecalControls.PushButton
Attribute btnClearMessageLog.VB_VarHelpID = -1

Public Enum eMessageTypes
    TYPE_NONE
    TYPE_PUBLIC_CHAT
    TYPE_FELLOWSHIP
    TYPE_TELL
End Enum

Private Sub Class_Terminate()
    Call Unload
End Sub

Public Function Init() As Boolean
On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    MyDebug INTERFACE_NAME & ".Init() -- Begin"
    
    'Create Controls
    
    Set lstMessageLog = g_MainView.Control("lstMessageLog")
    Set chkMonitorTells = g_MainView.Control("chkMonitorTells")
    Set chkMonitorPublic = g_MainView.Control("chkMonitorPublic")
    Set chkMonitorFellowship = g_MainView.Control("chkMonitorFellowship")
    Set btnClearMessageLog = g_MainView.Control("btnClearMessageLog")
    
    'Controls default settings
    Call lstMessageLog.Clear
    
    Init = True
    MyDebug INTERFACE_NAME & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage INTERFACE_NAME & ".Init - " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue() As Boolean
On Error GoTo Error_Handler
    
    MyDebug "[" & INTERFACE_NAME & "] Loading controls values"
    
    '***************************************************************************
    Call lstMessageLog.Clear
    chkMonitorTells.Checked = g_Settings.GetValue("chkMonitorTells", True)
    chkMonitorPublic.Checked = g_Settings.GetValue("chkMonitorPublic", True)
    chkMonitorFellowship.Checked = g_Settings.GetValue("chkMonitorFellowship", True)
    '***************************************************************************
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage INTERFACE_NAME & ".LoadControlsValue - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & INTERFACE_NAME & "] Saving controls settings"
    
    With g_Settings.Profile.Config
        .SaveCheckbox chkMonitorTells, "chkMonitorTells"
        .SaveCheckbox chkMonitorPublic, "chkMonitorPublic"
        .SaveCheckbox chkMonitorFellowship, "chkMonitorFellowship"
    End With
    
    bRet = True
Fin:
    SaveControlSettings = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage INTERFACE_NAME & ".SaveControlSettings - " & Err.Description
    Resume Fin
End Function

Public Function Unload() As Boolean
On Error GoTo Error_Handler
    
    MyDebug INTERFACE_NAME & ".Unload() -- Begin"
    'Unload initialization
    Unload = False

    Set lstMessageLog = Nothing
    Set chkMonitorTells = Nothing
    Set chkMonitorPublic = Nothing
    Set chkMonitorFellowship = Nothing
    Set btnClearMessageLog = Nothing

    'Unload complete
    Unload = True
    MyDebug INTERFACE_NAME & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage INTERFACE_NAME & ".Unload - " & Err.Description
    Resume Fin
End Function

Private Sub btnClearMessageLog_Accepted(ByVal nID As Long)
    PrintMessage "Clearing Logs..."
    Call lstMessageLog.Clear
End Sub


Public Sub AddLog(MsgType As eMessageTypes, SenderName As String, Message As String)
On Error GoTo Error_Handler

    Dim i As Integer
    Dim lineColor As Long
    Dim szMessageType As String

    If SenderName = "" Then
        Exit Sub
    End If
    
    'MyDebug "Adding Log to message log..."
    
    Select Case MsgType
        Case TYPE_TELL
            lineColor = vbYellow
            szMessageType = "T"
        Case TYPE_PUBLIC_CHAT
            lineColor = vbWhite
            szMessageType = "P"
        Case TYPE_FELLOWSHIP
            lineColor = vbGreen
            szMessageType = "F"
        Case Else
            lineColor = vbWhite
    End Select

    i = lstMessageLog.AddRow
    lstMessageLog.Data(0, i) = Time
    lstMessageLog.Color(0, i) = &H5500DD
    
    'lstMessageLog.Data(1, i) = szMessageType
    'lstMessageLog.Color(1, i) = lineColor
    
    lstMessageLog.Data(1, i) = SenderName
    lstMessageLog.Color(1, i) = &HDD4400
    
    lstMessageLog.Data(2, i) = Message
    lstMessageLog.Color(2, i) = lineColor
    
    lstMessageLog.ScrollPosition = i
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage INTERFACE_NAME & ".AddLog(" & CInt(MsgType) & ", " & SenderName & ", " & Message
    Resume Fin
End Sub

Private Sub lstMessageLog_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)

    Select Case nX
        
        'Clicked on SenderName...
        Case 1
            If Not g_Hooks.ChatState Then
                g_Core.SendTextToConsole "/tell " & lstMessageLog.Data(1, nY) & ", ", , False
            End If
        
        'Clicked on Message...
        Case 2
            PrintMessage " [" & lstMessageLog.Data(0, nY) & "] " & lstMessageLog.Data(1, nY) & " : " & lstMessageLog.Data(2, nY)
        
    End Select
End Sub
