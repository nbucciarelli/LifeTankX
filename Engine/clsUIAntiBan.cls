VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIAntiBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
' Options>AntiBan panel
Private Const INTERFACE_NAME = "UIAntiBan"
'===========================================

Option Explicit

'Options
Public chkAlertOnIRC As DecalControls.CheckBox
Public chkEnableAdminDetection As DecalControls.CheckBox
Public chkEnableAutoResponse As DecalControls.CheckBox
Public chkEnableUnfriendlyDetect As DecalControls.CheckBox
Public chkLogoutOnDetect As DecalControls.CheckBox
Public chkReportUnfriendly As DecalControls.CheckBox
Public WithEvents btnStopAlarm As DecalControls.PushButton
Attribute btnStopAlarm.VB_VarHelpID = -1

Public chkReportOpenChat As DecalControls.CheckBox
Public chkReportTells As DecalControls.CheckBox


'SRV
'Todo

Private Sub Class_Terminate()
    Call Unload
End Sub

Public Function Init() As Boolean
On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    MyDebug INTERFACE_NAME & ".Init() -- Begin"
    
    'Create Controls
    Set chkAlertOnIRC = g_MainView.Control("chkAlertOnIRC")
    Set chkEnableAdminDetection = g_MainView.Control("chkEnableAdminDetection")
    Set chkEnableAutoResponse = g_MainView.Control("chkEnableAutoResponse")
    Set chkLogoutOnDetect = g_MainView.Control("chkLogoutOnDetect")
    Set chkReportUnfriendly = g_MainView.Control("chkReportUnfriendly")
    Set chkEnableUnfriendlyDetect = g_MainView.Control("chkEnableUnfriendlyDetect")
    Set btnStopAlarm = g_MainView.Control("btnStopAlarm")
    
    Set chkReportOpenChat = g_MainView.Control("chkReportOpenChat")
    Set chkReportTells = g_MainView.Control("chkReportTells")

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
    
    MyDebug "[" & INTERFACE_NAME & "] Loading controls settings"
    
    '***************************************************************************
    
    chkAlertOnIRC.Checked = g_Settings.GetValue("chkAlertOnIRC", True)
    chkEnableAdminDetection.Checked = g_Settings.GetValue("chkEnableAdminDetection", True)
    chkEnableAutoResponse.Checked = g_Settings.GetValue("chkEnableAutoResponse", False)
    chkLogoutOnDetect.Checked = g_Settings.GetValue("chkLogoutOnDetect", False)
    chkReportUnfriendly.Checked = g_Settings.GetValue("chkReportUnfriendly", True)
    chkEnableUnfriendlyDetect.Checked = g_Settings.GetValue("chkEnableUnfriendlyDetect", True)
    
    chkReportOpenChat.Checked = g_Settings.GetValue("chkReportOpenChat", False)
    chkReportTells.Checked = g_Settings.GetValue("chkReportTells", False)
    
    
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
        .SaveCheckbox chkAlertOnIRC, "chkAlertOnIRC"
        .SaveCheckbox chkEnableAdminDetection, "chkEnableAdminDetection"
        .SaveCheckbox chkEnableAutoResponse, "chkEnableAutoResponse"
        .SaveCheckbox chkLogoutOnDetect, "chkLogoutOnDetect"
        .SaveCheckbox chkReportUnfriendly, "chkReportUnfriendly"
        .SaveCheckbox chkEnableUnfriendlyDetect, "chkEnableUnfriendlyDetect"
        .SaveCheckbox chkReportOpenChat, "chkReportOpenChat"
        .SaveCheckbox chkReportTells, "chkReportTells"

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

    Set chkLogoutOnDetect = Nothing
    Set chkEnableAdminDetection = Nothing
    Set chkEnableAutoResponse = Nothing
    Set chkAlertOnIRC = Nothing
    Set chkReportUnfriendly = Nothing
    Set chkEnableUnfriendlyDetect = Nothing
    Set btnStopAlarm = Nothing
    Set chkReportOpenChat = Nothing
    Set chkReportTells = Nothing
    
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

Private Sub btnStopAlarm_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler
    
    'stop admin alert
    If g_AntiBan.AdminDetected Or g_AntiBan.AlarmTriggered Then
        Call g_AntiBan.SetAdminAlert(False)
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage INTERFACE_NAME & ".btnStopAlarm_Accepted - " & Err.Description
    Resume Fin
End Sub

