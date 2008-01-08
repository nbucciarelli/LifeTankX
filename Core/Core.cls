VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Core Interface
Option Explicit

Implements ICore

Private WithEvents m_tmrMacro As Timer
Attribute m_tmrMacro.VB_VarHelpID = -1
Private m_engine As LTCoreInterface.IEngine
Attribute m_engine.VB_VarHelpID = -1


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                          Construction / Destruction
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

1    Set m_tmrMacro = frmTimer.tmrMacro
2    m_tmrMacro.Enabled = False

3    Set m_engine = CreateObject("LTEngine.Engine") 'New LTEngine.Engine

    LogEvent "Core.Class_Initialize - done"

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.Class_Initialize - " & Err.Description & " - Line: " & Erl
    Resume Fin
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    Set m_tmrMacro = Nothing
    Set m_engine = Nothing
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.Class_Terminate - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                              Friend Properties
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Friend Property Get Engine() As IEngine
    Set Engine = m_engine
End Property


Friend Property Get mg_Time() As Long
    mg_Time = g_Time / 1000
End Property



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                              Friend Methods
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Friend Function Initialize() As Boolean
On Error GoTo ErrorHandler

    MyDebug "Core.Initialize"
    
1    Initialize = m_engine.Initialize(Me)
    
Fin:
    Exit Function
ErrorHandler:
    Initialize = False
    PrintErrorMessage "Core.Initialize - " & Err.Description
    Resume Fin
End Function

Friend Sub Terminate()
On Error GoTo ErrorHandler

    MyDebug "Core.Terminate"
    
    If Valid(m_engine) Then
        Call m_engine.Terminate
    End If

    Set m_engine = Nothing

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.Terminate - " & Err.Description
    Resume Fin
End Sub

Friend Sub ForceMacroShutdown()
On Error GoTo ErrorHandler

    'Shutdown macro heartbeat (don't trust the Engine)
    m_tmrMacro.Enabled = False
    
    'Tell the engine to stfu
    Call m_engine.ForceMacroShutdown
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.ForceMacroShutdown - " & Err.Description
    Resume Fin
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                              Interface Impl
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'Start the macro heart beat
Private Sub ICore_OnStartMacro()
On Error GoTo ErrorHandler

    If m_engine.StartMacro() Then
        m_tmrMacro.Enabled = True
    End If
        
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.m_engine_OnStartMacro - " & Err.Description
    Resume Fin
End Sub

Private Property Get ICore_ACHooks() As Decal.ACHooks
    Set ICore_ACHooks = g_Hooks
End Property

Private Sub ICore_AppActivate()
    Call myAppActivate
End Sub

Private Function ICore_ExtractNextMessageFragment(sData As String) As String
    ICore_ExtractNextMessageFragment = shIrcUtils.ExtractNextMessageFragment(sData)
End Function

Private Sub ICore_FatalError(ByVal sMessage As String)
On Error GoTo ErrorHandler

    PrintErrorMessage "FATAL ERROR - " & sMessage
    Call g_Service.UnloadPlugin
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.FatalError - " & Err.Description
    Resume Fin
End Sub

Private Property Get ICore_InitComplete() As Boolean
    ICore_InitComplete = g_bInitComplete
End Property

Private Property Get ICore_MainView() As DecalPlugins.IView
    Set ICore_MainView = g_MainView
End Property

Private Sub ICore_MouseClick(ByVal xPos As Integer, ByVal yPos As Integer)
    Call myMouseClick(xPos, yPos)
End Sub

Private Function ICore_ParseIRCMode(ByVal sUserName As String, Optional sRetUserMode As String) As String
    ICore_ParseIRCMode = shIrcUtils.ParseMode(sUserName, sRetUserMode)
End Function

Private Property Get ICore_PluginSite() As DecalPlugins.PluginSite
    Set ICore_PluginSite = g_PluginSite
End Property

Private Property Get ICore_PluginSite2() As Decal.PluginSite2
    Set ICore_PluginSite2 = g_PluginSite2
End Property

Private Sub ICore_SendKey(ByVal lKey As Long)
    Call mySendKey(lKey)
End Sub

Private Sub ICore_SendKeyHold(ByVal lKey As Long)
    Call mySendKeyHold(lKey)
End Sub

Private Sub ICore_SendKeyRelease(ByVal lKey As Long)
    Call mySendKeyRelease(lKey)
End Sub

Private Sub ICore_SendTextToConsole(ByVal sText As String, Optional ByVal bForceSend As Boolean = False, Optional ByVal bSendMessage As Boolean = True)
    Call mySendTextToConsole(sText, bForceSend, bSendMessage)
End Sub

Private Property Get ICore_Ticking() As Boolean
    ICore_Ticking = m_tmrMacro.Enabled
End Property

Private Property Get ICore_Time() As Long
    ICore_Time = g_Time / 1000
End Property

Private Property Get ICore_ElapsedSeconds() As Long
    ICore_ElapsedSeconds = g_ElapsedTime
End Property

Private Property Let ICore_TimerInterval(ByVal RHS As Long)
    m_tmrMacro.Interval = RHS
End Property

Private Property Get ICore_TimerInterval() As Long
    ICore_TimerInterval = m_tmrMacro.Interval
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                              LTEngine Events
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub ICore_OnStopMacro()
On Error GoTo ErrorHandler

    m_tmrMacro.Enabled = False

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.m_engine_OnStopMacro - " & Err.Description
    Resume Fin
End Sub

Private Sub ICore_BuildLogFiles(ByVal sPath As String)
    Call xBuildPluginLogFiles(sPath)
End Sub

Private Sub ICore_OnDebugMessage(ByVal sMsg As String)
    MyDebug sMsg
End Sub

Private Sub ICore_OnErrorMessage(ByVal sMsg As String)
    PrintErrorMessage sMsg
End Sub

Private Sub ICore_OnLogChatMessage(ByVal sMessage As String)
    If Valid(g_chatLog) Then
        xWriteMessageToFile g_chatLog, sMessage
    Else
        WriteMessageToFile GetChatLogPath, sMessage
    End If
End Sub

Private Sub ICore_OnLogEvent(ByVal sMsg As String)
    LogEvent sMsg
End Sub

Private Sub ICore_OnPrintIrcMessage(ByVal sSenderName As String, ByVal sMessageBody As String, ByVal bPrivateMsg As Boolean)
    PrintIrcMessage sSenderName, sMessageBody, bPrivateMsg
End Sub

Private Sub ICore_OnPrintMessage(ByVal sMsg As String, ByVal lColor As Long)
    PrintMessage sMsg, lColor
End Sub

Private Sub ICore_OnWarningMessage(ByVal sMsg As String)
    PrintWarning sMsg
End Sub

Private Sub m_tmrMacro_Timer()
On Error GoTo ErrorHandler
    
    Call m_engine.OnTick

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Core.m_tmrMacro_Timer - " & Err.Description
    Resume Fin
End Sub
