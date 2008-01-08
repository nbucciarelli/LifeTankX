VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                                  Properties
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get PluginSite() As DecalPlugins.PluginSite
End Property

Public Property Get PluginSite2() As Decal.PluginSite2
End Property

Public Property Get ACHooks() As Decal.ACHooks
End Property

Public Property Get MainView() As DecalPlugins.IView
End Property

Public Property Get Time() As Long
End Property

Public Property Get ElapsedSeconds() As Long
End Property

Public Property Get Ticking() As Boolean
End Property

Public Property Get InitComplete() As Boolean
End Property

Public Property Get TimerInterval() As Long
End Property

Public Property Let TimerInterval(ByVal lVal As Long)
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                              Public Methods
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Sub SendKey(ByVal lKey As Long)
End Sub

Public Sub SendKeyHold(ByVal lKey As Long)
End Sub

Public Sub SendKeyRelease(ByVal lKey As Long)
End Sub

Public Sub SendTextToConsole(ByVal sText As String, Optional ByVal bForceSend As Boolean = False, Optional ByVal bSendMessage As Boolean = True)
End Sub

Public Sub MouseClick(ByVal xPos As Integer, ByVal yPos As Integer)
End Sub

Public Sub AppActivate()
End Sub

Public Function ParseIRCMode(ByVal sUserName As String, Optional ByRef sRetUserMode As String) As String
End Function

Public Function ExtractNextMessageFragment(ByRef sData As String) As String
End Function

Public Sub FatalError(ByVal sMessage As String)
End Sub

Public Sub OnDebugMessage(ByVal sMsg As String)
End Sub

Public Sub OnErrorMessage(ByVal sMsg As String)
End Sub

Public Sub OnWarningMessage(ByVal sMsg As String)
End Sub

Public Sub OnLogEvent(ByVal sMsg As String)
End Sub

Public Sub OnPrintMessage(ByVal sMsg As String, ByVal lColor As Long)
End Sub

Public Sub OnLogChatMessage(ByVal sMessage As String)
End Sub

Public Sub OnStartMacro()
End Sub

Public Sub OnStopMacro()
End Sub

'per toon Log Files
Public Sub BuildLogFiles(ByVal sPath As String)
End Sub

