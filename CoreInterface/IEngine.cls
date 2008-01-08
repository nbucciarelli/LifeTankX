VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get DebugMode() As Boolean
End Property

Property Get DisableLogs() As Boolean
End Property

Public Function Initialize(ByVal oCore As ICore) As Boolean
End Function

Public Sub Terminate()
End Sub

Public Function OnLogin() As Boolean
End Function

Public Function OnLoginComplete() As Boolean
End Function

Public Sub ForceMacroShutdown()
End Sub

Public Function HandleConsoleText(ByVal bstrMsg As String, ByVal pColor As Long) As Boolean
End Function

Public Function HandleConsoleCommand(ByVal bstrMsg As String) As Boolean
End Function

Public Function HandleStatusText(ByVal bstrMsg As String) As Boolean
End Function

Public Sub SavePluginConfiguration()
End Sub

Public Function HandleHotkey(ByVal lKey As Long, ByVal bCtrl As Boolean) As Boolean
End Function

Public Function HandleMouseClick(ByVal x As Long, ByVal y As Long) As Boolean
End Function

Public Sub HandleMouseMoveHUD(ByVal x As Long, ByVal y As Long)
End Sub

Public Function StartMacro() As Boolean
End Function

Public Sub OnTick()
End Sub

Public Sub UpdateClock()
End Sub
