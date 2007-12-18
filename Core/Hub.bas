Attribute VB_Name = "Hub"
Option Explicit

'Reference (not created by LTCore)
Public g_PluginSite As DecalPlugins.PluginSite
Public g_PluginSite2 As Decal.IPluginSite2
Public g_Hooks As Decal.ACHooks
Public g_MainView As DecalPlugins.IView
Public g_ds As DarksideFilter.Filter

'Objects created by LTCore
Public g_Core As Core
Public g_Service As clsServices
Public g_Events As clsACEvents
Public g_Clock As clsClock
Public g_Timers As clsTimers
Public g_Hotkeys As clsHotkeys

'Global var
Public g_bInitComplete As Boolean   'tell if the plugin initialized properly
Public g_bStopPlugin As Boolean
Public g_bObjectsLoaded As Boolean

Public g_Time As Long        'Current time in miliseconds
Public g_ElapsedTime As Long 'Elapsed seconds since start
Public g_LogPath As String

Public g_debugLog As TextStream
Public g_errorLog As TextStream
Public g_chatLog As TextStream
Public g_eventLog As TextStream


'Called at plugin class initialization
'Objects will be created before IPlugin_Initialize
Public Function CreatePluginObjects() As Boolean
On Error GoTo ErrorHandler

1    Call Randomize(DateTime.Timer)

2    Set g_Timers = New clsTimers 'g_Timers must be loaded before other modules
3    Set g_Service = New clsServices
4    Set g_Events = New clsACEvents
5    Set g_Clock = New clsClock
6    Set g_Hotkeys = New clsHotkeys
    
    CreatePluginObjects = True
    
Fin:
    Exit Function
ErrorHandler:
    CreatePluginObjects = False
    PrintErrorMessage "CreatePluginObjects - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

'Called at plugin_terminate
Public Function UnloadPluginObjects() As Boolean
On Error GoTo ErrorHandler

    'Log plugin exit in the main log path
    MyDebug "Plugin Objects Unload - Start "

    Set g_Hotkeys = Nothing
    Set g_Clock = Nothing
    Set g_Events = Nothing
    Set g_Service = Nothing
    Set g_Timers = Nothing

    MyDebug "Plugin Objects Unload - End"
    UnloadPluginObjects = True
    
Fin:
    Exit Function
ErrorHandler:
    UnloadPluginObjects = False
    PrintErrorMessage "UnloadPluginObjects - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function
