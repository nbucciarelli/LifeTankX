VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Decal.IPlugin2
Implements DecalPlugins.IWindowsMessageSink

Private m_sViewSchema As String

Private WithEvents m_ACHooks As Decal.ACHooks
Attribute m_ACHooks.VB_VarHelpID = -1
Private WithEvents m_DSFilter As DarksideFilter.Filter
Attribute m_DSFilter.VB_VarHelpID = -1

Private Sub Class_Initialize()
    
    'Init
    g_bInitComplete = False
    g_bStopPlugin = False
    
    'Build Log files
1    Call BuildPluginLogFiles
    
    'Create the global plugin objects
2    g_bObjectsLoaded = CreatePluginObjects
    
3    LogEvent "Main.Class_Initialize Schema file: " & Utils.GetDataFolder & "\" & PATH_DATA & "\" & FILE_PLUGIN_VIEW
    'Get the plugin UI string
4    m_sViewSchema = FileToString(PATH_DATA & "\" & FILE_PLUGIN_VIEW)
    
    LogEvent "Main.Class_Initialize - done"
    
End Sub


Private Function LoadDarksideFilter() As Boolean
On Error GoTo InvalidFilter
    Dim bRet As Boolean
    
    Set m_DSFilter = g_PluginSite.NetworkFilter("DarksideFilter.Filter")
    If Valid(m_DSFilter) Then
        Set g_ds = m_DSFilter
        Call g_Events.SetDarksideFilter(m_DSFilter)
    End If
    
    bRet = Valid(m_DSFilter)
   
Fin:
    LoadDarksideFilter = bRet
    Exit Function
InvalidFilter:
    bRet = False
    PrintErrorMessage "LoadDarksideFilter - " & Err.Description & " (line : " & Erl & ") - Err# " & Err.Number
    Set m_DSFilter = Nothing
    Resume Fin
End Function

Private Sub IPlugin2_Initialize(ByVal Site As Decal.PluginSite2)
On Error GoTo Error_Handler
    
    MyDebug ""
    MyDebug "*********************************"
    MyDebug ".....Initializing Plugin....."
    MyDebug "*********************************"
    
    g_bInitComplete = False
    
1    Call Site.RegisterSinks(Me)
    
2    Set g_PluginSite2 = Site
    
    'Save reference to PluginSite
3    Set g_PluginSite = Site.PluginSite
    
    'Setup ACHooks
4    Set m_ACHooks = g_PluginSite2.Hooks
5    Set g_Hooks = m_ACHooks
    
    MyDebug "ACHooks : " & Valid(g_Hooks)
    
    'Plugin UI
    MyDebug "IPlugin2_Initialize : loading Plugin View"
6    Set g_MainView = g_PluginSite.LoadView(m_sViewSchema)
    
7    If Not g_bObjectsLoaded Then
        PrintErrorMessage "Plugin Objects not loaded properly."
        GoTo Fin
    End If
     
8    If Not LoadDarksideFilter Then
        PrintErrorMessage "Could not link to Darkside Filter - Please make sure you're using the correct version."
        GoTo Fin
    End If
        
    'Call Engine Initialization
9    Set g_Core = New Core
10   If Not g_Core.Initialize Then
        PrintErrorMessage "Core Init Failed!"
        GoTo Fin
    End If
    
    'Start the main clock timer
11  Call g_Clock.StartTimer
    
    'Everything went fine
    g_bInitComplete = True

    MyDebug "Plugin Initialization Complete : " & g_bInitComplete
      
Fin:
    Exit Sub
Error_Handler:
    g_bInitComplete = False
    PrintErrorMessage "IPlugin2_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub IPlugin2_Terminate()
On Error GoTo ErrorHandler

    MyDebug "IPlugin2_Terminate - Start"
    
    'save plugin settings
    If Not g_bStopPlugin And g_bInitComplete Then
        Call g_Core.Engine.SavePluginConfiguration
    End If

    'Remove any references to allow objects to be destroyed
    If Valid(g_Core) Then
       Call g_Core.Terminate
       Set g_Core = Nothing
    End If
    
    'Unload objects created by the plugin
    Call UnloadPluginObjects
    
    'Zero everything else
    Set m_DSFilter = Nothing
    Set m_ACHooks = Nothing
    Set g_ds = Nothing
    Set g_Hooks = Nothing
    Set g_MainView = Nothing
    Set g_PluginSite = Nothing
    Set g_PluginSite2 = Nothing
    
    MyDebug "IPlugin2_Terminate - End"
    
    ' Close all the log files
    Call xCloseLogFiles
    
    Set g_debugLog = Nothing
    Set g_errorLog = Nothing
    Set g_chatLog = Nothing
    Set g_eventLog = Nothing

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "IPlugin2_Terminate - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub


Private Function IWindowsMessageSink_WindowMessage(ByVal hwnd As Long, ByVal uMsg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Boolean
On Error GoTo ErrorHandler

    'MyDebug "IWindowsMessageSink_Windowmessage"

    'ignore if plugin must be stopped
    If g_bStopPlugin Then Exit Function

    If Valid(g_Hotkeys) Then
        IWindowsMessageSink_WindowMessage = g_Hotkeys.HandleKeyboardMessages(hwnd, uMsg, wParam, lParam)
    End If

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "IWindowsMessageSink_WindowMessage - " & Err.Description
    Resume Fin
End Function

Private Sub IWindowsMessageSink_WindowMessageEnd()
    'dummy
End Sub


'----------------------

' Text scrolling thru chat area
Private Sub m_ACHooks_ChatTextIntercept(ByVal bstrText As String, ByVal lColor As Long, ByVal lTarget As Long, bEat As Boolean)
On Error GoTo ErrorHandler

    'MyDebug "LTCore.m_acHooks_ChatTextIntercept: " & bstrText

    'ignore if plugin must be stopped
    If g_bStopPlugin Then Exit Sub

    If g_bInitComplete And Valid(g_Core) Then
        bEat = g_Core.Engine.HandleConsoleText(bstrText, lColor)
    Else
        bEat = False
    End If

Fin:
    Exit Sub
ErrorHandler:
    bEat = False
    PrintErrorMessage "LTCore.m_acHooks_ChatTextIntercept - " & Err.Description
    Resume Fin
End Sub

' User typed a command
Private Sub m_acHooks_ChatParserIntercept(ByVal bstrText As String, bEat As Boolean)
On Error GoTo ErrorHandler

    'MyDebug "LTCore.m_acHooks_ChatParserIntercept: " & bstrText

    'ignore if plugin must be stopped
    If g_bStopPlugin Then Exit Sub

    If g_bInitComplete And Valid(g_Core) Then
        bEat = g_Core.Engine.HandleConsoleCommand(bstrText)
    Else
        bEat = False
    End If

Fin:
    Exit Sub
ErrorHandler:
    bEat = False
    PrintErrorMessage "LTCore.m_acHooks_ChatParserIntercept - " & Err.Description
    Resume Fin
End Sub

' Red text in center of screen
Private Sub m_ACHooks_StatusTextIntercept(ByVal bstrText As String, bEat As Boolean)
On Error GoTo ErrorHandler

    'MyDebug "LTCore.m_acHooks_StatusTextIntercept: " & bstrText

    'ignore if plugin must be stopped
    If g_bStopPlugin Then Exit Sub

    If g_bInitComplete And Valid(g_Core) Then
        bEat = g_Core.Engine.HandleStatusText(bstrText)
    Else
        bEat = False
    End If

Fin:
    Exit Sub
ErrorHandler:
    bEat = False
    PrintErrorMessage "LTCore.m_acHooks_StatusTextIntercept - " & Err.Description
    Resume Fin
End Sub


Public Function GetDataFolder() As String
On Error GoTo GenericFolder
   Dim PathName  As String
   Dim strPath   As String
   Dim lngReturn As Long
   Dim ReturnVal As Long

   strPath = String(260, 0)
   lngReturn = SHGetFolderPath(0, CSIDL_PERSONAL, 0, &H0, strPath)
   PathName = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)
   GetDataFolder = PathName & "\LifeTankX"
    
   Exit Function
GenericFolder:
   If Err.Number = 453 Then GetDataFolder = App.Path
End Function
