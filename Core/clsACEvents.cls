VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'AC Events
Option Explicit

Private WithEvents m_DSFilter As DarksideFilter.Filter
Attribute m_DSFilter.VB_VarHelpID = -1

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

   Call SetDarksideFilter(Nothing)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.Class_Initialize - " & Err.Description & " - " & Erl
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_DSFilter = Nothing
End Sub


'#####################################################################################
'#
'#                                  SETTERS
'#
'#####################################################################################

Friend Sub SetDarksideFilter(ByVal dsFilter As DarksideFilter.Filter)
On Error GoTo ErrorHandler

    MyDebug "clsACEvetns.SetDarksideFilter"
    
    Set m_DSFilter = dsFilter
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.SetDarksideFilter - " & Err.Description
    Resume Fin
End Sub

'Friend Sub SetACHooks(ByVal aHook As Decal.ACHooks)
'On Error GoTo ErrorHandler
'
'    MyDebug "clsACEvents.SetACHooks"
'
'    Set m_acHooks = aHook
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "ACEvents.SetACHooks - " & Err.Description
'    Resume Fin
'End Sub

'######################################################################################
'#
'#                           DARKSIDE FILTER EVENTS
'#
'######################################################################################

Private Sub m_DSFilter_OnLogin()
On Error GoTo ErrorHandler
    
    MyDebug "m_DSFilter_Onlogin -- start"
    
    If Not g_bInitComplete Then
    
        PrintErrorMessage "LifeTank could not initialize properly. Please check the Errors.txt log in your Lifetank\Logs\ directory for more information."
        Call SetDarksideFilter(Nothing)   'prevent us from receiving filter events
        Call g_Service.UnloadPlugin
        
    Else
        LogEvent "Logging in...", True
                
        'Setup IDLE timer
        Call g_Hooks.SetIdleTime(IDLE_TIMER)
        
        'Call Engine Method
        Call g_Core.Engine.OnLogin
    
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnLogin"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnLoginComplete()
On Error GoTo ErrorHandler

    If g_Core.Engine.OnLoginComplete Then
        
        'Add things to do when login is complete here
        'g_Hooks.AddChatTextRaw "[ ", COLOR_YELLOW, 0
        'g_Hooks.AddChatTextRaw PLUG_NAME, COLOR_BRIGHT_YELLOW, 0
        'g_Hooks.AddChatTextRaw " Core v" & App.Major & "." & App.Minor & " - Rev:" & App.Revision, COLOR_BLUE, 0
        'g_Hooks.AddChatTextRaw " Loaded ] ", COLOR_YELLOW, 0
        'g_Hooks.AddChatTextRaw "Plugin by ", COLOR_BLUE, 0
        'g_Hooks.AddChatTextRaw "Spax/Spk ", COLOR_YELLOW, 0
        'g_Hooks.AddChatText "of Wintersebb", COLOR_BLUE, 0
        'g_Hooks.addChatTextRaw "eXtra code by ", COLOR_BLUE, 0
        'g_Hooks.AddChatTextRaw "Xeon Xarid ", COLOR_YELLOW, 0
        'g_Hooks.AddChatText "of LeafCull", COLOR_BLUE, 0
        'g_Hooks.AddChatText "Original plugin by Spax/Spk of WintersEbb", COLOR_CYAN, 0
        'g_Hooks.AddChatText "Additional Code by C'relic and Swordz", COLOR_CYAN, 0
        
        'g_Hooks.AddChatText "[" & PLUG_NAME & "] Version " & App.Major & "." & App.Minor & " - Rev: " & App.Revision, COLOR_CYAN, 0
        g_Hooks.AddChatText "Developed by Xeon Xarid / Xeonx", COLOR_CYAN, 0
        
        'Print MOTD
        'PrintMessage MOTD
        PrintMessage USAGE_LINE

    Else
        PrintErrorMessage "Core: Engine failed OnLoginComplete"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnLoginComplete"
    Resume Fin
End Sub

