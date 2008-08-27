VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements LTCoreInterface.IEngine

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                                   LT ENGINE
'@
'@      This is the module used as an interface between LTCore and LTEngine
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private m_bEngineObjectsLoaded As Boolean

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                          Construction / Destruction
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    m_bEngineObjectsLoaded = False
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ERROR: Engine.Class_Initialize - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                           Internal private Methods
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Friend Sub FireDebugMessage(ByVal sMsg As String)
On Error Resume Next

    sMsg = "[Engine] " & sMsg
    Call g_Core.OnDebugMessage(sMsg)
End Sub

Friend Sub FireErrorMessage(ByVal sMsg As String)
On Error Resume Next

    sMsg = "[Engine] " & sMsg
    Call g_Core.OnErrorMessage(sMsg)
    
    If g_Core.InitComplete And Valid(g_RemoteCmd) Then
        Call g_RemoteCmd.RemoteRedirectChatToIRC("[LTx] ERROR: " & sMsg & vbCrLf)
    End If
    
End Sub

Friend Sub FireWarningMessage(ByVal sMsg As String)
On Error Resume Next

    sMsg = "[Engine] " & sMsg
    Call g_Core.OnWarningMessage(sMsg)
    
    If g_Core.InitComplete And Valid(g_RemoteCmd) Then
        Call g_RemoteCmd.RemoteRedirectChatToIRC("[LTx] WARNING: " & sMsg & vbCrLf)
    End If
    
End Sub

Friend Sub FireLogEvent(ByVal sMsg As String)
On Error Resume Next

    Call g_Core.OnLogEvent(sMsg)
End Sub

Friend Sub FirePrintMessage(ByVal sMsg As String, ByVal lColor As Long)
On Error Resume Next

    Call g_Core.OnPrintMessage(sMsg, lColor)
    
    If g_Core.InitComplete And Valid(g_RemoteCmd) Then
        Call g_RemoteCmd.RemoteRedirectChatToIRC("[LTx] " & sMsg & vbCrLf)
    End If
        
End Sub

Friend Sub FireStartMacro()
On Error Resume Next
    Call g_Core.OnStartMacro
End Sub

Friend Sub FireStopMacro()
On Error Resume Next
    Call g_Core.OnStopMacro
End Sub

Friend Sub FireLogChatMessage(ByVal sMessage As String)
On Error Resume Next
    Call g_Core.OnLogChatMessage(sMessage)
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@
'@                               private Interface
'@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Function IEngine_Initialize(ByVal oCore As LTCoreInterface.ICore) As Boolean
On Error GoTo ErrorHandler

    Dim bInitComplete As Boolean, _
        bFiltersOk As Boolean, _
        bSharedConfigOk As Boolean, _
        bUIInit As Boolean, _
        bSpellsOk As Boolean
        
    bInitComplete = False
        
    'Keep a handy reference to ourself
1    Set g_Engine = Me
        
    'Keep a reference to the core
2    Set g_Core = oCore

    'Create Engine objects
3    m_bEngineObjectsLoaded = CreateEngineObjects
4    If Not m_bEngineObjectsLoaded Then
        PrintErrorMessage "Engine.Initialize - Engine Objects Creation Failed"
        GoTo Fin
    End If
    
    MyDebug "----- Engine Initialization -----"
    
5    Set g_PluginSite = g_Core.PluginSite
6    Set g_PluginSite2 = g_Core.PluginSite2
7    Set g_Hooks = g_Core.ACHooks
8    Set g_MainView = g_Core.MainView

    'Load shared config file settings
9    bSharedConfigOk = g_Settings.LoadSharedConfig
    If bSharedConfigOk Then
        'g_Data.mDebugMode = g_Settings.GetValue("DebugMode", False, True)
        Hub.mWindowObj.Top = g_Settings.GetValue("WindowTop", 60, True)
        Hub.mWindowObj.Bottom = g_Settings.GetValue("WindowBottom", 420, True)
        Hub.mWindowObj.Left = g_Settings.GetValue("WindowLeft", 30, True)
        Hub.mWindowObj.Right = g_Settings.GetValue("WindowRight", 330, True)
        
        g_ui.Options.macroHUDposX = g_Settings.GetValue("macroHUDposX", 20, True)
        g_ui.Options.macroHUDposY = g_Settings.GetValue("macroHUDposY", 60, True)
        g_ui.Options.statsHUDposX = g_Settings.GetValue("statsHUDposX", 20, True)
        g_ui.Options.statsHUDposY = g_Settings.GetValue("statsHUDposY", 120, True)
        g_ui.Options.invHUDposX = g_Settings.GetValue("invHUDposX", 200, True)
        g_ui.Options.invHUDposY = g_Settings.GetValue("invHUDposY", 60, True)
        g_ui.Options.debugHUDposX = g_Settings.GetValue("debugHUDposX", 20, True)
        g_ui.Options.debugHUDposY = g_Settings.GetValue("debugHUDposY", 300, True)
        g_ui.Options.dotHUDposX = g_Settings.GetValue("dotHUDposX", 200, True)
        g_ui.Options.dotHUDposY = g_Settings.GetValue("dotHUDposY", 200, True)

    Else
        PrintErrorMessage "Error loading Shared Config file."
    End If
    
    'Load filters
    MyDebug "---------------------------------"
    bFiltersOk = g_Filters.LoadFilters
        
    'Init User Interface
    MyDebug "---------------------------------"
    bUIInit = g_ui.Init
    
    'Load spells
    MyDebug "---------------------------------"
    bSpellsOk = g_Spells.LoadSpells
    
    'init HUD
    Call g_HUD.Init
    
    'Check if everything went fine
    bInitComplete = bSharedConfigOk And bFiltersOk And bUIInit And bSpellsOk
    
    MyDebug "---------------------------------"
    MyDebug "...Shared Config : " & bSharedConfigOk
    MyDebug "...Filters : " & bFiltersOk
    MyDebug "...UI : " & bUIInit
    MyDebug "...Spells : " & bSpellsOk
    MyDebug "---------------------------------"
    MyDebug "Engine Init Complete : " & bInitComplete
    MyDebug "*********************************"
    MyDebug ""
    
Fin:
    IEngine_Initialize = bInitComplete
    Exit Function
ErrorHandler:
    bInitComplete = False
    PrintErrorMessage "Engine.Initialize - " & Err.Description
    Resume Fin
End Function

Private Sub IEngine_Terminate()
On Error GoTo ErrorHandler

    MyDebug "IEngine_Terminate"

    'Remove objects created by the LT Engine
    Call DeleteEngineObjects

    'Remove global references
    Set g_ds = Nothing
    Set g_Objects = Nothing
    Set g_ACConst = Nothing
    Set g_Filters = Nothing
    
    Set g_PluginSite = Nothing
    Set g_PluginSite2 = Nothing
    Set g_Hooks = Nothing
    Set g_MainView = Nothing
    Set g_Core = Nothing
    
    'Lastly remove the reference to ourself to IEngine destructor can be called
    m_bEngineObjectsLoaded = False
    Set g_Engine = Nothing
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "IEngine_Terminate - " & Err.Description
    Resume Fin
End Sub

Private Function IEngine_StartMacro() As Boolean
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Function
    
    Dim bStarted As Boolean
    
    'If mAuth.m_PluginEnabled = False Then
    '    PrintMessage g_String(mStrings.e_strUnableToAuth)
    '    bStarted = False
    '    GoTo Fin
    'End If
    
    bStarted = g_Macro.StartMacro()
    
    g_ui.Main.chkEnable.Checked = bStarted
    If bStarted Then
        PrintMessage "Macro Started [TurboMode : " & CStr(TurboMode) & "]"
        If Valid(g_HUD) Then
            Call g_HUD.createAllHuds
            Call g_HUD.StartHUDs
        End If
        If Valid(g_DOT) Then
            Call g_DOT.Init
        End If
    End If
  
Fin:
    IEngine_StartMacro = bStarted
    Exit Function
ErrorHandler:
    bStarted = False
    PrintErrorMessage "Engine.IEngine_StartMacro - " & Err.Description
    Resume Fin
End Function

Private Sub IEngine_OnTick()
On Error GoTo ErrorHandler
    
    If m_bEngineObjectsLoaded Then
        Call g_Macro.OnMacroTick
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "IEngine_OnTick - " & Err.Description
    Resume Fin
End Sub

Private Property Get IEngine_DebugMode() As Boolean
    If Not m_bEngineObjectsLoaded Then Exit Property
    If Valid(g_Data) Then
        IEngine_DebugMode = g_Data.mDebugMode
    Else
        IEngine_DebugMode = False
    End If
End Property

Private Property Get IEngine_DisableLogs() As Boolean
    If Not m_bEngineObjectsLoaded Then Exit Property
    If Valid(g_ui) Then
        IEngine_DisableLogs = g_ui.Options.chkDisableLogs.Checked
    Else
        IEngine_DisableLogs = False
    End If
End Property

Private Function IEngine_OnLogin() As Boolean
On Error GoTo ErrorHandler
    
    MyDebug "IEngine_OnLogin"

    If Not m_bEngineObjectsLoaded Then Exit Function
        
    'Load AC Keys
    Call g_Keys.LoadKeys
    
    IEngine_OnLogin = True
    
    MyDebug "*** Checking for Correct windows resolution ***"
    Dim acRect As Decal.tagRECT
    acRect = g_Hooks.ACWindowRect
    MyDebug "acRect dimensions: " & g_Hooks.ACWindowRect.Right & " x " & g_Hooks.ACWindowRect.Bottom
    'PrintMessage "Window: " & g_Hooks.ACWindowRect.Right & " x " & g_Hooks.ACWindowRect.Bottom
    If (g_Hooks.ACWindowRect.Right < 800) Then
        PrintErrorMessage "*************************************************"
        PrintErrorMessage "*** You should run at least 800x600 resolution ***"
        PrintErrorMessage "*************************************************"
        'Call DeleteEngineObjects
        'm_bEngineObjectsLoaded = False
    End If
    
    'init D3D
    Call g_D3D.Init

    mAuth.LoadSettingsAuth
    
    g_Data.mDebugMode = False
    Call g_Filters.dsFilter.SetDebugMode(False)
    
    MyDebug "IEngine_OnLogin all done"

    
Fin:
    Exit Function
ErrorHandler:
    IEngine_OnLogin = False
    PrintErrorMessage "Engine.OnLogin"
    Resume Fin
End Function

Private Function IEngine_OnLoginComplete() As Boolean
On Error GoTo ErrorHandler

    If Not m_bEngineObjectsLoaded Then Exit Function
    
    'LogEvent "Login Complete", True

    If g_Filters.playerName = "LocalPlayer" Then
        MyDebug "OnLoginComplete, but no name for: " & g_Filters.g_charFilter.Name & " : " & g_Filters.g_charFilter.LoginStatus
        PrintMessage "ERROR: === Could not find player name. Exiting LifeTank ==="
        IEngine_OnLoginComplete = False
        Call g_Service.UnloadPlugin
        Exit Function
    End If
    
    Call g_Filters.TestFilters
    Call g_Window.UpdateDimensions
    Call g_Spells.FindAndSetDefaultDamage
    
    MyDebug "IEngine_OnLoginComplete: before LoadProfile for: " & g_Filters.playerName
    
    g_bProfileLoaded = g_Settings.LoadProfile(g_Filters.playerName)
    
    If Not g_bProfileLoaded Then
        Call g_Core.fatalError("Error while loading " & g_Settings.Profile.Name & " profile.")
        IEngine_OnLoginComplete = False
    Else
        'Now that we have our character info, connect to irc
        Call g_ui.Irc.CheckAutoConnect
        
        'Print bad key settings
        Call g_Keys.PrintNotFound
        
        Call g_HUD.StartHUDs
        
        'Return ok
        IEngine_OnLoginComplete = True
    End If
    
    'Set View to saved position on screen
    g_MainView.Position.Top = Hub.mWindowObj.Top
    g_MainView.Position.Bottom = Hub.mWindowObj.Bottom
    g_MainView.Position.Left = Hub.mWindowObj.Left
    g_MainView.Position.Right = Hub.mWindowObj.Right
    g_MainView.Position = Hub.mWindowObj
    
    'authenticate!
    'mAuth.m_Auth (1)
    

Fin:
    Exit Function
ErrorHandler:
    IEngine_OnLoginComplete = False
    PrintErrorMessage "Engine.OnLoginComplete"
    Resume Fin
End Function

Private Sub IEngine_ForceMacroShutdown()
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Sub
    
    Call ToggleMacro(False)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Engine.ForceMacroShutdown - " & Err.Description
    Resume Fin
End Sub

'Must return true if the console text must be eaten
Private Function IEngine_HandleConsoleText(ByVal bstrMsg As String, ByVal pColor As Long) As Boolean
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Function
    
    IEngine_HandleConsoleText = g_Events.HandleConsoleMessage(bstrMsg, pColor)
    
Fin:
    Exit Function
ErrorHandler:
    IEngine_HandleConsoleText = False
    PrintErrorMessage "Engine.HandleConsoleText - " & Err.Description
    Resume Fin
End Function

'Must return true if the entered console text must be processed
Private Function IEngine_HandleConsoleCommand(ByVal bstrMsg As String) As Boolean
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Function
    
    IEngine_HandleConsoleCommand = Commands.HandleConsoleCommand(bstrMsg)
    
Fin:
    Exit Function
ErrorHandler:
    IEngine_HandleConsoleCommand = False
    PrintErrorMessage "Engine.HandleConsoleCommand - " & Err.Description
    Resume Fin
End Function

'Must return true if the Status text must be eaten
Private Function IEngine_HandleStatusText(ByVal bstrMsg As String) As Boolean
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Function
    
    IEngine_HandleStatusText = g_Events.HandleStatusMessage(bstrMsg)
    
Fin:
    Exit Function
ErrorHandler:
    IEngine_HandleStatusText = False
    PrintErrorMessage "Engine.HandleStatusText - " & Err.Description
    Resume Fin
End Function


Private Sub IEngine_SavePluginConfiguration()
On Error GoTo ErrorHandler
    
    If Not m_bEngineObjectsLoaded Then Exit Sub
    
    Call g_Settings.SavePluginConfiguration
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Engine.SavePluginConfiguration - " & Err.Description
    Resume Fin
End Sub

Private Sub IEngine_HandleMouseMoveHUD(ByVal x As Long, ByVal y As Long)
    MyDebug "Trying to move HUD to: " & x & " : " & y
    Call g_HUD.MoveHUD(x, y)
End Sub

Private Function IEngine_HandleMouseClick(ByVal x As Long, ByVal y As Long) As Boolean
On Error GoTo ErrorHandler
    
    Dim bEatKey As Boolean
    
    bEatKey = False
    
    If Not g_Macro.Ticking Then GoTo Fin
    
    If Not m_bEngineObjectsLoaded Then GoTo Fin

    Call g_Macro.Loot.SetSilentID(False)

    If Not Valid(g_HUD) Then GoTo Fin
    
    If g_ui.Options.chkLockHuds.Checked Then
        ' Do nothing
    Else
        ' Check to see if this is a click on one of the HUD windows
        ' X and Y are the coords of the mouse click
        bEatKey = g_HUD.getHUDAt(x, y)
        If bEatKey Then MyDebug "Found a HUD to move"
    End If
    

Fin:
    IEngine_HandleMouseClick = bEatKey
    Exit Function
ErrorHandler:
    bEatKey = False
    PrintErrorMessage "Engine.HandleMouseClick - " & Err.Description
    Resume Fin
End Function

'Handles the LT hotkeys - bCTRL = true when the CTRL key is pushed
'HandleHotkey must return true if the key must be eaten
Private Function IEngine_HandleHotkey(ByVal lKey As Long, ByVal bCtrl As Boolean) As Boolean
On Error GoTo ErrorHandler

    Dim bEatKey As Boolean
    Dim bHud As Boolean
    
    bEatKey = False
    
    If Not m_bEngineObjectsLoaded Then GoTo Fin
    
    'Hotkeys requiring CTRL to be pushed go here
    If bCtrl Then
    
        'assume we're going to hit an LT hotkey
        bEatKey = True
        
        Select Case lKey
            Case vbKeyF1
                Call ToggleMacro(Not g_Macro.Ticking)
            
            Case vbKeyF2
                If g_ui.Options.chkEnableHUD.Checked Then
                    g_ui.Options.chkEnableHUD.Checked = False
                    Call g_HUD.stopAllHUDs
                Else
                    g_ui.Options.chkEnableHUD.Checked = True
                    Call g_HUD.StartHUDs
                End If
                PrintMessage "HUD Display : " & g_ui.Options.chkEnableHUD.Checked
                
            Case vbKeyF3
                PrintMessage "Forcing Rebuff..."
                Call g_Macro.ForceRebuff
                
            Case vbKeyF4
                g_ui.Macro.chkEnableSticky.Checked = True
                Call g_Nav.SetSticky
            
            Case vbKeyF5
                If g_ui.Options.chkHUDDOT.Checked Then
                    g_ui.Options.chkHUDDOT.Checked = False
                    Call g_HUD.startDOTHUD
                Else
                    g_ui.Options.chkHUDDOT.Checked = True
                    Call g_HUD.stopDOTHUD
                End If
                PrintMessage "DOT Display : " & g_ui.Options.chkHUDDOT.Checked
                
            Case vbKeyF6
                Call PhatLoot.PutSelectionInUst
            
            Case vbKeyF7
                Call g_ui.NavEditor.ShowEditor
                
            Case Else
                bEatKey = False
                
        End Select
    End If
            
    'Other non CTRL hotkeys go here
    If Not bEatKey Then
        bEatKey = True
        
        Select Case lKey
            Case vbKeyPause
                Call TogglePause
                
            Case Else
                bEatKey = False
        End Select
    End If
    
Fin:
    IEngine_HandleHotkey = bEatKey
    Exit Function
ErrorHandler:
    bEatKey = False
    PrintErrorMessage "Engine.HandleHotkey - " & Err.Description
    Resume Fin
End Function

Private Sub IEngine_UpdateClock()
On Error GoTo ErrorHandler

    If m_bEngineObjectsLoaded Then
        Call g_Timers.Update
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Engine.UpdateClock - " & Err.Description
    Resume Fin
End Sub

