Attribute VB_Name = "Hub"
Option Explicit

'References & Shortcuts (not new'ed)
Public g_Engine As Engine
Public g_Core As ICore

Public g_PluginSite As DecalPlugins.PluginSite
Public g_PluginSite2 As Decal.IPluginSite2
Public g_Hooks As Decal.ACHooks
Public g_MainView As DecalPlugins.IView

Public g_ds As DarksideFilter.Filter
Public g_Objects As DarksideFilter.GameObjects
Public g_ACConst As DarksideFilter.acConst

Public mWindowObj As Decal.tagRECT   'window position

'Objects created by LTEngine at class initialization
'New'd in CreateEngineObjects
Public g_Timers As clsTimers
Public g_Filters As clsFilters
Public g_Web As clsAsyncWeb
Public g_Data As clsDatas
Public g_Service As clsServices
Public g_Events As clsACEvents
Public g_Settings As clsSettings
Public g_Spells As clsSpells
Public g_Buffer As clsBuffer
Public g_BuddyBuffer As clsBuddyBuffer
Public g_Macro As clsMacro
Public g_HUD As clsHUD
Public g_D3D As clsD3D
Public g_DOT As clsDOT
Public g_Nav As clsNav
Public g_AntiBan As clsAntiBan
Public g_RemoteCmd As clsRemoteCmd
Public g_Keys As clsACKeys
Public g_FellowList As clsFellowList
Public g_RareTracker As clsRareTracker

Public g_Window As clsACWindow
Public g_ui As clsPluginInterface

'Global var
Public g_bProfileLoaded As Boolean
Public g_manaItem As acObject
Public g_bFindingItem As Boolean
Public g_bLootRare As Boolean
Public g_currentEquip As acObject
Public g_currentArrow As acObject

'Corpse and kill trackers
Public g_TotalKilled As Long
Public g_TotalLooted As Long
Public g_buffBuddy As acObject


'Called from Engine.Initialize
'Paired with DeleteEngineObjects
Public Function CreateEngineObjects() As Boolean
On Error GoTo ErrorHandler

    MyDebug "CreateEngineObjects - Start"
    
    'Seed the randomizer
    Call Randomize(DateTime.Timer)

    Set g_Timers = New clsTimers
    Set g_Filters = New clsFilters
    Set g_Data = New clsDatas
    Set g_Web = New clsAsyncWeb
    Set g_Service = New clsServices
    Set g_Events = New clsACEvents
    Set g_Settings = New clsSettings
    Set g_Spells = New clsSpells
    Set g_Buffer = New clsBuffer
    Set g_BuddyBuffer = New clsBuddyBuffer
    Set g_Macro = New clsMacro
    Set g_HUD = New clsHUD
    Set g_D3D = New clsD3D
    Set g_DOT = New clsDOT
    Set g_AntiBan = New clsAntiBan
    Set g_RemoteCmd = New clsRemoteCmd
    Set g_Keys = New clsACKeys
    Set g_Window = New clsACWindow
    Set g_ui = New clsPluginInterface
    Set g_FellowList = New clsFellowList
    Set g_RareTracker = New clsRareTracker
    Set g_buffBuddy = Nothing
    
    MyDebug "CreateEngineObjects - End"
    
    CreateEngineObjects = True
    
Fin:
    Exit Function
ErrorHandler:
    CreateEngineObjects = False
    PrintErrorMessage "CreateEngineObjects - Error : " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

'Called from the Engine destructor
Public Function DeleteEngineObjects() As Boolean
On Error GoTo ErrorHandler

    'Log plugin exit in the main log path
    MyDebug "DeleteEngineObjects - Start"
    
    If Valid(g_Data) Then
        g_Data.mLogOutputPath = PATH_LOGS
    End If

    Set g_Events = Nothing
    Set g_Service = Nothing
    Set g_Spells = Nothing
    Set g_Data = Nothing
    Set g_Web = Nothing
    Set g_Settings = Nothing
    Set g_Buffer = Nothing
    Set g_BuddyBuffer = Nothing
    Set g_Macro = Nothing
    Set g_HUD = Nothing
    Set g_D3D = Nothing
    Set g_DOT = Nothing
    Set g_Nav = Nothing
    Set g_AntiBan = Nothing
    Set g_RemoteCmd = Nothing
    Set g_Keys = Nothing
    Set g_Window = Nothing
    Set g_ui = Nothing
    Set g_Filters = Nothing
    Set g_Timers = Nothing
    Set g_FellowList = Nothing
    Set g_RareTracker = Nothing
    
    Set g_manaItem = Nothing
    Set g_buffBuddy = Nothing
                
    DeleteEngineObjects = True
    
    MyDebug "DeleteEngineObjects - End"
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "DeleteEngineObjects - " & Err.Description
    DeleteEngineObjects = False
    Resume Fin
End Function
