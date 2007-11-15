VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPluginInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'HUB for clsUI* ingame interface classes
Option Explicit

Public Main As clsUIMain
Public Macro As clsUIMacro
Public Buffs As clsUIBuffs
Public Monsters As clsUIMonsters
Public Irc As clsUIIrc
Public MsgLog As clsUIMsgLog
Public Options As clsUIOptions
Public AntiBan As clsUIAntiBan
Public Loot As clsUILoot

'Special
Public NavEditor As clsNavEditor

Private Sub Class_Initialize()
    Set Main = New clsUIMain
    Set Macro = New clsUIMacro
    Set Buffs = New clsUIBuffs
    Set Monsters = New clsUIMonsters
    Set Irc = New clsUIIrc
    Set NavEditor = New clsNavEditor
    Set MsgLog = New clsUIMsgLog
    Set Options = New clsUIOptions
    Set AntiBan = New clsUIAntiBan
    Set Loot = New clsUILoot
End Sub

Private Sub Class_Terminate()
    Call Unload
End Sub

Public Function Init() As Boolean
On Error GoTo Error_Handler
    
    Dim bOk As Boolean
    bOk = True
    
    If Not Valid(g_MainView) Then
        PrintWarning "PluginInterface.Init - g_MainView empty, ignoring."
        Init = False
        Exit Function
    End If

    MyDebug "[Plugin Interface] Initializing User Interface"
    
    '****************************************************************************************
    If bOk Then bOk = Main.Init
    If bOk Then bOk = Macro.Init
    If bOk Then bOk = Buffs.Init
    If bOk Then bOk = Monsters.Init
    If bOk Then bOk = Irc.Init
    If bOk Then bOk = MsgLog.Init
    If bOk Then bOk = Options.Init
    If bOk Then bOk = AntiBan.Init
    If bOk Then bOk = Loot.Init
    '****************************************************************************************
    
Fin:
    MyDebug "[Plugin Interface] Initialization complete (Errors: " & CStr(Not bOk) & ")"
    Init = bOk
    Exit Function
Error_Handler:
    bOk = False
    PrintErrorMessage "PluginInterface.Init - " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue() As Boolean
On Error GoTo Error_Handler

    Dim bOk As Boolean
    bOk = True

    MyDebug "[Plugin Interface] Loading User Interface Settings..."
    
    '****************************************************************************************
    If bOk Then bOk = Main.LoadControlsValue
    If bOk Then bOk = Macro.LoadControlsValue
    If bOk Then bOk = Buffs.LoadControlsValue
    If bOk Then bOk = Monsters.LoadControlsValue
    If bOk Then bOk = Irc.LoadControlsValue
    If bOk Then bOk = MsgLog.LoadControlsValue
    If bOk Then bOk = Options.LoadControlsValue
    If bOk Then bOk = AntiBan.LoadControlsValue
    If bOk Then bOk = Loot.LoadControlsValue
    '****************************************************************************************
    
Fin:
    MyDebug "[Plugin Interface] Loading complete (Errors :" & CStr(Not bOk) & ")"
    LoadControlsValue = bOk
    Exit Function
Error_Handler:
    bOk = False
    PrintErrorMessage "PluginInterface.LoadControlsValue - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo ErrorHandler

    Dim bOk As Boolean
    bOk = True
    
    MyDebug "[Plugin Interface] Saving User Interface Settings..."
    
    '****************************************************************************************
    If bOk Then bOk = Main.SaveControlSettings
    If bOk Then bOk = Macro.SaveControlSettings
    If bOk Then bOk = Buffs.SaveControlSettings
    If bOk Then bOk = Monsters.SaveControlSettings
    If bOk Then bOk = Irc.SaveControlSettings
    If bOk Then bOk = MsgLog.SaveControlSettings
    If bOk Then bOk = Options.SaveControlSettings
    If bOk Then bOk = AntiBan.SaveControlSettings
    If bOk Then bOk = Loot.SaveControlSettings
    '****************************************************************************************
    
Fin:
    MyDebug "[Plugin Interface] Save complete (Errors :" & CStr(Not bOk) & ")"
    SaveControlSettings = bOk
    Exit Function
ErrorHandler:
    bOk = False
    PrintErrorMessage "PluginInterface.SaveControlSettings - " & Err.Description
    Resume Fin
End Function


Public Function Unload() As Boolean
    MyDebug "PluginInterface.Unload() -- Begin"
    Unload = False
    
    Set Main = Nothing
    Set Macro = Nothing
    Set Buffs = Nothing
    Set Monsters = Nothing
    Set Irc = Nothing
    Set NavEditor = Nothing
    Set MsgLog = Nothing
    Set Options = Nothing
    Set AntiBan = Nothing
    Set Loot = Nothing
    
    Unload = True
    MyDebug "PluginInterface.Unload() -- End"
End Function


